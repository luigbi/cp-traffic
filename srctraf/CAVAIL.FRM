VERSION 5.00
Begin VB.Form CAvail 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5145
   ClientLeft      =   750
   ClientTop       =   1485
   ClientWidth     =   7065
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
   ScaleHeight     =   5145
   ScaleWidth      =   7065
   Begin VB.PictureBox plcCalendar 
      ForeColor       =   &H00000000&
      Height          =   5115
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   7005
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7065
      Begin VB.PictureBox plcShow 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   165
         ScaleHeight     =   255
         ScaleWidth      =   2610
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4710
         Width           =   2610
         Begin VB.OptionButton rbcShow 
            Caption         =   "Percent"
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
            Left            =   1515
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   30
            Width           =   990
         End
         Begin VB.OptionButton rbcShow 
            Caption         =   "Counts"
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
            Left            =   585
            TabIndex        =   23
            Top             =   30
            Value           =   -1  'True
            Width           =   915
         End
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
         Left            =   75
         ScaleHeight     =   165
         ScaleWidth      =   90
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   4605
         Width           =   90
      End
      Begin VB.HScrollBar hbcDate 
         Height          =   255
         LargeChange     =   2
         Left            =   210
         Max             =   1
         Min             =   1
         TabIndex        =   17
         Top             =   4365
         Value           =   1
         Width           =   6645
      End
      Begin VB.PictureBox pbcPctNumber 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   5
         Left            =   4740
         ScaleHeight     =   330
         ScaleWidth      =   2115
         TabIndex        =   16
         Top             =   2775
         Width           =   2115
      End
      Begin VB.PictureBox pbcPctDraw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1155
         Index           =   5
         Left            =   4725
         ScaleHeight     =   1125
         ScaleWidth      =   2115
         TabIndex        =   20
         Top             =   3120
         Width           =   2145
      End
      Begin VB.PictureBox pbcPctNumber 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   4
         Left            =   2460
         ScaleHeight     =   330
         ScaleWidth      =   2115
         TabIndex        =   14
         Top             =   2775
         Width           =   2115
      End
      Begin VB.PictureBox pbcPctDraw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1155
         Index           =   4
         Left            =   2460
         ScaleHeight     =   1125
         ScaleWidth      =   2115
         TabIndex        =   15
         Top             =   3120
         Width           =   2145
      End
      Begin VB.PictureBox pbcPctNumber 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   225
         ScaleHeight     =   330
         ScaleWidth      =   2115
         TabIndex        =   12
         Top             =   2775
         Width           =   2115
      End
      Begin VB.PictureBox pbcPctDraw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1155
         Index           =   3
         Left            =   195
         ScaleHeight     =   1125
         ScaleWidth      =   2115
         TabIndex        =   13
         Top             =   3120
         Width           =   2145
      End
      Begin VB.PictureBox pbcPctNumber 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   2
         Left            =   4725
         ScaleHeight     =   330
         ScaleWidth      =   2115
         TabIndex        =   9
         Top             =   945
         Width           =   2115
      End
      Begin VB.PictureBox pbcPctDraw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1155
         Index           =   2
         Left            =   4740
         ScaleHeight     =   1125
         ScaleWidth      =   2115
         TabIndex        =   10
         Top             =   1290
         Width           =   2145
      End
      Begin VB.PictureBox pbcPctNumber 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   2475
         ScaleHeight     =   330
         ScaleWidth      =   2115
         TabIndex        =   7
         Top             =   945
         Width           =   2115
      End
      Begin VB.PictureBox pbcPctDraw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1155
         Index           =   1
         Left            =   2475
         ScaleHeight     =   1125
         ScaleWidth      =   2115
         TabIndex        =   19
         Top             =   1290
         Width           =   2145
      End
      Begin VB.PictureBox pbcPctNumber 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   210
         ScaleHeight     =   330
         ScaleWidth      =   2115
         TabIndex        =   4
         Top             =   945
         Width           =   2115
      End
      Begin VB.PictureBox pbcPctDraw 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1155
         Index           =   0
         Left            =   210
         ScaleHeight     =   1125
         ScaleWidth      =   2115
         TabIndex        =   5
         Top             =   1290
         Width           =   2145
      End
      Begin VB.ComboBox cbcPrgDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   6675
      End
      Begin VB.CommandButton cmcDone 
         Appearance      =   0  'Flat
         Caption         =   "&Done"
         Height          =   285
         HelpContextID   =   1
         Left            =   3150
         TabIndex        =   18
         Top             =   4695
         Width           =   795
      End
      Begin VB.Label lacCompetitive 
         Appearance      =   0  'Flat
         Caption         =   "Competitive Sellout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2640
         TabIndex        =   11
         Top             =   2535
         Width           =   1755
      End
      Begin VB.Label lacDate 
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
         Height          =   225
         Index           =   2
         Left            =   4740
         TabIndex        =   8
         Top             =   690
         Width           =   2115
      End
      Begin VB.Label lacDate 
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
         Height          =   225
         Index           =   1
         Left            =   2460
         TabIndex        =   6
         Top             =   705
         Width           =   2115
      End
      Begin VB.Label lacDate 
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
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   690
         Width           =   2115
      End
      Begin VB.Label lacAvails 
         Appearance      =   0  'Flat
         Caption         =   "Total Sellout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3000
         TabIndex        =   2
         Top             =   480
         Width           =   1155
      End
   End
End
Attribute VB_Name = "CAvail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CAVAIL.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Copy.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Calendar input screen code
Option Explicit
Option Compare Text
'Library calendar
Dim hmLcf As Integer            'Log calendar library file handle
Dim tmLcf As LCF               'LCF record image
Dim tmLcfSrchKey As LCFKEY0     'LCF key record image
Dim imLcfRecLen As Integer         'LCF record length
Dim hmCHF As Integer            'Contract header file handle
Dim tmChf As CHF               'CHF record image
Dim imCHFRecLen As Integer         'CHF record length
Dim tmClf As CLF               'CHF record image
Dim hmRdf As Integer            'Rate card program file handle
Dim tmRdf As RDF               'RDF record image
Dim imRdfRecLen As Integer         'RDF record length
Dim hmSdf As Integer            'Spot file handle
Dim tmSdf As SDF               'SDF record image
Dim imSdfRecLen As Integer         'SDF record length
Dim hmAsf As Integer
'Spot summary
Dim hmSsf As Integer        'Spot summary file handle
'Dim imLineIndex As Integer
Dim imLnRowNo As Integer
Dim imTerminate As Integer
Dim smPctAvail(0 To 2, 0 To 6) As String    'Percent sold out:Index 1 is for week, Index 2 is for day within week
Dim smPctComp(0 To 2, 0 To 6) As String     'Percent of spots with same competitive: Index 1 is for week, Index 2 is for day within week
Dim smSAvail(0 To 2, 0 To 6) As String    'sold:Index 1 is for week, Index 2 is for day within week
Dim smSComp(0 To 2, 0 To 6) As String     'spots with same competitive: Index 1 is for week, Index 2 is for day within week
Dim smIAvail(0 To 2, 0 To 6) As String    'Inventory:Index 1 is for week, Index 2 is for day within week
Dim smIComp(0 To 2, 0 To 6) As String     'Inventories with same competitive: Index 1 is for week, Index 2 is for day within week
Dim lmWeekDate(0 To 2) As Long
Dim lmCEarliestDate As Long 'Earliest current date
Dim lmCLatestDate As Long   'Latest current date
Dim lmLastLogDate As Long
Dim lmNowDatePlus1 As Long
Dim lmMoNowDate As Long
Dim imFirstActivate As Integer
Dim imVefCode As Integer
Dim imVpfIndex As Integer   'Vehicle option index
Dim imRdfCode As Integer
Dim imIgnoreChg As Integer
Dim imEvtType(0 To 14) As Integer
Dim tmEvt() As EVTINFO  'Event information
Dim tmVcf0() As VCF      'M-F Vcf records
Dim tmVcf6() As VCF      'Sa Vcf records
Dim tmVcf7() As VCF      'Su Vcf records
Private Sub cbcPrgDescription_Change()
    Dim ilLineNo As Integer
    Dim slLineNo As String
    Dim slStr As String
    Dim ilClf As Integer
    Dim ilRet As Integer
    'slStr = cbcPrgDescription.List(cbcPrgDescription.ListIndex)
    'ilPos = InStr(slStr, ":")
    'If ilPos > 0 Then
    '    slStr = Left$(slStr, ilPos - 1)
    '    ilPos = InStr(slStr, " ")
    '    slStr = Mid$(slStr, ilPos + 1)
    '    ilLineNo = Val(slStr)
    '    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
    '        If (tgClfCntr(ilClf).ClfRec.iLine = ilLineNo) Then
    '            imLineIndex = ilClf
    '            mInitLineValues
    '            Exit For
    '        End If
    '    Next ilClf
    'End If
    slStr = cbcPrgDescription.List(cbcPrgDescription.ListIndex)
    ilRet = gParseItem(slStr, 1, " ", slLineNo)
    ilLineNo = Val(slLineNo)
    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) Step 1
        If tgClfCntr(ilClf).ClfRec.iLine = ilLineNo Then
            imLnRowNo = ilClf + 1
            mInitLineValues
            Exit For
        End If
    Next ilClf
End Sub
Private Sub cbcPrgDescription_Click()
    cbcPrgDescription_Change
End Sub
Private Sub cmcDone_Click()
    mTerminate
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
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        mTerminate
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    Erase tmEvt
    Erase tmVcf0
    Erase tmVcf6
    Erase tmVcf7
    If tgUrf(0).iRemoteUserID > 0 Then
        ilRet = btrClose(hmAsf)
        btrDestroy hmAsf
    End If
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmRdf)
    btrDestroy hmRdf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    
    Set CAvail = Nothing   'Remove data segment
    
End Sub

Private Sub hbcDate_Change()
    Dim ilLoop As Integer
    If imIgnoreChg = True Then
        imIgnoreChg = False
        Exit Sub
    End If
    For ilLoop = 0 To 5 Step 1
        pbcPctNumber(ilLoop).Cls
        pbcPctDraw(ilLoop).Cls
    Next ilLoop
    mBuildWeeks
    For ilLoop = 0 To 5 Step 1
        pbcPctNumber_Paint ilLoop
        pbcPctDraw_Paint ilLoop
    Next ilLoop
End Sub
Private Sub lacAvails_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub lacCompetitive_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub lacDate_Click(Index As Integer)
    pbcClickFocus.SetFocus
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBuildWeeks                     *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build Week Data                 *
'*                                                     *
'*******************************************************
Private Sub mBuildWeeks()
    Dim llDate As Long
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilDay As Integer
    Dim ilWeek As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llFlStartDate As Long
    Dim llFlEndDate As Long
    Dim llGenDate As Long
    Dim ilUnitInv As Integer
    Dim llUnitSold As Long
    Dim ilCompUnitSold As Integer
    Dim ilNoAvailsForBreak As Integer  'Number of avails for Not in same break(Competitive)
    Dim ilNoAvailsForNBB As Integer 'Number of avails for Not Back To Back (competitive)
    Dim slDateShow As String
    Dim slDate As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llCompTime As Long
    Dim slLength As String
    Dim ilTime0 As Integer
    Dim ilTime1 As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim ilLen As Integer
    Dim ilAvails As Integer
    Dim ilSpots As Integer
    Dim ilNoDays As Integer
    Dim ilNoWkDays As Integer
    Dim ilFirstDay As Integer
    Dim ilCff As Integer
    ReDim ilDays(0 To 6) As Integer
    Dim ilAdfCode As Integer
    Dim ilAdfCodeCount As Integer
    Dim ilCompAvailCount As Integer
    Dim ilSvToggleAvail As Integer
    
    ilSvToggleAvail = imToggleAvail
    imToggleAvail = 1
    'If ((tgClfCntr(imLineIndex).ClfRec.iStartTime(0) <> 1) Or (tgClfCntr(imLineIndex).ClfRec.iStartTime(1) <> 0)) And ((tgClfCntr(imLineIndex).ClfRec.iEndTime(0) <> 1) Or (tgClfCntr(imLineIndex).ClfRec.iEndTime(1) <> 0)) Then
        'gUnpackTimeLong tgClfCntr(imLineIndex).ClfRec.iStartTime(0), tgClfCntr(imLineIndex).ClfRec.iStartTime(1), False, llOvStartTime
        'gUnpackTimeLong tgClfCntr(imLineIndex).ClfRec.iEndTime(0), tgClfCntr(imLineIndex).ClfRec.iEndTime(1), True, llOvEndTime
    If imLnSave(4, imLnRowNo) = 1 Then
        If smLnSave(2, imLnRowNo) <> "" Then
            gPackTime smLnSave(2, imLnRowNo), ilTime0, ilTime1
            gUnpackTimeLong ilTime0, ilTime1, False, llOvStartTime
            gPackTime smLnSave(3, imLnRowNo), ilTime0, ilTime1
            gUnpackTimeLong ilTime0, ilTime1, False, llOvEndTime
        End If
    Else
        llOvStartTime = 0
        llOvEndTime = 0
    End If
    ilLen = Val(smLnSave(16, imLnRowNo))
    If tgVpf(imVpfIndex).sSCompType = "T" Then  'By time
        gUnpackLength tgVpf(imVpfIndex).iSCompLen(0), tgVpf(imVpfIndex).iSCompLen(1), "3", False, slLength
        llCompTime = CLng(gLengthToCurrency(slLength))
    Else
        llCompTime = 0
    End If
    llDate = 7 * (hbcDate.Value - 1) + lmCEarliestDate
    For ilWeek = LBound(smPctAvail, 1) To UBound(smPctAvail, 1) Step 1
        'Build week
        ilFound = False
        For ilLoop = LBound(smPctAvail, 1) To UBound(smPctAvail, 1) Step 1
            If lmWeekDate(ilLoop) = llDate Then
                ilFound = True
                For ilDay = 0 To 6 Step 1
                    smPctAvail(ilWeek, ilDay) = smPctAvail(ilLoop, ilDay)
                    smPctComp(ilWeek, ilDay) = smPctComp(ilLoop, ilDay)
                    smSAvail(ilWeek, ilDay) = smSAvail(ilLoop, ilDay)
                    smSComp(ilWeek, ilDay) = smSComp(ilLoop, ilDay)
                    smIAvail(ilWeek, ilDay) = smIAvail(ilLoop, ilDay)
                    smIComp(ilWeek, ilDay) = smIComp(ilLoop, ilDay)
                Next ilDay
                Exit For
            End If
        Next ilLoop
        lmWeekDate(ilWeek) = llDate
        If Not ilFound Then
            llStartDate = llDate
            llEndDate = llStartDate + 6
            For llGenDate = llStartDate To llEndDate Step 1
                'Set ilDay
                For ilDay = 0 To 6 Step 1
                    ilDays(ilDay) = True    'Include all days was False
                Next ilDay
                ilDay = gWeekDayLong(llGenDate)
                ilSpots = 0
                ilNoWkDays = 0
                ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
                Do While ilCff <> -1
                    If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                        gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llFlStartDate    'Week Start date
                        gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), llFlEndDate    'Week Start date
                        If (llGenDate >= llFlStartDate) And (llGenDate <= llFlEndDate) Then
                            If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                                ilSpots = ilSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                                If tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0 Then
                                    ilDays(ilDay) = True
                                End If
                            Else
                                ilSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                                'Distribute spots
                                ilNoDays = 0
                                ilFirstDay = -1
                                For ilLoop = 0 To 6 Step 1
                                    If (tgCffCntr(ilCff).CffRec.iDay(ilLoop) > 0) Or (tgCffCntr(ilCff).CffRec.sXDay(ilLoop) = "1") Then
                                        ilNoDays = ilNoDays + 1
                                        If ilLoop <= ilDay Then
                                            ilNoWkDays = ilNoWkDays + 1
                                        End If
                                        If ilFirstDay = -1 Then
                                            ilFirstDay = ilLoop
                                        End If
                                    End If
                                Next ilLoop
                                If (tgCffCntr(ilCff).CffRec.iDay(ilDay) > 0) Or (tgCffCntr(ilCff).CffRec.sXDay(ilDay) = "1") Then
                                    ilDays(ilDay) = True
                                End If
                                If ilNoDays - 1 > 0 Then
                                    If ilSpots <= ilNoDays Then
                                        If ilNoWkDays <= ilSpots Then
                                            ilSpots = 1
                                        Else
                                            ilSpots = 0
                                        End If
                                    Else
                                        If ilDay = ilFirstDay Then
                                            ilSpots = ilSpots - ((ilSpots \ ilNoDays) * (ilNoDays - 1))
                                        Else
                                            ilSpots = ilSpots \ ilNoDays
                                        End If
                                    End If
                                End If
                            End If
                            Exit Do
                        End If
                    End If
                    ilCff = tgCffCntr(ilCff).iNextCff
                Loop
                ilAdfCode = 0
                mAvailCount hmSsf, hmSdf, hmLcf, hmAsf, tgClfCntr(imLnRowNo - 1).ClfRec.iLine, imVefCode, imRdfCode, llOvStartTime, llOvEndTime, ilLen, ilDays(), llGenDate, llGenDate, llUnitSold, ilAvails, ilUnitInv, ilCompUnitSold, ilNoAvailsForBreak, ilNoAvailsForNBB, llStartTime, llEndTime, True, ilAdfCode, ilAdfCodeCount, ilCompAvailCount
                llUnitSold = llUnitSold + ilSpots
                ilCompUnitSold = ilCompUnitSold + ilSpots
                'slDate = Format$(llGenDate, "m/d/yy")
                'mGetLineTimes slDate, llStartTime, llEndTime
                'ilUnitInv = 0
                'llSecInv = 0
                'llUnitSold = 0
                'llSecSold = 0
                'llCompSecSold = 0
                'ilCompUnitSold = 0
                'ilNoAvailsForBreak = 0
                'ilNoAvailsForNBB = 0
                'If llGenDate <= lmCLatestDate Then
                '    slDate = Format$(llGenDate, "m/d/yy")
                '    gPackDate slDate, ilLogDate0, ilLogDate1
                '    mGetLineTimes slDate, llStartTime, llEndTime
                '    ilDay = gWeekDayLong(llGenDate)
                '    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                '    If (tmSsf.iStartTime(0) <> 0) Or (tmSsf.iStartTime(1) <> 0) Then
                '        tmSsfSrchKey.sType = "O"
                '        tmSsfSrchKey.iVefCode = imVefCode
                '        tmSsfSrchKey.iDate(0) = ilLogDate0
                '        tmSsfSrchKey.iDate(1) = ilLogDate1
                '        tmSsfSrchKey.iStartTime(0) = 0
                '        tmSsfSrchKey.iStartTime(1) = 0
                '        ilRet = btrGetEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                '    Else
                '        ilRet = BTRV_ERR_NONE
                '    End If
                '    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = "O") And (tmSsf.iVefCode = imVefCode) And (tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1)
                '        ilEvt = 1
                '        Do While ilEvt <= tmSsf.iCount
                '           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                '            If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Avail
                '                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                '                llAvailTime = CLng(gTimeToCurrency(slTime, False))
                '                If (llAvailTime >= llStartTime) And (llAvailTime <= llEndTime) Then
                '                    ilUnitInv = ilUnitInv + (tmAvail.iAvInfo And &H1F)
                '                    llSecInv = llSecInv + tmAvail.iLen
                '                    ilNoAvailsForBreak = ilNoAvailsForBreak + 1
                '                    For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                '                        ilEvt = ilEvt + 1
                '                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                '                        'Don't count PSA or Promos
                '                        'If (tmEvt(ilEvt).sCntrType <> "S") And (tmEvt(ilEvt).sCntrType <> "M") Then
                '                            llUnitSold = llUnitSold + 1
                '                            llSecSold = llSecSold + (tmSpot.iPosLen And &HFFF) 'units contains spot length
                '                            If (igComp1Code > 0) And (igComp1Code = tmSpot.iMnfComp(0)) Then
                '                                ilCompUnitSold = ilCompUnitSold + 1
                '                                llCompSecSold = llCompSecSold + (tmSpot.iPosLen And &HFFF) 'units contains spot length
                '                                If tmAvail.iAvInfo And &H1F >= 3 Then
                '                                    ilNoAvailsForNBB = ilNoAvailsForNBB + 1
                '                                End If
                '                            Else
                '                                If (igComp1Code > 0) And (igComp1Code = tmSpot.iMnfComp(1)) Then
                '                                    ilCompUnitSold = ilCompUnitSold + 1
                '                                    llCompSecSold = llCompSecSold + (tmSpot.iPosLen And &HFFF) 'units contains spot length
                '                                    If tmAvail.iAvInfo And &H1F >= 3 Then
                '                                        ilNoAvailsForNBB = ilNoAvailsForNBB + 1
                '                                    End If
                '                                Else
                '                                    If (igComp2Code > 0) And (igComp2Code = tmSpot.iMnfComp(0)) Then
                '                                        ilCompUnitSold = ilCompUnitSold + 1
                '                                        llCompSecSold = llCompSecSold + (tmSpot.iPosLen And &HFFF) 'units contains spot length
                '                                        If tmAvail.iAvInfo And &H1F >= 3 Then
                '                                            ilNoAvailsForNBB = ilNoAvailsForNBB + 1
                '                                        End If
                '                                    Else
                '                                        If (igComp2Code > 0) And (igComp2Code = tmSpot.iMnfComp(1)) Then
                '                                            ilCompUnitSold = ilCompUnitSold + 1
                '                                            llCompSecSold = llCompSecSold + (tmSpot.iPosLen And &HFFF) 'units contains spot length
                '                                            If tmAvail.iAvInfo And &H1F >= 3 Then
                '                                                ilNoAvailsForNBB = ilNoAvailsForNBB + 1
                '                                            End If
                '                                        Else
                '                                            ilNoAvailsForNBB = ilNoAvailsForNBB + 1
                '                                        End If
                '                                    End If
                '                                End If
                '                            End If
                '                        'End If
                '                    Next ilSpot
                '                End If
                '            End If
                '            ilEvt = ilEvt + 1
                '        Loop
                '        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                '        ilRet = btrGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                '    Loop
                'End If
                'Compute %
                'smPctAvail(ilWeek, ilDay + 1) = "0" 'using 0 as place holder
                'smPctComp(ilWeek, ilDay + 1) = "0"
                'smSAvail(ilWeek, ilDay + 1) = "0" 'using 0 as place holder
                'smSComp(ilWeek, ilDay + 1) = "0"
                'smIAvail(ilWeek, ilDay + 1) = "0" 'using 0 as place holder
                'smIComp(ilWeek, ilDay + 1) = "0"
                smPctAvail(ilWeek, ilDay) = "0" 'using 0 as place holder
                smPctComp(ilWeek, ilDay) = "0"
                smSAvail(ilWeek, ilDay) = "0"  'using 0 as place holder
                smSComp(ilWeek, ilDay) = "0"
                smIAvail(ilWeek, ilDay) = "0" 'using 0 as place holder
                smIComp(ilWeek, ilDay) = "0"
                If ilUnitInv > 0 Then
                    'smPctAvail(ilWeek, ilDay + 1) = Trim$(str$(Int((CLng(llUnitSold) * 100) / ilUnitInv)))
                    'smSAvail(ilWeek, ilDay + 1) = Trim$(str$(llUnitSold))
                    'smIAvail(ilWeek, ilDay + 1) = Trim$(str$(ilUnitInv))
                    smPctAvail(ilWeek, ilDay) = Trim$(str$(Int((CLng(llUnitSold) * 100) / ilUnitInv)))
                    smSAvail(ilWeek, ilDay) = Trim$(str$(llUnitSold))
                    smIAvail(ilWeek, ilDay) = Trim$(str$(ilUnitInv))
                    If tgVpf(imVpfIndex).sSCompType = "T" Then  'By time
                        '# of Avails= Take smallest of number of avails or Book length/Separation Length
                        If llCompTime > 0 Then
                            If (llEndTime - llStartTime) / llCompTime < ilNoAvailsForBreak Then
                                'smPctComp(ilWeek, ilDay + 1) = Trim$(str$(Int((CLng(ilCompUnitSold) * 100) / ((llEndTime - llStartTime) / llCompTime))))
                                'smSComp(ilWeek, ilDay + 1) = Trim$(str$(ilCompUnitSold))
                                'smIComp(ilWeek, ilDay + 1) = Trim$(str$((llEndTime - llStartTime) / llCompTime))
                                smPctComp(ilWeek, ilDay) = Trim$(str$(Int((CLng(ilCompUnitSold) * 100) / ((llEndTime - llStartTime) / llCompTime))))
                                smSComp(ilWeek, ilDay) = Trim$(str$(ilCompUnitSold))
                                smIComp(ilWeek, ilDay) = Trim$(str$((llEndTime - llStartTime) / llCompTime))
                            Else
                                If ilNoAvailsForBreak > 0 Then
                                    'smPctComp(ilWeek, ilDay + 1) = Trim$(str$(Int((CLng(ilCompUnitSold) * 100) / ilNoAvailsForBreak)))
                                    'smSComp(ilWeek, ilDay + 1) = Trim$(str$(ilCompUnitSold))
                                    'smIComp(ilWeek, ilDay + 1) = Trim$(str$(ilNoAvailsForBreak))
                                    smPctComp(ilWeek, ilDay) = Trim$(str$(Int((CLng(ilCompUnitSold) * 100) / ilNoAvailsForBreak)))
                                    smSComp(ilWeek, ilDay) = Trim$(str$(ilCompUnitSold))
                                    smIComp(ilWeek, ilDay) = Trim$(str$(ilNoAvailsForBreak))
                                End If
                            End If
                        Else
                            If ilNoAvailsForBreak > 0 Then
                                'smPctComp(ilWeek, ilDay + 1) = Trim$(str$(Int((CLng(ilCompUnitSold) * 100) / ilNoAvailsForBreak)))
                                'smSComp(ilWeek, ilDay + 1) = Trim$(str$(ilCompUnitSold))
                                'smIComp(ilWeek, ilDay + 1) = Trim$(str$(ilNoAvailsForBreak))
                                smPctComp(ilWeek, ilDay) = Trim$(str$(Int((CLng(ilCompUnitSold) * 100) / ilNoAvailsForBreak)))
                                smSComp(ilWeek, ilDay) = Trim$(str$(ilCompUnitSold))
                                smIComp(ilWeek, ilDay) = Trim$(str$(ilNoAvailsForBreak))
                            End If
                        End If
                    ElseIf tgVpf(imVpfIndex).sSCompType = "B" Then  'Not in same break
                        If ilNoAvailsForBreak > 0 Then
                            'smPctComp(ilWeek, ilDay + 1) = Trim$(str$(Int((CLng(ilCompUnitSold) * 100) / ilNoAvailsForBreak)))
                            'smSComp(ilWeek, ilDay + 1) = Trim$(str$(ilCompUnitSold))
                            'smIComp(ilWeek, ilDay + 1) = Trim$(str$(ilNoAvailsForBreak))
                            smPctComp(ilWeek, ilDay) = Trim$(str$(Int((CLng(ilCompUnitSold) * 100) / ilNoAvailsForBreak)))
                            smSComp(ilWeek, ilDay) = Trim$(str$(ilCompUnitSold))
                            smIComp(ilWeek, ilDay) = Trim$(str$(ilNoAvailsForBreak))
                        End If
                    ElseIf tgVpf(imVpfIndex).sSCompType = "N" Then  'Not Back to Back
                        If ilNoAvailsForNBB > 0 Then
                            'smPctComp(ilWeek, ilDay + 1) = Trim$(str$(Int((CLng(ilCompUnitSold) * 100) / ilNoAvailsForNBB)))
                            'smSComp(ilWeek, ilDay + 1) = Trim$(str$(ilCompUnitSold))
                            'smIComp(ilWeek, ilDay + 1) = Trim$(str$(ilNoAvailsForNBB))
                            smPctComp(ilWeek, ilDay) = Trim$(str$(Int((CLng(ilCompUnitSold) * 100) / ilNoAvailsForNBB)))
                            smSComp(ilWeek, ilDay) = Trim$(str$(ilCompUnitSold))
                            smIComp(ilWeek, ilDay) = Trim$(str$(ilNoAvailsForNBB))
                        End If
                    End If
                End If
            Next llGenDate
        End If
        slDate = Format$(llDate, "m/d/yy")
        slDateShow = Format$(slDate, "m/d/yy") & "-" & Format$(gDateValue(slDate) + 6, "m/d/yy")
        lacDate(ilWeek).Caption = slDateShow
        llDate = llDate + 7
    Next ilWeek
    imToggleAvail = ilSvToggleAvail
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetDateSpan                    *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get LCF span                    *
'*                                                     *
'*******************************************************
Private Sub mGetDateSpan()
    Dim ilType As Integer
    Dim ilRet As Integer
    Dim slDate As String
    imVpfIndex = gVpfFind(CAvail, imVefCode)
    If (tgVpf(imVpfIndex).iLLD(0) <> 0) Or (tgVpf(imVpfIndex).iLLD(1) <> 0) Then
        gUnpackDate tgVpf(imVpfIndex).iLLD(0), tgVpf(imVpfIndex).iLLD(1), slDate
        lmLastLogDate = gDateValue(slDate)
    Else
        lmLastLogDate = -1
    End If
    lmCEarliestDate = -1
    lmCLatestDate = -1
    tmLcfSrchKey.iType = 0
    ilType = 0
    tmLcfSrchKey.sStatus = "C"
    tmLcfSrchKey.iVefCode = imVefCode
    tmLcfSrchKey.iLogDate(0) = 257  'Year 1/1/2100
    tmLcfSrchKey.iLogDate(1) = 2100
    tmLcfSrchKey.iSeqNo = 1
    ilRet = btrGetLessOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    If (ilRet = BTRV_ERR_NONE) And (tmLcf.sStatus = "C") And (tmLcf.iVefCode = imVefCode) And (tmLcf.iType = ilType) Then
        gUnpackDate tmLcf.iLogDate(0), tmLcf.iLogDate(1), slDate
        lmCLatestDate = gDateValue(gObtainNextSunday(slDate))
        tmLcfSrchKey.iType = ilType
        tmLcfSrchKey.sStatus = "C"
        tmLcfSrchKey.iVefCode = imVefCode
        tmLcfSrchKey.iLogDate(0) = 257  'Year 1/1/1900
        tmLcfSrchKey.iLogDate(1) = 1900
        tmLcfSrchKey.iSeqNo = 1
        ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If (ilRet = BTRV_ERR_NONE) And (tmLcf.sStatus = "C") And (tmLcf.iVefCode = imVefCode) And (tmLcf.iType = ilType) Then
            gUnpackDate tmLcf.iLogDate(0), tmLcf.iLogDate(1), slDate
            lmCEarliestDate = gDateValue(gObtainPrevMonday(slDate))
        Else
            slDate = Format$(gNow(), "m/d/yy")
            slDate = gObtainNextMonday(slDate)
            slDate = gDecOneWeek(slDate)
            lmCEarliestDate = gDateValue(slDate)
        End If
    Else
        slDate = Format$(gNow(), "m/d/yy")
        slDate = gObtainNextMonday(slDate)
        lmCEarliestDate = gDateValue(slDate)
        lmCLatestDate = lmCEarliestDate + 6
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
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
    Dim slDate As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    'Log calendar file
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imTerminate = False
    hmLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", CAvail
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf)  'Get and save ADF record length
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", CAvail
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)  'Get and save ADF record length
    hmRdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rdf.Btr)", CAvail
    On Error GoTo 0
    imRdfRecLen = Len(tmRdf)  'Get and save ADF record length
    hmSdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sdf.Btr)", CAvail
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)  'Get and save ADF record length
    hmSsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Ssf.Btr)", CAvail
    On Error GoTo 0
    If tgUrf(0).iRemoteUserID > 0 Then
        hmAsf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmAsf, "", sgDBPath & "ASF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: ASF.Btr)", CAvail
        On Error GoTo 0
    End If
    slDate = Format$(gNow(), "m/d/yy")
    lmNowDatePlus1 = gDateValue(slDate) + 1
    slDate = gObtainPrevMonday(slDate)
    lmMoNowDate = gDateValue(slDate)
    For ilLoop = 0 To 14 Step 1
        imEvtType(ilLoop) = False
    Next ilLoop
    'For ilLoop = 2 To 9 Step 1  'Avail only
    imEvtType(2) = True
    For ilLoop = 6 To 9 Step 1  'Avail only
        imEvtType(ilLoop) = True
    Next ilLoop
    ReDim tmVcf0(0 To 0) As VCF      'M-F Vcf records
    ReDim tmVcf6(0 To 0) As VCF      'Sa Vcf records
    ReDim tmVcf7(0 To 0) As VCF      'Su Vcf records
    mLnPop
    gCenterModalForm CAvail
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDraw                           *
'*                                                     *
'*             Created:10/08/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show made-up data              *
'*                                                     *
'*******************************************************
Private Sub mInitLineValues()
'
'   mInit
'   Where:
'
    Dim ilWeek As Integer
    Dim ilDay As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slStartDate As String
    Dim llMoStartDate As Long
    Dim ilCff As Integer
    Dim slLineType As String
    Dim slNameCode As String
    Dim slVefCode As String
    For ilLoop = 0 To 5 Step 1
        pbcPctNumber(ilLoop).Cls
        pbcPctDraw(ilLoop).Cls
    Next ilLoop
    For ilWeek = LBound(smPctAvail, 1) To UBound(smPctAvail, 1) Step 1
        For ilDay = 0 To 6 Step 1
            smPctAvail(ilWeek, ilDay) = "0"
            smPctComp(ilWeek, ilDay) = "0"
            smSAvail(ilWeek, ilDay) = "0"
            smSComp(ilWeek, ilDay) = "0"
            smIAvail(ilWeek, ilDay) = "0"
            smIComp(ilWeek, ilDay) = "0"
        Next ilDay
        lmWeekDate(ilWeek) = 0
    Next ilWeek
    If imLnRowNo - 1 < 0 Then
        Exit Sub
    End If
    'imVefCode = tgClfCntr(imLnRowNo - 1).ClfRec.iVefCode
    'tmRdfSrchKey.iCode = tgClfCntr(imLnRowNo - 1).ClfRec.iRdfCode
    'ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
    imVefCode = -1
    slLineType = mGetLineType(imLnRowNo)
    If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
        gFindMatch smLnSave(1, imLnRowNo), 0, Contract.lbcLnVehicle(igTabMapIndex)
        If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
            slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
            ilRet = gParseItem(slNameCode, 2, "\", slVefCode)
            imVefCode = Val(slVefCode)
        End If
    End If
    If imVefCode <= 0 Then
        Exit Sub
    End If
    imRdfCode = imLnSave(1, imLnRowNo)
    tmClf = tgClfCntr(imLnRowNo - 1).ClfRec
    llMoStartDate = 0
    ilCff = tgClfCntr(imLnRowNo - 1).iFirstCff
    If ilCff <> -1 Then
        Do While ilCff <> -1
            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate    'Week Start date
                llMoStartDate = gDateValue(gObtainPrevMonday(slStartDate))
                Exit Do
            End If
            ilCff = tgCffCntr(ilCff).iNextCff
        Loop
    End If
    If llMoStartDate = 0 Then
        slStartDate = Format$(gNow(), "m/d/yy")
        llMoStartDate = gDateValue(gObtainPrevMonday(slStartDate))
    End If
    'Obtain the earliest LCF date and Latest LCF Date
    mGetDateSpan
    imIgnoreChg = True
    hbcDate.Min = 1
    imIgnoreChg = True
    If (lmCLatestDate - lmCEarliestDate) \ 7 + 1 - hbcDate.LargeChange < 1 Then
        hbcDate.Max = 1
    Else
        hbcDate.Max = (lmCLatestDate - lmCEarliestDate) \ 7 + 1 - hbcDate.LargeChange
    End If
    imIgnoreChg = False
    If (llMoStartDate >= lmCEarliestDate) And (llMoStartDate <= lmCLatestDate) And (llMoStartDate >= lmMoNowDate) Then
        If hbcDate.Value <> (llMoStartDate - lmCEarliestDate) \ 7 + 1 Then
            '3/25/16: Verify within boundarys of the scroll
            'hbcDate.Value = (llMoStartDate - lmCEarliestDate) \ 7 + 1
            If (llMoStartDate - lmCEarliestDate) \ 7 + 1 >= hbcDate.Max Then
                hbcDate.Value = hbcDate.Max
            Else
                If (llMoStartDate - lmCEarliestDate) \ 7 + 1 >= hbcDate.Min Then
                    hbcDate.Value = (llMoStartDate - lmCEarliestDate) \ 7 + 1
                Else
                    hbcDate.Value = hbcDate.Min
                End If
            End If
        Else
            hbcDate_Change
        End If
    Else
        If (lmMoNowDate >= lmCEarliestDate) And (lmMoNowDate <= lmCLatestDate) Then
            If hbcDate.Value <> (lmMoNowDate - lmCEarliestDate) \ 7 + 1 Then
                '3/25/16: Verify within boundarys of the scroll
                'hbcDate.Value = (lmMoNowDate - lmCEarliestDate) \ 7 + 1
                If (lmMoNowDate - lmCEarliestDate) \ 7 + 1 >= hbcDate.Max Then
                    hbcDate.Value = hbcDate.Max
                Else
                    If (lmMoNowDate - lmCEarliestDate) \ 7 + 1 >= hbcDate.Min Then
                        hbcDate.Value = (lmMoNowDate - lmCEarliestDate) \ 7 + 1
                    Else
                        hbcDate.Value = hbcDate.Min
                    End If
                End If
            Else
                hbcDate_Change
            End If
        Else
            If hbcDate.Value <> 1 Then
                hbcDate.Value = 1
            Else
                hbcDate_Change
            End If
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mLnPop                          *
'*                                                     *
'*             Created:10/08/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                      Taken from: mPopAirLine        *
'*                                                     *
'*******************************************************
Private Sub mLnPop()
    'Dim ilClf As Integer
    'Dim ilLoop As Integer
    'Dim slStr As String
    'Dim ilRow As Integer
    'Dim slLineType As String
    Dim ilClf As Integer
    Dim slLineType As String
    Dim ilLineIndex As Integer
    Dim slNameCode As String
    Dim slVefName As String
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilMax As Integer
    Dim ilAddLine As Integer
    Dim slCode As String
    Dim ilVef As Integer
    'ilRow = 0
    'cbcPrgDescription.Clear
    'For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
    '    If (tgClfCntr(ilClf).iStatus = 0) Or (tgClfCntr(ilClf).iStatus = 1) Then   '(0=New; 1=Old & retain; 2=Old & delete;-1=ignore)
    '        slStr = ""
    '        slLineType = mGetLineType(ilClf + 1)
    '        If (slLineType <> "O") And (slLineType <> "A") Then
    '            'For ilLoop = 0 To Contract!lbcLnProgram(igTabMapIndex).ListCount - 1 Step 1
    '            '    If Contract!lbcPrg(igTabMapIndex).ItemData(ilLoop) = tgClfCntr(ilClf).ClfRec.iRdfCode Then
    '            '        slStr = Contract!lbcLnProgram(igTabMapIndex).List(ilLoop)
    '            '        Exit For
    '            '    End If
    '            'Next ilLoop
    '            slStr = mMakeDPName(ilClf + 1)  'mMakePrgName(tgClfCntr(ilClf).ClfRec.iRdfCode)
    '            cbcPrgDescription.AddItem "Line " & Trim$(Str$(tgClfCntr(ilClf).ClfRec.iLine)) & ": " & slStr
    '            If tgClfCntr(ilClf).ClfRec.iLine = igLineNo Then
    '                ilRow = cbcPrgDescription.ListCount - 1
    '            End If
    '        End If
    '    End If
    'Next ilClf
    'cbcPrgDescription.ListIndex = ilRow
    cbcPrgDescription.Clear
    ilLineIndex = -1
    ilMax = UBound(tgClfCntr) - 1
    If (smLnSave(1, ilMax + 2) <> "") And (imLnSave(1, ilMax + 2) > 0) And (Val(smLnSave(16, ilMax + 2)) > 0) Then
        ilMax = UBound(tgClfCntr)
    End If
    For ilClf = LBound(tgClfCntr) To ilMax Step 1
        ilAddLine = True
        If tgClfCntr(ilClf).iCancel Then
            'If imLnRowNo <> ilClf + 1 Then
                ilAddLine = False
            'End If
        End If
        If ilAddLine Then
            slLineType = mGetLineType(ilClf + 1)
            If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                gFindMatch smLnSave(1, ilClf + 1), 0, Contract.lbcLnVehicle(igTabMapIndex)
                If gLastFound(Contract.lbcLnVehicle(igTabMapIndex)) >= 0 Then
                    slNameCode = tmVehicleCode(gLastFound(Contract.lbcLnVehicle(igTabMapIndex))).sKey    'lbcVehicle.List(gLastFound(lbcLnVehicle(igTabMapIndex)))
                    ilRet = gParseItem(slNameCode, 1, "\", slVefName)
                    ilRet = gParseItem(slVefName, 3, "|", slVefName)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilVef = gBinarySearchVef(Val(slCode))
                    If ilVef <> -1 Then
                        If tgMVef(ilVef).sType <> "G" Then
                    'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                    '    If imLnSave(1, ilClf + 1) = tgMRdf(ilRdf).iCode Then
                            slStr = mMakeDPName(ilClf + 1)  'mMakePrgName(tgMRdf(ilRdf).iCode)
                            If tgClfCntr(ilClf).iCancel Then
                                cbcPrgDescription.AddItem Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine)) & " " & Trim$(slVefName) & " " & slStr & " " & smLnSave(16, ilClf + 1) & "s " & "Canceled"
                            Else
                                cbcPrgDescription.AddItem Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine)) & " " & Trim$(slVefName) & " " & slStr & " " & smLnSave(16, ilClf + 1) & "s " & smLnSave(14, ilClf + 1) & "-" & smLnSave(15, ilClf + 1)
                            End If
                            If tgClfCntr(ilClf).ClfRec.iLine = igLineNo Then
                                ilLineIndex = cbcPrgDescription.ListCount - 1
                            End If
                    '        Exit For
                    '    End If
                    'Next ilRdf
                        End If
                    End If
                End If
            Else
                'For ilVeh = LBound(tgPackVeh) To UBound(tgPackVeh) - 1 Step 1
                '    If StrComp(Trim$(smLnSave(1, ilClf + 1)), Trim$(tgPackVeh(ilVeh).sName), 1) = 0 Then
                '        slVefName = tgPackVeh(ilVeh).sName
                '        'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                '        '    If StrComp(smLnSave(20, ilClf + 1), Trim$(tgMRdf(ilRdf).sName), 1) = 0 Then
                '                slStr = mMakeDPName(ilClf + 1)  'mMakePrgName(tgMRdf(ilRdf).iCode)
                '                If tgClfCntr(ilClf).iCancel Then
                '                    cbcPrgDescription.AddItem Trim$(Str$(tgClfCntr(ilClf).ClfRec.iLine)) & " " & Trim$(slVefName) & " " & slStr & " " & smLnSave(16, ilClf + 1) & "s " & "Canceled"
                '                Else
                '                    cbcPrgDescription.AddItem Trim$(Str$(tgClfCntr(ilClf).ClfRec.iLine)) & " " & Trim$(slVefName) & " " & slStr & " " & smLnSave(16, ilClf + 1) & "s " & smLnSave(14, ilClf + 1) & "-" & smLnSave(15, ilClf + 1)
                '                End If
                 '               If imLnRowNo = ilClf + 1 Then
                '                    ilLineIndex = cbcPrgDescription.ListCount - 1
                '                End If
                '        '        Exit For
                '        '    End If
                '        'Next ilRdf
                '    End If
                'Next ilVeh
            End If
        End If
    Next ilClf
    If ilLineIndex >= 0 Then
        cbcPrgDescription.ListIndex = ilLineIndex
    End If
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
    igManUnload = YES
    Unload CAvail
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcPctDraw_Paint(Index As Integer)
    Dim ilLoop As Integer
    Dim fStartX As Single
    Dim fStartY As Single
    Dim ilPct As Integer
    If Index <= 2 Then
        For ilLoop = 0 To 6 Step 1
            ilPct = Val(smPctAvail(Index, ilLoop))
            fStartY = pbcPctDraw(Index).Height - (pbcPctDraw(Index).Height * ilPct) / 100
            'fStartX = 300 * (ilLoop - 1) + 15
            fStartX = 300 * (ilLoop) + 15
            pbcPctDraw(Index).Line (fStartX, fStartY)-(fStartX + 270, pbcPctDraw(Index).Height), GREEN, BF
        Next ilLoop
    Else
        For ilLoop = 0 To 6 Step 1
            ilPct = Val(smPctComp(Index - 3, ilLoop))
            fStartY = pbcPctDraw(Index).Height - (pbcPctDraw(Index).Height * ilPct) / 100
            'fStartX = 300 * (ilLoop - 1) + 15
            fStartX = 300 * (ilLoop) + 15
            pbcPctDraw(Index).Line (fStartX, fStartY)-(fStartX + 270, pbcPctDraw(Index).Height), MAGENTA, BF
        Next ilLoop
    End If
End Sub
Private Sub pbcPctNumber_Paint(Index As Integer)
    Dim ilLoop As Integer
    If rbcShow(0).Value Then
        If Index <= 2 Then
            For ilLoop = 0 To 6 Step 1
                'pbcPctNumber(Index).CurrentX = 300 * (ilLoop - 1) + 30
                pbcPctNumber(Index).CurrentX = 300 * (ilLoop) + 30
                pbcPctNumber(Index).CurrentY = -15
                pbcPctNumber(Index).Print smSAvail(Index, ilLoop) & "/"
                'pbcPctNumber(Index).CurrentX = 300 * (ilLoop - 1) + 30
                pbcPctNumber(Index).CurrentX = 300 * (ilLoop) + 30
                pbcPctNumber(Index).CurrentY = 150
                pbcPctNumber(Index).Print smIAvail(Index, ilLoop)
            Next ilLoop
        Else
            For ilLoop = 0 To 6 Step 1
                'pbcPctNumber(Index).CurrentX = 300 * (ilLoop - 1) + 30
                pbcPctNumber(Index).CurrentX = 300 * (ilLoop) + 30
                pbcPctNumber(Index).CurrentY = -15
                pbcPctNumber(Index).Print smSComp(Index - 3, ilLoop) & "/"
                'pbcPctNumber(Index).CurrentX = 300 * (ilLoop - 1) + 30
                pbcPctNumber(Index).CurrentX = 300 * (ilLoop) + 30
                pbcPctNumber(Index).CurrentY = 150
                pbcPctNumber(Index).Print smIComp(Index - 3, ilLoop)
            Next ilLoop
        End If
    Else
        If Index <= 2 Then
            For ilLoop = 0 To 6 Step 1
                'pbcPctNumber(Index).CurrentX = 300 * (ilLoop - 1) + 30
                pbcPctNumber(Index).CurrentX = 300 * (ilLoop) + 30
                pbcPctNumber(Index).CurrentY = -15
                pbcPctNumber(Index).Print smPctAvail(Index, ilLoop)
            Next ilLoop
        Else
            For ilLoop = 0 To 6 Step 1
                'pbcPctNumber(Index).CurrentX = 300 * (ilLoop - 1) + 30
                pbcPctNumber(Index).CurrentX = 300 * (ilLoop) + 30
                pbcPctNumber(Index).CurrentY = -15
                pbcPctNumber(Index).Print smPctComp(Index - 3, ilLoop)
            Next ilLoop
        End If
    End If
End Sub
Private Sub plcCalendar_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub rbcShow_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcShow(Index).Value
    'End of coded added
    Dim ilLoop As Integer
    If Value Then
        For ilLoop = 0 To 5 Step 1
            pbcPctNumber(ilLoop).Cls
            pbcPctDraw(ilLoop).Cls
        Next ilLoop
        For ilLoop = 0 To 5 Step 1
            pbcPctNumber_Paint ilLoop
            pbcPctDraw_Paint ilLoop
        Next ilLoop
    End If
End Sub
Private Sub plcShow_Paint()
    plcShow.CurrentX = 0
    plcShow.CurrentY = 0
    plcShow.Print "Show"
End Sub
