VERSION 5.00
Begin VB.Form CCancel 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4860
   ClientLeft      =   600
   ClientTop       =   2085
   ClientWidth     =   9525
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
   ScaleHeight     =   4860
   ScaleWidth      =   9525
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
      Left            =   0
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4650
      Width           =   120
   End
   Begin VB.ComboBox cbcProgram 
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
      Height          =   300
      Left            =   3885
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.ListBox lbcPrg 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2445
      Sorted          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   2520
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   2520
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2820
      TabIndex        =   7
      Top             =   4485
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4230
      TabIndex        =   8
      Top             =   4485
      Width           =   1050
   End
   Begin VB.PictureBox plcInv 
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
      Height          =   4020
      Left            =   75
      ScaleHeight     =   3960
      ScaleWidth      =   9285
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   375
      Width           =   9345
      Begin VB.PictureBox pbcLbcLines 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2715
         Left            =   150
         ScaleHeight     =   2715
         ScaleWidth      =   9030
         TabIndex        =   12
         Top             =   585
         Width           =   9030
      End
      Begin VB.CheckBox ckcAll 
         Caption         =   "All Lines"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3465
         TabIndex        =   6
         Top             =   3705
         Width           =   1170
      End
      Begin VB.ListBox lbcLines 
         Appearance      =   0  'Flat
         Height          =   2760
         ItemData        =   "CCANCEL.frx":0000
         Left            =   135
         List            =   "CCANCEL.frx":0002
         MultiSelect     =   2  'Extended
         TabIndex        =   5
         Top             =   570
         Width           =   9030
      End
      Begin VB.ComboBox cbcTerminateDay 
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
         Height          =   300
         Left            =   2700
         TabIndex        =   3
         Top             =   150
         Width           =   780
      End
      Begin VB.ComboBox cbcTerminateDate 
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
         Height          =   300
         Left            =   3540
         TabIndex        =   4
         Top             =   150
         Width           =   2985
      End
      Begin VB.Label lacTerminate 
         Appearance      =   0  'Flat
         Caption         =   "Terminate Date"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1290
         TabIndex        =   2
         Top             =   180
         Width           =   1350
      End
   End
End
Attribute VB_Name = "CCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CCANCEL.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CCancel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract terminate screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imChgMode As Integer
Dim imBSMode As Integer
Dim imDateSelectedIndex As Integer
Dim imDaySelectedIndex As Integer
Dim imComboBoxIndex As Integer
Dim imSetAll As Integer 'True=Set list box; False= don't change list box
Dim imAllClicked As Integer  'True=All box clicked (don't call ckcAll within lbcLines)
Dim lmEarliestDate As Long  'Earliest allowed date to be entered
Dim lmLatestLLD As Long
Dim lmCntrStartDate As Long
Dim lmCntrEndDate As Long
Dim imIncludedCBS As Integer
'Dim imListField(1 To 6) As Integer
Dim imListField(0 To 6) As Integer
Dim imLBCtrls As Integer


Private Sub cbcTerminateDate_Change()
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilWeekDay As Integer
    Dim slDate As String
    Dim ilSelStart As Integer
    Dim ilDay As Integer
    Dim ilNoWeeks As Integer
    If imChgMode = False Then
        imChgMode = True
        ilSelStart = cbcTerminateDate.SelStart
        ilRet = gOptionLookAhead(cbcTerminateDate, imBSMode, slStr)
        If ilRet > 1 Then
            'Reset dates
            If gValidDate(slStr) Then
                If gDateValue(slStr) <= lmEarliestDate Then
                    Beep
                    slStr = Format$(lmEarliestDate + 1, "m/d/yy")
                    ilDay = imDaySelectedIndex
                    Do Until gWeekDayStr(slStr) = ilDay
                        slStr = gIncOneDay(slStr)
                    Loop
                End If
                ilNoWeeks = (lmCntrEndDate - gDateValue(slStr)) \ 7 + 1
                gDatePop slStr, ilNoWeeks, cbcTerminateDate
                slDate = Format$(slStr, "m/d/yy")
                If imIncludedCBS Then
                    cbcTerminateDate.AddItem "[Canceled before Start]", 0
                End If
                gFindMatch slDate, 0, cbcTerminateDate
                If gLastFound(cbcTerminateDate) >= 0 Then
                    cbcTerminateDate.ListIndex = gLastFound(cbcTerminateDate)
                Else
                    cbcTerminateDate.Text = slDate
                End If
                ilWeekDay = gWeekDayStr(slDate)
                cbcTerminateDay.ListIndex = ilWeekDay
                imDaySelectedIndex = ilWeekDay
                If ilSelStart < Len(cbcTerminateDate.Text) Then
                    cbcTerminateDate.SelStart = ilSelStart
                    cbcTerminateDate.SelLength = 14
                Else
                    cbcTerminateDate.SelStart = 0
                    cbcTerminateDate.SelLength = 14
                End If
            Else
                imDateSelectedIndex = cbcTerminateDate.ListIndex
                imChgMode = False
                Exit Sub
            End If
        ElseIf ilRet = 1 Then
            imDateSelectedIndex = cbcTerminateDate.ListIndex
            imChgMode = False
            Exit Sub
        Else
            If ilSelStart < Len(cbcTerminateDate.Text) Then
                cbcTerminateDate.SelStart = ilSelStart
                cbcTerminateDate.SelLength = 25
            Else
                cbcTerminateDate.SelStart = 0
                cbcTerminateDate.SelLength = 25
            End If
        End If
        imDateSelectedIndex = cbcTerminateDate.ListIndex
        imChgMode = False
    End If
End Sub
Private Sub cbcTerminateDate_Click()
    cbcTerminateDate_Change
End Sub
Private Sub cbcTerminateDate_GotFocus()
    If cbcTerminateDate.Text = "" Then
        If imIncludedCBS Then
            cbcTerminateDate.ListIndex = 0
        Else
            cbcTerminateDate.ListIndex = cbcTerminateDate.ListCount - 1
        End If
    End If
    gCtrlGotFocus cbcTerminateDate
End Sub
Private Sub cbcTerminateDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcTerminateDate_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcTerminateDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcTerminateDay_Change()

    Dim ilSelStart As Integer
    If imChgMode = False Then
        imChgMode = True
        ilSelStart = cbcTerminateDay.SelStart
        If cbcTerminateDay.Text <> "" Then
            gManLookAhead cbcTerminateDay, imBSMode, imComboBoxIndex
        End If
        imDaySelectedIndex = cbcTerminateDay.ListIndex

        mSetDates
        imDateSelectedIndex = -1
        If ilSelStart < Len(cbcTerminateDay.Text) Then
            cbcTerminateDay.SelStart = ilSelStart
            cbcTerminateDay.SelLength = 2
        Else
            cbcTerminateDay.SelStart = 0
            cbcTerminateDay.SelLength = 2
        End If
        imChgMode = False
    End If
End Sub
Private Sub cbcTerminateDay_Click()
    imComboBoxIndex = cbcTerminateDay.ListIndex
    cbcTerminateDay_Change
End Sub
Private Sub cbcTerminateDay_GotFocus()
    If cbcTerminateDay.Text = "" Then
        cbcTerminateDay.ListIndex = 6
    End If
    gCtrlGotFocus cbcTerminateDay
    imComboBoxIndex = cbcTerminateDay.ListIndex
End Sub
Private Sub cbcTerminateDay_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcTerminateDay_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcTerminateDay.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub ckcAll_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Dim ilClf As Integer

    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    Dim ilLoop As Integer
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        If ilValue Then
            For ilLoop = lbcLines.ListCount - 1 To 0 Step -1
                ilClf = lbcLines.ItemData(ilLoop)
                If mOKToSelect(ilClf) Then
                    lbcLines.Selected(ilLoop) = ilValue
                End If
            Next ilLoop
        Else
'            For ilLoop = lbcLines.ListCount - 1 To 0 Step -1
'                lbcLines.Selected(ilLoop) = ilValue
'            Next ilLoop
            If lbcLines.ListCount > 0 Then
                llRg = CLng(lbcLines.ListCount - 1) * &H10000 + 0
                llRet = SendMessageByNum(lbcLines.hWnd, LB_SELITEMRANGE, False, llRg)
            End If
        End If
        imAllClicked = False
        pbcLbcLines_Paint
    End If
End Sub
Private Sub ckcAll_GotFocus()
    gCtrlGotFocus ckcAll
End Sub
Private Sub cmcCancel_Click()
    igTerminateReturn = 0   'Cancel selected
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    gCtrlGotFocus cmcCancel
End Sub
Private Sub cmcDone_Click()
    Dim ilLoop As Integer
    Dim ilCff As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llCancelDate As Long
    Dim llMoCancelDate As Long
    Dim ilDay As Integer
    Dim ilIndex As Integer
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim ilVef As Integer
    Dim ilCgf As Integer
    Dim llAirDate As Long
    Dim slVefType As String
    Dim llDate As Long
    Dim slStr As String
    Dim llCffDate As Long
    Dim llRow As Long

    If imDateSelectedIndex < 0 Then
        Beep
        Exit Sub
    End If
    llCancelDate = 0
    If imIncludedCBS Then
        If imDateSelectedIndex > 0 Then
            slDate = cbcTerminateDate.Text
            llCancelDate = gDateValue(slDate)
        End If
    Else
        slDate = cbcTerminateDate.Text
        llCancelDate = gDateValue(slDate)
        'Check if cancel between todays date and the last log date
    End If
    If Contract.grdCPMBuy.Visible Then
        Contract.grdCPMBuy.Redraw = False
        gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbHourglass
        CCancel.MousePointer = vbHourglass
        For ilLoop = lbcLines.ListCount - 1 To 0 Step -1
            If lbcLines.Selected(ilLoop) Then
                llRow = lbcLines.ItemData(ilLoop)
                slStartDate = Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE)
                slEndDate = Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)
                If (slStartDate <> "CBS") And (slStartDate <> "Deleted") Then
                    If ((imDateSelectedIndex = 0) And (imIncludedCBS)) Or (llCancelDate < gDateValue(slStartDate)) Then  '[Cancel before Start]
                        imCPMChg = True
                        If Contract.grdCPMBuy.TextMatrix(llRow, CPMPCFCODE) = "" Then
                            Contract.grdCPMBuy.RemoveItem llRow
                        Else
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) = "CBS"
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE) = ""
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMIMPRESSIONS) = ""
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMCPMVALUE) = ""
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMCOST) = ""
                        End If
                    Else
                        If (llCancelDate >= gDateValue(slStartDate)) And (llCancelDate < gDateValue(slEndDate)) Then
                            Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE) = slDate
                            imCPMChg = True
                        End If
                    End If
                End If
            End If
        Next ilLoop
        mSetFGColorForRows
        Contract.grdCPMBuy.Redraw = False
        mComputeTotals
        CCancel.MousePointer = vbDefault
        Contract.grdCPMBuy.Redraw = True
        gSetMousePointer Contract.grdCPMBuy, Contract.grdCPMBuy, vbDefault
    Else
        mGetLLD
        If lmLatestLLD <> 0 Then
            If llCancelDate <= lmLatestLLD Then
                ilRet = MsgBox("Logs must be re-generated for days being canceled, Ok to proceed?", vbQuestion + vbYesNo, "Cancel")
                If ilRet = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        If llCancelDate <> 0 Then
            llMoCancelDate = llCancelDate
            Do While gWeekDayLong(llMoCancelDate) <> 0
                llMoCancelDate = llMoCancelDate - 1
            Loop
        Else
            llMoCancelDate = 0
        End If
        For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
            If lbcLines.Selected(ilLoop) Then
                ilClf = lbcLines.ItemData(ilLoop)
                If (tgClfCntr(ilClf).iFirstCff <> -1) And (Not tgClfCntr(ilClf).iCancel) Then
                    ilVef = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
                    If ilVef <> -1 Then
                        slVefType = tgMVef(ilVef).sType
                    Else
                        slVefType = ""
                    End If
                    If slVefType = "G" Then
                        ilCgf = tgClfCntr(ilClf).iFirstCgf
                        Do While ilCgf <> -1
                            gUnpackDateLong tgCgfCntr(ilCgf).CgfRec.iAirDate(0), tgCgfCntr(ilCgf).CgfRec.iAirDate(1), llAirDate
                            If llCancelDate < llAirDate Then
                                gUnpackDate tgCgfCntr(ilCgf).CgfRec.iAirDate(0), tgCgfCntr(ilCgf).CgfRec.iAirDate(1), slStr
                                llDate = gDateValue(gObtainPrevMonday(slStr))
                                ilCff = tgClfCntr(ilClf).iFirstCff
                                Do While ilCff <> -1
                                    'Date
                                    gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), llCffDate
                                    If llDate = llCffDate Then
                                        tgCffCntr(ilCff).CffRec.iSpotsWk = tgCffCntr(ilCff).CffRec.iSpotsWk - tgCgfCntr(ilCgf).CgfRec.iNoSpots
                                        Exit Do
                                    End If
                                    ilCff = tgCffCntr(ilCff).iNextCff
                                Loop
                                tgCgfCntr(ilCgf).CgfRec.iNoSpots = 0
                            End If
                            ilCgf = tgCgfCntr(ilCgf).iNextCgf
                        Loop
    
                    End If
                    ilCff = tgClfCntr(ilClf).iFirstCff
                    gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate
                    ilCff = tgClfCntr(ilClf).iFirstCff
                    If ((imDateSelectedIndex = 0) And (imIncludedCBS)) Or (llCancelDate < gDateValue(slStartDate)) Then  '[Cancel before Start]
                        tgClfCntr(ilClf).iCancel = True
                        slEndDate = gDecOneDay(slStartDate)
                        gPackDate slEndDate, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)
                        If tgCffCntr(ilCff).iStatus = -1 Then
                            tgCffCntr(ilCff).iStatus = 0    'Reset to new
                        ElseIf tgCffCntr(ilCff).iStatus = 2 Then
                            tgCffCntr(ilCff).iStatus = 1    'Reset old
                        End If
                        Do While tgCffCntr(ilCff).iNextCff <> -1
                            ilCff = tgCffCntr(ilCff).iNextCff
                            If tgCffCntr(ilCff).iStatus = 0 Then
                                tgCffCntr(ilCff).iStatus = -1   'New to not used
                            ElseIf tgCffCntr(ilCff).iStatus = 1 Then
                                tgCffCntr(ilCff).iStatus = 2    'Old to delete
                            End If
                        Loop
                    Else
                        'Cancel future weeks
                        ilCff = tgClfCntr(ilClf).iFirstCff
                        Do While ilCff <> -1
                            gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate
                            gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), slEndDate
                            If llCancelDate < gDateValue(slStartDate) Then
                                If tgCffCntr(ilCff).iStatus = 0 Then
                                    tgCffCntr(ilCff).iStatus = -1   'New to not used
                                ElseIf tgCffCntr(ilCff).iStatus = 1 Then
                                    tgCffCntr(ilCff).iStatus = 2    'Old to delete
                                End If
                            End If
                            ilCff = tgCffCntr(ilCff).iNextCff
                        Loop
                        ilCff = tgClfCntr(ilClf).iFirstCff
                        Do While ilCff <> -1
                            gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate
                            gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), slEndDate
                            If (llCancelDate >= gDateValue(slStartDate)) And (llCancelDate <= gDateValue(slEndDate)) Then
                                If gDateValue(slStartDate) < llMoCancelDate Then
                                    mSplitFlight ilCff, llMoCancelDate, llCancelDate
                                    Exit Do
                                Else
                                    slEndDate = Format$(llCancelDate, "m/d/yy")
                                    gPackDate slEndDate, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)
                                    ilDay = gWeekDayLong(llCancelDate)
                                    For ilIndex = ilDay + 1 To 6 Step 1
                                        tgCffCntr(ilCff).CffRec.iDay(ilIndex) = 0
                                        tgCffCntr(ilCff).CffRec.sXDay(ilIndex) = " "
                                    Next ilIndex
                                    Exit Do
                                End If
                            End If
                            ilCff = tgCffCntr(ilCff).iNextCff
                        Loop
                    End If
                End If
            End If
        Next ilLoop
        For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
            If lbcLines.Selected(ilLoop) Then
                ilClf = lbcLines.ItemData(ilLoop)
                igLnSPChg = True
                tgClfCntr(ilClf).iSPChg = True
                tgClfCntr(ilClf).iResearch = False
                tgClfCntr(ilClf).lAvgAud = 0
                tgClfCntr(ilClf).iAvgRating = 0
                If tgClfCntr(ilClf).iFirstCff <> -1 Then
                    ilCff = tgClfCntr(ilClf).iFirstCff
                    gUnpackDateLong tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), tgCffCntr(ilCff).lStartDate
                    gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), tgCffCntr(ilCff).lEndDate
                    Do While ilCff <> -1
                        If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                            gUnpackDateLong tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), tgCffCntr(ilCff).lEndDate
                        End If
                        ilCff = tgCffCntr(ilCff).iNextCff
                    Loop
                End If
            End If
        Next ilLoop
    End If
    igTerminateReturn = 1   'Done selected
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus cmcDone
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If cbcTerminateDate.Text = "" Then
        If imIncludedCBS Then
            cbcTerminateDate.ListIndex = 0
        Else
            cbcTerminateDate.ListIndex = cbcTerminateDate.ListCount - 1
        End If
    End If
'    If cbcTerminateDate.Text = "" Then 'Setting combos in mInit required along within here
'        cbcTerminateDate.ListIndex = 0
        cbcTerminateDay.ListIndex = 6
'    End If
    Me.KeyPreview = True
    CCancel.Refresh
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
        cmcCancel_Click
    End If
End Sub
Private Sub lbcLines_Click()
    Dim slLineType As String
    Dim ilLoop As Integer
    Dim ilClf As Integer
    Dim ilLp As Integer
    Dim ilIndex As Integer
    Dim ilLine As Integer
    Dim ilClfLp As Integer
    Dim llRow As Long
    Dim ilTest As Integer

    If Not imAllClicked Then
        imSetAll = False
        ckcAll.Value = vbUnchecked
        imSetAll = True
        For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
            If lbcLines.Selected(ilLoop) Then
                ilClf = lbcLines.ItemData(ilLoop)
                If Not mOKToSelect(ilClf) Then
                    Beep
                    lbcLines.Selected(ilLoop) = False
                End If
            End If
        Next ilLoop
        'If package, then select all lines within package
        If Contract.grdCPMBuy.Visible Then
            For ilLine = 0 To lbcLines.ListCount - 1 Step 1
                If lbcLines.Selected(ilLine) Then
                    ilClf = lbcLines.ItemData(ilLine)
                    If Contract.grdCPMBuy.TextMatrix(ilClf, CPMLNTYPE) = "P" Then
                        For ilTest = ilLine + 1 To lbcLines.ListCount - 1 Step 1
                            llRow = lbcLines.ItemData(ilTest)
                            If llRow <> ilClf Then
                                If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
                                    If Contract.grdCPMBuy.TextMatrix(llRow, CPMPACKAGEID) = Contract.grdCPMBuy.TextMatrix(ilClf, CPMID) Then
                                        If Not lbcLines.Selected(ilTest) Then
                                            imAllClicked = True
                                            lbcLines.Selected(ilTest) = True
                                            imAllClicked = False
                                        End If
                                    End If
                                End If
                            End If
                        Next ilTest
                    End If
                End If
            Next ilLine
        Else
            For ilLine = 0 To lbcLines.ListCount - 1 Step 1
                If lbcLines.Selected(ilLine) Then
                    ilClf = lbcLines.ItemData(ilLine)
                    slLineType = mGetLineType(ilClf + 1)
                    If (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
                        'For ilClf = 0 To lbcLines.ListCount - 1 Step 1
                            'If tgClfCntr(ilClf).ClfRec.iPkLineNo = tgClfCntr(ilLoop).ClfRec.iLine Then
                            '    lbcLines.Selected(ilClf) = True
                            'End If
                        'Next ilClf
                        For ilLoop = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
                            If tgClfCntr(ilClf).ClfRec.iLine = tgPackLn(ilLoop).iLineNo Then
                                ilIndex = tgPackLn(ilLoop).iIndex
                                Do While ilIndex >= 0
                                    For ilLp = 0 To lbcLines.ListCount - 1 Step 1
                                        ilClfLp = lbcLines.ItemData(ilLp)
                                        If tgPackLn(ilIndex).iLineNo = tgClfCntr(ilClfLp).ClfRec.iLine Then
                                            If Not lbcLines.Selected(ilLp) Then
                                                imAllClicked = True
                                                lbcLines.Selected(ilLp) = True
                                                imAllClicked = False
                                            End If
                                            Exit For
                                        End If
                                    Next ilLp
                                    ilIndex = tgPackLn(ilIndex).iIndex
                                Loop
                                Exit For
                            End If
                        Next ilLoop
                    End If
                End If
            Next ilLine
        End If
    End If
    pbcLbcLines_Paint
End Sub
Private Sub lbcLines_GotFocus()
    gCtrlGotFocus lbcLines
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
    Dim ilTest As Integer
    Dim slStr As String
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilCff As Integer
    Dim slDate As String
    Dim slLineType As String
    Dim ilVef As Integer
    Dim slVefType As String
    Dim llRow As Long

    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imTerminate = False
    imLBCtrls = 1
    imListField(1) = 15
    imListField(2) = 6 * igAlignCharWidth
    'imListField(3) = 30 * igAlignCharWidth
    'imListField(4) = 55 * igAlignCharWidth
    'imListField(5) = 60 * igAlignCharWidth
    'imListField(6) = 100 * igAlignCharWidth
    imListField(3) = 50 * igAlignCharWidth
    imListField(4) = 75 * igAlignCharWidth
    imListField(5) = 80 * igAlignCharWidth
    imListField(6) = 110 * igAlignCharWidth
    imSetAll = True
    imAllClicked = False
    imDateSelectedIndex = -1
    imDaySelectedIndex = -1
    gDayPop cbcTerminateDay
    slDate = ""
    lmEarliestDate = gDateValue(gNow())
    gUnpackDate tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1), slStartDate
    lmCntrStartDate = gDateValue(slStartDate)
    gUnpackDate tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1), slEndDate
    lmCntrEndDate = gDateValue(slEndDate)

    imIncludedCBS = False
    If lmCntrStartDate > lmEarliestDate Then
        imIncludedCBS = True
    End If
    cbcTerminateDay.ListIndex = 6
    mSetDates

    If Contract.grdCPMBuy.Visible Then
        For llRow = Contract.grdCPMBuy.FixedRows To Contract.grdCPMBuy.Rows - 1 Step 1
            If (Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE) <> "") And (Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART) <> "") Then
                slStr = Contract.grdCPMBuy.TextMatrix(llRow, CPMID)
                If Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) = "CBS" Then
                    slStr = slStr & "| |" & "Canceled before Start"
                ElseIf Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) = "Deleted" Then
                    slStr = slStr & "| |" & "Deleted"
                Else
                    slStr = slStr & "|" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMVEHICLE))
                    slStr = slStr & "|" & Trim$(Contract.grdCPMBuy.TextMatrix(llRow, CPMDAYPART))
                    slStr = slStr & "|" & ""    'Length
                    slStr = slStr & "|" & Contract.grdCPMBuy.TextMatrix(llRow, CPMSTARTDATE) & "-" & Contract.grdCPMBuy.TextMatrix(llRow, CPMENDDATE)
                End If
                lbcLines.AddItem slStr
                lbcLines.ItemData(lbcLines.NewIndex) = llRow
            End If
        Next llRow
    Else
        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            ilVef = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
            If ilVef <> -1 Then
                slVefType = tgMVef(ilVef).sType
            Else
                slVefType = ""
            End If
            'Show Game line but disallow lines from being selected.
            'If slVefType <> "G" Then
                slStr = Trim$(str$(tgClfCntr(ilClf).ClfRec.iLine))
                If tgClfCntr(ilClf).iCancel Then
                    slStr = slStr & "| |" & "Canceled before Start"
                Else
                    slRecCode = Trim$(str$(tgClfCntr(ilClf).ClfRec.iVefCode))
                    slLineType = mGetLineType(ilClf + 1)
                    If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                        For ilTest = 0 To UBound(tmVehicleCode) - 1 Step 1  'Contract!lbcVehicle.ListCount - 1 Step 1
                            slNameCode = tmVehicleCode(ilTest).sKey    'Contract!lbcVehicle.List(ilTest)
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            On Error GoTo mInitErr
                            gCPErrorMsg ilRet, "mInit (gParseItem field 2)", CCancel
                            On Error GoTo 0
                            If slRecCode = slCode Then
                                ilRet = gParseItem(slNameCode, 1, "\", slName)
                                On Error GoTo mInitErr
                                gCPErrorMsg ilRet, "mInit (gParseItem field 2)", CCancel
                                On Error GoTo 0
                                ilRet = gParseItem(slName, 3, "|", slName)
                                On Error GoTo mInitErr
                                gCPErrorMsg ilRet, "mInit (gParseItem field 2)", CCancel
                                On Error GoTo 0
                                slStr = slStr & "|" & slName
                                Exit For
                            End If
                        Next ilTest
                    Else
                        If (tgClfCntr(ilClf).ClfRec.iVefCode = 0) Then
                            slStr = slStr & "|" & Trim$(smLnSave(1, ilClf + 1))
                        Else
                            'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            '    If tgClfCntr(ilClf).ClfRec.iVefCode = tgMVef(ilVef).iCode Then
                                ilVef = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
                                If ilVef <> -1 Then
                                    slStr = slStr & "|" & Trim$(tgMVef(ilVef).sName)
                            '        Exit For
                                End If
                            'Next ilVef
                        End If
                    End If
                    'mPrgPop (ilClf)
                    'If imTerminate Then
                    '    Exit Sub
                    'End If
                    'slLineType = mGetLineType(ilClf + 1)
                    'If (slLineType <> "O") And (slLineType <> "A") Then
                    '    For ilLoop = 0 To cbcProgram.ListCount - 1 Step 1
                    '        If lbcPrg(igTabMapIndex).ItemData(ilLoop) = tgClfCntr(ilClf).ClfRec.iRdfCode Then
                    '            slStr = slStr & "|" & cbcProgram.List(ilLoop)
                    '            Exit For
                    '        End If
                    '    Next ilLoop
                    'Else
                    '    For ilLoop = LBound(tgPackDP) To UBound(tgPackDP) - 1 Step 1
                    '        If tgPackDP(ilLoop).iCode = tgClfCntr(ilClf).ClfRec.iRdfCode Then
                    '            slStr = slStr & "|" & Trim$(tgPackDP(ilLoop).sName) 'cbcProgram.List(ilLoop)
                    '            Exit For
                    '        End If
                    '    Next ilLoop
                    'End If
                    slStr = slStr & "|" & mMakeDPName(ilClf + 1)  '
                    slStr = slStr & "|" & Trim$(str$(tgClfCntr(ilClf).ClfRec.iLen))
                    If tgClfCntr(ilClf).iFirstCff <> -1 Then
                        slStartDate = ""
                        slEndDate = ""
                        ilCff = tgClfCntr(ilClf).iFirstCff
                        Do While ilCff <> -1
                            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                                If slStartDate = "" Then
                                    gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate
                                End If
                                gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), slEndDate
                            End If
                            ilCff = tgCffCntr(ilCff).iNextCff
                        Loop
                        slStr = slStr & "|" & slStartDate & "-" & slEndDate
                    End If
                End If
                'lbcLines.AddItem gAlignStringByPixel(slStr, "|", imListField(), imListFieldChar())
                lbcLines.AddItem slStr
                lbcLines.ItemData(lbcLines.NewIndex) = ilClf
            'End If
        Next ilClf
    End If
    CCancel.height = cmcDone.Top + 5 * cmcDone.height / 3
    gCenterModalForm CCancel
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOKToSelect                     *
'*                                                     *
'*             Created:8/31/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if schedule line can *
'*                      be selected for termination    *
'*                                                     *
'*******************************************************
Private Function mOKToSelect(ilClf As Integer) As Integer
    Dim slTest As String
    Dim ilCff As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slDate As String
    Dim slLnLLD As String
    Dim llAllowedChgWkStartDate As Long
    Dim slVefType As String
    Dim ilVef As Integer

    mOKToSelect = False
    If Contract.grdCPMBuy.Visible Then
        slStartDate = Contract.grdCPMBuy.TextMatrix(ilClf, CPMSTARTDATE)
        slEndDate = Contract.grdCPMBuy.TextMatrix(ilClf, CPMENDDATE)
        'If gDateValue(smCPMLLD) < gDateValue(slEndDate) Then
        If gDateValue(mCPMAllowedDate(CLng(ilClf))) < gDateValue(slEndDate) Then
            If imIncludedCBS Then
                If imDateSelectedIndex > 0 Then
                    slDate = cbcTerminateDate.List(imDateSelectedIndex)
                    If gDateValue(slDate) > gDateValue(smCPMLLD) Then
                        mOKToSelect = True
                    Else
                        If gDateValue(smCPMLLD) < gDateValue(slStartDate) Then
                            mOKToSelect = True
                        End If
                    End If
                Else
                    If gDateValue(smCPMLLD) < gDateValue(slStartDate) Then
                        mOKToSelect = True
                    End If
                End If
            Else
                If imDateSelectedIndex >= 0 Then
                    slDate = cbcTerminateDate.List(imDateSelectedIndex)
                    If gDateValue(slDate) > gDateValue(smCPMLLD) Then
                        mOKToSelect = True
                    Else
                        If gDateValue(smCPMLLD) < gDateValue(slStartDate) Then
                            mOKToSelect = True
                        End If
                    End If
                End If
            End If
        End If
    Else
        slVefType = ""
        ilVef = gBinarySearchVef(tgClfCntr(ilClf).ClfRec.iVefCode)
        If ilVef <> -1 Then
            slVefType = tgMVef(ilVef).sType
        End If
        '1/18/08:  Allow games to be canceled
        'If (Not tgClfCntr(ilClf).iCancel) And (slVefType <> "G") Then
        If (Not tgClfCntr(ilClf).iCancel) Then
            'slTest = gFindVehicleLatestDate(CCancel, tgClfCntr(ilClf).ClfRec.iVefCode)  'Find First allowed scheduling date (Last log date or Now)
            llAllowedChgWkStartDate = mLnDate(ilClf + 1, slLnLLD, False)
            If llAllowedChgWkStartDate <> 0 Then
                slTest = Format$(llAllowedChgWkStartDate, "m/d/yy")
            Else
                slTest = "1/1/80"
            End If
            slStartDate = ""
            slEndDate = ""
            If tgClfCntr(ilClf).iFirstCff <> -1 Then
                ilCff = tgClfCntr(ilClf).iFirstCff
                Do
                    If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                        If slStartDate = "" Then
                            gUnpackDate tgCffCntr(ilCff).CffRec.iStartDate(0), tgCffCntr(ilCff).CffRec.iStartDate(1), slStartDate
                        End If
                        gUnpackDate tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1), slEndDate
                    End If
                    ilCff = tgCffCntr(ilCff).iNextCff
                Loop While ilCff <> -1
            End If
            If (slStartDate <> "") And (slEndDate <> "") Then
                If gDateValue(slTest) < gDateValue(slEndDate) Then
                    If imIncludedCBS Then
                        If imDateSelectedIndex > 0 Then
                            slDate = cbcTerminateDate.List(imDateSelectedIndex)
                            If gDateValue(slDate) > gDateValue(slTest) Then
                                mOKToSelect = True
                            Else
                                If gDateValue(slTest) < gDateValue(slStartDate) Then
                                    mOKToSelect = True
                                End If
                            End If
                        Else
                            If gDateValue(slTest) < gDateValue(slStartDate) Then
                                mOKToSelect = True
                            End If
                        End If
                    Else
                        If imDateSelectedIndex >= 0 Then
                            slDate = cbcTerminateDate.List(imDateSelectedIndex)
                            If gDateValue(slDate) > gDateValue(slTest) Then
                                mOKToSelect = True
                            Else
                                If gDateValue(slTest) < gDateValue(slStartDate) Then
                                    mOKToSelect = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function
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
    Unload CCancel
    Set CCancel = Nothing   'Remove data segment
End Sub

Private Sub lbcLines_Scroll()
    pbcLbcLines_Paint
End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcLbcLines_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLinesEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    Dim slFields(0 To 4) As String
    Dim llFgColor As Long
    Dim slLineType As String
    Dim ilClf As Integer
    Dim ilFieldIndex As Integer
    
    ilLinesEnd = lbcLines.TopIndex + lbcLines.height \ fgListHtArial825
    If ilLinesEnd > lbcLines.ListCount Then
        ilLinesEnd = lbcLines.ListCount
    End If
    If lbcLines.ListCount <= lbcLines.height \ fgListHtArial825 Then
        llWidth = lbcLines.Width - 30
    Else
        llWidth = lbcLines.Width - igScrollBarWidth - 30
    End If
    pbcLbcLines.Width = llWidth
    pbcLbcLines.Cls
    llFgColor = pbcLbcLines.ForeColor
    For ilLoop = lbcLines.TopIndex To ilLinesEnd - 1 Step 1
        pbcLbcLines.ForeColor = llFgColor
        If lbcLines.MultiSelect = 0 Then
            If lbcLines.ListIndex = ilLoop Then
                gPaintArea pbcLbcLines, CSng(0), CSng((ilLoop - lbcLines.TopIndex) * fgListHtArial825), CSng(pbcLbcLines.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcLines.ForeColor = vbWhite
            End If
        Else
            If lbcLines.Selected(ilLoop) Then
                gPaintArea pbcLbcLines, CSng(0), CSng((ilLoop - lbcLines.TopIndex) * fgListHtArial825), CSng(pbcLbcLines.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcLines.ForeColor = vbWhite
            End If
        End If
        ilClf = lbcLines.ItemData(ilLoop)
        If Contract.grdCPMBuy.Visible Then
            If Contract.grdCPMBuy.TextMatrix(ilClf, CPMLNTYPE) = "H" Then
                pbcLbcLines.FontStrikethru = True
                slLineType = "H"
            End If
        Else
            slLineType = mGetLineType(ilClf + 1)
            If slLineType = "H" Then
                pbcLbcLines.FontStrikethru = True
            End If
        End If
        ilFieldIndex = 0
        slStr = lbcLines.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        For ilField = imLBCtrls To UBound(imListField) - 1 Step 1
            pbcLbcLines.CurrentX = imListField(ilField)
            pbcLbcLines.CurrentY = (ilLoop - lbcLines.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilFieldIndex)
            ilFieldIndex = ilFieldIndex + 1
            gAdjShowLen pbcLbcLines, slStr, imListField(ilField + 1) - imListField(ilField)
            If (slLineType = "H") And ((ilField = 1) Or (ilField = 2)) Then
                If ilField = 1 Then
                    pbcLbcLines.FontStrikethru = False
                Else
                    pbcLbcLines.CurrentX = pbcLbcLines.CurrentX + pbcLbcLines.TextWidth("  ")
                End If
                pbcLbcLines.Print slStr
                pbcLbcLines.FontStrikethru = True
            Else
                pbcLbcLines.Print slStr
            End If
        Next ilField
        pbcLbcLines.FontStrikethru = False
        pbcLbcLines.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub plcInv_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    If Contract.grdCPMBuy.Visible Then
        plcScreen.Print "Terminate Contract CPM"
    Else
        plcScreen.Print "Terminate Contract Lines"
    End If
End Sub

Private Sub mSetDates()
    Dim ilDay As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim ilNoWeeks As Integer

    ilDay = cbcTerminateDay.ListIndex

    If imIncludedCBS Then
        llDate = lmCntrStartDate
    Else
        llDate = lmEarliestDate
    End If
    Do While gWeekDayLong(llDate) <> ilDay
        llDate = llDate + 1
    Loop
    ilNoWeeks = (lmCntrEndDate - llDate) \ 7 + 1
    slDate = Format$(llDate, "m/d/yy")
    gDatePop slDate, ilNoWeeks, cbcTerminateDate
    If imIncludedCBS Then
        cbcTerminateDate.AddItem "[Canceled before Start]", 0
    End If
End Sub

Private Sub mGetLLD()
    Dim ilLoop As Integer
    Dim slLineType As String
    Dim llDate As Long
    Dim ilVpfIndex As Integer
    Dim ilClf As Integer

    lmLatestLLD = 0
    If Contract.grdCPMBuy.Visible Then
        lmLatestLLD = gDateValue(smCPMLLD)
    Else
        For ilLoop = 0 To lbcLines.ListCount - 1 Step 1
            If lbcLines.Selected(ilLoop) Then
                ilClf = lbcLines.ItemData(ilLoop)
                If Not tgClfCntr(ilClf).iCancel Then
                    slLineType = mGetLineType(ilClf + 1)
                    If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                        ilVpfIndex = gVpfFind(CCancel, tgClfCntr(ilClf).ClfRec.iVefCode)
                        If ilVpfIndex >= 0 Then
                            gUnpackDateLong tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), llDate
                            If (llDate > 0) And (llDate >= lmEarliestDate) Then
                                If llDate > lmLatestLLD Then
                                    lmLatestLLD = llDate
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next ilLoop
    End If
End Sub

Private Sub mSplitFlight(ilCff As Integer, llStartDate As Long, llEndDate As Long)
    Dim ilCffIndex As Integer
    Dim ilIndex As Integer
    Dim slDate As String
    Dim ilDay As Integer

    'Split flight- The links are retained because of the LSet
    ilCffIndex = UBound(tgCffCntr)
    tgCffCntr(ilCffIndex) = tgCffCntr(ilCff)
    tgCffCntr(ilCffIndex).iStatus = 0
    tgCffCntr(ilCffIndex).lRecPos = 0
    tgCffCntr(ilCff).iNextCff = ilCffIndex
    ReDim Preserve tgCffCntr(0 To UBound(tgCffCntr) + 1) As CFFLIST
    tgCffCntr(UBound(tgCffCntr)).iStatus = -1 'Not Used
    tgCffCntr(UBound(tgCffCntr)).lRecPos = 0
    tgCffCntr(UBound(tgCffCntr)).iNextCff = -1
    slDate = Format$(llStartDate - 1, "m/d/yy")
    gPackDate slDate, tgCffCntr(ilCff).CffRec.iEndDate(0), tgCffCntr(ilCff).CffRec.iEndDate(1)   'Week Start date
    tgCffCntr(ilCff).lEndDate = gDateValue(slDate)
    slDate = Format$(llStartDate, "m/d/yy")
    gPackDate slDate, tgCffCntr(ilCffIndex).CffRec.iStartDate(0), tgCffCntr(ilCffIndex).CffRec.iStartDate(1)   'Week Start date
    tgCffCntr(ilCffIndex).lStartDate = gDateValue(slDate)
    slDate = Format$(llEndDate, "m/d/yy")
    gPackDate slDate, tgCffCntr(ilCffIndex).CffRec.iEndDate(0), tgCffCntr(ilCffIndex).CffRec.iEndDate(1)   'Week Start date
    tgCffCntr(ilCffIndex).lEndDate = gDateValue(slDate)
    ilDay = gWeekDayLong(llEndDate)
    For ilIndex = ilDay + 1 To 6 Step 1
        tgCffCntr(ilCffIndex).CffRec.iDay(ilIndex) = 0
        tgCffCntr(ilCffIndex).CffRec.sXDay(ilIndex) = " "
    Next ilIndex
End Sub



