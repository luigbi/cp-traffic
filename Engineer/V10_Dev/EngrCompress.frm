VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form EngrCompress 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5085
   ClientLeft      =   2970
   ClientTop       =   1440
   ClientWidth     =   8115
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5085
   ScaleWidth      =   8115
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   285
      Left            =   4575
      TabIndex        =   10
      Top             =   4665
      Width           =   945
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   4260
      Visible         =   0   'False
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmcProcess 
      Appearance      =   0  'Flat
      Caption         =   "&Process"
      Height          =   285
      Left            =   2580
      TabIndex        =   9
      Top             =   4665
      Width           =   945
   End
   Begin VB.PictureBox plcCompress 
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
      Height          =   3495
      Left            =   180
      ScaleHeight     =   3435
      ScaleWidth      =   7635
      TabIndex        =   1
      Top             =   270
      Width           =   7695
      Begin VB.CheckBox ckcList 
         Caption         =   "Remove History from Audio, Bus, Event Type, Follows, etc........"
         Height          =   345
         Left            =   120
         TabIndex        =   21
         Top             =   1575
         Value           =   1  'Checked
         Width           =   5505
      End
      Begin VB.CheckBox ckcT2 
         Caption         =   "Assign Title 2 to Libraries/Templates............................................."
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   2295
         Value           =   1  'Checked
         Width           =   5505
      End
      Begin VB.CheckBox ckcSetBusNames 
         Caption         =   "Set Bus Names into Libraries............................................................"
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Value           =   1  'Checked
         Width           =   5475
      End
      Begin VB.CheckBox ckcUnusedComments 
         Caption         =   "Remove Unused Comments............................................................."
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   2655
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.CheckBox ckcPurgeOld 
         Caption         =   "Remove Libraries/Templates/Schedules prior to........................."
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   495
         Value           =   1  'Checked
         Width           =   5475
      End
      Begin VB.CheckBox ckcComments 
         Caption         =   "Remove Duplicate Title 1 Comments in Libraries/Templates...."
         Height          =   345
         Left            =   405
         TabIndex        =   5
         Top             =   3300
         Visible         =   0   'False
         Width           =   5505
      End
      Begin VB.CheckBox ckcTemplates 
         Caption         =   "Remove Templates History................................................................"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   1215
         Value           =   1  'Checked
         Width           =   5505
      End
      Begin VB.CheckBox ckcLibraries 
         Caption         =   "Remove Libraries History..................................................................."
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   855
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.Label lacList 
         Height          =   240
         Left            =   5655
         TabIndex        =   22
         Top             =   1575
         Width           =   1740
      End
      Begin VB.Label lacT2Count 
         Height          =   240
         Left            =   5655
         TabIndex        =   18
         Top             =   2295
         Width           =   1740
      End
      Begin VB.Label lacBusNamesCount 
         Height          =   240
         Left            =   5655
         TabIndex        =   17
         Top             =   1905
         Width           =   1740
      End
      Begin VB.Label lacUnusedComments 
         Height          =   240
         Left            =   5655
         TabIndex        =   19
         Top             =   2640
         Width           =   1740
      End
      Begin VB.Label lacRemoveOld 
         Height          =   240
         Left            =   5655
         TabIndex        =   13
         Top             =   480
         Width           =   1740
      End
      Begin VB.Label lacCount 
         Alignment       =   2  'Center
         Caption         =   "Count"
         Height          =   165
         Left            =   5685
         TabIndex        =   12
         Top             =   60
         Width           =   1740
      End
      Begin VB.Label lacCommCount 
         Height          =   240
         Left            =   5940
         TabIndex        =   16
         Top             =   3300
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label lacTempCount 
         Height          =   240
         Left            =   5655
         TabIndex        =   15
         Top             =   1215
         Width           =   1740
      End
      Begin VB.Label lacLibCount 
         Height          =   240
         Left            =   5655
         TabIndex        =   14
         Top             =   825
         Width           =   1740
      End
   End
   Begin VB.Label lacTask 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   195
      TabIndex        =   20
      Top             =   3960
      Width           =   7650
   End
   Begin VB.Label plcScreen 
      Caption         =   "File Clean-up"
      Height          =   240
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2070
   End
End
Attribute VB_Name = "EngrCompress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EngrCompress.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the password screen code
Option Explicit
Option Compare Text
Private tmDHE() As DHE
Private tmT2CTE() As CTE
Private tmT2CTE_Name() As DEECTE
Private tmDee() As DEE
Private smDateTimeStamp As String
Private lmDeleteCode() As Long
Private lmTotalCount As Long
Private lmCurrentCount As Long
Private lmPercent As Long
Private imCompress As Integer
Private imTerminate As Integer
Private smPurgeDate As String
Private tmDEECTE() As DEECTE
Private tmCurrDBE() As DBE
Private smCurrDBE As String
Private smNowDate As String
Private smNowTime As String

Private hmCTE As Integer
Private tmCTE As CTE

Private Sub cmcCancel_Click()
    If imCompress Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub

Private Sub cmcProcess_Click()
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    imCompress = True
    cmcCancel.Caption = "&Cancel"
    lmTotalCount = 0
    lmCurrentCount = 0
    lacTask.Caption = ""
    If ckcPurgeOld.Value = vbChecked Then
        ilRet = mPurgeOld()
    End If
    If imTerminate Then
        lacTask.Caption = ""
        imCompress = False
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    plcGauge.Value = 0
    plcGauge.Visible = True
    If ckcLibraries.Value = vbChecked Then
        lmTotalCount = lmTotalCount + gGetCount("SELECT count(dheCode) FROM DHE_Day_Header_Info WHERE dheCurrent = 'N' And dheType = 'L'", "EngrCompress- GetCount")
    End If
    If imTerminate Then
        lacTask.Caption = ""
        imCompress = False
        Screen.MousePointer = vbDefault
        plcGauge.Visible = False
        Exit Sub
    End If
    If ckcTemplates.Value = vbChecked Then
        lmTotalCount = lmTotalCount + gGetCount("SELECT count(dheCode) FROM DHE_Day_Header_Info WHERE dheCurrent = 'N' And dheType = 'T'", "EngrCompress- GetCount")
    End If
    If imTerminate Then
        lacTask.Caption = ""
        imCompress = False
        Screen.MousePointer = vbDefault
        plcGauge.Visible = False
        Exit Sub
    End If
    If ckcList.Value = vbChecked Then
        lmTotalCount = lmTotalCount + 17
    End If
    'If ckcComments.Value = vbChecked Then
    '    lmTotalCount = lmTotalCount + gGetCount("SELECT count(dheCode) FROM DHE_Day_Header_Info WHERE dheCurrent = 'Y'", "EngrCompress- GetCount")
    'End If
    'If imTerminate Then
    '    lacTask.Caption = ""
    '    imCompress = False
    '    Screen.MousePointer = vbDefault
    '    plcGauge.Visible = False
    '    Exit Sub
    'End If
    If ckcSetBusNames.Value = vbChecked Then
        lmTotalCount = lmTotalCount + gGetCount("SELECT count(dheCode) FROM DHE_Day_Header_Info WHERE dheCurrent = 'Y' And dheType = 'L'", "EngrCompress- GetCount")
    End If
    If imTerminate Then
        lacTask.Caption = ""
        imCompress = False
        Screen.MousePointer = vbDefault
        plcGauge.Visible = False
        Exit Sub
    End If
    If ckcT2.Value = vbChecked Then
        lmTotalCount = lmTotalCount + gGetCount("SELECT count(dheCode) FROM DHE_Day_Header_Info WHERE dheCurrent = 'Y'", "EngrCompress- GetCount")
    End If
    If imTerminate Then
        lacTask.Caption = ""
        imCompress = False
        Screen.MousePointer = vbDefault
        plcGauge.Visible = False
        Exit Sub
    End If
    If ckcLibraries.Value = vbChecked Then
        lacTask.Caption = "Removing Library History"
        DoEvents
        ilRet = mCompressLibraries()
    End If
    If imTerminate Then
        lacTask.Caption = ""
        imCompress = False
        Screen.MousePointer = vbDefault
        plcGauge.Visible = False
        Exit Sub
    End If
    If ckcTemplates.Value = vbChecked Then
        lacTask.Caption = "Removing Template History"
        DoEvents
        ilRet = mCompressTemplates()
    End If
    If imTerminate Then
        lacTask.Caption = ""
        imCompress = False
        Screen.MousePointer = vbDefault
        plcGauge.Visible = False
        Exit Sub
    End If
    If ckcList.Value = vbChecked Then
        lacTask.Caption = "Removing List History"
        DoEvents
        ilRet = mCompressList()
    End If
    If imTerminate Then
        lacTask.Caption = ""
        imCompress = False
        Screen.MousePointer = vbDefault
        plcGauge.Visible = False
        Exit Sub
    End If
    'If ckcComments.Value = vbChecked Then
    '    lacTask.Caption = "Removing Duplicate Comments"
    '    DoEvents
    '    ilRet = mCompressComments()
    'End If
    If imTerminate Then
        lacTask.Caption = ""
        imCompress = False
        Screen.MousePointer = vbDefault
        plcGauge.Visible = False
        Exit Sub
    End If
    If ckcSetBusNames.Value = vbChecked Then
        lacTask.Caption = "Set Bus Names"
        DoEvents
        mSetBusNames
    End If
    If imTerminate Then
        lacTask.Caption = ""
        imCompress = False
        Screen.MousePointer = vbDefault
        plcGauge.Visible = False
        Exit Sub
    End If
    If ckcT2.Value = vbChecked Then
        lacTask.Caption = "Assigning T2 to Libraries/Templates"
        DoEvents
        ilRet = mAssignT2()
    End If
    If imTerminate Then
        lacTask.Caption = ""
        imCompress = False
        Screen.MousePointer = vbDefault
        plcGauge.Visible = False
        Exit Sub
    End If
    If ckcUnusedComments.Value = vbChecked Then
        lacTask.Caption = "Removing Unused Comments"
        DoEvents
        mRemoveUnusedComments
    End If
    If imTerminate Then
        lacTask.Caption = ""
        imCompress = False
        Screen.MousePointer = vbDefault
        plcGauge.Visible = False
        Exit Sub
    End If
    lacTask.Caption = "Clean-up Completed"
    plcGauge.Value = 100
    cmcCancel.Caption = "&Done"
    imCompress = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    Me.Refresh
End Sub

Private Sub Form_Load()
    mInit
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilRet As Integer
    imCompress = False
    imTerminate = False
    smPurgeDate = Format$(gNow(), "ddddd")
    smPurgeDate = DateAdd("d", -tgSOE.iDaysRetainAsAir, smPurgeDate)
    ckcPurgeOld.Caption = "Remove Libraries/Templates/Schedules prior to " & smPurgeDate & "...."
    EngrCompress.Height = cmcProcess.Top + 5 * cmcProcess.Height / 3
    EngrCompress.Move (Screen.Width - EngrCompress.Width) \ 2, (Screen.Height - EngrCompress.Height) \ 2 + 175 '+ Screen.Height \ 10
    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrCompress-Minit:Get Bus Definitions", tgCurrBDE())
    hmCTE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCTE, "", sgDBPath & "CTE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim ilRet As Integer
    Unload EngrCompress
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    btrDestroy hmCTE
    
    Erase tmDEECTE
    Erase lmDeleteCode
    Erase tmCurrDBE
    Erase tmT2CTE
    Erase tmT2CTE_Name
    Erase tmDee
    Set EngrCompress = Nothing   'Remove data segment
End Sub

Private Function mCompressLibraries() As Integer
    Dim ilRet As Integer
    Dim llRow As Long
    Dim llCount As Long
    
    gLogMsg "Removing Library History", "EngrFileCleanUp.Txt", False
    llCount = 0
    mCompressLibraries = True
    smDateTimeStamp = ""
    ReDim tmDHE(0 To 0) As DHE
    ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("H", "L", smDateTimeStamp, "EngrCompress-mCompressLibraries", tmDHE())
    If ilRet Then
        For llRow = 0 To UBound(tmDHE) - 1 Step 1
            If gGetCount_SEE_For_DHE(tmDHE(llRow).lCode, "EngrCompress-mCompressLibraries: GetCount") = 0 Then
                ilRet = gPutDelete_DHE_DayHeaderInfo(tmDHE(llRow).lCode, "EngrCompress-mCompressLibraries: Delete DHE")
                llCount = llCount + 1
                If llCount < 1000 Then
                    lacLibCount.Caption = llCount
                Else
                    lacLibCount.Caption = Format$(llCount, "0,000")
                End If
            End If
            DoEvents
            If imTerminate Then
                mCompressLibraries = False
                Exit Function
            End If
            lmCurrentCount = lmCurrentCount + 1
            lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
            If lmPercent >= 100 Then
                lmPercent = 100
            End If
            If plcGauge.Value <> lmPercent Then
                plcGauge.Value = lmPercent
            End If
        Next llRow
    End If
    If llCount < 1000 Then
        gLogMsg "Removed " & llCount & " unused Libraries", "EngrFileCleanUp.Txt", False
    Else
        gLogMsg "Removed " & Format$(llCount, "0,000") & " unused Libraries", "EngrFileCleanUp.Txt", False
    End If
End Function
Private Function mCompressTemplates() As Integer
    Dim ilRet As Integer
    Dim llRow As Long
    Dim llCount As Long
    
    gLogMsg "Removing Template History", "EngrFileCleanUp.Txt", False
    llCount = 0
    mCompressTemplates = True
    smDateTimeStamp = ""
    ReDim tmDHE(0 To 0) As DHE
    ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("H", "T", smDateTimeStamp, "EngrCompress-mCompressTemplates", tmDHE())
    If ilRet Then
        For llRow = 0 To UBound(tmDHE) - 1 Step 1
            If gGetCount_SEE_For_DHE(tmDHE(llRow).lCode, "EngrCompress-mCompressTemplates: GetCount") = 0 Then
                ilRet = gPutDelete_DHE_DayHeaderInfo(tmDHE(llRow).lCode, "EngrCompress-mCompressTemplates: Delete DHE")
                llCount = llCount + 1
                If llCount < 1000 Then
                    lacTempCount.Caption = llCount
                Else
                    lacTempCount.Caption = Format$(llCount, "0,000")
                End If
            End If
            DoEvents
            If imTerminate Then
                mCompressTemplates = False
                Exit Function
            End If
            lmCurrentCount = lmCurrentCount + 1
            lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
            If lmPercent >= 100 Then
                lmPercent = 100
            End If
            If plcGauge.Value <> lmPercent Then
                plcGauge.Value = lmPercent
            End If
        Next llRow
    End If
    If llCount < 1000 Then
        gLogMsg "Removed " & llCount & " unused Templates", "EngrFileCleanUp.Txt", False
    Else
        gLogMsg "Removed " & Format$(llCount, "0,000") & " unused Templates", "EngrFileCleanUp.Txt", False
    End If
End Function
Private Function mCompressComments() As Integer
    Dim llDHE As Long
    Dim ilRet As Long
    Dim llRowOuter As Long
    Dim llRowInner As Long
    Dim llCount As Long
    
    gLogMsg "Removing Duplicate Title 1 Comments", "EngrFileCleanUp.Txt", False
    llCount = 0
    mCompressComments = True
    smDateTimeStamp = ""
    ReDim tmDHE(0 To 0) As DHE
    ReDim tmDEECTE(0 To 0) As DEECTE
    ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("C", "B", smDateTimeStamp, "EngrCompress-mCompressComments", tmDHE())
    If ilRet Then
        For llDHE = 0 To UBound(tmDHE) - 1 Step 1
            ilRet = gGetTypeOfRecs_DEECTE_ForDHE(tmDHE(llDHE).lCode, "EngrCompress- mCompressComments", tmDEECTE)
            If ilRet Then
                For llRowOuter = 0 To UBound(tmDEECTE) - 1 Step 1
                    If tmDEECTE(llRowOuter).lCteCode > 0 Then
                        For llRowInner = llRowOuter + 1 To UBound(tmDEECTE) - 1 Step 1
                            If tmDEECTE(llRowInner).lCteCode > 0 Then
                                If StrComp(UCase(Trim$(tmDEECTE(llRowInner).sComment)), UCase(Trim$(tmDEECTE(llRowOuter).sComment)), vbBinaryCompare) = 0 Then
                                    'Replace
                                    ilRet = gPutUpdate_DEE_CTECode(1, tmDEECTE(llRowInner).lDeeCode, tmDEECTE(llRowOuter).lCteCode, "EngrCompress- mCompressComments: Update Comment")
                                    'Remove
                                    ilRet = gPutDelete_CTE_CommtsTitle(tmDEECTE(llRowInner).lCteCode, "EngrCompress- mCompressComments: Remove Comment")
                                    tmDEECTE(llRowInner).lCteCode = -tmDEECTE(llRowInner).lCteCode
                                    llCount = llCount + 1
                                    If llCount < 1000 Then
                                        lacCommCount.Caption = llCount
                                    Else
                                        lacCommCount.Caption = Format$(llCount, "0,000")
                                    End If
                                Else
                                    Exit For
                                End If
                            End If
                        Next llRowInner
                    End If
                Next llRowOuter
            End If
            DoEvents
            If imTerminate Then
                mCompressComments = False
                Exit Function
            End If
            lmCurrentCount = lmCurrentCount + 1
            lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
            If lmPercent >= 100 Then
                lmPercent = 100
            End If
            If plcGauge.Value <> lmPercent Then
                plcGauge.Value = lmPercent
            End If
        Next llDHE
    End If
    If llCount < 1000 Then
        gLogMsg "Removed " & llCount & " duplicate Title 1 comments", "EngrFileCleanUp.Txt", False
    Else
        gLogMsg "Removed " & Format$(llCount, "0,000") & " duplicate Title 1 comments", "EngrFileCleanUp.Txt", False
    End If
End Function

Private Function mPurgeOld() As Integer
    Dim ilRet As Integer
    Dim slStr As String
    
    gLogMsg "Removing Old Schedules", "EngrFileCleanUp.Txt", False
    lacTask.Caption = "Removing Old Schedules"
    DoEvents
    'ilRet = gSchdAndAsAiredDelete(smPurgeDate, "Delete Schedule and As Aired Prior to " & smPurgeDate)
    ilRet = gSchdAndAsAiredDelete(smPurgeDate, "Delete Schedule Prior to " & smPurgeDate)
    DoEvents
    If imTerminate Then
        mPurgeOld = False
        Exit Function
    End If
    If lgPurgeCount < 1000 Then
        slStr = lgPurgeCount
    Else
        slStr = Format$(lgPurgeCount, "0,000")
    End If
    gLogMsg "Removed " & slStr & " Old Schedules", "EngrFileCleanUp.Txt", False
    lacRemoveOld.Caption = "....Schedule: " & slStr
    gLogMsg "Removing Old Libraries", "EngrFileCleanUp.Txt", False
    lacTask.Caption = "Removing Old Libraries"
    DoEvents
    ilRet = gLibraryDelete(smPurgeDate, "Delete Library Prior to " & smPurgeDate)
    DoEvents
    If imTerminate Then
        mPurgeOld = False
        Exit Function
    End If
    If lgPurgeCount < 1000 Then
        slStr = lgPurgeCount
    Else
        slStr = Format$(lgPurgeCount, "0,000")
    End If
    gLogMsg "Removed " & slStr & " Old Libraries", "EngrFileCleanUp.Txt", False
    lacRemoveOld.Caption = lacRemoveOld.Caption & "; " & "Libraries: " & slStr
    gLogMsg "Removing Old Templates", "EngrFileCleanUp.Txt", False
    lacTask.Caption = "Removing Old Template Schedule Dates"
    DoEvents
    ilRet = gTemplateSchdDelete(smPurgeDate, "Delete Template Schedule Prior to " & smPurgeDate)
    DoEvents
    If imTerminate Then
        mPurgeOld = False
        Exit Function
    End If
    If lgPurgeCount < 1000 Then
        slStr = lgPurgeCount
    Else
        slStr = Format$(lgPurgeCount, "0,000")
    End If
    gLogMsg "Removed " & slStr & " Old Templates", "EngrFileCleanUp.Txt", False
    lacRemoveOld.Caption = lacRemoveOld.Caption & "; " & "Templates: " & slStr
    'lacTask.Caption = "Removing Unused Comments"
    'DoEvents
    'ilRet = gCommentDelete("Delete Comment")
    'DoEvents
    'If imTerminate Then
    '    mPurgeOld = False
    '    Exit Function
    'End If
    'If lgPurgeCount < 1000 Then
    '    slStr = lgPurgeCount
    'Else
    '    slStr = Format$(lgPurgeCount, "0,000")
    'End If
    'lacRemoveOld.Caption = lacRemoveOld.Caption & "; " & "Comments: " & slStr
    mPurgeOld = True
End Function

Private Function mRemoveUnusedComments() As Integer
    Dim ilRet As Integer
    Dim slStr As String
    gLogMsg "Removing Unused Comments", "EngrFileCleanUp.Txt", False
    ilRet = gCommentDelete("Delete Comment")
    DoEvents
    If imTerminate Then
        mRemoveUnusedComments = False
        Exit Function
    End If
    If lgPurgeCount < 1000 Then
        slStr = lgPurgeCount
    Else
        slStr = Format$(lgPurgeCount, "0,000")
    End If
    gLogMsg "Removed " & slStr & " unused Comments", "EngrFileCleanUp.Txt", False
    lacUnusedComments.Caption = slStr
    mRemoveUnusedComments = True
End Function

Private Function mSetBusNames() As Integer

    Dim ilRet As Integer
    Dim llRow As Long
    Dim llCount As Long
    Dim ilDBE As Integer
    Dim ilBDE As Integer
    Dim slStr As String
    Dim llNewAgedDHECode As Long
    
    gLogMsg "Set Bus Names", "EngrFileCleanUp.Txt", False
    llCount = 0
    mSetBusNames = True
    smDateTimeStamp = ""
    ReDim tmDHE(0 To 0) As DHE
    ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("C", "L", smDateTimeStamp, "EngrCompress-mSetBusNames: Get Current Libraries", tmDHE())
    If ilRet Then
        For llRow = 0 To UBound(tmDHE) - 1 Step 1
            smCurrDBE = ""
            ilRet = gGetRecs_DBE_DayBusSel(smCurrDBE, tmDHE(llRow).lCode, "EngrCompress-mSetBusNames: Get Header Bus Names", tmCurrDBE())
            slStr = ""
            For ilDBE = 0 To UBound(tmCurrDBE) - 1 Step 1
                If tmCurrDBE(ilDBE).sType = "B" Then
                    'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                    '    If tmCurrDBE(ilDBE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                        ilBDE = gBinarySearchBDE(tmCurrDBE(ilDBE).iBdeCode, tgCurrBDE())
                        If ilBDE <> -1 Then
                            If slStr = "" Then
                                slStr = Trim$(tgCurrBDE(ilBDE).sName)
                            Else
                                slStr = slStr & "," & Trim$(tgCurrBDE(ilBDE).sName)
                            End If
                    '        Exit For
                        End If
                    'Next ilBDE
                End If
            Next ilDBE
            tmDHE(llRow).sBusNames = slStr
            ilRet = gPutUpdate_DHE_DayHeaderInfo(6, tmDHE(llRow), "EngrCompress-mSetBusNames: Update Bus Names", llNewAgedDHECode)
            DoEvents
            If imTerminate Then
                mSetBusNames = False
                Exit Function
            End If
            llCount = llCount + 1
            If llCount < 1000 Then
                lacBusNamesCount.Caption = llCount
            Else
                lacBusNamesCount.Caption = Format$(llCount, "0,000")
            End If
            DoEvents
            If imTerminate Then
                mSetBusNames = False
                Exit Function
            End If
            lmCurrentCount = lmCurrentCount + 1
            lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
            If lmPercent >= 100 Then
                lmPercent = 100
            End If
            If plcGauge.Value <> lmPercent Then
                plcGauge.Value = lmPercent
            End If
            DoEvents
            If imTerminate Then
                mSetBusNames = False
                Exit Function
            End If
        Next llRow
    End If
    If llCount < 1000 Then
        gLogMsg "Total number of Bus Names Set " & llCount, "EngrFileCleanUp.Txt", False
    Else
        gLogMsg "Total number of Bus Names Set " & Format$(llCount, "0,000"), "EngrFileCleanUp.Txt", False
    End If

End Function

Private Function mAssignT2() As Integer
    Dim ilRet As Integer
    Dim llRow As Long
    Dim llCount As Long
    Dim slStamp As String
    Dim llDHE As Long
    Dim llDee As Long
    Dim llIndex As Long
    Dim llCTE As Long
    Dim ilPass As Integer
    
    gLogMsg "Assigning Title 2", "EngrFileCleanUp.Txt", False
    llCount = 0
    mAssignT2 = True
    slStamp = ""
    ilRet = gGetTypeOfRecs_CTE_CommtsTitle("C", "T2", slStamp, "EngrCompress-mAssignT2", tmT2CTE())
    If UBound(tmT2CTE) <= LBound(tmT2CTE) Then
        gLogMsg "No Title 2 to Assign", "EngrFileCleanUp.Txt", False
        Exit Function
    End If
    For ilPass = 0 To 1 Step 1
        smDateTimeStamp = ""
        ReDim tmDHE(0 To 0) As DHE
        If ilPass = 1 Then
            ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("B", "T", smDateTimeStamp, "EngrCompress-mAssignT2: Get DHE", tmDHE())
        Else
            ilRet = gGetTypeOfRecs_DHE_DayHeaderInfo("B", "L", smDateTimeStamp, "EngrCompress-mAssignT2: Get DHE", tmDHE())
        End If
        If ilRet Then
            For llDHE = 0 To UBound(tmDHE) - 1 Step 1
                ReDim tmT2CTE_Name(0 To 0) As DEECTE
                slStamp = ""
                ilRet = gGetRecs_DEE_DayEvent(slStamp, tmDHE(llDHE).lCode, "EngrCompress-mAssignT2: Get DEE", tmDee())
                For llDee = 0 To UBound(tmDee) - 1 Step 1
                    If tmDee(llDee).l2CteCode > 0 Then
                        llIndex = gBinarySearchCTE(tmDee(llDee).l2CteCode, tmT2CTE())
                        If llIndex <> -1 Then
                            tmDee(llDee).l2CteCode = 0
                            For llCTE = 0 To UBound(tmT2CTE_Name) - 1 Step 1
                                If StrComp(UCase(Trim$(tmT2CTE_Name(llCTE).sComment)), UCase(Trim$(tmT2CTE(llIndex).sComment)), vbBinaryCompare) = 0 Then
                                    tmDee(llDee).l2CteCode = tmT2CTE_Name(llCTE).lCteCode
                                    Exit For
                                End If
                            Next llCTE
                            If tmDee(llDee).l2CteCode = 0 Then
                                mSetCTE tmT2CTE(llIndex).sComment, "T2"
                                ilRet = gPutInsert_CTE_CommtsTitle(0, tmCTE, "EngrCompress-mAssignT2: Insert CTE", hmCTE)
                                If ilRet Then
                                    tmDee(llDee).l2CteCode = tmCTE.lCode
                                    tmT2CTE_Name(UBound(tmT2CTE_Name)).sComment = tmCTE.sComment
                                    tmT2CTE_Name(UBound(tmT2CTE_Name)).lCteCode = tmCTE.lCode
                                    tmT2CTE_Name(UBound(tmT2CTE_Name)).lDheCode = tmDHE(llDHE).lCode
                                    ReDim Preserve tmT2CTE_Name(0 To UBound(tmT2CTE_Name) + 1) As DEECTE
                                End If
                            End If
                            If tmDee(llDee).l2CteCode <> 0 Then
                                llCount = llCount + 1
                                If llCount < 1000 Then
                                    lacT2Count.Caption = llCount
                                Else
                                    lacT2Count.Caption = Format$(llCount, "0,000")
                                End If
                                ilRet = gPutUpdate_DEE_CTECode(2, tmDee(llDee).lCode, tmDee(llDee).l2CteCode, "EngrCompress- mAssignT2: Update Comment")
                            End If
                        End If
                    End If
                Next llDee
                DoEvents
                If imTerminate Then
                    mAssignT2 = False
                    Exit Function
                End If
                lmCurrentCount = lmCurrentCount + 1
                lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
                If lmPercent >= 100 Then
                    lmPercent = 100
                End If
                If plcGauge.Value <> lmPercent Then
                    plcGauge.Value = lmPercent
                End If
            Next llDHE
        End If
    Next ilPass
    
    If llCount < 1000 Then
        gLogMsg "Reassigned " & llCount & " Title 2's", "EngrFileCleanUp.Txt", False
    Else
        gLogMsg "Reassigned " & Format$(llCount, "0,000") & " Title 2's", "EngrFileCleanUp.Txt", False
    End If
End Function
Private Sub mSetCTE(slComment As String, slType As String)
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    tmCTE.lCode = 0
    tmCTE.sComment = slComment
    tmCTE.sState = "A"
    tmCTE.sType = slType    '"DH" or "T1"
    tmCTE.sUsedFlag = "Y"
    tmCTE.iVersion = 0
    tmCTE.lOrigCteCode = 0
    tmCTE.sCurrent = "Y"
    'tmCTE.sEnteredDate = smNowDate
    'tmCTE.sEnteredTime = smNowTime
    tmCTE.sEnteredDate = Format(Now, sgShowDateForm) 'smNowDate
    tmCTE.sEnteredTime = Format(Now, sgShowTimeWSecForm) 'smNowTime
    tmCTE.iUieCode = tgUIE.iCode
    tmCTE.sUnused = ""

End Sub

Private Function mCompressList() As Integer
    Dim ilRet As Integer
    Dim llCount As Long
    
    llCount = 0
    ilRet = gGetListInfo("SELECT aneCode, aneCurrent, aneOrigAneCode, aneVersion, aneState FROM ANE_Audio_Name ORDER BY aneOrigAneCode, aneVersion")
    If ilRet Then
        'Fix ANE
        ilRet = mCheckCurrent("aneCode", "aneCurrent", "ANE_Audio_Name")
        ilRet = mChangeRefs("asePriAneCode", "ASE_Audio_Source")
        ilRet = mChangeRefs("aseBkupAneCode", "ASE_Audio_Source")
        ilRet = mChangeRefs("aseProtAneCode", "ASE_Audio_Source")
        ilRet = mChangeRefs("deeBkupAneCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("deeProtAneCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("seeBkupAneCode", "SEE_Schedule_Events")
        ilRet = mChangeRefs("seeProtAneCode", "SEE_Schedule_Events")
        ilRet = mDeleteUnused("aneCode", "ANE_Audio_Name")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT aseCode, aseCurrent, aseOrigAseCode, aseVersion, aseState FROM ASE_Audio_Source ORDER BY aseOrigAseCode, aseVersion")
    If ilRet Then
        'Fix ASE
        ilRet = mCheckCurrent("aseCode", "aseCurrent", "ASE_Audio_Source")
        ilRet = mChangeRefs("bdeAseCode", "BDE_Bus_Definition")
        ilRet = mChangeRefs("deeAudioAseCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("seeAudioAseCode", "SEE_Schedule_Events")
        ilRet = mDeleteUnused("aseCode", "ASE_Audio_Source")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT ateCode, ateCurrent, ateOrigAteCode, ateVersion, ateState FROM ATE_Audio_Type ORDER BY ateOrigAteCode, ateVersion")
    If ilRet Then
        'Fix ATE
        ilRet = mCheckCurrent("ateCode", "ateCurrent", "ATE_Audio_Type")
        ilRet = mChangeRefs("aneAteCode", "ANE_Audio_Name")
        ilRet = mDeleteUnused("ateCode", "ATE_Audio_Type")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT bdeCode, bdeCurrent, bdeOrigbdeCode, bdeVersion, bdeState FROM BDE_Bus_Definition ORDER BY bdeOrigBdeCode, bdeVersion")
    If ilRet Then
        'Fix BDE
        ilRet = mCheckCurrent("bdeCode", "bdeCurrent", "BDE_Bus_Definition")
        ilRet = mChangeRefs("bseBdeCode", "BSE_Bus_Sel_Group")
        ilRet = mChangeRefs("dbeBdeCode", "DBE_Day_Bus_Sel")
        ilRet = mChangeRefs("ebeBdeCode", "EBE_Event_Bus_Sel")
        ilRet = mChangeRefs("seeBdeCode", "SEE_Schedule_Events")
        ilRet = mChangeRefs("tseBdeCode", "TSE_Template_Schd")
        ilRet = mDeleteUnused("bdeCode", "BDE_Bus_Definition")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT bgeCode, bgeCurrent, bgeOrigbgeCode, bgeVersion, bgeState FROM BGE_Bus_Group ORDER BY bgeOrigBgeCode, bgeVersion")
    If ilRet Then
        'Fix BGE
        ilRet = mCheckCurrent("bgeCode", "bgeCurrent", "BGE_Bus_Group")
        ilRet = mChangeRefs("bseBgeCode", "BSE_Bus_Sel_Group")
        ilRet = mChangeRefs("dbeBgeCode", "DBE_Day_Bus_Sel")
        ilRet = mDeleteUnused("bgeCode", "BGE_Bus_Group")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT cceCode, cceCurrent, cceOrigCceCode, cceVersion, cceState FROM CCE_Control_Char ORDER BY cceOrigCceCode, cceVersion")
    If ilRet Then
        'Fix CCE
        ilRet = mCheckCurrent("cceCode", "cceCurrent", "CCE_Control_Char")
        ilRet = mChangeRefs("aneCceCode", "ANE_Audio_Name")
        ilRet = mChangeRefs("asePriCceCode", "ASE_Audio_Source")
        ilRet = mChangeRefs("aseBkupCceCode", "ASE_Audio_Source")
        ilRet = mChangeRefs("aseProtCceCode", "ASE_Audio_Source")
        ilRet = mChangeRefs("bdeCceCode", "BDE_Bus_Definition")
        ilRet = mChangeRefs("deeCceCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("deeAudioCceCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("deeBkupCceCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("deeProtCceCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("seeBusCceCode", "SEE_Schedule_Events")
        ilRet = mChangeRefs("seeAudioCceCode", "SEE_Schedule_Events")
        ilRet = mChangeRefs("seeBkupCceCode", "SEE_Schedule_Events")
        ilRet = mChangeRefs("seeProtCceCode", "SEE_Schedule_Events")
        ilRet = mDeleteUnused("cceCode", "CCE_Control_Char")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT dneCode, dneCurrent, dneOrigDneCode, dneVersion, dneState FROM DNE_Day_Name ORDER BY dneOrigDneCode, dneVersion")
    If ilRet Then
        'Fix DNE
        ilRet = mCheckCurrent("dneCode", "dneCurrent", "DNE_Day_Name")
        ilRet = mChangeRefs("dheDneCode", "DHE_Day_Header_Info")
        ilRet = mChangeRefs("nneDneCode", "NNE_Netcue_Name")
        ilRet = mDeleteUnused("dneCode", "DNE_Day_Name")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT dseCode, dseCurrent, dseOrigDseCode, dseVersion, dseState FROM DSE_Day_SubName ORDER BY dseOrigDseCode, dseVersion")
    If ilRet Then
        'Fix DSE
        ilRet = mCheckCurrent("dseCode", "dseCurrent", "DSE_Day_SubName")
        ilRet = mChangeRefs("dheDseCode", "DHE_Day_Header_Info")
        ilRet = mDeleteUnused("dseCode", "DSE_Day_SubName")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT eteCode, eteCurrent, eteOrigEteCode, eteVersion, eteState FROM ETE_Event_Type ORDER BY eteOrigEteCode, eteVersion")
    If ilRet Then
        'Fix ETE
        ilRet = mCheckCurrent("eteCode", "eteCurrent", "ETE_Event_Type")
        ilRet = mChangeRefs("deeEteCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("epeEteCode", "EPE_Event_Properties")
        ilRet = mChangeRefs("seeEteCode", "SEE_Schedule_Events")
        ilRet = mDeleteUnused("eteCode", "ETE_Event_Type")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT fneCode, fneCurrent, fneOrigFneCode, fneVersion, fneState FROM FNE_Follow_Name ORDER BY fneOrigFneCode, fneVersion")
    If ilRet Then
        'Fix FNE
        ilRet = mCheckCurrent("fneCode", "fneCurrent", "FNE_Follow_Name")
        ilRet = mChangeRefs("deeFneCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("seeFneCode", "SEE_Schedule_Events")
        ilRet = mDeleteUnused("fneCode", "FNE_Follow_Name")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT mteCode, mteCurrent, mteOrigMteCode, mteVersion, mteState FROM MTE_Material_Type ORDER BY mteOrigMteCode, mteVersion")
    If ilRet Then
        'Fix MTE
        ilRet = mCheckCurrent("mteCode", "mteCurrent", "MTE_Material_Type")
        ilRet = mChangeRefs("deeMteCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("seeMteCode", "SEE_Schedule_Events")
        ilRet = mDeleteUnused("mteCode", "MTE_Material_Type")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT nneCode, nneCurrent, nneOrigNneCode, nneVersion, nneState FROM NNE_Netcue_Name ORDER BY nneOrigNneCode, nneVersion")
    If ilRet Then
        'Fix NNE
        ilRet = mCheckCurrent("nneCode", "nneCurrent", "NNE_Netcue_Name")
        ilRet = mChangeRefs("deeStartNneCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("deeEndNneCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("seeStartNneCode", "SEE_Schedule_Events")
        ilRet = mChangeRefs("seeEndNneCode", "SEE_Schedule_Events")
        ilRet = mDeleteUnused("nneCode", "NNE_Netcue_Name")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT rneCode, rneCurrent, rneOrigRneCode, rneVersion, rneState FROM RNE_Relay_Name ORDER BY rneOrigRneCode, rneVersion")
    If ilRet Then
        'Fix RNE
        ilRet = mCheckCurrent("rneCode", "rneCurrent", "RNE_Relay_Name")
        ilRet = mChangeRefs("dee1RneCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("dee2RneCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("see1RneCode", "SEE_Schedule_Events")
        ilRet = mChangeRefs("see2RneCode", "SEE_Schedule_Events")
        ilRet = mDeleteUnused("rneCode", "RNE_Relay_Name")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT sceCode, sceCurrent, sceOrigSceCode, sceVersion, sceState FROM SCE_Silence_Char ORDER BY sceOrigSceCode, sceVersion")
    If ilRet Then
        'Fix SCE
        ilRet = mCheckCurrent("sceCode", "sceCurrent", "SCE_Silence_Char")
        ilRet = mChangeRefs("dee1SceCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("dee2SceCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("dee3SceCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("dee4SceCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("see1SceCode", "SEE_Schedule_Events")
        ilRet = mChangeRefs("see2SceCode", "SEE_Schedule_Events")
        ilRet = mChangeRefs("see3SceCode", "SEE_Schedule_Events")
        ilRet = mChangeRefs("see4SceCode", "SEE_Schedule_Events")
        ilRet = mDeleteUnused("sceCode", "SCE_Silence_Char")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
    ilRet = gGetListInfo("SELECT tteCode, tteCurrent, tteOrigTteCode, tteVersion, tteState FROM TTE_Time_Type ORDER BY tteOrigTteCode, tteVersion")
    If ilRet Then
        'Fix TTE
        ilRet = mCheckCurrent("tteCode", "tteCurrent", "TTE_Time_Type")
        ilRet = mChangeRefs("deeStartTteCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("deeEndTteCode", "DEE_Day_Event_Info")
        ilRet = mChangeRefs("seeStartTteCode", "SEE_Schedule_Events")
        ilRet = mChangeRefs("seeEndTteCode", "SEE_Schedule_Events")
        ilRet = mDeleteUnused("tteCode", "TTE_Time_Type")
        DoEvents
        If imTerminate Then
            mCompressList = False
            Exit Function
        End If
    End If
    llCount = llCount + 1
    If llCount < 1000 Then
        lacList.Caption = llCount
    Else
        lacList.Caption = Format$(llCount, "0,000")
    End If
    lmCurrentCount = lmCurrentCount + 1
    lmPercent = (lmCurrentCount * CSng(100)) / lmTotalCount
    If lmPercent >= 100 Then
        lmPercent = 100
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
    End If
End Function

Private Function mCheckCurrent(slCodeName As String, slFieldName As String, slFileName As String) As Integer
    Dim llRow As Long
    Dim ilRet As Integer
    
    For llRow = 0 To UBound(tgListInfo) - 1 Step 1
        If llRow <> UBound(tgListInfo) - 1 Then
            If tgListInfo(llRow).lOrigCode = tgListInfo(llRow + 1).lOrigCode Then
                If tgListInfo(llRow).sCurrent = "Y" Then
                    'Change to N
                    ilRet = gExecGenSQLCall("UPDATE " & slFileName & " SET " & slFieldName & " = 'N'" & " WHERE " & slCodeName & " = " & tgListInfo(llRow).lCode)
                    tgListInfo(llRow).sCurrent = "N"
                End If
            Else
                If tgListInfo(llRow).sCurrent <> "Y" Then
                    'Change to Y
                    ilRet = gExecGenSQLCall("UPDATE " & slFileName & " SET " & slFieldName & " = 'Y'" & " WHERE " & slCodeName & " = " & tgListInfo(llRow).lCode)
                    tgListInfo(llRow).sCurrent = "Y"
                End If
            End If
        Else
            If tgListInfo(llRow).sCurrent <> "Y" Then
                'Change to Y
                ilRet = gExecGenSQLCall("UPDATE " & slFileName & " SET " & slFieldName & " = 'Y'" & " WHERE " & slCodeName & " = " & tgListInfo(llRow).lCode)
                tgListInfo(llRow).sCurrent = "Y"
            End If
        End If
    Next llRow
    mCheckCurrent = True
End Function

Private Function mChangeRefs(slFieldName As String, slFileName As String) As Integer
    Dim llRowOuter As Long
    Dim llRowInner
    Dim ilRet As Integer
    
    For llRowOuter = 0 To UBound(tgListInfo) - 1 Step 1
        If tgListInfo(llRowOuter).sCurrent <> "Y" Then
            For llRowInner = llRowOuter + 1 To UBound(tgListInfo) - 1 Step 1
                If tgListInfo(llRowInner).sCurrent = "Y" Then
                    'Fix references
                    ilRet = gExecGenSQLCall("UPDATE " & slFileName & " SET " & slFieldName & " = " & tgListInfo(llRowInner).lCode & " WHERE " & slFieldName & " = " & tgListInfo(llRowOuter).lCode)
                    Exit For
                End If
            Next llRowInner
        End If
    Next llRowOuter
    mChangeRefs = True
End Function


Private Function mDeleteUnused(slCodeName As String, slFileName As String) As Integer
    Dim llRow As Long
    Dim ilRet As Integer

    For llRow = 0 To UBound(tgListInfo) - 1 Step 1
        If tgListInfo(llRow).sCurrent <> "Y" Then
            ilRet = gExecGenSQLCall("DELETE FROM " & slFileName & " WHERE " & slCodeName & " = " & tgListInfo(llRow).lCode)
        End If
    Next llRow
    mDeleteUnused = True
End Function
