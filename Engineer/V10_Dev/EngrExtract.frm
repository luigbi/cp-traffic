VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form EngrExtract 
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   9825
   ControlBox      =   0   'False
   Icon            =   "EngrExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frcExtract 
      Caption         =   "Extract"
      Height          =   2340
      Left            =   210
      TabIndex        =   1
      Top             =   300
      Width           =   9345
      Begin VB.ListBox lbcANE 
         Height          =   1425
         ItemData        =   "EngrExtract.frx":08CA
         Left            =   6735
         List            =   "EngrExtract.frx":08CC
         MultiSelect     =   2  'Extended
         TabIndex        =   12
         Top             =   720
         Width           =   2355
      End
      Begin VB.CheckBox ckcAudio 
         Caption         =   "Match Audio"
         Height          =   300
         Left            =   6750
         TabIndex        =   11
         Top             =   330
         Width           =   2100
      End
      Begin VB.TextBox edcOffsets 
         Height          =   285
         Left            =   3165
         TabIndex        =   8
         Text            =   "00:00-60:00"
         Top             =   1080
         Width           =   3330
      End
      Begin V10EngineeringDev.CSI_HourPicker hpcHours 
         Height          =   270
         Left            =   3180
         TabIndex        =   10
         Top             =   1455
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   476
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BorderStyle     =   1
         CSI_ShowSelectRangeButtons=   -1  'True
         CSI_AllowMultiSelection=   -1  'True
         CSI_ShowDayPartButtons=   0   'False
         CSI_ShowDropDownOnFocus=   -1  'True
         CSI_InputBoxBoxAlignment=   0
         CSI_HourOnColor =   4638790
         CSI_HourOffColor=   -2147483633
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
      Begin V10EngineeringDev.CSI_TimeLength tlcEndTime 
         Height          =   255
         Left            =   3780
         TabIndex        =   6
         Top             =   705
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   450
         Text            =   "00:00:00"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BorderStyle     =   1
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
      Begin V10EngineeringDev.CSI_TimeLength tlcStartTime 
         Height          =   255
         Left            =   3780
         TabIndex        =   4
         Top             =   330
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   450
         Text            =   "00:00:00"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BorderStyle     =   1
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
      Begin VB.ListBox lbcBDE 
         Height          =   1815
         ItemData        =   "EngrExtract.frx":08CE
         Left            =   270
         List            =   "EngrExtract.frx":08D0
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   330
         Width           =   1935
      End
      Begin VB.Label lacHours 
         Caption         =   "Hours"
         Height          =   240
         Left            =   2460
         TabIndex        =   9
         Top             =   1485
         Width           =   675
      End
      Begin VB.Label lacOffsets 
         Caption         =   "Offsets"
         Height          =   255
         Left            =   2460
         TabIndex        =   7
         Top             =   1095
         Width           =   1005
      End
      Begin VB.Label lacEndTime 
         Caption         =   "End Time"
         Height          =   240
         Left            =   2460
         TabIndex        =   5
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lacStartTime 
         Caption         =   "Start Time"
         Height          =   240
         Left            =   2460
         TabIndex        =   3
         Top             =   345
         Width           =   1305
      End
   End
   Begin VB.Frame frcFrom 
      Caption         =   "From"
      Height          =   1980
      Left            =   240
      TabIndex        =   13
      Top             =   2775
      Width           =   9345
      Begin VB.FileListBox lbcFiles 
         Height          =   1455
         Left            =   6255
         MultiSelect     =   2  'Extended
         Pattern         =   "*.sch"
         TabIndex        =   16
         Top             =   330
         Width           =   2835
      End
      Begin VB.DirListBox lbcFolder 
         Height          =   1440
         Left            =   2520
         TabIndex        =   15
         Top             =   330
         Width           =   3525
      End
      Begin VB.DriveListBox cbcDrive 
         Height          =   315
         Left            =   315
         TabIndex        =   14
         Top             =   330
         Width           =   1965
      End
   End
   Begin VB.ListBox lbcError 
      Height          =   1230
      ItemData        =   "EngrExtract.frx":08D2
      Left            =   300
      List            =   "EngrExtract.frx":08D4
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4875
      Width           =   9255
   End
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5130
      TabIndex        =   19
      Top             =   6645
      Width           =   1245
   End
   Begin VB.CommandButton cmcExtract 
      Caption         =   "Extract"
      Height          =   315
      Left            =   3555
      TabIndex        =   18
      Top             =   6645
      Width           =   1245
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8865
      Top             =   6150
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7110
      FormDesignWidth =   9825
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   225
      Left            =   7410
      TabIndex        =   22
      Top             =   6705
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lacFiles 
      Caption         =   "File"
      Height          =   210
      Index           =   1
      Left            =   6675
      TabIndex        =   23
      Top             =   6705
      Width           =   525
   End
   Begin VB.Label lacScreen 
      Caption         =   "Import Engineering Files"
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4590
   End
   Begin VB.Label lacFiles 
      Caption         =   "Files"
      Height          =   210
      Index           =   0
      Left            =   300
      TabIndex        =   21
      Top             =   6705
      Width           =   2820
   End
   Begin VB.Label lacMsg 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   330
      TabIndex        =   20
      Top             =   6300
      Width           =   9195
   End
End
Attribute VB_Name = "EngrExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  EngrExtract - displays import csv information
'*
'*  Created Aug,1998 by Dick LeVine
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private hmCTE As Integer

Private imTerminate As Integer
Private imExtracting As Integer
Private hmFrom As Integer
Private hmMsg As Integer
Private smMsgFile As String
Private smNowDate As String
Private smNowTime As String
Private lmTotalNoBytes As Long
Private lmProcessedNoBytes As Long
Private smCurDir As String
Private lmFloodPercent As Long
Private imImportSelection As Integer
Private imStartTimeChgd As Integer
Private imEndTimeChgd As Integer
Private smImportPath As String
Private smFromFileName() As String
Private tmExtract() As SCHDEXTRACT
Private tmDayExtract() As SCHDEXTRACT
Private smBothANEStamp As String
Private tmBothANE() As ANE
Private tmANE As ANE
Private smBothNNEStamp As String
Private tmBothNNE() As NNE
Private tmNNE As NNE
Private smBothRNEStamp As String
Private tmBothRNE() As RNE
Private tmRNE As RNE
Private smBothBDEStamp As String
Private tmBothBDE() As BDE
Private tmBDE As BDE
Private smBothASEStamp As String
Private tmBothASE() As ASE
Private tmASE As ASE
Private smBothAudioCCEStamp As String
Private tmBothAudioCCE() As CCE
Private smBothBusCCEStamp As String
Private tmBothBusCCE() As CCE
Private tmCCE As CCE
Private smBothCTEStamp As String
Private tmBothCTE() As CTE
Private tmCTE As CTE
Private smBothFNEStamp As String
Private tmBothFNE() As FNE
Private tmFNE As FNE
Private smBothMTEStamp As String
Private tmBothMTE() As MTE
Private tmMTE As MTE
Private smBothSCEStamp As String
Private tmBothSCE() As SCE
Private tmSCE As SCE
Private smBothEndTTEStamp As String
Private tmBothEndTTE() As TTE
Private smBothStartTTEStamp As String
Private tmBothStartTTE() As TTE
Private tmTTE As TTE

























'*******************************************************
'*                                                     *
'*      Procedure Name:mReadFileCP                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Private Function mExtractRecs(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilLoop As Integer
    Dim slCode As String
    Dim ilFound As Integer
    Dim slMsg As String
    Dim llPercent As Long
    Dim slStr As String
    Dim ilExtract As Integer
    Dim ilUse As Integer
    Dim ilBus As Integer
    Dim ilAudio As Integer
    Dim llTime As Long
    Dim ilPos As Integer
    Dim slHour As String
    Dim ilHour As Integer
    Dim slOffSet As String
    Dim llOffset As Long
    Dim ilOffset As Integer
    Dim ilDay As Integer
    Dim ilDayExtract As Integer
    Dim llUpperStart As Long
    Dim llUpper As Long
    ReDim tlDayExtract(0 To 0) As SCHDEXTRACT
        
    ilRet = 0
    llUpperStart = UBound(tmDayExtract)
    ReDim tmExtract(0 To 0) As SCHDEXTRACT
    On Error GoTo mExtractRecsErr:
    hmFrom = FreeFile
    Open smImportPath & slFromFile For Input Access Read As hmFrom
    If ilRet <> 0 Then
        lbcError.AddItem "Open " & slFromFile & " error#" & Str$(ilRet)
        Print #hmMsg, "Open " & slFromFile & " error#" & Str$(ilRet)
        Close hmFrom
        mExtractRecs = False
        Exit Function
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    lmProcessedNoBytes = 0
    Do
        ilRet = 0
        On Error GoTo mExtractRecsErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet <> 0 Then
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Print #hmMsg, "User Cancelled: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
            Close hmFrom
            mExtractRecs = False
            Exit Function
        End If
        lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2
        slLine = Trim$(slLine)
        If (Len(slLine) > 0) Then
            'Extract Fields
            llUpper = UBound(tmExtract)
            tmExtract(llUpper).sEventType = mExtractField(slLine, tgStartColAFE.iEventType, tgNoCharAFE.iEventType)
            tmExtract(llUpper).sBus = mExtractField(slLine, tgStartColAFE.iBus, tgNoCharAFE.iBus)
            tmExtract(llUpper).sBusCtrl = mExtractField(slLine, tgStartColAFE.iBusControl, tgNoCharAFE.iBusControl)
            tmExtract(llUpper).sTime = mExtractField(slLine, tgStartColAFE.iTime, tgNoCharAFE.iTime)
            tmExtract(llUpper).sStartType = mExtractField(slLine, tgStartColAFE.iStartType, tgNoCharAFE.iStartType)
            tmExtract(llUpper).sEndType = mExtractField(slLine, tgStartColAFE.iEndType, tgNoCharAFE.iEndType)
            tmExtract(llUpper).sDuration = mExtractField(slLine, tgStartColAFE.iDuration, tgNoCharAFE.iDuration)
            tmExtract(llUpper).sMaterialType = mExtractField(slLine, tgStartColAFE.iMaterialType, tgNoCharAFE.iMaterialType)
            tmExtract(llUpper).sAudioName = mExtractField(slLine, tgStartColAFE.iAudioName, tgNoCharAFE.iAudioName)
            tmExtract(llUpper).sAudioID = mExtractField(slLine, tgStartColAFE.iAudioItemID, tgNoCharAFE.iAudioItemID)
            tmExtract(llUpper).sAudioISCI = mExtractField(slLine, tgStartColAFE.iAudioISCI, tgNoCharAFE.iAudioISCI)
            tmExtract(llUpper).sAudioCtrl = mExtractField(slLine, tgStartColAFE.iAudioControl, tgNoCharAFE.iAudioControl)
            tmExtract(llUpper).sBackupName = mExtractField(slLine, tgStartColAFE.iBkupAudioName, tgNoCharAFE.iBkupAudioName)
            tmExtract(llUpper).sBackupCtrl = mExtractField(slLine, tgStartColAFE.iBkupAudioControl, tgNoCharAFE.iBkupAudioControl)
            tmExtract(llUpper).sProtName = mExtractField(slLine, tgStartColAFE.iProtAudioName, tgNoCharAFE.iProtAudioName)
            tmExtract(llUpper).sProtItemID = mExtractField(slLine, tgStartColAFE.iProtItemID, tgNoCharAFE.iProtItemID)
            tmExtract(llUpper).sProtISCI = mExtractField(slLine, tgStartColAFE.iProtISCI, tgNoCharAFE.iProtISCI)
            tmExtract(llUpper).sProtCtrl = mExtractField(slLine, tgStartColAFE.iProtAudioControl, tgNoCharAFE.iProtAudioControl)
            tmExtract(llUpper).sRelay1 = mExtractField(slLine, tgStartColAFE.iRelay1, tgNoCharAFE.iRelay1)
            tmExtract(llUpper).sRelay2 = mExtractField(slLine, tgStartColAFE.iRelay2, tgNoCharAFE.iRelay2)
            tmExtract(llUpper).sFollow = mExtractField(slLine, tgStartColAFE.iFollow, tgNoCharAFE.iFollow)
            tmExtract(llUpper).sSilenceTime = mExtractField(slLine, tgStartColAFE.iSilenceTime, tgNoCharAFE.iSilenceTime)
            tmExtract(llUpper).sSilence1 = mExtractField(slLine, tgStartColAFE.iSilence1, tgNoCharAFE.iSilence1)
            tmExtract(llUpper).sSilence2 = mExtractField(slLine, tgStartColAFE.iSilence2, tgNoCharAFE.iSilence2)
            tmExtract(llUpper).sSilence3 = mExtractField(slLine, tgStartColAFE.iSilence3, tgNoCharAFE.iSilence3)
            tmExtract(llUpper).sSilence4 = mExtractField(slLine, tgStartColAFE.iSilence4, tgNoCharAFE.iSilence4)
            tmExtract(llUpper).sNetcue1 = mExtractField(slLine, tgStartColAFE.iStartNetcue, tgNoCharAFE.iStartNetcue)
            tmExtract(llUpper).sNetcue2 = mExtractField(slLine, tgStartColAFE.iStopNetcue, tgNoCharAFE.iStopNetcue)
            tmExtract(llUpper).sTitle1 = mExtractField(slLine, tgStartColAFE.iTitle1, tgNoCharAFE.iTitle1)
            tmExtract(llUpper).sTitle2 = mExtractField(slLine, tgStartColAFE.iTitle2, tgNoCharAFE.iTitle2)
            tmExtract(llUpper).sFixedTime = mExtractField(slLine, tgStartColAFE.iFixedTime, tgNoCharAFE.iFixedTime)
            tmExtract(llUpper).sDate = mExtractField(slLine, tgStartColAFE.iDate, tgNoCharAFE.iDate)
            tmExtract(llUpper).sEndTime = mExtractField(slLine, tgStartColAFE.iEndTime, tgNoCharAFE.iEndTime)
            tmExtract(llUpper).sEventID = mExtractField(slLine, tgStartColAFE.iEventID, tgNoCharAFE.iEventID)
            If ((Left(tmExtract(llUpper).sAudioID, 1) = "C") Or (Left(tmExtract(llUpper).sAudioID, 1) = "P")) And (gStrLengthInTenthToLong(tmExtract(llUpper).sDuration) >= gStrLengthInTenthToLong("00:30")) Then
                tmExtract(llUpper).sEventType = "A"
                tmExtract(llUpper).sAudioID = ""
                tmExtract(llUpper).sProtItemID = ""
                tmExtract(llUpper).sProtISCI = ""
                tmExtract(llUpper).sTitle1 = ""
                tmExtract(llUpper).sTitle2 = ""
            Else
                tmExtract(llUpper).sEventType = "P"
            End If
            If Trim$(tmExtract(llUpper).sFixedTime) = "" Then
                tmExtract(llUpper).sFixedTime = "N"
            Else
                tmExtract(llUpper).sFixedTime = "Y"
            End If
            tmExtract(llUpper).sHours = String(24, "N")
            tmExtract(llUpper).sDays = String(7, "N")
            tmExtract(llUpper).lLinkBus = -1
            'Use Record
            llTime = gStrLengthInTenthToLong(tmExtract(llUpper).sTime)
            tmExtract(llUpper).lRunningTime = llTime
            ilUse = False
            If (llTime >= lgExtractStartTime) And (llTime < lgExtractEndTime) Then
                ilPos = InStr(1, tmExtract(llUpper).sTime, ":", vbTextCompare)
                If ilPos > 0 Then
                    slHour = Left$(tmExtract(llUpper).sTime, ilPos - 1)
                    ilHour = Val(slHour) + 1
                    slOffSet = Mid$(tmExtract(llUpper).sTime, ilPos + 1)
                    llOffset = gStrLengthInTenthToLong(slOffSet)
                    If Mid$(sgExtractHours, ilHour, 1) = "Y" Then
                        For ilOffset = LBound(lgExtractOffsetStart) To UBound(lgExtractOffsetStart) Step 1
                            If (llOffset >= lgExtractOffsetStart(ilOffset)) And (llOffset <= lgExtractOffsetEnd(ilOffset)) Then
                                For ilBus = 0 To UBound(sgExtractBusNames) - 1 Step 1
                                    If StrComp(Trim$(tmExtract(llUpper).sBus), Trim$(sgExtractBusNames(ilBus)), vbTextCompare) = 0 Then
                                        If ckcAudio.Value = vbChecked Then
                                            For ilAudio = 0 To UBound(sgExtractAudios) - 1 Step 1
                                                If StrComp(Trim$(tmExtract(llUpper).sAudioName), Trim$(sgExtractAudios(ilAudio)), vbTextCompare) = 0 Then
                                                    ilUse = True
                                                    Exit For
                                                End If
                                            Next ilAudio
                                        Else
                                            ilUse = True
                                        End If
                                        Exit For
                                    End If
                                Next ilBus
                            Exit For
                            End If
                        Next ilOffset
                    End If
                End If
            End If
            If ilUse Then
                'Merge
                If sgExtractType = "T" Then
                    'Ignore bus and Bus Ctrl with templates
                    tmExtract(llUpper).sBus = ""
                    tmExtract(llUpper).sBusCtrl = ""
                    'Adjust time
                    llTime = gStrLengthInTenthToLong(tmExtract(llUpper).sTime) - lgExtractStartTime
                    slStr = gLongToStrLengthInTenth(llTime, True)
                    ilPos = InStr(1, slStr, ":", vbTextCompare)
                    If ilPos > 0 Then
                        slHour = Left$(slStr, ilPos - 1)
                        ilHour = Val(slHour) + 1
                        tmExtract(llUpper).sOffset = Mid$(slStr, ilPos + 1)
                        tmExtract(llUpper).lOffset = gStrLengthInTenthToLong(tmExtract(llUpper).sOffset)
                    Else
                        ilUse = False
                    End If
                Else
                    tmExtract(llUpper).sOffset = slOffSet
                    tmExtract(llUpper).lOffset = gStrLengthInTenthToLong(tmExtract(llUpper).sOffset)
                End If
                slStr = Mid$(tmExtract(llUpper).sDate, 5, 2) & "/" & Mid$(tmExtract(llUpper).sDate, 7, 2) & "/" & Left(tmExtract(llUpper).sDate, 4)
                ilDay = Weekday(slStr, vbMonday)
            End If
            If ilUse Then
                Mid(sgExtractDays, ilDay, 1) = "Y"
                Mid(tmExtract(llUpper).sHours, ilHour, 1) = "Y"
                Mid(tmExtract(llUpper).sDays, ilDay, 1) = "Y"
                If tmExtract(llUpper).sEventType = "A" Then
                    For ilExtract = 0 To UBound(tmExtract) - 1 Step 1
                        If tmExtract(ilExtract).sEventType = "A" Then
                            If mCompareExtract(True, False, tmExtract(ilExtract), tmExtract(llUpper)) Then
                                If tmExtract(ilExtract).sHours = tmExtract(llUpper).sHours Then
                                    tmExtract(ilExtract).lRunningTime = tmExtract(ilExtract).lRunningTime + gStrLengthInTenthToLong(tmExtract(llUpper).sDuration)
                                    tmExtract(ilExtract).sDuration = gLongToStrLengthInTenth(gStrLengthInTenthToLong(tmExtract(ilExtract).sDuration) + gStrLengthInTenthToLong(tmExtract(llUpper).sDuration), True)
                                    If Trim$(tmExtract(llUpper).sNetcue1) <> "" Then
                                        tmExtract(ilExtract).sNetcue1 = tmExtract(llUpper).sNetcue1
                                    End If
                                    If Trim$(tmExtract(llUpper).sNetcue2) <> "" Then
                                        tmExtract(ilExtract).sNetcue2 = tmExtract(llUpper).sNetcue2
                                    End If
                                    ilUse = False
                                    Exit For
                                End If
                            End If
                        End If
                    Next ilExtract
                End If
                If ilUse Then
                    tmExtract(llUpper).lRunningTime = tmExtract(llUpper).lRunningTime + gStrLengthInTenthToLong(tmExtract(llUpper).sDuration)
                    ReDim Preserve tmExtract(0 To UBound(tmExtract) + 1) As SCHDEXTRACT
                End If
            End If
            llPercent = (lmProcessedNoBytes / lmTotalNoBytes) * 100
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmFloodPercent <> llPercent Then
                lmFloodPercent = llPercent
                plcGauge.Value = llPercent
            End If
        End If
        ilRet = 0
    Loop
    Close hmFrom
    'Merge Hours
    For ilExtract = 0 To UBound(tmExtract) - 1 Step 1
        ilFound = False
        For ilDayExtract = 0 To UBound(tlDayExtract) - 1 Step 1
            If mCompareExtract(True, True, tlDayExtract(ilDayExtract), tmExtract(ilExtract)) Then
                For ilHour = 1 To 24 Step 1
                    If Mid(tmExtract(ilExtract).sHours, ilHour, 1) = "Y" Then
                        Mid(tlDayExtract(ilDayExtract).sHours, ilHour, 1) = "Y"
                    End If
                Next ilHour
                For ilDay = 1 To 7 Step 1
                    If Mid(tmExtract(ilExtract).sDays, ilDay, 1) = "Y" Then
                        Mid(tlDayExtract(ilDayExtract).sDays, ilDay, 1) = "Y"
                    End If
                Next ilDay
                ilFound = True
                Exit For
            End If
        Next ilDayExtract
        If Not ilFound Then
            LSet tlDayExtract(UBound(tlDayExtract)) = tmExtract(ilExtract)
            ReDim Preserve tlDayExtract(0 To UBound(tlDayExtract) + 1) As SCHDEXTRACT
        End If
    Next ilExtract
    'Link buses
    For ilExtract = 0 To UBound(tlDayExtract) - 1 Step 1
        If sgExtractType <> "T" Then
            For ilDayExtract = llUpperStart To UBound(tmDayExtract) - 1 Step 1
                If tmDayExtract(ilDayExtract).lLinkBus = -1 Then
                    If (mCompareExtract(False, True, tmDayExtract(ilDayExtract), tlDayExtract(ilExtract))) And (StrComp(Trim$(tmDayExtract(ilDayExtract).sBus), Trim$(tlDayExtract(ilExtract).sBus), vbTextCompare) <> 0) And (tmDayExtract(ilDayExtract).sHours = tlDayExtract(ilExtract).sHours) Then
                        tmDayExtract(ilDayExtract).lLinkBus = UBound(tmDayExtract)
                        Exit For
                    End If
                End If
            Next ilDayExtract
        End If
        LSet tmDayExtract(UBound(tmDayExtract)) = tlDayExtract(ilExtract)
        ReDim Preserve tmDayExtract(0 To UBound(tmDayExtract) + 1) As SCHDEXTRACT
    Next ilExtract
    mExtractRecs = True
    lmFloodPercent = 100
    Print #hmMsg, ""
    Exit Function
mExtractRecsErr:
    ilRet = Err.Number
    Resume Next
End Function







'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Private Function mOpenMsgFile() As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer

    On Error GoTo mOpenMsgFileErr:
    slToFile = sgMsgDirectory & smMsgFile
    slNowDate = Format$(gNow(), sgShowDateForm)
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        Kill slToFile
        On Error GoTo 0
        ilRet = 0
        On Error GoTo mOpenMsgFileErr:
        hmMsg = FreeFile
        Open slToFile For Output As hmMsg
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            MsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        On Error GoTo mOpenMsgFileErr:
        hmMsg = FreeFile
        Open slToFile For Output As hmMsg
        If ilRet <> 0 Then
            Close hmMsg
            hmMsg = -1
            MsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    If sgExtractType <> "T" Then
        Print #hmMsg, "** Extract Libraries: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Else
        Print #hmMsg, "** Extract Templates: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    End If
    Print #hmMsg, ""
    mOpenMsgFile = True
    Exit Function
mOpenMsgFileErr:
    ilRet = 1
    Resume Next
End Function







Private Sub cbcDrive_Change()
    lbcFolder.Path = cbcDrive.Drive
End Sub

Private Sub ckcAudio_Click()
    Dim slStr As String
    Dim ilANE As Integer
    Dim ilIndex As Integer
    Dim llRg As Long
    Dim ilValue As Integer
    Dim llRet As Long
    
    ilValue = False
    llRg = CLng(lbcANE.ListCount - 1) * &H10000 Or 0
    llRet = SendMessageByNum(lbcANE.hwnd, LB_SELITEMRANGE, ilValue, llRg)
    If ckcAudio.Value = vbChecked Then
        For ilANE = 0 To UBound(sgExtractAudios) - 1 Step 1
            slStr = Trim$(sgExtractAudios(ilANE))
            ilIndex = gListBoxFind(lbcANE, slStr)
            If ilIndex >= 0 Then
                lbcANE.Selected(ilIndex) = True
            End If
        Next ilANE
    End If
End Sub

Private Sub cmcCancel_Click()
    If imExtracting Then
        imTerminate = True
        Exit Sub
    End If
    igReturnCallStatus = CALLCANCELLED
    Unload EngrExtract
End Sub

Private Sub cmcExtract_Click()
    Dim ilRet As Integer
    Dim slDate As String
    Dim ilVefCode As Integer
    Dim ilSelected As Integer
    Dim ilLoop As Integer
    Dim slHours As String
    Dim slOffsets As String
    Dim ilPos As Integer
    Dim slStartOffset As String
    Dim slEndOffset As String
    Dim slStr As String
    
    ReDim tgExtract(0 To 0) As SCHDEXTRACT
    ReDim smFromFileName(0 To 0) As String
    lacFiles(0).Caption = ""
    lacMsg.Caption = ""
    plcGauge.Value = 0
    smImportPath = lbcFolder.Path
    If Right$(smImportPath, 1) <> "\" Then
        smImportPath = smImportPath & "\"
    End If
    For ilLoop = 0 To lbcFiles.ListCount - 1 Step 1
        If lbcFiles.Selected(ilLoop) Then
            smFromFileName(UBound(smFromFileName)) = lbcFiles.List(ilLoop)
            ReDim Preserve smFromFileName(0 To UBound(smFromFileName) + 1) As String
        End If
    Next ilLoop
    If UBound(smFromFileName) <= LBound(smFromFileName) Then
        MsgBox "Extract File must be specified.", vbOKOnly
        Exit Sub
    End If
    ReDim sgExtractBusNames(0 To 0) As String
    For ilLoop = 0 To lbcBDE.ListCount - 1 Step 1
        If lbcBDE.Selected(ilLoop) Then
            sgExtractBusNames(UBound(sgExtractBusNames)) = lbcBDE.List(ilLoop)
            ReDim Preserve sgExtractBusNames(0 To UBound(sgExtractBusNames) + 1) As String
        End If
    Next ilLoop
    If UBound(sgExtractBusNames) <= LBound(sgExtractBusNames) Then
        MsgBox "Bus Names must be specified.", vbOKOnly
        lbcBDE.SetFocus
        Exit Sub
    End If
    ReDim sgExtractAudios(0 To 0) As String
    If ckcAudio.Value = vbChecked Then
        For ilLoop = 0 To lbcANE.ListCount - 1 Step 1
            If lbcANE.Selected(ilLoop) Then
                sgExtractAudios(UBound(sgExtractAudios)) = lbcANE.List(ilLoop)
                ReDim Preserve sgExtractAudios(0 To UBound(sgExtractAudios) + 1) As String
            End If
        Next ilLoop
        If UBound(sgExtractAudios) <= LBound(sgExtractAudios) Then
            MsgBox "Audio Names must be specified.", vbOKOnly
            lbcANE.SetFocus
            Exit Sub
        End If
    End If
    sgExtractStartTime = tlcStartTime.text
    If (Not gIsLength(sgExtractStartTime)) Or (sgExtractStartTime = "") Then
        MsgBox "Invalid Start Time specified.", vbOKOnly
        tlcStartTime.SetFocus
        Exit Sub
    End If
    lgExtractStartTime = gStrLengthInTenthToLong(sgExtractStartTime)
    sgExtractEndTime = tlcEndTime.text
    If (Not gIsLength(sgExtractEndTime)) Or (sgExtractEndTime = "") Then
        MsgBox "Invalid End Time specified.", vbOKOnly
        tlcEndTime.SetFocus
        Exit Sub
    End If
    lgExtractEndTime = gStrLengthInTenthToLong(sgExtractEndTime)
    If lgExtractEndTime < lgExtractStartTime Then
        MsgBox "End Time prior to Start Time not allowed.", vbOKOnly
        tlcEndTime.SetFocus
        Exit Sub
    End If
    slHours = hpcHours.text
    If slHours = "" Then
        MsgBox "Hours must be specified.", vbOKOnly
        hpcHours.SetFocus
        Exit Sub
    End If
    sgExtractHours = gCreateHourStr(slHours)
    ilPos = InStr(1, sgExtractHours, "Y", vbTextCompare)
    If ilPos <= 0 Then
        MsgBox "No Hours specified.", vbOKOnly
        hpcHours.SetFocus
        Exit Sub
    End If
    slOffsets = Trim$(edcOffsets.text)
    If slOffsets = "" Then
        MsgBox "Offsets must be specified.", vbOKOnly
        edcOffsets.SetFocus
        Exit Sub
    End If
    gParseCDFields slOffsets, False, sgExtractOffsets()
    ReDim lgExtractOffsetStart(LBound(sgExtractOffsets) To UBound(sgExtractOffsets)) As Long
    ReDim lgExtractOffsetEnd(LBound(sgExtractOffsets) To UBound(sgExtractOffsets)) As Long
    For ilLoop = LBound(sgExtractOffsets) To UBound(sgExtractOffsets) Step 1
        slStr = sgExtractOffsets(ilLoop)
        ilPos = InStr(1, slStr, "-", vbTextCompare)
        If ilPos <= 0 Then
            MsgBox "Invalid Offset Times specified.", vbOKOnly
            edcOffsets.SetFocus
            Exit Sub
        End If
        slStartOffset = Left$(slStr, ilPos - 1)
        slEndOffset = Mid$(slStr, ilPos + 1)
        lgExtractOffsetStart(ilLoop) = gStrLengthInTenthToLong(slStartOffset)
        lgExtractOffsetEnd(ilLoop) = gStrLengthInTenthToLong(slEndOffset)
        If lgExtractOffsetEnd(ilLoop) < lgExtractOffsetStart(ilLoop) Then
            MsgBox "Offset End Time prior to Offset Start Time not allowed.", vbOKOnly
            edcOffsets.SetFocus
            Exit Sub
        End If
    Next ilLoop
    sgExtractDays = String(7, "N")
    lmProcessedNoBytes = 0
    lbcError.Clear
    imExtracting = True
    If sgExtractType <> "T" Then
        smMsgFile = "ExtractLibrary.Txt"
    Else
        smMsgFile = "ExtractTemplate.Txt"
    End If
    smNowDate = Format(gNow(), sgShowDateForm)
    smNowTime = Format(gNow(), sgShowTimeWSecForm)
    ilRet = mOpenMsgFile()
    Screen.MousePointer = vbHourglass
    ReDim tmDayExtract(0 To 0) As SCHDEXTRACT
    For ilLoop = 0 To UBound(smFromFileName) - 1 Step 1
        plcGauge.Value = 0
        lacFiles(0).Caption = "Files: " & Trim$(Str$(ilLoop + 1)) & " of " & Trim$(Str$(UBound(smFromFileName)))
        Print #hmMsg, "Extracting File: " & smFromFileName(ilLoop)
        lacMsg.Caption = "Processing: " & smFromFileName(ilLoop)
        ilRet = mExtractRecs(smFromFileName(ilLoop))
        If imTerminate Then
            Screen.MousePointer = vbDefault
            Close hmMsg
            igReturnCallStatus = CALLCANCELLED
            Unload EngrExtract
            Exit Sub
        End If
    Next ilLoop
    mMergeDays
    mAddRecs
    Print #hmMsg, "** Extract Finished: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    Close hmMsg
    Screen.MousePointer = vbDefault
    lacMsg.Caption = "See " & smMsgFile & " for Messages"
    'cmcImport.Enabled = False
    imExtracting = False
    igReturnCallStatus = CALLDONE
    Unload EngrExtract
    Exit Sub

End Sub




Private Sub edcOffsets_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts EngrExtract
    gCenterFormModal EngrExtract
End Sub

Private Sub Form_Load()
    Dim iUpper As Integer
    
    smCurDir = CurDir
    Screen.MousePointer = vbHourglass
    mInit
    imExtracting = False
    imTerminate = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    btrDestroy hmCTE
    
    Erase smFromFileName
    Erase tmDayExtract
    Erase tmExtract
    Erase tmBothANE
    Erase tmBothNNE
    Erase tmBothRNE
    Erase tmBothBDE
    Erase tmBothASE
    Erase tmBothAudioCCE
    Erase tmBothBusCCE
    Erase tmBothCTE
    Erase tmBothFNE
    Erase tmBothMTE
    Erase tmBothSCE
    Erase tmBothEndTTE
    Erase tmBothStartTTE
    If InStr(1, smCurDir, ":") > 0 Then
        ChDrive Left$(smCurDir, 1)
        ChDir smCurDir
    End If
    Set EngrExtract = Nothing
End Sub






Private Sub mInit()
    Dim ilPos As Integer
    Dim slDrive As String
    Dim slPath As String
    Dim ilBDE As Integer
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilHour As Integer
    Dim ilRet As Integer
    
    imStartTimeChgd = False
    imEndTimeChgd = False
    lacScreen.Caption = sgExtractName
    hmCTE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCTE, "", sgDBPath & "CTE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If sgImportDirectory <> "" Then
        ilPos = InStr(sgImportDirectory, ":")
        If ilPos > 0 Then
            slDrive = Left$(sgImportDirectory, ilPos)
            slPath = Mid$(sgImportDirectory, ilPos + 1)
            If Right$(slPath, 1) = "/" Then
                slPath = Left$(slPath, Len(slPath) - 1)
            End If
            cbcDrive.Drive = slDrive
            lbcFolder.Path = slPath
            lbcFiles.fileName = slPath
        End If
    End If
    mPopulateCurrent
    mPopulateBoth
    For ilBDE = 0 To UBound(sgExtractBusNames) - 1 Step 1
        slStr = Trim$(sgExtractBusNames(ilBDE))
        ilIndex = gListBoxFind(lbcBDE, slStr)
        If ilIndex >= 0 Then
            lbcBDE.Selected(ilIndex) = True
        End If
    Next ilBDE
    hpcHours.text = gHourMap(sgExtractHours)
    For ilHour = 1 To 24 Step 1
        If Mid$(sgExtractHours, ilHour, 1) = "Y" Then
            If ilHour - 1 <= 9 Then
                tlcStartTime.text = "0" & ilHour - 1 & ":00:00"
            Else
                tlcStartTime.text = ilHour - 1 & ":00:00"
            End If
            imStartTimeChgd = True
            Exit For
        End If
    Next ilHour
    For ilHour = 24 To 1 Step -1
        If Mid$(sgExtractHours, ilHour, 1) = "Y" Then
            If ilHour <= 9 Then
                tlcEndTime.text = "0" & ilHour & ":00:00"
            Else
                tlcEndTime.text = ilHour & ":00:00"
            End If
            imEndTimeChgd = True
            Exit For
        End If
    Next ilHour
    ReDim tgExtract(0 To 0) As SCHDEXTRACT
    mSetCommands
End Sub

Private Sub mPopulateBoth()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    ilRet = gGetTypeOfRecs_ASE_AudioSource("B", smBothASEStamp, "EngrExtract-mPopulate Audio Source", tmBothASE())
    ilRet = gGetTypeOfRecs_ANE_AudioName("B", smBothANEStamp, "EngrExtract-mPopulate Audio Audio Names", tmBothANE())
    ilRet = gGetTypeOfRecs_BDE_BusDefinition("B", smBothBDEStamp, "EngrExtract-mPopulate Bus Definition", tmBothBDE())
    ilRet = gGetTypeOfRecs_CCE_ControlChar("B", "A", smBothAudioCCEStamp, "EngrExtract-mPopulate Control Character", tmBothAudioCCE())
    ilRet = gGetTypeOfRecs_CCE_ControlChar("B", "B", smBothBusCCEStamp, "EngrExtract-mPopulate Control Character", tmBothBusCCE())
    ilRet = gGetTypeOfRecs_CTE_CommtsTitle("B", "T2", smBothCTEStamp, "EngrExtract-mPopulate Title 2", tmBothCTE())
    ilRet = gGetTypeOfRecs_FNE_FollowName("B", smBothFNEStamp, "EngrExtract-mPopulate Follow", tmBothFNE())
    ilRet = gGetTypeOfRecs_MTE_MaterialType("B", smBothMTEStamp, "EngrExtract-mPopulate Material Type", tmBothMTE())
    ilRet = gGetTypeOfRecs_NNE_NetcueName("B", smBothNNEStamp, "EngrExtract-mPopulate Netcue", tmBothNNE())
    ilRet = gGetTypeOfRecs_RNE_RelayName("B", smBothRNEStamp, "EngrExtract-mPopulate Relay", tmBothRNE())
    ilRet = gGetTypeOfRecs_SCE_SilenceChar("B", smBothSCEStamp, "EngrExtract-mPopulate Silence Character", tmBothSCE())
    ilRet = gGetTypeOfRecs_TTE_TimeType("B", "E", smBothEndTTEStamp, "EngrExtract-mPopulate End Type", tmBothEndTTE())
    ilRet = gGetTypeOfRecs_TTE_TimeType("B", "S", smBothStartTTEStamp, "EngrExtract-mPopulate Start Type", tmBothStartTTE())
End Sub
Private Sub mPopulateCurrent()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrLibDef-mPopASE Audio Source", tgCurrASE())
    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrLibDef-mPopASE Audio Audio Names", tgCurrANE())
    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrLibDef-mPopBDE Bus Definition", tgCurrBDE())
    
    lbcBDE.Clear
    For ilLoop = 0 To UBound(tgCurrBDE) - 1 Step 1
        lbcBDE.AddItem Trim$(tgCurrBDE(ilLoop).sName)
        lbcBDE.ItemData(lbcBDE.NewIndex) = tgCurrBDE(ilLoop).iCode
    Next ilLoop
    lbcANE.Clear
    For ilLoop = 0 To UBound(tgCurrANE) - 1 Step 1
        lbcANE.AddItem Trim$(tgCurrANE(ilLoop).sName)
        lbcANE.ItemData(lbcANE.NewIndex) = tgCurrANE(ilLoop).iCode
    Next ilLoop
End Sub

Private Sub hpcHours_OnChange()
    mSetCommands
End Sub

Private Sub lbcFiles_Click()
    mSetCommands
End Sub

Private Sub lbcFolder_Change()
    Dim slStr As String
    slStr = lbcFolder.Path
    'If Right$(slStr, 1) <> "\" Then
    '    slStr = slStr & "\"
    'End If
    lbcFiles.fileName = slStr
    'If (igBrowserType And Not SHIFT8) = 0 Then
    '    edcBrowserFile.Text = "*.bmp"
    '    lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.bmp"
    'ElseIf (igBrowserType And Not SHIFT8) = 1 Then
    '    edcBrowserFile.Text = "*.csv"
    '    lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.csv"
    'ElseIf (igBrowserType And Not SHIFT8) = 2 Then
    '    edcBrowserFile.Text = "*.txt"
    '    lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.txt"
    'ElseIf (igBrowserType And Not SHIFT8) = 3 Then
    '    edcBrowserFile.Text = "*.rec"
    '    lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.rec"
    'ElseIf (igBrowserType And Not SHIFT8) = 4 Then
    '    edcBrowserFile.Text = "*.rt?"
    '    lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.rt?"
    'ElseIf (igBrowserType And Not SHIFT8) = 5 Then
    '    edcBrowserFile.Text = "*.ct?"
    '    lbcBrowserFile(imBrowserIndex).fileName = slStr & "*.ct?"
    'ElseIf (igBrowserType And Not SHIFT8) = 6 Then
    '    edcBrowserFile.Text = "Tape*.*"
    '    lbcBrowserFile(imBrowserIndex).fileName = slStr & "Tape*.*"
    'ElseIf (igBrowserType And Not SHIFT8) = 7 Then
    '    edcBrowserFile.Text = sgBrowseMaskFile
    '    lbcBrowserFile(imBrowserIndex).fileName = slStr & sgBrowseMaskFile
    'End If

End Sub

Public Sub mSetCommands()
    Dim ilLoop As Integer
    Dim ilFileSelected As Integer
    
    ilFileSelected = False
    For ilLoop = 0 To lbcFiles.ListCount - 1 Step 1
        If lbcFiles.Selected(ilLoop) Then
            ilFileSelected = True
            Exit For
        End If
    Next ilLoop
    If (lbcBDE.SelCount > 0) And (ilFileSelected) And (Trim$(hpcHours.text) <> "") And (Trim$(edcOffsets.text) <> "") And (imStartTimeChgd) And (imEndTimeChgd) And (tlcStartTime.text <> "") And (tlcEndTime.text <> "") Then
        cmcExtract.Enabled = True
    Else
        cmcExtract.Enabled = False
    End If
End Sub

Private Sub tlcEndTime_OnChange()
    imEndTimeChgd = True
    mSetCommands
End Sub

Private Sub tlcStartTime_OnChange()
    imStartTimeChgd = True
    mSetCommands
End Sub

Private Function mExtractField(slLine As String, ilStartCol As Integer, ilNoChars As Integer) As String
    mExtractField = ""
    If (ilStartCol <= 0) Or (ilNoChars <= 0) Or (ilStartCol > Len(slLine)) Then
        Exit Function
    End If
    mExtractField = Trim$(Mid$(slLine, ilStartCol, ilNoChars))
End Function

Private Function mCompareExtract(ilTestBus As Integer, ilTreatAasP As Integer, tlExtract1 As SCHDEXTRACT, tlExtract2 As SCHDEXTRACT) As Integer
    mCompareExtract = False
    If tlExtract1.sEventType <> tlExtract2.sEventType Then
        Exit Function
    End If
    If ilTestBus Then
        If StrComp(Trim$(tlExtract1.sBus), Trim$(tlExtract2.sBus), vbTextCompare) <> 0 Then
            Exit Function
        End If
    End If
    If StrComp(Trim$(tlExtract1.sBusCtrl), Trim$(tlExtract2.sBusCtrl), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If (tlExtract2.sEventType = "P") Or (ilTreatAasP) Then
        If (tlExtract1.lOffset <> tlExtract2.lOffset) Then
            Exit Function
        End If
    Else
        If (tlExtract1.lRunningTime <> tlExtract2.lRunningTime) Then
            Exit Function
        End If
    End If
    If StrComp(Trim$(tlExtract1.sStartType), Trim$(tlExtract2.sStartType), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sEndType), Trim$(tlExtract2.sEndType), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If (tlExtract2.sEventType = "P") Or (ilTreatAasP) Then
        If gStrLengthInTenthToLong(tlExtract1.sDuration) <> gStrLengthInTenthToLong(tlExtract2.sDuration) Then
            Exit Function
        End If
    End If
    If StrComp(Trim$(tlExtract1.sMaterialType), Trim$(tlExtract2.sMaterialType), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sAudioName), Trim$(tlExtract2.sAudioName), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sAudioID), Trim$(tlExtract2.sAudioID), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sAudioISCI), Trim$(tlExtract2.sAudioISCI), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sAudioCtrl), Trim$(tlExtract2.sAudioCtrl), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sBackupName), Trim$(tlExtract2.sBackupName), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sBackupCtrl), Trim$(tlExtract2.sBackupCtrl), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sProtName), Trim$(tlExtract2.sProtName), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sProtItemID), Trim$(tlExtract2.sProtItemID), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sProtISCI), Trim$(tlExtract2.sProtISCI), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sProtCtrl), Trim$(tlExtract2.sProtCtrl), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sRelay1), Trim$(tlExtract2.sRelay1), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sRelay2), Trim$(tlExtract2.sRelay2), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sFollow), Trim$(tlExtract2.sFollow), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If gLengthToLong(tlExtract1.sSilenceTime) <> gLengthToLong(tlExtract2.sSilenceTime) Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sSilence1), Trim$(tlExtract2.sSilence1), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sSilence2), Trim$(tlExtract2.sSilence2), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sSilence3), Trim$(tlExtract2.sSilence3), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sSilence4), Trim$(tlExtract2.sSilence4), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If (tlExtract2.sEventType = "P") Or (ilTreatAasP) Then
        If StrComp(Trim$(tlExtract1.sNetcue1), Trim$(tlExtract2.sNetcue1), vbTextCompare) <> 0 Then
            Exit Function
        End If
    Else
        If (Trim$(tlExtract1.sNetcue1) <> "") And (Trim$(tlExtract2.sNetcue1) <> "") Then
            If StrComp(Trim$(tlExtract1.sNetcue1), Trim$(tlExtract2.sNetcue1), vbTextCompare) <> 0 Then
                Exit Function
            End If
        End If
    End If
    If (tlExtract2.sEventType = "P") Or (ilTreatAasP) Then
        If StrComp(Trim$(tlExtract1.sNetcue2), Trim$(tlExtract2.sNetcue2), vbTextCompare) <> 0 Then
            Exit Function
        End If
    Else
        If (Trim$(tlExtract1.sNetcue2) <> "") And (Trim$(tlExtract2.sNetcue2) <> "") Then
            If StrComp(Trim$(tlExtract1.sNetcue2), Trim$(tlExtract2.sNetcue2), vbTextCompare) <> 0 Then
                Exit Function
            End If
        End If
    End If
    If StrComp(Trim$(tlExtract1.sTitle1), Trim$(tlExtract2.sTitle1), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sTitle2), Trim$(tlExtract2.sTitle2), vbTextCompare) <> 0 Then
        Exit Function
    End If
    If StrComp(Trim$(tlExtract1.sFixedTime), Trim$(tlExtract2.sFixedTime), vbTextCompare) <> 0 Then
        Exit Function
    End If
    mCompareExtract = True
End Function

Private Sub mMergeDays()
    Dim ilDayExtract As Integer
    Dim ilFound As Integer
    Dim ilHour As Integer
    Dim ilDay As Integer
    Dim ilExtract As Integer
    Dim ilIndex As Integer
    Dim llLinkBus As Long
    
    lacMsg.Caption = "Merging Days"
    For ilDayExtract = 0 To UBound(tmDayExtract) - 1 Step 1
        If tmDayExtract(ilDayExtract).lLinkBus <> 0 Then
            ilFound = False
            For ilExtract = 0 To UBound(tgExtract) - 1 Step 1
                If mCompareExtract(True, True, tmDayExtract(ilDayExtract), tgExtract(ilExtract)) And (tmDayExtract(ilDayExtract).sHours = tgExtract(ilExtract).sHours) Then
                    For ilDay = 1 To 7 Step 1
                        If Mid(tmDayExtract(ilDayExtract).sDays, ilDay, 1) = "Y" Then
                            Mid(tgExtract(ilExtract).sDays, ilDay, 1) = "Y"
                        End If
                    Next ilDay
                    ilFound = True
                    Exit For
                End If
            Next ilExtract
            If Not ilFound Then
                If sgExtractType <> "T" Then
                    For ilExtract = 0 To UBound(tgExtract) - 1 Step 1
                        If tgExtract(ilExtract).lLinkBus = -1 Then
                            If (mCompareExtract(False, True, tmDayExtract(ilDayExtract), tgExtract(ilExtract))) And (StrComp(Trim$(tmDayExtract(ilExtract).sBus), Trim$(tgExtract(ilExtract).sBus), vbTextCompare) <> 0) And (tmDayExtract(ilDayExtract).sHours = tgExtract(ilExtract).sHours) Then
                                tgExtract(ilExtract).lLinkBus = UBound(tgExtract)
                                Exit For
                            End If
                        End If
                    Next ilExtract
                End If
                LSet tgExtract(UBound(tgExtract)) = tmDayExtract(ilDayExtract)
                ilExtract = UBound(tgExtract)
                tgExtract(ilExtract).lLinkBus = -1
                ReDim Preserve tgExtract(0 To UBound(tgExtract) + 1) As SCHDEXTRACT
                ilIndex = tmDayExtract(ilDayExtract).lLinkBus
                Do While ilIndex <> -1
                    tgExtract(ilExtract).lLinkBus = UBound(tgExtract)
                    LSet tgExtract(UBound(tgExtract)) = tmDayExtract(ilIndex)
                    ilExtract = UBound(tgExtract)
                    tgExtract(ilExtract).lLinkBus = -1
                    ReDim Preserve tgExtract(0 To UBound(tgExtract) + 1) As SCHDEXTRACT
                    llLinkBus = tmDayExtract(ilIndex).lLinkBus
                    tmDayExtract(ilIndex).lLinkBus = 0
                    ilIndex = llLinkBus
                Loop
            End If
        End If
    Next ilDayExtract
End Sub
Private Sub mMoveAudioToRec(slAudioName As String, ilAddASE As Integer)
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    ilFound = False
    For ilLoop = LBound(tmBothANE) To UBound(tmBothANE) - 1 Step 1
        If StrComp(Trim$(tmBothANE(ilLoop).sName), slAudioName, vbTextCompare) = 0 Then
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        tmANE.iCode = 0
        tmANE.sName = slAudioName
        tmANE.sDescription = ""
        tmANE.iCceCode = 0
        tmANE.iAteCode = 0
        tmANE.sState = "A"
        tmANE.sUsedFlag = "N"
        tmANE.iVersion = 0
        tmANE.iOrigAneCode = tmANE.iCode
        tmANE.sCurrent = "Y"
        tmANE.sEnteredDate = Format(Now, sgShowDateForm)
        tmANE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
        tmANE.iUieCode = tgUIE.iCode
        tmANE.sCheckConflicts = "Y"
        tmANE.sUnused = ""
        ilRet = gPutInsert_ANE_AudioName(0, tmANE, "Audio Name-mImportFile: Insert ANE")
        LSet tmBothANE(UBound(tmBothANE)) = tmANE
        ReDim Preserve tmBothANE(LBound(tmBothANE) To UBound(tmBothANE) + 1) As ANE
        If ilAddASE Then
            mMoveASEToRec
            ilRet = gPutInsert_ASE_AudioSource(0, tmASE, "Audio Source-ImportFile: Insert ASE")
        End If
        Print #hmMsg, "Audio Name: " & slAudioName & " added"
        sgCurrANEStamp = ""
    End If
End Sub
Private Sub mMoveNetcueToRec(slNetcueName As String)
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    ilFound = False
    For ilLoop = LBound(tmBothNNE) To UBound(tmBothNNE) - 1 Step 1
        If StrComp(Trim$(tmBothNNE(ilLoop).sName), slNetcueName, vbTextCompare) = 0 Then
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        tmNNE.iCode = 0
        tmNNE.sName = slNetcueName
        tmNNE.sDescription = ""
        tmNNE.lDneCode = 0
        tmNNE.sState = "A"
        tmNNE.sUsedFlag = "N"
        tmNNE.iVersion = 0
        tmNNE.iOrigNneCode = tmNNE.iCode
        tmNNE.sCurrent = "Y"
        tmNNE.sEnteredDate = Format(Now, sgShowDateForm)
        tmNNE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
        tmNNE.iUieCode = tgUIE.iCode
        tmNNE.sUnused = ""
        ilRet = gPutInsert_NNE_NetcueName(0, tmNNE, "Netcue-mImportFile: Insert NNE")
        LSet tmBothNNE(UBound(tmBothNNE)) = tmNNE
        ReDim Preserve tmBothNNE(LBound(tmBothNNE) To UBound(tmBothNNE) + 1) As NNE
        Print #hmMsg, "Netcue Name: " & slNetcueName & " added"
        sgCurrNNEStamp = ""
    End If
End Sub

Private Sub mMoveRelayToRec(slRelayName As String)
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    ilFound = False
    For ilLoop = LBound(tmBothRNE) To UBound(tmBothRNE) - 1 Step 1
        If StrComp(Trim$(tmBothRNE(ilLoop).sName), slRelayName, vbTextCompare) = 0 Then
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        tmRNE.iCode = 0
        tmRNE.sName = slRelayName
        tmRNE.sDescription = ""
        tmRNE.sState = "A"
        tmRNE.sUsedFlag = "N"
        tmRNE.iVersion = 0
        tmRNE.iOrigRneCode = tmRNE.iCode
        tmRNE.sCurrent = "Y"
        tmRNE.sEnteredDate = Format(Now, sgShowDateForm)
        tmRNE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
        tmRNE.iUieCode = tgUIE.iCode
        tmRNE.sUnused = ""
        ilRet = gPutInsert_RNE_RelayName(0, tmRNE, "Relay-mImportFile: Insert RNE")
        LSet tmBothRNE(UBound(tmBothRNE)) = tmRNE
        ReDim Preserve tmBothRNE(LBound(tmBothRNE) To UBound(tmBothRNE) + 1) As RNE
        Print #hmMsg, "Relay Name: " & slRelayName & " added"
        sgCurrRNEStamp = ""
    End If
End Sub

Private Sub mMoveBusToRec(slBusName As String, slAudioName As String)
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    ilFound = False
    For ilLoop = LBound(tmBothBDE) To UBound(tmBothBDE) - 1 Step 1
        If StrComp(Trim$(tmBothBDE(ilLoop).sName), slBusName, vbTextCompare) = 0 Then
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        tmBDE.iCode = 0
        tmBDE.sName = slBusName
        tmBDE.sDescription = ""
        tmBDE.iCceCode = 0
        tmBDE.sChannel = ""
        tmBDE.iAseCode = 0
        slStr = slAudioName
        For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
            For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                    If StrComp(Trim$(tgCurrANE(ilANE).sName), slStr, vbTextCompare) = 0 Then
                        tmBDE.iAseCode = tgCurrASE(ilASE).iCode
                        Exit For
                    End If
                End If
            Next ilANE
            If tmBDE.iAseCode <> 0 Then
                Exit For
            End If
        Next ilASE
        tmBDE.sState = "A"
        tmBDE.sUsedFlag = "N"
        tmBDE.iVersion = 0
        tmBDE.iOrigBdeCode = tmBDE.iCode
        tmBDE.sCurrent = "Y"
        tmBDE.sEnteredDate = Format(Now, sgShowDateForm)
        tmBDE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
        tmBDE.iUieCode = tgUIE.iCode
        tmBDE.sUnused = ""
        ilRet = gPutInsert_BDE_BusDefinition(0, tmBDE, "Bus-mImportFile: Insert BDE")
        LSet tmBothBDE(UBound(tmBothBDE)) = tmBDE
        ReDim Preserve tmBothBDE(LBound(tmBothBDE) To UBound(tmBothBDE) + 1) As BDE
        Print #hmMsg, "Bus Name: " & slBusName & " added"
        sgCurrBDEStamp = ""
    End If
End Sub

Private Sub mMoveASEToRec()
    Dim ilCCE As Integer
    Dim ilANE As Integer
    Dim ilASE As Integer
    Dim slStr As String
        
    tmASE.iCode = 0
    tmASE.iPriAneCode = tmANE.iCode
    tmASE.iPriCceCode = 0
    tmASE.sDescription = ""
    tmASE.iBkupAneCode = 0
    tmASE.iBkupCceCode = 0
    tmASE.iProtAneCode = 0
    tmASE.iProtCceCode = 0
    tmASE.sState = "A"
    tmASE.sUsedFlag = "N"
    tmASE.iVersion = 0
    tmASE.iOrigAseCode = tmASE.iCode
    tmASE.sCurrent = "Y"
    tmASE.sEnteredDate = Format(Now, sgShowDateForm)
    tmASE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
    tmASE.iUieCode = tgUIE.iCode
    tmASE.sUnused = ""
End Sub

Private Sub mMoveAutoCharToRec(slAutoChar As String, slType As String)
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    ilFound = False
    If slType = "A" Then
        For ilLoop = LBound(tmBothAudioCCE) To UBound(tmBothAudioCCE) - 1 Step 1
            If StrComp(Trim$(tmBothAudioCCE(ilLoop).sAutoChar), slAutoChar, vbTextCompare) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
    Else
        For ilLoop = LBound(tmBothBusCCE) To UBound(tmBothBusCCE) - 1 Step 1
            If StrComp(Trim$(tmBothBusCCE(ilLoop).sAutoChar), slAutoChar, vbTextCompare) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
    End If
    If Not ilFound Then
        tmCCE.iCode = 0
        tmCCE.sType = slType
        tmCCE.sAutoChar = slAutoChar
        tmCCE.sDescription = ""
        tmCCE.sState = "A"
        tmCCE.sUsedFlag = "N"
        tmCCE.iVersion = 0
        tmCCE.iOrigCceCode = tmCCE.iCode
        tmCCE.sCurrent = "Y"
        tmCCE.sEnteredDate = Format(Now, sgShowDateForm)
        tmCCE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
        tmCCE.iUieCode = tgUIE.iCode
        tmCCE.sUnused = ""
        ilRet = gPutInsert_CCE_ControlChar(0, tmCCE, "Auto Char-mImportFile: Insert CCE")
        If slType = "A" Then
            LSet tmBothAudioCCE(UBound(tmBothAudioCCE)) = tmCCE
            ReDim Preserve tmBothAudioCCE(LBound(tmBothAudioCCE) To UBound(tmBothAudioCCE) + 1) As CCE
            Print #hmMsg, "Audio Auto Character: " & slAutoChar & " added"
            sgCurrAudioCCEStamp = ""
        Else
            LSet tmBothBusCCE(UBound(tmBothBusCCE)) = tmCCE
            ReDim Preserve tmBothBusCCE(LBound(tmBothBusCCE) To UBound(tmBothBusCCE) + 1) As CCE
            Print #hmMsg, "Bus Auto Character: " & slAutoChar & " added"
            sgCurrBusCCEStamp = ""
        End If
    End If
End Sub

Private Sub mMoveTitle2ToRec(slTitle2 As String)
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    ilFound = False
    For ilLoop = LBound(tmBothCTE) To UBound(tmBothCTE) - 1 Step 1
        If StrComp(Trim$(tmBothCTE(ilLoop).sComment), slTitle2, vbTextCompare) = 0 Then
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        tmCTE.lCode = 0
        tmCTE.sType = "T2"
        tmCTE.sComment = slTitle2
        tmCTE.sState = "A"
        tmCTE.sUsedFlag = "N"
        tmCTE.iVersion = 0
        tmCTE.lOrigCteCode = tmCTE.lCode
        tmCTE.sCurrent = "Y"
        tmCTE.sEnteredDate = Format$(Now, sgShowDateForm)
        tmCTE.sEnteredTime = Format$(Now, sgShowTimeWSecForm)
        tmCTE.iUieCode = tgUIE.iCode
        tmCTE.sUnused = ""
        ilRet = gPutInsert_CTE_CommtsTitle(0, tmCTE, "Title 2-mImportFile: Insert CTE", hmCTE)
        LSet tmBothCTE(UBound(tmBothCTE)) = tmCTE
        ReDim Preserve tmBothCTE(LBound(tmBothCTE) To UBound(tmBothCTE) + 1) As CTE
        Print #hmMsg, "Title 2: " & slTitle2 & " added"
        sgCurrCTEStamp = ""
    End If
End Sub
Private Sub mMoveFollowToRec(slFollow As String)
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    ilFound = False
    For ilLoop = LBound(tmBothFNE) To UBound(tmBothFNE) - 1 Step 1
        If StrComp(Trim$(tmBothFNE(ilLoop).sName), slFollow, vbTextCompare) = 0 Then
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        tmFNE.iCode = 0
        tmFNE.sName = slFollow
        tmFNE.sDescription = ""
        tmFNE.sState = "A"
        tmFNE.sUsedFlag = "N"
        tmFNE.iVersion = 0
        tmFNE.iOrigFneCode = tmFNE.iCode
        tmFNE.sCurrent = "Y"
        tmFNE.sEnteredDate = Format$(Now, sgShowDateForm)
        tmFNE.sEnteredTime = Format$(Now, sgShowTimeWSecForm)
        tmFNE.iUieCode = tgUIE.iCode
        tmFNE.sUnused = ""
        ilRet = gPutInsert_FNE_FollowName(0, tmFNE, "Follow-mImportFile: Insert FNE")
        LSet tmBothFNE(UBound(tmBothFNE)) = tmFNE
        ReDim Preserve tmBothFNE(LBound(tmBothFNE) To UBound(tmBothFNE) + 1) As FNE
        Print #hmMsg, "Follow: " & slFollow & " added"
        sgCurrFNEStamp = ""
    End If
End Sub
Private Sub mMoveMaterialToRec(slMaterial As String)
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    ilFound = False
    For ilLoop = LBound(tmBothMTE) To UBound(tmBothMTE) - 1 Step 1
        If StrComp(Trim$(tmBothMTE(ilLoop).sName), slMaterial, vbTextCompare) = 0 Then
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        tmMTE.iCode = 0
        tmMTE.sName = slMaterial
        tmMTE.sDescription = ""
        tmMTE.sState = "A"
        tmMTE.sUsedFlag = "N"
        tmMTE.iVersion = 0
        tmMTE.iOrigMteCode = tmMTE.iCode
        tmMTE.sCurrent = "Y"
        tmMTE.sEnteredDate = Format$(Now, sgShowDateForm)
        tmMTE.sEnteredTime = Format$(Now, sgShowTimeWSecForm)
        tmMTE.iUieCode = tgUIE.iCode
        tmMTE.sUnused = ""
        ilRet = gPutInsert_MTE_MaterialType(0, tmMTE, "Material Type-mImportFile: Insert MTE")
        LSet tmBothMTE(UBound(tmBothMTE)) = tmMTE
        ReDim Preserve tmBothMTE(LBound(tmBothMTE) To UBound(tmBothMTE) + 1) As MTE
        Print #hmMsg, "Material Type: " & slMaterial & " added"
        sgCurrMTEStamp = ""
    End If
End Sub

Private Sub mMoveSilenceToRec(slAutoChar As String)
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    ilFound = False
    For ilLoop = LBound(tmBothSCE) To UBound(tmBothSCE) - 1 Step 1
        If StrComp(Trim$(tmBothSCE(ilLoop).sAutoChar), slAutoChar, vbTextCompare) = 0 Then
            ilFound = True
            Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        tmSCE.iCode = 0
        tmSCE.sAutoChar = slAutoChar
        tmSCE.sDescription = ""
        tmSCE.sState = "A"
        tmSCE.sUsedFlag = "N"
        tmSCE.iVersion = 0
        tmSCE.iOrigSceCode = tmSCE.iCode
        tmSCE.sCurrent = "Y"
        tmSCE.sEnteredDate = Format$(Now, sgShowDateForm)
        tmSCE.sEnteredTime = Format$(Now, sgShowTimeWSecForm)
        tmSCE.iUieCode = tgUIE.iCode
        tmSCE.sUnused = ""
        ilRet = gPutInsert_SCE_SilenceChar(0, tmSCE, "Silence Char Type-mImportFile: Insert SCE")
        LSet tmBothSCE(UBound(tmBothSCE)) = tmSCE
        ReDim Preserve tmBothSCE(LBound(tmBothSCE) To UBound(tmBothSCE) + 1) As SCE
        Print #hmMsg, "Silence Char: " & slAutoChar & " added"
        sgCurrSCEStamp = ""
    End If
End Sub
Private Sub mMoveTimeTypeToRec(slName As String, slType As String)
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    ilFound = False
    If slType = "S" Then
        For ilLoop = LBound(tmBothStartTTE) To UBound(tmBothStartTTE) - 1 Step 1
            If StrComp(Trim$(tmBothStartTTE(ilLoop).sName), slName, vbTextCompare) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
    Else
        For ilLoop = LBound(tmBothEndTTE) To UBound(tmBothEndTTE) - 1 Step 1
            If StrComp(Trim$(tmBothEndTTE(ilLoop).sName), slName, vbTextCompare) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
    End If
    If Not ilFound Then
        tmTTE.iCode = 0
        tmTTE.sType = slType
        tmTTE.sName = slName
        tmTTE.sDescription = ""
        tmTTE.sState = "A"
        tmTTE.sUsedFlag = "N"
        tmTTE.iVersion = 0
        tmTTE.iOrigTteCode = tmTTE.iCode
        tmTTE.sCurrent = "Y"
        tmTTE.sEnteredDate = Format(Now, sgShowDateForm)
        tmTTE.sEnteredTime = Format(Now, sgShowTimeWSecForm)
        tmTTE.iUieCode = tgUIE.iCode
        tmTTE.sUnused = ""
        ilRet = gPutInsert_TTE_TimeType(0, tmTTE, "Time Type-mImportFile: Insert TTE")
        If slType = "S" Then
            LSet tmBothStartTTE(UBound(tmBothStartTTE)) = tmTTE
            ReDim Preserve tmBothStartTTE(LBound(tmBothStartTTE) To UBound(tmBothStartTTE) + 1) As TTE
            Print #hmMsg, "Start Time Type: " & slName & " added"
            sgCurrStartTTEStamp = ""
        Else
            LSet tmBothEndTTE(UBound(tmBothEndTTE)) = tmTTE
            ReDim Preserve tmBothEndTTE(LBound(tmBothEndTTE) To UBound(tmBothEndTTE) + 1) As TTE
            Print #hmMsg, "End Time Type: " & slName & " added"
            sgCurrEndTTEStamp = ""
        End If
    End If
End Sub
Private Sub mAddRecs()
    Dim ilLoop As Integer
    
    lacMsg.Caption = "Adding Undefined Items"
    For ilLoop = 0 To UBound(tgExtract) - 1 Step 1
        If Trim$(tgExtract(ilLoop).sAudioName) <> "" Then
            mMoveAudioToRec Trim$(tgExtract(ilLoop).sAudioName), True
        End If
        If Trim$(tgExtract(ilLoop).sProtName) <> "" Then
            mMoveAudioToRec Trim$(tgExtract(ilLoop).sProtName), False
        End If
        If Trim$(tgExtract(ilLoop).sBackupName) <> "" Then
            mMoveAudioToRec Trim$(tgExtract(ilLoop).sBackupName), False
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgExtract) - 1 Step 1
        If Trim$(tgExtract(ilLoop).sBus) <> "" Then
            mMoveBusToRec Trim$(tgExtract(ilLoop).sBus), Trim$(tgExtract(ilLoop).sAudioName)
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgExtract) - 1 Step 1
        If Trim$(tgExtract(ilLoop).sRelay1) <> "" Then
            mMoveRelayToRec Trim$(tgExtract(ilLoop).sRelay1)
        End If
        If Trim$(tgExtract(ilLoop).sRelay2) <> "" Then
            mMoveRelayToRec Trim$(tgExtract(ilLoop).sRelay2)
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgExtract) - 1 Step 1
        If Trim$(tgExtract(ilLoop).sNetcue1) <> "" Then
            mMoveNetcueToRec Trim$(tgExtract(ilLoop).sNetcue1)
        End If
        If Trim$(tgExtract(ilLoop).sNetcue2) <> "" Then
            mMoveNetcueToRec Trim$(tgExtract(ilLoop).sNetcue2)
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgExtract) - 1 Step 1
        If Trim$(tgExtract(ilLoop).sAudioCtrl) <> "" Then
            mMoveAutoCharToRec Trim$(tgExtract(ilLoop).sAudioCtrl), "A"
        End If
        If Trim$(tgExtract(ilLoop).sProtCtrl) <> "" Then
            mMoveAutoCharToRec Trim$(tgExtract(ilLoop).sProtCtrl), "A"
        End If
        If Trim$(tgExtract(ilLoop).sBackupCtrl) <> "" Then
            mMoveAutoCharToRec Trim$(tgExtract(ilLoop).sBackupCtrl), "A"
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgExtract) - 1 Step 1
        If Trim$(tgExtract(ilLoop).sBusCtrl) <> "" Then
            mMoveAutoCharToRec Trim$(tgExtract(ilLoop).sBusCtrl), "B"
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgExtract) - 1 Step 1
        If Trim$(tgExtract(ilLoop).sTitle2) <> "" Then
            mMoveTitle2ToRec Trim$(tgExtract(ilLoop).sTitle2)
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgExtract) - 1 Step 1
        If Trim$(tgExtract(ilLoop).sFollow) <> "" Then
            mMoveFollowToRec Trim$(tgExtract(ilLoop).sFollow)
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgExtract) - 1 Step 1
        If Trim$(tgExtract(ilLoop).sMaterialType) <> "" Then
            mMoveMaterialToRec Trim$(tgExtract(ilLoop).sMaterialType)
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgExtract) - 1 Step 1
        If Trim$(tgExtract(ilLoop).sSilence1) <> "" Then
            mMoveSilenceToRec Trim$(tgExtract(ilLoop).sSilence1)
        End If
        If Trim$(tgExtract(ilLoop).sSilence2) <> "" Then
            mMoveSilenceToRec Trim$(tgExtract(ilLoop).sSilence2)
        End If
        If Trim$(tgExtract(ilLoop).sSilence3) <> "" Then
            mMoveSilenceToRec Trim$(tgExtract(ilLoop).sSilence3)
        End If
        If Trim$(tgExtract(ilLoop).sSilence4) <> "" Then
            mMoveSilenceToRec Trim$(tgExtract(ilLoop).sSilence4)
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgExtract) - 1 Step 1
        If Trim$(tgExtract(ilLoop).sStartType) <> "" Then
            mMoveTimeTypeToRec Trim$(tgExtract(ilLoop).sStartType), "S"
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(tgExtract) - 1 Step 1
        If Trim$(tgExtract(ilLoop).sEndType) <> "" Then
            mMoveTimeTypeToRec Trim$(tgExtract(ilLoop).sEndType), "E"
        End If
    Next ilLoop
End Sub
