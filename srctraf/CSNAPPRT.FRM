VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CSnapPrt 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demo Summary"
   ClientHeight    =   4200
   ClientLeft      =   1845
   ClientTop       =   2940
   ClientWidth     =   6810
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
   ScaleHeight     =   4200
   ScaleWidth      =   6810
   Begin VB.Frame plcDemo 
      Caption         =   "Demo Summary Selection"
      Height          =   1695
      Left            =   360
      TabIndex        =   17
      Top             =   1560
      Width           =   6105
      Begin VB.PictureBox plcPrice 
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
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   3060
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Width           =   3060
         Begin VB.OptionButton rbcPrice 
            Caption         =   "Yes"
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
            Left            =   1365
            TabIndex        =   20
            Top             =   0
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.OptionButton rbcPrice 
            Caption         =   "No"
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
            Left            =   2115
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   585
         End
      End
   End
   Begin VB.Frame frcCopies 
      Caption         =   "Printing"
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2520
      TabIndex        =   13
      Top             =   120
      Width           =   3900
      Begin VB.TextBox edcCopies 
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
         Left            =   1065
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "1"
         Top             =   330
         Width           =   345
      End
      Begin VB.CommandButton cmcSetup 
         Appearance      =   0  'Flat
         Caption         =   "Printer Setup"
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Top             =   825
         Width           =   1260
      End
      Begin VB.Label lacCopies 
         Appearance      =   0  'Flat
         Caption         =   "# Copies"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frcFile 
      Caption         =   "Save to File"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   3900
      Begin VB.CommandButton cmcBrowse 
         Appearance      =   0  'Flat
         Caption         =   "Browse"
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox cbcFileType 
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
         Height          =   315
         Left            =   780
         TabIndex        =   9
         Top             =   270
         Width           =   2925
      End
      Begin VB.TextBox edcFileName 
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
         Left            =   780
         TabIndex        =   8
         Top             =   615
         Width           =   2925
      End
      Begin VB.Label lacFileType 
         Appearance      =   0  'Flat
         Caption         =   "Format"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lacFileName 
         Appearance      =   0  'Flat
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   645
         Width           =   645
      End
   End
   Begin VB.Frame frcOutput 
      Caption         =   "Report Destination"
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1890
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Save to File"
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   840
         Width           =   1275
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Print"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   555
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton rbcOutput 
         Caption         =   "Display"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   3480
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
      Left            =   -30
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1620
      Width           =   120
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "&Generate"
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   3480
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   5280
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".pdf"
      Filter          =   $"CSNAPPRT.frx":0000
      FilterIndex     =   2
      FontSize        =   0
      MaxFileSize     =   256
   End
End
Attribute VB_Name = "CSnapPrt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CSNAPPRT.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CSnapPrt.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract Snap Shot Print screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim imFTSelectedIndex As Integer
Dim imComboBoxIndex As Integer
Dim imBSMode As Integer
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim hmCbf As Integer            'Contract BR file handle
Dim imCbfRecLen As Integer      'CBF record length
Dim tmCbf As CBF
Dim imTerminate As Integer
'Dim smScreenTitle As String

Private Sub cbcFileType_Change()
    If imChgMode = False Then
        imChgMode = True
        If cbcFileType.Text <> "" Then
            gManLookAhead cbcFileType, imBSMode, imComboBoxIndex
        End If
        imFTSelectedIndex = cbcFileType.ListIndex
        imChgMode = False
    End If
End Sub

Private Sub cbcFileType_Click()
    imComboBoxIndex = cbcFileType.ListIndex
    imFTSelectedIndex = cbcFileType.ListIndex
End Sub

Private Sub cbcFileType_GotFocus()
    If cbcFileType.Text = "" Then
        cbcFileType.ListIndex = 0
    End If
    imComboBoxIndex = cbcFileType.ListIndex
    gCtrlGotFocus cbcFileType
End Sub

Private Sub cbcFileType_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cbcFileType_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the left of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcFileType.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cmcBrowse_Click()
    'dan M 11/30/2011
    gAdjustCDCFilter imFTSelectedIndex, cdcSetup
    cdcSetup.flags = cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt
    cdcSetup.fileName = edcFileName.Text
    cdcSetup.InitDir = Left$(sgRptSavePath, Len(sgRptSavePath) - 1)
    cdcSetup.Action = 2    'DLG_FILE_SAVE
    edcFileName.Text = cdcSetup.fileName
    gChDrDir        '3-25-03
    ChDir sgCurDir
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcGen_Click()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilFoundCnt As Integer
    Dim slDate As String            'generation date
    Dim slTime As String            'generation time
    Dim slDetSumBoth As String      '0=det, 1 = sum, 2 = both
    Dim slInclRates As String       'Show rates on printed contract Y/N
    Dim llSeqCount As Long
    Dim slNameCode As String
    Dim slName As String
    Dim ilClf As Integer
    Dim ilUnFmt As Integer
    Dim slvalue As String
    Dim ilDay As Integer
    ReDim ilGenTime(0 To 1) As Integer
    Dim slOutputTo As String * 1
    Dim slExportIndex As String
    Dim slSaveToFileName As String
    'Dim tlBR As BRSELECTIONS        'BR options
    ReDim ilOVDays(0 To 6) As Integer  '0=Yes; 0=No; 2=Leave (Mix)

    'gObtainCorpCal             'corp calendar (if used), has been retrieved in mInitReport , stored in tgMCof
    If rbcOutput(2).Value = True And Len(edcFileName) = 0 Then
        gMsgBox "Exported file must have a name"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If imShowDiff = 2 Then
        'parm 1:  Application name^Test or Prod
        'parm 2:  user name
        'parm 3:  Generation date
        'parm 4:  Generation time
        'parm 5:  Summary #
        'parm 6:  Include Rate fields
        'Parm 7: Display(0) or Print(1) (always 1)
        ilUnFmt = UNFMTDEFAULT
        slDetSumBoth = Trim$(str$(imDBTotals))
        If rbcPrice(0).Value Then
            slInclRates = "Y"
        Else
            slInclRates = "N"
        End If
        slDate = Format$(gNow(), "m/d/yy")
        gPackDate slDate, tmCbf.iGenDate(0), tmCbf.iGenDate(1)
        slTime = Format$(gNow(), "h:mm:ssAM/PM")
        gPackTime slTime, ilGenTime(0), ilGenTime(1)
        gUnpackTimeLong ilGenTime(0), ilGenTime(1), False, tmCbf.lGenTime
        tmCbf.lContrNo = tgChfCntr.lCntrNo
        If Contract.lbcAdvt.ListIndex >= 1 Then
            slNameCode = tmCntAdvertiser(Contract.lbcAdvt.ListIndex - 1).sKey  'Traffic!lbcAdvertiser.List(lbcAdvt.ListIndex - 1)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            tmCbf.sSurvey = slName
        Else
            tmCbf.sSurvey = ""
        End If
        tmCbf.sProduct = tgChfCntr.sProduct
        If Contract.lbcDBSocEco.ListIndex >= 1 Then
            tmCbf.sDemos = Trim$(Contract.edcDBSocEco.Text) & " " & Trim$(Contract.edcDBDemo.Text)
        Else
            tmCbf.sDemos = Trim$(Contract.edcDBDemo.Text)
        End If
        ilFoundCnt = False
        Select Case imDBTotals
            Case 5
                For ilLoop = imLB1Or2 To UBound(smLnSumShow, 2) - 1 Step 1
                    If Trim$(smLnSumShow(12, ilLoop)) = "T" Then
                        ilFoundCnt = True
                        Exit For
                    End If
                Next ilLoop
            Case 6
                ilFoundCnt = False
            Case 7
                ilFoundCnt = False
            Case 8
                For ilLoop = imLB1Or2 To UBound(smVSumShow, 2) - 1 Step 1
                    If Trim$(smVSumShow(12, ilLoop)) = "T" Then
                        ilFoundCnt = True
                        Exit For
                    End If
                Next ilLoop
            Case 9
                For ilLoop = imLB1Or2 To UBound(smDPSumShow, 2) - 1 Step 1
                    If Trim$(smDPSumShow(12, ilLoop)) = "T" Then
                        ilFoundCnt = True
                        Exit For
                    End If
                Next ilLoop
        End Select
        llSeqCount = 0
        Select Case imDBTotals
            Case 5
                For ilLoop = imLB1Or2 To UBound(smLnSumShow, 2) - 1 Step 1
                    If (Not ilFoundCnt) Or ((Trim$(smLnSumShow(12, ilLoop)) = "T") And (ilFoundCnt)) Then
                        For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                            If Val(smLnSumShow(1, ilLoop)) = tgClfCntr(ilClf).ClfRec.iLine Then
                                llSeqCount = llSeqCount + 1
                                tmCbf.lLineNo = llSeqCount
                                tmCbf.sSortField1 = smLnSave(1, ilClf + 1)
                                For ilDay = 0 To 6 Step 1
                                    ilOVDays(ilDay) = imLnSave(12 + ilDay, ilClf + 1)
                                Next ilDay
                                tmCbf.sSortField2 = mSetDPShowName(ilClf + 1, -1, ilOVDays())
                                tmCbf.lGRP = gStrDecToLong(smLnSumShow(4, ilLoop), 1)
                                tmCbf.iPctDist = Val(smLnSumShow(5, ilLoop))
'                                If tgSpf.sSAudData <> "H" Then
'                                    tmCbf.lGrImp = Val(smLnSumShow(6, ilLoop))
'                                Else
'                                    tmCbf.lGrImp = gStrDecToLong(smLnSumShow(6, ilLoop), 1)
'                                End If
                                If tgSpf.sSAudData = "H" Then
                                    tmCbf.lGrImp = gStrDecToLong(smLnSumShow(6, ilLoop), 1)
                                ElseIf tgSpf.sSAudData = "N" Then
                                    tmCbf.lGrImp = gStrDecToLong(smLnSumShow(6, ilLoop), 2)
                                ElseIf tgSpf.sSAudData = "U" Then
                                    tmCbf.lGrImp = gStrDecToLong(smLnSumShow(6, ilLoop), 3)
                                Else
                                    tmCbf.lGrImp = Val(smLnSumShow(6, ilLoop))
                                End If
                                tmCbf.iPctTrade = Val(smLnSumShow(7, ilLoop))
                                tmCbf.lCPM = gStrDecToLong(smLnSumShow(8, ilLoop), 2)
                                tmCbf.lCPP = Val(smLnSumShow(9, ilLoop))
                                slStr = smLnSumShow(10, ilLoop)
                                gUnformatStr slStr, ilUnFmt, slvalue
                                tmCbf.lRate = gStrDecToLong(slvalue, 2)
                                slStr = smLnSumShow(11, ilLoop)
                                gUnformatStr slStr, ilUnFmt, slvalue
                                tmCbf.lExtra4Byte = gStrDecToLong(slvalue, 2)
                                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                                Exit For
                            End If
                        Next ilClf
                    End If
                Next ilLoop
            Case 6
                For ilLoop = imLB1Or2 To UBound(smWkSumShow, 2) - 1 Step 1
                    llSeqCount = llSeqCount + 1
                    tmCbf.lLineNo = llSeqCount
                    tmCbf.sSortField1 = smWkSumShow(1, ilLoop)
                    tmCbf.sSortField2 = ""
                    tmCbf.lGRP = gStrDecToLong(smWkSumShow(2, ilLoop), 1)
                    tmCbf.iPctDist = Val(smWkSumShow(3, ilLoop))
'                    If tgSpf.sSAudData <> "H" Then
'                        tmCbf.lGrImp = Val(smWkSumShow(4, ilLoop))
'                    Else
'                        tmCbf.lGrImp = gStrDecToLong(smWkSumShow(4, ilLoop), 1)
'                    End If
                    If tgSpf.sSAudData = "H" Then
                        tmCbf.lGrImp = gStrDecToLong(smWkSumShow(4, ilLoop), 1)
                    ElseIf tgSpf.sSAudData = "N" Then
                        tmCbf.lGrImp = gStrDecToLong(smWkSumShow(4, ilLoop), 2)
                    ElseIf tgSpf.sSAudData = "U" Then
                        tmCbf.lGrImp = gStrDecToLong(smWkSumShow(4, ilLoop), 3)
                    Else
                        tmCbf.lGrImp = Val(smWkSumShow(4, ilLoop))
                    End If
                    tmCbf.iPctTrade = Val(smWkSumShow(5, ilLoop))
                    tmCbf.lCPM = gStrDecToLong(smWkSumShow(6, ilLoop), 2)
                    tmCbf.lCPP = Val(smWkSumShow(7, ilLoop))
                    slStr = smWkSumShow(8, ilLoop)
                    gUnformatStr slStr, ilUnFmt, slvalue
                    tmCbf.lRate = gStrDecToLong(slvalue, 2)
                    slStr = smWkSumShow(9, ilLoop)
                    gUnformatStr slStr, ilUnFmt, slvalue
                    tmCbf.lExtra4Byte = gStrDecToLong(slvalue, 2)
                    ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                Next ilLoop
            Case 7
                For ilLoop = imLB1Or2 To UBound(smQSumShow, 2) - 1 Step 1
                    llSeqCount = llSeqCount + 1
                    tmCbf.lLineNo = llSeqCount
                    tmCbf.sSortField1 = smQSumShow(1, ilLoop)
                    tmCbf.sSortField2 = ""
                    'GRP
                    tmCbf.lGRP = gStrDecToLong(smQSumShow(2, ilLoop), 1)
                    '%GRP
                    tmCbf.iPctDist = Val(smQSumShow(3, ilLoop))
                    'GrImp
'                    If tgSpf.sSAudData <> "H" Then
'                        tmCbf.lGrImp = Val(smQSumShow(4, ilLoop))
'                    Else
'                        tmCbf.lGrImp = gStrDecToLong(smQSumShow(4, ilLoop), 1)
'                    End If
                    If tgSpf.sSAudData = "H" Then
                        tmCbf.lGrImp = gStrDecToLong(smQSumShow(4, ilLoop), 1)
                    ElseIf tgSpf.sSAudData = "N" Then
                        tmCbf.lGrImp = gStrDecToLong(smQSumShow(4, ilLoop), 2)
                    ElseIf tgSpf.sSAudData = "U" Then
                        tmCbf.lGrImp = gStrDecToLong(smQSumShow(4, ilLoop), 3)
                    Else
                        tmCbf.lGrImp = Val(smQSumShow(4, ilLoop))
                    End If
                    '% GrImp
                    tmCbf.iPctTrade = Val(smQSumShow(5, ilLoop))
                    'CPM
                    tmCbf.lCPM = gStrDecToLong(smQSumShow(6, ilLoop), 2)
                    'CPP
                    tmCbf.lCPP = Val(smQSumShow(7, ilLoop))
                    'Avg Price
                    slStr = smQSumShow(8, ilLoop)
                    gUnformatStr slStr, ilUnFmt, slvalue
                    tmCbf.lRate = gStrDecToLong(slvalue, 2)
                    'Total Cost
                    slStr = smQSumShow(9, ilLoop)
                    gUnformatStr slStr, ilUnFmt, slvalue
                    tmCbf.lExtra4Byte = gStrDecToLong(slvalue, 2)
                    ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                Next ilLoop
            Case 8
                For ilLoop = imLB1Or2 To UBound(smVSumShow, 2) - 1 Step 1
                    If (Not ilFoundCnt) Or ((Trim$(smVSumShow(12, ilLoop)) = "T") And (ilFoundCnt)) Then
                        llSeqCount = llSeqCount + 1
                        tmCbf.lLineNo = llSeqCount
                        tmCbf.sSortField1 = smVSumShow(1, ilLoop)
                        tmCbf.sSortField2 = ""
                        'GRP
                        tmCbf.lGRP = gStrDecToLong(smVSumShow(2, ilLoop), 1)
                        '%GRP
                        tmCbf.iPctDist = Val(smVSumShow(3, ilLoop))
                        'GrImp
'                        If tgSpf.sSAudData <> "H" Then
'                            tmCbf.lGrImp = Val(smVSumShow(4, ilLoop))
'                        Else
'                            tmCbf.lGrImp = gStrDecToLong(smVSumShow(4, ilLoop), 1)
'                        End If
                        If tgSpf.sSAudData = "H" Then
                            tmCbf.lGrImp = gStrDecToLong(smVSumShow(4, ilLoop), 1)
                        ElseIf tgSpf.sSAudData = "N" Then
                            tmCbf.lGrImp = gStrDecToLong(smVSumShow(4, ilLoop), 2)
                        ElseIf tgSpf.sSAudData = "U" Then
                            tmCbf.lGrImp = gStrDecToLong(smVSumShow(4, ilLoop), 3)
                        Else
                            tmCbf.lGrImp = Val(smVSumShow(4, ilLoop))
                        End If
                        '% GrImp
                        tmCbf.iPctTrade = Val(smVSumShow(5, ilLoop))
                        'CPM
                        tmCbf.lCPM = gStrDecToLong(smVSumShow(6, ilLoop), 2)
                        'CPP
                        tmCbf.lCPP = Val(smVSumShow(7, ilLoop))
                        'Avg Price
                        slStr = smVSumShow(8, ilLoop)
                        gUnformatStr slStr, ilUnFmt, slvalue
                        tmCbf.lRate = gStrDecToLong(slvalue, 2)
                        'Total Cost
                        slStr = smVSumShow(9, ilLoop)
                        gUnformatStr slStr, ilUnFmt, slvalue
                        tmCbf.lExtra4Byte = gStrDecToLong(slvalue, 2)
                        ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                    End If
                Next ilLoop
            Case 9
                For ilLoop = imLB1Or2 To UBound(smDPSumShow, 2) - 1 Step 1
                    If (Not ilFoundCnt) Or ((Trim$(smDPSumShow(12, ilLoop)) = "T") And (ilFoundCnt)) Then
                        llSeqCount = llSeqCount + 1
                        tmCbf.lLineNo = llSeqCount
                        tmCbf.sSortField1 = smDPSumShow(1, ilLoop)
                        tmCbf.sSortField2 = ""
                        'GRP
                        tmCbf.lGRP = gStrDecToLong(smDPSumShow(2, ilLoop), 1)
                        '%GRP
                        tmCbf.iPctDist = Val(smDPSumShow(3, ilLoop))
                        'GrImp
'                        If tgSpf.sSAudData <> "H" Then
'                            tmCbf.lGrImp = Val(smDPSumShow(4, ilLoop))
'                        Else
'                            tmCbf.lGrImp = gStrDecToLong(smDPSumShow(4, ilLoop), 1)
'                        End If
                        If tgSpf.sSAudData = "H" Then
                            tmCbf.lGrImp = gStrDecToLong(smDPSumShow(4, ilLoop), 1)
                        ElseIf tgSpf.sSAudData = "N" Then
                            tmCbf.lGrImp = gStrDecToLong(smDPSumShow(4, ilLoop), 2)
                        ElseIf tgSpf.sSAudData = "U" Then
                            tmCbf.lGrImp = gStrDecToLong(smDPSumShow(4, ilLoop), 3)
                        Else
                            tmCbf.lGrImp = Val(smDPSumShow(4, ilLoop))
                        End If
                        '% GrImp
                        tmCbf.iPctTrade = Val(smDPSumShow(5, ilLoop))
                        'CPM
                        tmCbf.lCPM = gStrDecToLong(smDPSumShow(6, ilLoop), 2)
                        'CPP
                        tmCbf.lCPP = Val(smDPSumShow(7, ilLoop))
                        'Avg Price
                        slStr = smDPSumShow(8, ilLoop)
                        gUnformatStr slStr, ilUnFmt, slvalue
                        tmCbf.lRate = gStrDecToLong(slvalue, 2)
                        'Total Cost
                        slStr = smDPSumShow(9, ilLoop)
                        gUnformatStr slStr, ilUnFmt, slvalue
                        tmCbf.lExtra4Byte = gStrDecToLong(slvalue, 2)
                        ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                    End If
                Next ilLoop
        End Select
        tmCbf.lLineNo = 99999
        tmCbf.sSortField1 = ""
        tmCbf.sSortField2 = ""
        'GRP
        tmCbf.lGRP = gStrDecToLong(smTSumShow(1), 1)
        '%GRP
        tmCbf.iPctDist = Val(smTSumShow(2))
        'GrImp
'        If tgSpf.sSAudData <> "H" Then
'            tmCbf.lGrImp = Val(smTSumShow(3))
'        Else
'            tmCbf.lGrImp = gStrDecToLong(smTSumShow(3), 1)
'        End If
        If tgSpf.sSAudData = "H" Then
            tmCbf.lGrImp = gStrDecToLong(smTSumShow(3), 1)
        ElseIf tgSpf.sSAudData = "N" Then
            tmCbf.lGrImp = gStrDecToLong(smTSumShow(3), 2)
        ElseIf tgSpf.sSAudData = "U" Then
            tmCbf.lGrImp = gStrDecToLong(smTSumShow(3), 3)
        Else
            tmCbf.lGrImp = Val(smTSumShow(3))
        End If
        '% GrImp
        tmCbf.iPctTrade = Val(smTSumShow(4))
        'CPM
        tmCbf.lCPM = gStrDecToLong(smTSumShow(5), 2)
        'CPP
        tmCbf.lCPP = Val(smTSumShow(6))
        'Avg Price
        slStr = smTSumShow(7)
        gUnformatStr slStr, ilUnFmt, slvalue
        tmCbf.lRate = gStrDecToLong(slvalue, 2)
        'Total Cost
        slStr = smTSumShow(8)
        gUnformatStr slStr, ilUnFmt, slvalue
        tmCbf.lExtra4Byte = gStrDecToLong(slvalue, 2)
        ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
        If igTestSystem Then
            slStr = "Traffic^Test^NOHELP\" & sgUserName
        Else
            slStr = "Traffic^Prod^NOHELP\" & sgUserName
        End If

        slSaveToFileName = ""
        If rbcOutput(0).Value = True Then
            slOutputTo = "0"
        ElseIf rbcOutput(1).Value = True Then
            slOutputTo = "1"
        Else
            slOutputTo = "2"
            slExportIndex = Trim$(str(imFTSelectedIndex))
            slSaveToFileName = Trim$(edcFileName)      'save to location
        End If
        slStr = slStr & "\" & slDate & "\" & slTime & "\" & slDetSumBoth & "\" & slInclRates & "\" & slOutputTo & "\" & slExportIndex & "\" & slSaveToFileName
        'ilShell = Shell(sgExePath & "RptSelDB.Exe " & slStr, 1)
        'While GetModuleUsage(ilShell) > 0
        '    ilRet = DoEvents()
        'Wend
        sgCommandStr = slStr
        RptSelDB.Show vbModal
    Else
        'tlBR.sSnapshot = smSnapshot     'flag to indicate snapshot rather than from reports
        'tlBR.iShowProof = False     'force to not show Proof mode
        'tlBR.iCorpOrStd = True      'forc standard monthly summaries
        'tlBR.iPrintables = False
        'tlBR.iDiffOnly = False
        'tlBR.iThisCntMod = False
        '
        'If ckcDiff.Value Then                  'presence of code indicates to do differences
        '    tlBR.iDiffOnly = True
        '    tlBR.iThisCntMod = True
        'End If
        '
        ' If CSnapPrt!ckcRating.Value Then
        '    tlBR.iShowRates = True
        ' End If
        ' If CSnapPrt!ckcResearch.Value Then
        '    tlBR.iShowResearch = True
        ' End If
        ' slInclProof = "N"
        ' If CSnapPrt!ckcProof.Value Then
        '    tlBR.iShowProof = True
        ' End If
        ' If CSnapPrt!rbcShow(0).Value Or CSnapPrt!rbcShow(2).Value Then         'is selection Detail or Both?, if so set include detail
        '    tlBR.iDetail = True
        ' End If
        ' If CSnapPrt!rbcShow(1).Value Or CSnapPrt!rbcShow(2).Value Then         'is selection summary or both, if so set include summary
        '    tlBR.iSummary = True
        ' End If
        ' If CSnapPrt!rbcDemo(0).Value Then          'Primary demos (vs all)
         '   tlBR.iAllDemos = False
        ' Else
        '    tlBR.iAllDemos = True
        ' End If
        '
        ' tlBR.iWhichSort = 0         'force to sort by advt, contr#
        '
        'llCurrCode = tgChfCntr.lCode
        'llPrevCode = 0
        'ilPrevSpots = 0
        'llPrevGross = 0
        'ilCurrTotalWks = 0
        'ilCurrAirWks = 0
        '
        'tlBR.iPropOrOrder = False               'assume order (HOGN)
        'If tgChfCntr.sStatus = "W" Or tgChfCntr.sStatus = "D" Or tgChfCntr.sStatus = "C" Or tgChfCntr.sStatus = "I" Then
        '    tlBR.iPropOrOrder = True
        'End If
        '
        'If tlBR.sSnapshot = " " Then                    'user is taking a snapshot of a version just viewed
        '    ilChfRecLen = Len(tmChf)
        '    tmChfSrchKey1.lCntrNo = tgChfCntr.lCntrNo
        '    tmChfSrchKey1.iCntRevNo = 32000
        '    tmChfSrchKey1.iPropVer = 32000
        '    ilRet = btrGetGreaterOrEqual(hmCtrPrtChf, tmChf, ilChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        '    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChfCntr.lCntrNo) And (tgChfCntr.lCode <> tmChf.lCode) Then
        '        tlBR.sSnapshot = "P"            'previous version
        '    End If
        'End If
        '
        'If tlBR.iDiffOnly Then                            'differences for Mods
        '    ilChfRecLen = Len(tmChf)
        '    tmChfSrchKey1.lCntrNo = tgChfCntr.lCntrNo
        '    tmChfSrchKey1.iCntRevNo = tgChfCntr.iCntRevNo
        '    tmChfSrchKey1.iPropVer = 32000
        '    ilRet = btrGetGreaterOrEqual(hmCtrPrtChf, tmChf, ilChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        '    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChfCntr.lCntrNo) And (tmChf.sSchStatus <> "F")
        '        ilRet = btrGetNext(hmCtrPrtChf, tmChf, ilChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        '    Loop
        '    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChfCntr.lCntrNo) And (tmChf.sSchStatus = "F") Then
        '        'tmchf has original contract
        '        'differences only on this contract required - compare 2 contracts and return differences in the contract
        '
        '        'Save the current tgChfCntr, tgclf, & tgcff so that when the snapshot returns,
        '        'the original image on the screen will be restored since the Differences mode
        '        'fakes out a contract for the differences only into the 'tg' arrays
        '        tmTChf = tgChf
        '        ReDim tmTClf(0 To 0) As CLFLIST
        '        For ilLoop = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        '            tmTClf(ilLoop) = tgClfCntr(ilLoop)
        '            ReDim Preserve tmTClf(0 To UBound(tmTClf) + 1)  As CLFLIST
        '        Next ilLoop
        '        ReDim tmTCff(0 To 0) As CFFLIST
        '        For ilLoop = LBound(tgCffCntr) To UBound(tgCffCntr) - 1 Step 1
        '            tmTCff(ilLoop) = tgCffCntr(ilLoop)
        '            ReDim Preserve tmTCff(0 To UBound(tmTCff) + 1)  As CFFLIST
        '        Next ilLoop
        '        llPrevCode = tmChf.lCode
        '        gGenDiff CONTRACTSJOB, hmCtrPrtChf, hmCtrPrtClf, hmCtrPrtCff, llCurrCode, llPrevCode, ilPrevSpots, llPrevGross, ilCurrTotalWks, ilCurrAirWks
        '    Else
        '        MsgBox "Previous version does not exist", vbOkOnly + vbExclamation, "Traffic"
        '        mTerminate
        '        Exit Sub
        '    End If
        'End If
        ''ProcessFlag 0 = open files, then process cnt, 1 = close files and exit
        'tlBR.iPrevSpots = ilPrevSpots
        'tlBR.lPrevGross = llPrevGross
        'tlBR.iCurrTotWks = ilCurrTotalWks
        'tlBR.iCurrAirWks = ilCurrAirWks
        '
        'slDate = Format$(gNow(), "m/d/yy")
        'gPackDate slDate, tlBR.iGenDate(0), tlBR.iGenDate(1)
        'slTime = Format$(gNow(), "h:mm:ssAM/PM")
        'gPackTime slTime, tlBR.iGenTime(0), tlBR.iGenTime(1)
        '
        'gProcessBR tlBR, 0, CONTRACTSJOB            'open files
        'gProcessBR tlBR, 1, CONTRACTSJOB           'process one contract
        'gProcessBR tlBR, 2, CONTRACTSJOB           'close files
        ''Call Report print module
        ''If (Not igStdAloneMode) And (imShowHelpMsg) Then
        ''    If igTestSystem Then
        ''        slStr = "Traffic^Test\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode)) & "\" & slStartDate & "\" & Trim$(Str$(tgSel(ilLoop).lEndDate - tgSel(ilLoop).lStartDate + 1)) & "\" & slStartTime & "\" & slEndTime & "\" & Trim$(Str$(tmVef.iCode)) & "\" & Trim$(Str$(tgSel(ilLoop).iZone)) & "\" & slOutput
        ''    Else
        ''        slStr = "Traffic^Prod\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType)) & "\" & Trim$(Str$(tgUrf(0).iCode)) & "\" & slStartDate & "\" & Trim$(Str$(tgSel(ilLoop).lEndDate - tgSel(ilLoop).lStartDate + 1)) & "\" & slStartTime & "\" & slEndTime & "\" & Trim$(Str$(tmVef.iCode)) & "\" & Trim$(Str$(tgSel(ilLoop).iZone)) & "\" & slOutput
        ''    End If
        ''Else
        '    If igTestSystem Then
        '        slStr = "Traffic^Test^NOHELP\" & sgUserName
        '    Else
        '        slStr = "Traffic^Prod^NOHELP\" & sgUserName
        '    End If
        ''End If
        '
        ''Set up the command line string to pass to print the contract
        'If tlBR.iDetail And tlBR.iSummary Then
        '    slDetSumBoth = "2"
        'ElseIf tlBR.iDetail And Not tlBR.iSummary Then
        '    slDetSumBoth = "0"              'detail only
        'Else
        '    slDetSumBoth = "1"              'summary only
        'End If
        'If tlBR.iShowRates Then
        '    slInclRates = "Y"
        'Else
        '    slInclRates = "N"
        'End If
        'If tlBR.iShowResearch Then
        '    slInclResearch = "Y"
        'Else
        '    slInclResearch = "N"
        'End If
        'If tlBR.iShowProof Then
        '    slInclProof = "Y"
        'Else
        '    slInclProof = "N"
        'End If
        ''parm 1:  Application name^Test or Prod
        ''parm 2:  user name
        ''parm 3:  Generation date
        ''parm 4:  Generation time
        ''parm 5:  Detail(0), Summary(1), Both (2)
        ''parm 6: Includes Rates Y/N
        ''Parm 7: include Research Y/n
        ''Parm 8: include Proof
        ''Parm 9: Display(0) or Print(1) (always 1)
        'slStr = slStr & "\" & slDate & "\" & slTime & "\" & slDetSumBoth & "\" & slInclRates & "\" & slInclResearch & "\" & slInclProof & "\1"
        'ilShell = Shell(sgExePath & "RptSelBR.Exe " & slStr, 1)
        'While GetModuleUsage(ilShell) > 0
        '    ilRet = DoEvents()
        'Wend
        '
        ''Restore the original contents of the screen
        'If tlBR.iDiffOnly Then
        '    tgChf = tmTChf
        '    ReDim tgClfCntr(0 To 0) As CLFLIST
        '    For ilLoop = LBound(tmTClf) To UBound(tmTClf) - 1 Step 1
        '        tgClfCntr(ilLoop) = tmTClf(ilLoop)
        '        ReDim Preserve tgClfCntr(0 To UBound(tgClfCntr) + 1)  As CLFLIST
        '    Next ilLoop
        '    ReDim tgCffCntr(0 To 0) As CFFLIST
        '    For ilLoop = LBound(tmTCff) To UBound(tmTCff) - 1 Step 1
        '        tgCffCntr(ilLoop) = tmTCff(ilLoop)
        '        ReDim Preserve tgCffCntr(0 To UBound(tgCffCntr) + 1)  As CFFLIST
        '    Next ilLoop
        'End If
    End If
    Screen.MousePointer = vbDefault
    mTerminate
End Sub

Private Sub cmcSetup_Click()
    cdcSetup.flags = cdlPDPrintSetup
    cdcSetup.ShowPrinter
End Sub

Private Sub edcCopies_GotFocus()
    gCtrlGotFocus edcCopies
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
        cmcCancel_Click
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
    Dim ilRet As Integer
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imTerminate = False
    If (imShowDiff = 0) Or (imShowDiff = 1) Then
        'plcBR.Visible = True
        'plcSum.Visible = False
        'If tgChfCntr.lCntrNo <> 0 Then
        '    plcScreen.Caption = "BR Snapshot for" & Str$(tgChfCntr.lCntrNo)
        'Else
        '    plcScreen.Caption = "BR Snapshot"
        'End If
        'If imShowDiff = 0 Then
        '    ckcDiff.Visible = False
        'End If
    Else
        'plcBR.Visible = False
        'plcSum.Visible = True
        If tgChfCntr.lCntrNo <> 0 Then
            'smScreenTitle = "Demo Bar Snapshot for" & Str$(tgChfCntr.lCntrNo)
            CSnapPrt.Caption = "Demo Bar Snapshot for" & str$(tgChfCntr.lCntrNo)
        Else
            'smScreenTitle = "Demo Bar Snapshot"
            CSnapPrt.Caption = "Demo Bar Snapshot"
        End If
        ilRet = gPopExportTypes(cbcFileType)     'populate export types for crystal
        'plcScreen.Cls
        'plcScreen.CurrentX = 0
        'plcScreen.CurrentY = 0
        'plcScreen.Print smScreenTitle
        hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Cbf.Btr)", Contract
        On Error GoTo 0
        imCbfRecLen = Len(tmCbf)
    End If
    'CSnapPrt.Height = cmcGen.Top + 5 * cmcGen.Height / 3
    gCenterModalForm CSnapPrt
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
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
    Dim ilRet As Integer
    Screen.MousePointer = vbDefault
    Unload CSnapPrt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    If imShowDiff = 2 Then
        ilRet = btrClose(hmCbf)
        btrDestroy hmCbf
    End If
    
    Set CSnapPrt = Nothing   'Remove data segment

End Sub

Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcPrice_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcPrice_Paint()
    plcPrice.CurrentX = 0
    plcPrice.CurrentY = 0
    plcPrice.Print "Include Prices"
End Sub

Private Sub rbcOutput_Click(Index As Integer)
'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcOutput(Index).Value
    'End of coded added
    If rbcOutput(Index).Value Then
        Select Case Index
            Case 0  'Display
                frcFile.Enabled = False
                frcCopies.Visible = False   'Print Box
                frcFile.Visible = False     'Save to File Box
                frcCopies.Enabled = False
            Case 1  'Print
                frcFile.Visible = False
                frcFile.Enabled = False
                frcCopies.Enabled = True
                frcCopies.Visible = True
            Case 2  'File
                frcCopies.Visible = False
                frcCopies.Enabled = False
                frcFile.Enabled = True
                frcFile.Visible = True
        End Select
    End If
End Sub
