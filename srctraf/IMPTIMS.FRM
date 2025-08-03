VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImptIMS 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5700
   ClientLeft      =   360
   ClientTop       =   1440
   ClientWidth     =   8835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5700
   ScaleWidth      =   8835
   Begin VB.CommandButton cmcBrowse 
      Caption         =   ".."
      Height          =   285
      Left            =   6480
      TabIndex        =   4
      Top             =   1290
      Width           =   375
   End
   Begin VB.ListBox lbcDemo 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "IMPTIMS.frx":0000
      Left            =   8295
      List            =   "IMPTIMS.frx":0002
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2985
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox plcDefault 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      ScaleHeight     =   255
      ScaleWidth      =   6075
      TabIndex        =   22
      Top             =   2730
      Width           =   6075
      Begin VB.CheckBox ckcDefault 
         Caption         =   "Reallocation Book"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   4020
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox ckcDefault 
         Caption         =   "Rating Book Name"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   2070
         TabIndex        =   11
         Top             =   15
         Width           =   1935
      End
   End
   Begin VB.ComboBox cbcDPNames 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   4050
      TabIndex        =   9
      Top             =   2385
      Width           =   2730
   End
   Begin VB.ComboBox cbcDays 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   7455
      TabIndex        =   10
      Top             =   2385
      Width           =   1140
   End
   Begin VB.ComboBox cbcVehicle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   750
      TabIndex        =   8
      Top             =   2385
      Width           =   2370
   End
   Begin VB.TextBox edcBookDate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      MaxLength       =   10
      TabIndex        =   7
      Top             =   2025
      Width           =   1275
   End
   Begin VB.TextBox edcBookName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      MaxLength       =   30
      TabIndex        =   6
      Top             =   1680
      Width           =   2670
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   60
      ScaleHeight     =   270
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   30
      Width           =   1815
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5340
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7695
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4470
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6615
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5205
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcErrors 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   1680
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3450
      Visible         =   0   'False
      Width           =   5340
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   7305
      Top             =   5010
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.CommandButton cmcFrom 
      Appearance      =   0  'Flat
      Caption         =   "Import Browser.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6975
      TabIndex        =   5
      Top             =   1290
      Width           =   1725
   End
   Begin VB.PictureBox plcFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1140
      ScaleHeight     =   345
      ScaleWidth      =   5235
      TabIndex        =   3
      Top             =   1260
      Width           =   5265
      Begin VB.TextBox edcFrom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   60
         TabIndex        =   16
         Top             =   60
         Width           =   4185
      End
   End
   Begin VB.CommandButton cmcFileConv 
      Appearance      =   0  'Flat
      Caption         =   "Convert &Files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   14
      Top             =   5310
      Width           =   1830
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4905
      TabIndex        =   15
      Top             =   5310
      Width           =   1050
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      Top             =   2985
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label lacDaypart 
      Appearance      =   0  'Flat
      Caption         =   "Daypart"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3285
      TabIndex        =   20
      Top             =   2430
      Width           =   690
   End
   Begin VB.Label lacDays 
      Appearance      =   0  'Flat
      Caption         =   "Days"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6945
      TabIndex        =   21
      Top             =   2430
      Width           =   510
   End
   Begin VB.Label lacVehicle 
      Appearance      =   0  'Flat
      Caption         =   "Vehicle"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   75
      TabIndex        =   19
      Top             =   2415
      Width           =   690
   End
   Begin VB.Label lacBookDate 
      Appearance      =   0  'Flat
      Caption         =   "Book Date"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   75
      TabIndex        =   18
      Top             =   2070
      Width           =   1095
   End
   Begin VB.Label lacBookName 
      Appearance      =   0  'Flat
      Caption         =   "Book Name"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   75
      TabIndex        =   17
      Top             =   1725
      Width           =   1095
   End
   Begin VB.Label lacMsg 
      Appearance      =   0  'Flat
      Caption         =   $"IMPTIMS.frx":0004
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   120
      TabIndex        =   1
      Top             =   375
      Width           =   8520
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   5205
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacFileType 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   735
      TabIndex        =   23
      Top             =   3045
      Width           =   2190
   End
   Begin VB.Label lbcFrom 
      Appearance      =   0  'Flat
      Caption         =   "From File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   75
      TabIndex        =   2
      Top             =   1365
      Width           =   810
   End
End
Attribute VB_Name = "ImptIMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of IMPTIMS.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImptIMS.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the import Arbitron Data input screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim imTestAddStdDemo As Integer
Dim hmFrom As Integer   'From file hanle
Dim hmTo As Integer   'From file hanle
Dim hmDnf As Integer    'file handle
Dim tmDnf As DNF
Dim imDnfRecLen As Integer  'Record length
Dim hmDrf As Integer    'file handle
Dim tmDrf As DRF
Dim tmDrfPop As DRF
Dim imDrfRecLen As Integer  'Record length
Dim tmDrfSrchKey As DRFKEY0
Dim tmDrfSrchKey2 As LONGKEY0
Dim hmVef As Integer    'file handle
Dim tmVef As VEF
Dim imVefRecLen As Integer  'Record length
Dim tmVefSrchKey As INTKEY0
Dim imVefCodeInDnf() As Integer   'Array if vehicles code which are vehicles with data in dnf
Dim hmMnf As Integer    'file handle
Dim tmMnf As MNF        'Record structure
Dim imMnfRecLen As Integer  'Record length
'Dim smFieldValues(1 To 30) As String    '25 fields generated in a record
Dim smFieldValues(0 To 30) As String    '25 fields generated in a record
Dim smDays(0 To 6) As String * 1
Dim tmRdfInfo() As RDFINFO
Dim smDPStamp As String
Dim imTerminate As Integer
Dim imConverting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smNowDate As String
Dim lmNowDate As Long
Dim imNowYear As Integer
Dim smSyncDate As String
Dim smSyncTime As String
Dim imVehSelectedIndex As Integer
Dim imDaySelectedIndex As Integer
Dim imDPSelectedIndex As Integer
Dim imComboBoxIndex As Integer
Dim imBSMode As Integer
Dim imChgMode As Integer
Dim imVefCode As Integer
Dim smDataForm As String
Dim smBookForm As String
Dim bmResearchSaved As Boolean

Dim tmNameCode() As SORTCODE
Dim smNameCodeTag As String

' MsgBox parameters
'Const vbOkOnly = 0                 ' OK button only
'Const vbCritical = 16          ' Critical message
'Const vbApplicationModal = 0
'Const INDEXKEY0 = 0

'*******************************************************
'*                                                     *
'*      Procedure Name:mAddStdDemo                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add Standard Demos              *
'*                                                     *
'*******************************************************
Private Function mAddStdDemo() As Integer
'
'   ilRet = mAddStdDemo ()
'   Where:
'       ilRet (O)- True = populated; False = error
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilAddMissingOnly As Integer

    If Not imTestAddStdDemo Then
        mAddStdDemo = True
        Exit Function
    End If
    imTestAddStdDemo = False
    ReDim ilfilter(0 To 1) As Integer
    ReDim slFilter(0 To 1) As String
    ReDim ilOffSet(0 To 1) As Integer
    ilfilter(0) = CHARFILTER
    slFilter(0) = "D"
    ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
    ilfilter(1) = INTEGERFILTER
    slFilter(1) = "0"
    ilOffSet(1) = gFieldOffset("Mnf", "MnfGroupNo") '2
    lbcDemo.Clear
    ilRet = gIMoveListBox(ImptIMS, lbcDemo, tmNameCode(), smNameCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    smNameCodeTag = ""
    If lbcDemo.ListCount > 0 Then
        'Test if 20 exist
        For ilLoop = 1 To lbcDemo.ListCount - 1 Step 1
            If InStr(1, lbcDemo.List(ilLoop), "20", vbTextCompare) > 0 Then
                mAddStdDemo = True
                Exit Function
            End If
        Next ilLoop
        'Add in missing demos
        ilAddMissingOnly = True
    Else
        ilAddMissingOnly = False
    End If
    lbcDemo.Clear
    gDemoPop lbcDemo   'Get demo names
    gGetSyncDateTime slSyncDate, slSyncTime
    For ilLoop = 1 To lbcDemo.ListCount - 1 Step 1
        ilFound = False
        If ilAddMissingOnly Then
            For ilIndex = LBound(tmNameCode) To UBound(tmNameCode) - 1 Step 1
                If InStr(1, Trim$(tmNameCode(ilIndex).sKey), Trim$(lbcDemo.List(ilLoop)), vbTextCompare) > 0 Then
                    ilFound = True
                    Exit For
                End If
            Next ilIndex
        End If
        If Not ilFound Then
            tmMnf.iCode = 0
            tmMnf.sType = "D"
            tmMnf.sName = lbcDemo.List(ilLoop)
            tmMnf.sRPU = ""
            tmMnf.sUnitType = ""
            tmMnf.iMerge = 0
            tmMnf.iGroupNo = 0
            tmMnf.sCodeStn = ""
            tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmMnf.iAutoCode = tmMnf.iCode
            ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
            Do
                'tmMnfSrchKey.iCode = tmMnf.iCode
                'ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
                tmMnf.iAutoCode = tmMnf.iCode
                gPackDate slSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
                gPackTime slSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
                ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        End If
    Next ilLoop
    mAddStdDemo = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPrepassDemo                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if 16 buckets or     *
'*                      18 buckets defined             *
'*                                                     *
'*******************************************************
Private Function mPrepassDemo(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    ilRet = 0
    'On Error GoTo mPrepassDemoErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        mPrepassDemo = False
    End If
    ilRet = 0
    err.Clear
    slLine = ""
    Do
        'On Error GoTo mPrepassDemoErr:
        If EOF(hmFrom) Then
            Exit Do
        End If
        Line Input #hmFrom, slLine
        On Error GoTo 0
        ilRet = err.Number
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                If (InStr(1, Trim$(slLine), "-20", 1) > 0) Or (InStr(1, Trim$(slLine), "- 20", 1) > 0) Then
                    If (InStr(1, Trim$(slLine), "18-", 1) > 0) Or (InStr(1, Trim$(slLine), "18 -", 1) > 0) Then
                        smDataForm = "8"
                        mPrepassDemo = True
                        Close hmFrom
                        Exit Function
                    End If
                End If
                If (InStr(1, Trim$(slLine), "21-", 1) > 0) Or (InStr(1, Trim$(slLine), "21 -", 1) > 0) Then
                    If (InStr(1, Trim$(slLine), "-24", 1) > 0) Or (InStr(1, Trim$(slLine), "- 24", 1) > 0) Then
                        smDataForm = "8"
                        mPrepassDemo = True
                        Close hmFrom
                        Exit Function
                    End If
                End If
                If (InStr(1, Trim$(slLine), "21+", 1) > 0) Or (InStr(1, Trim$(slLine), "21 +", 1) > 0) Then
                    smDataForm = "8"
                    mPrepassDemo = True
                    Close hmFrom
                    Exit Function
                End If
            End If
        End If
    Loop Until ilEof
    Close hmFrom
    mPrepassDemo = True
    smDataForm = "6"
    Exit Function
'mPrepassDemoErr:
'    ilRet = Err.Number
'    Resume Next

End Function

Private Sub cbcDays_Change()
    Dim ilLoopCount As Integer
    '  imChgMode is used to avoid entering this routine multiple times
    '            if a vehicle selection change occurs during the
    '            processing of a "change"
    If imChgMode = False Then
        imChgMode = True
        ilLoopCount = 0
        Screen.MousePointer = vbHourglass  'Wait
        Do
            If ilLoopCount > 0 Then
                If cbcDays.ListIndex >= 0 Then
                    cbcDays.Text = cbcDays.List(cbcDays.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            ' If there are characters in the combobox, look ahead
            '    to see if you can find a match
            If cbcDays.Text <> "" Then
                gManLookAhead cbcDays, imBSMode, imComboBoxIndex
            End If
            'imVehSelectedIndex is used to hold the index
            '   because VB has a bug
            imDaySelectedIndex = cbcDays.ListIndex
            ' this function uses imVehSelectedIndex to find the vehicles
            '      option table vehicle index and returns imVpfIndex
        Loop While imDaySelectedIndex <> cbcDays.ListIndex
        Screen.MousePointer = vbDefault    'Default
        imChgMode = False
    End If
End Sub
Private Sub cbcDays_Click()
    imComboBoxIndex = cbcDays.ListIndex
    cbcDays_Change
End Sub
Private Sub cbcDays_GotFocus()
    lacFileType.Caption = ""
    'plcGauge.FloodPercent = 0
    plcGauge.Value = 0
    gCtrlGotFocus cbcDays
    imComboBoxIndex = cbcDays.ListIndex
    imDaySelectedIndex = imComboBoxIndex
End Sub
Private Sub cbcDays_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcDays_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcDays.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcDPNames_Change()
    Dim ilRdf As Integer
    Dim ilDay As Integer
    Dim ilRow As Integer
    Dim slStr As String
    Dim ilLoopCount As Integer
    '  imChgMode is used to avoid entering this routine multiple times
    '            if a vehicle selection change occurs during the
    '            processing of a "change"
    If imChgMode = False Then
        imChgMode = True
        ilLoopCount = 0
        Screen.MousePointer = vbHourglass  'Wait
        Do
            If ilLoopCount > 0 Then
                If cbcDPNames.ListIndex >= 0 Then
                    cbcDPNames.Text = cbcDPNames.List(cbcDPNames.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            ' If there are characters in the combobox, look ahead
            '    to see if you can find a match
            If cbcDPNames.Text <> "" Then
                gManLookAhead cbcDPNames, imBSMode, imComboBoxIndex
            End If
            'imVehSelectedIndex is used to hold the index
            '   because VB has a bug
            imDPSelectedIndex = cbcDPNames.ListIndex
            ' this function uses imVehSelectedIndex to find the vehicles
            '      option table vehicle index and returns imVpfIndex
            slStr = ""
            If imDPSelectedIndex > 0 Then
                ilRdf = tmRdfInfo(imDPSelectedIndex - 1).iRdfIndex
                For ilRow = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                    If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
                        For ilDay = 1 To 7 Step 1
                            If tgMRdf(ilRdf).sWkDays(ilRow, ilDay - 1) <> "Y" Then
                                slStr = slStr & "N"
                            Else
                                slStr = slStr & "Y"
                            End If
                        Next ilDay
                        Exit For
                    End If
                Next ilRow
            End If
            If slStr = "" Then
                slStr = "YYYYYYY"
            End If
            Select Case slStr
                Case "YYYYYNN"
                    cbcDays.ListIndex = 0
                Case "NNNNNYN"
                    cbcDays.ListIndex = 1
                Case "NNNNNNY"
                    cbcDays.ListIndex = 2
                Case "YYYYYYN"
                    cbcDays.ListIndex = 3
                Case "YYYYYYY"
                    cbcDays.ListIndex = 4
                Case "NNNNNYY"
                    cbcDays.ListIndex = 5
                Case "NYYYYYY"
                    cbcDays.ListIndex = 6
                Case "NYYYYNN"
                    cbcDays.ListIndex = 7
                Case "NNYYYYY"
                    cbcDays.ListIndex = 8
                Case "YNNNNNN"
                    cbcDays.ListIndex = 9
                Case "NYNNNNN"
                    cbcDays.ListIndex = 10
                Case "NNYNNNN"
                    cbcDays.ListIndex = 11
                Case "NNNYNNN"
                    cbcDays.ListIndex = 12
                Case "NNNNYNN"
                    cbcDays.ListIndex = 13
            End Select
        Loop While imDPSelectedIndex <> cbcDPNames.ListIndex
        Screen.MousePointer = vbDefault    'Default
        imChgMode = False
    End If
End Sub
Private Sub cbcDPNames_Click()
    imComboBoxIndex = cbcDays.ListIndex
    cbcDPNames_Change
End Sub
Private Sub cbcDPNames_DropDown()
    mDPPop
End Sub
Private Sub cbcDPNames_GotFocus()
    mDPPop
    lacFileType.Caption = ""
    'plcGauge.FloodPercent = 0
    plcGauge.Value = 0
    gCtrlGotFocus cbcDPNames
    imComboBoxIndex = cbcDPNames.ListIndex
    imDPSelectedIndex = imComboBoxIndex
End Sub
Private Sub cbcDPNames_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcDPNames_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcDPNames.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub cbcVehicle_Change()
    Dim ilLoopCount As Integer
    '  imChgMode is used to avoid entering this routine multiple times
    '            if a vehicle selection change occurs during the
    '            processing of a "change"
    If imChgMode = False Then
        imChgMode = True
        ilLoopCount = 0
        Screen.MousePointer = vbHourglass  'Wait
        Do
            If ilLoopCount > 0 Then
                If cbcVehicle.ListIndex >= 0 Then
                    cbcVehicle.Text = cbcVehicle.List(cbcVehicle.ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            ' If there are characters in the combobox, look ahead
            '    to see if you can find a match
            If cbcVehicle.Text <> "" Then
                gManLookAhead cbcVehicle, imBSMode, imComboBoxIndex
            End If
            'imVehSelectedIndex is used to hold the index
            '   because VB has a bug
            imVehSelectedIndex = cbcVehicle.ListIndex
            ' this function uses imVehSelectedIndex to find the vehicles
            '      option table vehicle index and returns imVpfIndex
        Loop While imVehSelectedIndex <> cbcVehicle.ListIndex
        Screen.MousePointer = vbDefault    'Default
        imChgMode = False
    End If
End Sub
Private Sub cbcVehicle_Click()
    imComboBoxIndex = cbcVehicle.ListIndex
    cbcVehicle_Change
End Sub
Private Sub cbcVehicle_GotFocus()
    lacFileType.Caption = ""
    'plcGauge.FloodPercent = 0
    plcGauge.Value = 0
    gCtrlGotFocus cbcVehicle
    imComboBoxIndex = cbcVehicle.ListIndex
    imVehSelectedIndex = imComboBoxIndex
End Sub
Private Sub cbcVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub cbcVehicle_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcVehicle.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cmcBrowse_Click()

    'TTP 10339 - Automation Import - use windows File Browser (to see Local Drives and mapped drives from RDP)
    CMDialogBox.DialogTitle = "Import From File"
    CMDialogBox.Filter = "CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc|Text (*.txt)|*.Txt|Print (*.prn)|*.Prn|All (*.*)|*.*"
    CMDialogBox.InitDir = sgImportPath
    CMDialogBox.DefaultExt = "csv"
    CMDialogBox.Action = 1 'Open dialog
    edcFrom.Text = CMDialogBox.fileName
    
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
End Sub

Private Sub cmcCancel_Click()
    If imConverting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcFileConv_Click()
    Dim slFromName As String
    Dim slBookName As String
    Dim slBookDate As String
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilDay As Integer
    Dim ilSY As Integer
    Dim ilEY As Integer
    lacFileType.Caption = ""
    lbcErrors.Clear
    lbcErrors.Visible = True
    slFromName = Trim$(edcFrom.Text)
    If slFromName = "" Then
        MsgBox "From Name Must be Defined", vbExclamation, "Name Error"
        edcFrom.SetFocus
        Exit Sub
    End If
    'Test if Book Name Exist
    slBookName = Trim$(edcBookName.Text)
    If slBookName = "" Then
        MsgBox "Book Name Must be Defined", vbExclamation, "Name Error"
        edcBookName.SetFocus
        Exit Sub
    End If
    slBookDate = Trim$(edcBookDate.Text)
    If slBookDate = "" Then
        MsgBox "Book Date Must be Defined", vbExclamation, "Name Error"
        edcBookDate.SetFocus
        Exit Sub
    End If
    If Not gValidDate(slBookDate) Then
        MsgBox "Invalid Date", vbExclamation, "Date Error"
        edcBookName.SetFocus
        Exit Sub
    End If
    If imVehSelectedIndex < 0 Then
        MsgBox "Vehicle Must be Defined", vbExclamation, "Name Error"
        cbcVehicle.SetFocus
        Exit Sub
    End If
    slNameCode = tgUserVehicle(imVehSelectedIndex).sKey    'lbcMster.List(ilLoop)
    ilRet = gParseItem(slNameCode, 1, "\", slName)
    ilRet = gParseItem(slName, 3, "|", slName)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    imVefCode = Val(slCode)
    If imDaySelectedIndex < 0 Then
        MsgBox "Days Must be Defined", vbExclamation, "Name Error"
        cbcDays.SetFocus
        Exit Sub
    End If
    For ilDay = 0 To 6 Step 1
        smDays(ilDay) = "N"
    Next ilDay
    Select Case imDaySelectedIndex
        Case 0  'M-F
            ilSY = 0
            ilEY = 4
        Case 1  'Sa
            ilSY = 5
            ilEY = 5
        Case 2  'Su
            ilSY = 6
            ilEY = 6
        Case 3  'M-Sa
            ilSY = 0
            ilEY = 5
        Case 4  'M-Su
            ilSY = 0
            ilEY = 6
        Case 5  'Sa-Su
            ilSY = 5
            ilEY = 6
        Case 6  'Tu-Su
            ilSY = 1
            ilEY = 6
        Case 7  'Tu-Fr
            ilSY = 1
            ilEY = 4
        Case 8  'We-Su
            ilSY = 2
            ilEY = 6
        Case 9  'Mo
            ilSY = 0
            ilEY = 0
        Case 10  'Tu
            ilSY = 1
            ilEY = 1
        Case 11  'We
            ilSY = 2
            ilEY = 2
        Case 12  'Th
            ilSY = 3
            ilEY = 3
        Case 13  'Fr
            ilSY = 4
            ilEY = 4
    End Select
    For ilDay = ilSY To ilEY Step 1
        smDays(ilDay) = "Y"
    Next ilDay
    'ilRet = mBookNameUsed(slBookName, slBookDate, imVefCode, smDays(), ilPrevDnfCode)
    ''If ilRet = 1 Then
    ''    MsgBox "Book Name Previously Used", vbExclamation, "Name Error"
    ''    edcBookName.SetFocus
    ''    Exit Sub
    ''End If
    ''If ilRet = 2 Then
    ''    Screen.MousePointer = vbDefault
    ''    ilRet = MsgBox(slBookName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
    ''    If ilRet = vbNo Then
    ''        cmcCancel.SetFocus
    ''        Exit Sub
    ''    End If
    ''    'Remove records associated with book previously imported
    ''    ilRet = mRemovePrevDnf(ilPrevDnfCode)
    ''    If Not ilRet Then
    ''        cmcCancel.SetFocus
    ''        Exit Sub
    ''    End If
    ''End If
    'If ilRet = 3 Then
    '    Screen.MousePointer = vbDefault
     '   ilRet = MsgBox(slBookName & " for " & slName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
     '   If ilRet = vbNo Then
     ''       cmcCancel.SetFocus
     '       Exit Sub
     '   End If
     '   'Remove records associated with book previously imported
     '   ilRet = mRemovePrevDnf(ilPrevDnfCode, llPrevDrfCode)
     '   If Not ilRet Then
     '       cmcCancel.SetFocus
     '       Exit Sub
      '  End If
    'End If
    slBookDate = Trim$(edcBookDate.Text)
    'Check file names
    gGetSyncDateTime smSyncDate, smSyncTime
    'plcGauge.FloodPercent = 0
    plcGauge.Value = 0
    lmProcessedNoBytes = 0
    ilRet = 0
    'On Error GoTo cmcFileConvErr:
    'hmFrom = FreeFile
    'Open slFromName For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromName, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        MsgBox "Unable to find " & slFromName, vbExclamation, "Name Error"
        edcFrom.SetFocus
        Exit Sub
    End If
    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    If Not mTestColumnTitle() Then
        Close hmFrom
        MsgBox "Column Titles (Persons & Population) Missing from " & slFromName, vbExclamation, "Name Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
    Close hmFrom
    'hmTo = FreeFile
    'Open sgDBPath & "Messages\" & "ImptCSV.Txt" For Output As hmTo
    ilRet = gFileOpen(sgDBPath & "Messages\" & "ImptCSV.Txt", "Output", hmTo)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Open " & sgDBPath & "Messages\" & "ImptCSV.Txt" & ", Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Print #hmTo, "Import CSV on " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmTo, ""
    'tmDnf.iCode = 0
    'tmDnf.sBookName = slBookName
    'gPackDate slBookDate, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
    'gPackDate smNowDate, tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
    'tmDnf.iUrfCode = tgUrf(0).iCode
    'tmDnf.sType = "I"
    'ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
    'If ilRet <> BTRV_ERR_NONE Then
    '    Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet)
    '    Close hmTo
    '    lbcErrors.AddItem "Error Adding DNF"
    '    imConverting = False
    '    mTerminate
    'End If
    ''lmCount = 0
    lacFileType.Caption = "Processing " & slFromName
    imConverting = True
    If Not mPrepassDemo(slFromName) Then
        Print #hmTo, "Import CSV failed because unable to proceess " & slFromName & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Close hmTo
        imConverting = False
        mTerminate
        Exit Sub
    End If
    If Not mConvFile(slFromName, slBookName, slBookDate, slName) Then
        Print #hmTo, "Import CSV terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Close hmTo
        imConverting = False
        mTerminate
        Exit Sub
    End If
    If ckcDefault(0).Value Or ckcDefault(1).Value Then
        Do
            tmVefSrchKey.iCode = imVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            If ckcDefault(0).Value Then
                tmVef.iDnfCode = tmDnf.iCode
            End If
            If ckcDefault(1).Value Then
                tmVef.iReallDnfCode = tmDnf.iCode
            End If
            'tmVef.iSourceID = tgUrf(0).iRemoteUserID
            'gPackDate smSyncDate, tmVef.iSyncDate(0), tmVef.iSyncDate(1)
            'gPackTime smSyncTime, tmVef.iSyncTime(0), tmVef.iSyncTime(1)
            ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        ilRet = gBinarySearchVef(tmVef.iCode)
        If ilRet <> -1 Then
            tgMVef(ilRet) = tmVef
        End If
        '11/26/17
        gFileChgdUpdate "vef.btr", False
        
    End If
    Print #hmTo, "Import CSV successfully completed on " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Close hmTo
    lacFileType.Caption = "Done"
    ilRet = mObtainBookName()
    imConverting = False
    bmResearchSaved = True
    cmcCancel.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
'cmcFileConvErr:
'    ilRet = Err.Number
'    Resume Next
End Sub
Private Sub cmcFrom_Click()
    lacFileType.Caption = ""
    'plcGauge.FloodPercent = 0
    plcGauge.Value = 0
    'CMDialogBox.DialogTitle = "From File"
    'CMDialogBox.Filter = "Blank|*.|ASC|*.Asc|Text|*.Txt|Print|*.Prn|All|*.*"
    'CMDialogBox.InitDir = Left$(sgImportPath, Len(sgImportPath) - 1)
    'CMDialogBox.Filename = ""
    'CMDialogBox.DefaultExt = ""
    'CMDialogBox.Action = 1 'Open dialog
    'edcFrom.Text = CMDialogBox.Filename
    'ChDir sgCurDir
    
    igBrowserType = 1 '1=CSV
    Browser.Show vbModal
    If igBrowserReturn = 1 Then
        edcFrom.Text = sgBrowserFile
    End If
    
    DoEvents
    edcFrom.SetFocus
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
End Sub

Private Sub edcBookDate_GotFocus()
    lacFileType.Caption = ""
    'plcGauge.FloodPercent = 0
    plcGauge.Value = 0
    gCtrlGotFocus edcBookDate
End Sub
Private Sub edcBookName_GotFocus()
    lacFileType.Caption = ""
    'plcGauge.FloodPercent = 0
    plcGauge.Value = 0
    gCtrlGotFocus edcBookName
End Sub
Private Sub edcFrom_GotFocus()
    lacFileType.Caption = ""
    'plcGauge.FloodPercent = 0
    plcGauge.Value = 0
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
        'mInitDDE
        'mSendHelpMsg "BT"
    End If
    gCtrlGotFocus edcFrom
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub Form_Activate()
    If tgSpf.sCAudPkg <> "Y" Then
        ckcDefault(1).Visible = False
    End If
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

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
    End If

End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmNameCode
    Erase tgVef
    Erase tgDnfBook
    Erase tgMnfCDemo
    Erase imVefCodeInDnf
    Erase tmRdfInfo
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    ilRet = btrClose(hmDnf)
    btrDestroy hmDnf
    
    Set ImptIMS = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mBookNameUsed                   *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test if book name used before  *
'*                                                     *
'*******************************************************
Private Function mBookNameUsed(slBookName As String, slBookDate As String, ilVefCode As Integer, slDays() As String * 1, slInfoType As String, ilRdfCode As Integer, ilCustomDemo As Integer, ilPrevDnfCode As Integer, llPrevDrfCode As Long) As Integer
    'Dim llNoRec As Long         'Number of records in Sof
    'Dim slName As String
    Dim llDate As Long
    'Dim ilExtLen As Integer
    'Dim llRecPos As Long        'Record location
    'Dim ilRet As Integer
    'Dim ilOffset As Integer
    Dim llTestDate As Long
    'Dim tlDnf As DNF
    Dim ilLoop As Integer
    Dim ilDay As Integer
    Dim ilRet As Integer
    Dim ilMatch As Integer
    Dim tlDrf As DRF

    ilPrevDnfCode = 0
    llPrevDrfCode = 0
    llTestDate = gDateValue(slBookDate)
    'ilExtLen = Len(tlDnf)  'Extract operation record size
    'llNoRec = gExtNoRec(ilExtLen)'btrRecords(hmDnf) 'Obtain number of records
    'btrExtClear hmDnf   'Clear any previous extend operation
    'ilRet = btrGetFirst(hmDnf, tlDnf, imDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'If ilRet = BTRV_ERR_END_OF_FILE Then
    '    mBookNameUsed = False
    '    Exit Function
    'End If
    'Call btrExtSetBounds(hmDnf, llNoRec, -1, "UC") 'Set extract limits (all records including first)
    'ilOffset = 0
    'ilRet = btrExtAddField(hmDnf, ilOffset, imDnfRecLen)  'Extract iCode field
    'If ilRet <> BTRV_ERR_NONE Then
    '    mBookNameUsed = False
    '    Exit Function
    'End If
    ''ilRet = btrExtGetNextExt(hmDnf)    'Extract record
    'ilRet = btrExtGetNext(hmDnf, tlDnf, ilExtLen, llRecPos)
    'If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
    '    If ilRet <> BTRV_ERR_NONE Then
    '        mBookNameUsed = False
    '        Exit Function
    '    End If
    '    ilExtLen = Len(tlDnf)  'Extract operation record size
    '    'ilRet = btrExtGetFirst(hmDnf, tlDnfExt, ilExtLen, llRecPos)
    '    Do While ilRet = BTRV_ERR_REJECT_COUNT
    '        ilRet = btrExtGetNext(hmDnf, tlDnf, ilExtLen, llRecPos)
    '    Loop
    '    Do While ilRet = BTRV_ERR_NONE
    '        gUnpackDateLong tlDnf.iBookDate(0), tlDnf.iBookDate(1), llDate
    '        If (StrComp(Trim$(tlDnf.sBookName), Trim$(slBookName), 1) = 0) And (llDate = llTestDate) Then
    '            mBookNameUsed = True
    '            ilPrevDnfCode = tlDnf.iCode
    '            Exit Function
    '        End If
    '        ilRet = btrExtGetNext(hmDnf, tlDnf, ilExtLen, llRecPos)
    '        Do While ilRet = BTRV_ERR_REJECT_COUNT
    '            ilRet = btrExtGetNext(hmDnf, tlDnf, ilExtLen, llRecPos)
    '        Loop
    '    Loop
    'End If
    'mBookNameUsed = False
    mBookNameUsed = 0   'No
    For ilLoop = LBound(tgDnfBook) To UBound(tgDnfBook) - 1 Step 1
        If StrComp(Trim$(slBookName), Trim$(tgDnfBook(ilLoop).sBookName), 1) = 0 Then
            mBookNameUsed = 1
            gUnpackDateLong tgDnfBook(ilLoop).iBookDate(0), tgDnfBook(ilLoop).iBookDate(1), llDate
            If (llDate = llTestDate) Then
                tmDnf.iCode = tgDnfBook(ilLoop).iCode
                mBookNameUsed = 2
                ilPrevDnfCode = tgDnfBook(ilLoop).iCode
                smBookForm = Trim$(tgDnfBook(ilLoop).sForm)
                If smBookForm = "" Then
                    smBookForm = "6"
                End If
                'Test if DRF exist
                tmDrfSrchKey.iDnfCode = tgDnfBook(ilLoop).iCode
                tmDrfSrchKey.sDemoDataType = "D"
                tmDrfSrchKey.iMnfSocEco = 0
                tmDrfSrchKey.iVefCode = ilVefCode
                tmDrfSrchKey.sInfoType = slInfoType '"V"
                tmDrfSrchKey.iRdfCode = ilRdfCode   '0
                ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tgDnfBook(ilLoop).iCode) And (tlDrf.sDemoDataType = "D") And (tlDrf.iVefCode = ilVefCode) And (tlDrf.sInfoType = slInfoType) And (tlDrf.iRdfCode = ilRdfCode)
                    ilMatch = True
                    For ilDay = 0 To 6 Step 1
                        If tlDrf.sDay(ilDay) <> slDays(ilDay) Then
                            ilMatch = False
                            Exit For
                        End If
                    Next ilDay
                    If ilMatch Then
                        If ilCustomDemo And (tlDrf.sDataType <> "B") Then
                            ilMatch = False
                        ElseIf Not ilCustomDemo And (tlDrf.sDataType <> "A") Then
                            ilMatch = False
                        End If
                    End If
                    If ilMatch Then
                        mBookNameUsed = 3
                        llPrevDrfCode = tlDrf.lCode
                        Exit Function
                    End If
                    ilRet = btrGetNext(hmDrf, tlDrf, imDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                Loop
            End If
            Exit Function
        End If
    Next ilLoop
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mConvFile                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert File                   *
'*                                                     *
'*******************************************************
Private Function mConvFile(slFromFile As String, slBookName As String, slBookDate As String, slName As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilHeaderFd As Integer
    Dim ilPopIndex As Integer
    Dim ilAvgQHIndex As Integer
    Dim ilAvgQH As Integer
    Dim ilPop As Integer
    Dim ilDemoGender As Integer
    Dim slDemoAge As String
    Dim ilPos As Integer
    Dim ilDay As Integer
    Dim ilLoop As Integer
    Dim ilEof As Integer
    Dim llPercent As Long
    Dim slChar As String
    Dim slSexChar As String
    Dim ilIndex As Integer
    Dim ilCustomDemo As Integer
    Dim ilPrevDnfCode As Integer
    Dim llPrevDrfCode As Long
    Dim ilRdf As Integer
    Dim ilRow As Integer
    Dim tlDrf As DRF
    ilRet = 0
    'On Error GoTo mConvFileErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        MsgBox "Open " & slFromFile & ", Error #" & Str$(ilRet), vbExclamation, "Open Error"
        edcFrom.SetFocus
        mConvFile = False
        Exit Function
    End If
    DoEvents
    If imTerminate Then
        Close hmFrom
        mTerminate
        mConvFile = False
        Exit Function
    End If
    ilHeaderFd = False
    slLine = ""
    tmDrfPop.lCode = 0
    tmDrfPop.iDnfCode = tmDnf.iCode
    tmDrfPop.sDemoDataType = "P"
    tmDrfPop.iMnfSocEco = 0
    tmDrfPop.iVefCode = 0
    tmDrfPop.sInfoType = ""
    tmDrfPop.iRdfCode = 0
    tmDrfPop.sProgCode = ""
    tmDrfPop.iStartTime(0) = 1
    tmDrfPop.iStartTime(1) = 0
    tmDrfPop.iEndTime(0) = 1
    tmDrfPop.iEndTime(1) = 0
    tmDrfPop.iStartTime2(0) = 1
    tmDrfPop.iStartTime2(1) = 0
    tmDrfPop.iEndTime2(0) = 1
    tmDrfPop.iEndTime2(1) = 0
    For ilDay = 0 To 6 Step 1
        tmDrfPop.sDay(ilDay) = "Y"
    Next ilDay
    tmDrfPop.iQHIndex = 0
    tmDrfPop.iCount = 0
    tmDrfPop.sExStdDP = "N"
    tmDrfPop.sExRpt = "N"
    tmDrfPop.sDataType = "A"
    For ilLoop = 1 To 16 Step 1
        tmDrfPop.lDemo(ilLoop - 1) = 0
    Next ilLoop
    tmDrf.lCode = 0
    tmDrf.iDnfCode = tmDnf.iCode
    tmDrf.sDemoDataType = "D"
    tmDrf.iMnfSocEco = 0
    tmDrf.iVefCode = imVefCode
    If imDPSelectedIndex <= 0 Then
        tmDrf.sInfoType = "V"
        tmDrf.iRdfCode = 0
        gPackTime "12AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    Else
        tmDrf.sInfoType = "D"
        ilRdf = tmRdfInfo(imDPSelectedIndex - 1).iRdfIndex
        tmDrf.iRdfCode = tgMRdf(ilRdf).iCode
        gPackTime "12AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
        For ilRow = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
            If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
                tmDrf.iStartTime(0) = tgMRdf(ilRdf).iStartTime(0, ilRow)
                tmDrf.iStartTime(1) = tgMRdf(ilRdf).iStartTime(1, ilRow)
                tmDrf.iEndTime(0) = tgMRdf(ilRdf).iEndTime(0, ilRow)
                tmDrf.iEndTime(1) = tgMRdf(ilRdf).iEndTime(1, ilRow)
                Exit For
            End If
        Next ilRow
    End If
    tmDrf.sProgCode = ""
    tmDrf.iStartTime2(0) = 1
    tmDrf.iStartTime2(1) = 0
    tmDrf.iEndTime2(0) = 1
    tmDrf.iEndTime2(1) = 0
    For ilDay = 0 To 6 Step 1
        tmDrf.sDay(ilDay) = smDays(ilDay)
    Next ilDay
    tmDrf.iQHIndex = 0
    tmDrf.iCount = 0
    tmDrf.sExStdDP = "N"
    tmDrf.sExRpt = "N"
    tmDrf.sDataType = "A"
    For ilLoop = 1 To 16 Step 1
        tmDrf.lDemo(ilLoop - 1) = 0
    Next ilLoop
    ilCustomDemo = False
    Do
        ilRet = 0
        err.Clear
        'On Error GoTo mConvFileErr:
        If EOF(hmFrom) Then
            Exit Do
        End If
        Line Input #hmFrom, slLine
        On Error GoTo 0
        ilRet = err.Number
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                DoEvents
                If imTerminate Then
                    Close hmFrom
                    mTerminate
                    mConvFile = False
                    Exit Function
                End If
                gParseCDFields slLine, False, smFieldValues()
                '11/10/16: Move fields from 0 to X to 1 to X +1
                For ilLoop = UBound(smFieldValues) - 1 To LBound(smFieldValues) Step -1
                    smFieldValues(ilLoop + 1) = smFieldValues(ilLoop)
                Next ilLoop
                smFieldValues(0) = ""
                'Determine field Type
                'Header Record
                '   ,,Average,,Average,,,,,,,,,,
                '   ,,Persons,,Rating,,Population,,
                'Demo Record
                '   Males   12-17,,12,,0.1,,"21,847",,,
                '   Females 12-17,,.....
                '   Men     18-24,...
                '   Women   18-24,...
                '
                If Not ilHeaderFd Then
                    For ilAvgQH = LBound(smFieldValues) + 1 To UBound(smFieldValues) Step 1
                        If StrComp(Trim$(smFieldValues(ilAvgQH)), "Persons", 1) = 0 Then
                            For ilPop = ilAvgQH + 1 To UBound(smFieldValues) Step 1
                                If StrComp(Trim$(smFieldValues(ilPop)), "Population", 1) = 0 Then
                                    ilHeaderFd = True
                                    ilAvgQHIndex = ilAvgQH
                                    ilPopIndex = ilPop
                                    Exit For
                                End If
                            Next ilPop
                            If ilHeaderFd Then
                                Exit For
                            End If
                        End If
                    Next ilAvgQH
                Else
                    ilDemoGender = -1
                    ilPos = -1
                    smFieldValues(1) = Trim$(smFieldValues(1))
                    'Test if Custom Demo
                    For ilLoop = LBound(tgMnfCDemo) To UBound(tgMnfCDemo) Step 1
                        If StrComp(Trim$(tgMnfCDemo(ilLoop).sName), smFieldValues(1), 1) = 0 Then
                            ilDemoGender = tgMnfCDemo(ilLoop).iGroupNo
                            ilCustomDemo = True
                            Exit For
                        End If
                    Next ilLoop
                    If ilDemoGender = -1 Then
                        slSexChar = UCase$(Left$(smFieldValues(1), 1))
                        If (slSexChar = "M") Or (slSexChar = "B") Or (slSexChar = "W") Or (slSexChar = "F") Or (slSexChar = "G") Then
                            'Scan for xx-yy
                            ilIndex = 2
                            Do While ilIndex < Len(smFieldValues(1))
                                slChar = Mid$(smFieldValues(1), ilIndex, 1)
                                If (slChar >= "0") And (slChar <= "9") Then
                                    ilPos = ilIndex
                                    Exit Do
                                End If
                                ilIndex = ilIndex + 1
                            Loop
                        End If
                        If ilPos > 0 Then
                            If smDataForm <> "8" Then
                                If (slSexChar = "M") Or (slSexChar = "B") Then
                                    ilDemoGender = 0
                                    slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
                                Else
                                    ilDemoGender = 8
                                    slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
                                End If
                                Select Case slDemoAge
                                    Case "12-17"
                                        ilDemoGender = ilDemoGender + 1
                                    Case "18-24"
                                        ilDemoGender = ilDemoGender + 2
                                    Case "25-34"
                                        ilDemoGender = ilDemoGender + 3
                                    Case "35-44"
                                        ilDemoGender = ilDemoGender + 4
                                    Case "45-49"
                                        ilDemoGender = ilDemoGender + 5
                                    Case "50-54"
                                        ilDemoGender = ilDemoGender + 6
                                    Case "55-64"
                                        ilDemoGender = ilDemoGender + 7
                                    Case "65+"
                                        ilDemoGender = ilDemoGender + 8
                                    Case Else
                                        ilDemoGender = -1
                                End Select
                            Else
                                If (slSexChar = "M") Or (slSexChar = "B") Then
                                    ilDemoGender = 0
                                    slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
                                Else
                                    ilDemoGender = 9
                                    slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
                                End If
                                Select Case slDemoAge
                                    Case "12-17"
                                        ilDemoGender = ilDemoGender + 1
                                    Case "18-20"
                                        ilDemoGender = ilDemoGender + 2
                                    Case "21-24"
                                        ilDemoGender = ilDemoGender + 3
                                    Case "25-34"
                                        ilDemoGender = ilDemoGender + 4
                                    Case "35-44"
                                        ilDemoGender = ilDemoGender + 5
                                    Case "45-49"
                                        ilDemoGender = ilDemoGender + 6
                                    Case "50-54"
                                        ilDemoGender = ilDemoGender + 7
                                    Case "55-64"
                                        ilDemoGender = ilDemoGender + 8
                                    Case "65+"
                                        ilDemoGender = ilDemoGender + 9
                                    Case Else
                                        ilDemoGender = -1
                                End Select
                            End If
                            If ilDemoGender >= 0 Then
                                tmDrfPop.lDemo(ilDemoGender - 1) = CLng(smFieldValues(ilPopIndex))
                                tmDrf.lDemo(ilDemoGender - 1) = CLng(smFieldValues(ilAvgQHIndex))
                            Else
                                Print #hmTo, "Unable to find Demo " & smFieldValues(1)
                                lbcErrors.AddItem "Unable to find Demo " & smFieldValues(1)
                            End If
                        End If
                    Else
                        tmDrfPop.lDemo(ilDemoGender - 1) = CLng(smFieldValues(ilPopIndex))
                        tmDrf.lDemo(ilDemoGender - 1) = CLng(smFieldValues(ilAvgQHIndex))
                    End If
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            'If plcGauge.FloodPercent <> llPercent Then
            If plcGauge.Value <> llPercent Then
                'plcGauge.FloodPercent = llPercent
                plcGauge.Value = llPercent
            End If
        End If
    Loop Until ilEof
    Close hmFrom
    'plcGauge.FloodPercent = 100
    plcGauge.Value = 100
    If ilCustomDemo Then
        tmDrfPop.sDataType = "B"
        tmDrf.sDataType = "B"
    End If
    ilRet = mBookNameUsed(slBookName, slBookDate, imVefCode, smDays(), tmDrf.sInfoType, tmDrf.iRdfCode, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
    If (ilRet = 0) Or (ilRet = 1) Then
        tmDnf.iCode = 0
        tmDnf.sBookName = slBookName
        gPackDate slBookDate, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
        gPackDate smNowDate, tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
        tmDnf.iUrfCode = tgUrf(0).iCode
        tmDnf.sType = "I"
        tmDnf.sForm = smDataForm
        tmDnf.sExactTime = "N"
        tmDnf.sSource = "C"
        tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
        tmDnf.iAutoCode = tmDnf.iCode
        ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet)
            lbcErrors.AddItem "Error Adding DNF" & Str$(ilRet)
            mConvFile = False
            Exit Function
        End If
        'If tgSpf.sRemoteUsers = "Y" Then
            Do
                'tmDnfSrchKey.iCode = tmDnf.iCode
                'ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
                tmDnf.iAutoCode = tmDnf.iCode
                gPackDate smSyncDate, tmDnf.iSyncDate(0), tmDnf.iSyncDate(1)
                gPackTime smSyncTime, tmDnf.iSyncTime(0), tmDnf.iSyncTime(1)
                ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        'End If
        ilPrevDnfCode = tmDnf.iCode
        tmDrfPop.iDnfCode = ilPrevDnfCode
        tmDrfPop.sForm = smDataForm
        tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
        tmDrfPop.lAutoCode = tmDrfPop.lCode
        ilRet = btrInsert(hmDrf, tmDrfPop, imDrfRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet)
            lbcErrors.AddItem "Error Adding DRF" & Str$(ilRet)
            mConvFile = False
            Exit Function
        End If
        'If tgSpf.sRemoteUsers = "Y" Then
            Do
                'tmDrfSrchKey2.lCode = tmDrfPop.lCode
                'ilRet = btrGetEqual(hmDrf, tmDrfPop, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
                tmDrfPop.lAutoCode = tmDrfPop.lCode
                gPackDate smSyncDate, tmDrfPop.iSyncDate(0), tmDrfPop.iSyncDate(1)
                gPackTime smSyncTime, tmDrfPop.iSyncTime(0), tmDrfPop.iSyncTime(1)
                ilRet = btrUpdate(hmDrf, tmDrfPop, imDrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        'End If
    Else
        If StrComp(smBookForm, smDataForm, vbTextCompare) <> 0 Then
            If (smDataForm = "8") Or (smBookForm = "8") Then
                mConvFile = False
                Print #hmTo, slBookName & " for " & slName & " previously defined with different format then current import data"
                lbcErrors.AddItem "Error in Book Forms" & " for " & slName
                Exit Function
            End If
        End If
        If ilRet = 3 Then
            Screen.MousePointer = vbDefault
            ilRet = MsgBox(slBookName & " for " & slName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
            If ilRet = vbNo Then
                mConvFile = False
                Exit Function
            End If
            Print #hmTo, "Replaced Demo Data (DRF)"
        End If
    End If
    tmDrf.iDnfCode = ilPrevDnfCode
    If llPrevDrfCode = 0 Then
        tmDrf.sForm = smDataForm
        tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
        tmDrf.lAutoCode = tmDrf.lCode
        ilRet = btrInsert(hmDrf, tmDrf, imDrfRecLen, INDEXKEY2)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet)
            lbcErrors.AddItem "Error Adding DRF" & Str$(ilRet)
            mConvFile = False
            Exit Function
        End If
        'If tgSpf.sRemoteUsers = "Y" Then
            Do
                'tmDrfSrchKey2.lCode = tmDrf.lCode
                'ilRet = btrGetEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmDrf.lAutoCode = tmDrf.lCode
                gPackDate smSyncDate, tmDrf.iSyncDate(0), tmDrf.iSyncDate(1)
                gPackTime smSyncTime, tmDrf.iSyncTime(0), tmDrf.iSyncTime(1)
                ilRet = btrUpdate(hmDrf, tmDrf, imDrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        'End If
    Else
        Do
            tmDrfSrchKey2.lCode = llPrevDrfCode
            ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            tmDrf.lCode = tlDrf.lCode
            gPackDate smSyncDate, tmDrf.iSyncDate(0), tmDrf.iSyncDate(1)
            gPackTime smSyncTime, tmDrf.iSyncTime(0), tmDrf.iSyncTime(1)
            ilRet = btrUpdate(hmDrf, tmDrf, imDrfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
    End If
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet)
        lbcErrors.AddItem "Error Adding DRF" & Str$(ilRet)
        mConvFile = False
        Exit Function
    End If
    mConvFile = True
    MousePointer = vbDefault
    Exit Function
'mConvFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mDaysPop                        *
'*                                                     *
'*             Created:6/4/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Day list box with     *
'*                      standard days allowed          *
'*                                                     *
'*******************************************************
Private Sub mDaysPop()
    cbcDays.Clear
    cbcDays.AddItem "Mo-Fr"
    cbcDays.AddItem "Sa"
    cbcDays.AddItem "Su"
    cbcDays.AddItem "Mo-Sa"
    cbcDays.AddItem "Mo-Su"
    cbcDays.AddItem "Sa-Su"
    cbcDays.AddItem "Tu-Su"
    cbcDays.AddItem "Tu-Fr"
    cbcDays.AddItem "We-Su"
    cbcDays.AddItem "Mo"
    cbcDays.AddItem "Tu"
    cbcDays.AddItem "We"
    cbcDays.AddItem "Th"
    cbcDays.AddItem "Fr"
    cbcDays.ListIndex = 4
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDPPop                          *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mDPPop()
    Dim ilRcf As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilTest As Integer
    Dim slSTime As String
    Dim slETime As String
    Dim ilDay As Integer
    Dim slDayName As String
    Dim slEDays As String
    Dim ilUpper As Integer
    Dim ilRow As Integer
    Dim slTime As String
    Dim slStr As String
    Dim ilTime As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim llSec As Long
    Dim llMin As Long
    Dim llHour As Long
    Dim llTime As Long
    Dim llDate As Long
    Dim llTstDate As Long
    Dim ilVefCode As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilPass As Integer
    ReDim ilDays(0 To 6) As Integer
    ReDim slDays(0 To 6) As String * 1
    If imVehSelectedIndex >= 0 Then
        slNameCode = tgUserVehicle(imVehSelectedIndex).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 3, "|", slName)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilVefCode = Val(slCode)
    Else
        ilVefCode = 0
    End If
    If smDPStamp = Trim$(Str$(ilVefCode)) Then
        Exit Sub
    End If
    smDPStamp = Trim$(Str$(ilVefCode))
    Screen.MousePointer = vbHourglass
    ReDim tmRdfInfo(0 To 0) As RDFINFO
    cbcDPNames.Clear
    'Determine Rate Card
    ilRcf = -1
    For ilPass = 0 To 1 Step 1
        For ilLoop = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
            If (tgMRcf(ilLoop).iVefCode = ilVefCode) And (tgMRcf(ilLoop).iYear = imNowYear + ilPass) And (Left$(tgMRcf(ilLoop).sName, 1) <> "~") Then
                If ilRcf = -1 Then
                    ilRcf = ilLoop
                    gUnpackDateLong tgMRcf(ilLoop).iStartDate(0), tgMRcf(ilLoop).iStartDate(1), llDate
                Else
                    gUnpackDateLong tgMRcf(ilLoop).iStartDate(0), tgMRcf(ilLoop).iStartDate(1), llTstDate
                    If ilPass = 0 Then
                        If llTstDate > llDate Then
                            llDate = llTstDate
                            ilRcf = ilLoop
                        End If
                    Else
                        If llTstDate < llDate Then
                            llDate = llTstDate
                            ilRcf = ilLoop
                        End If
                    End If
                End If
            End If
        Next ilLoop
        If ilRcf = -1 Then
            For ilLoop = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                If (tgMRcf(ilLoop).iVefCode = 0) And (tgMRcf(ilLoop).iYear = imNowYear + ilPass) And (Left$(tgMRcf(ilLoop).sName, 1) <> "~") Then
                    If ilRcf = -1 Then
                        ilRcf = ilLoop
                        gUnpackDateLong tgMRcf(ilLoop).iStartDate(0), tgMRcf(ilLoop).iStartDate(1), llDate
                    Else
                        gUnpackDateLong tgMRcf(ilLoop).iStartDate(0), tgMRcf(ilLoop).iStartDate(1), llTstDate
                        If ilPass = 0 Then
                            If llTstDate > llDate Then
                                llDate = llTstDate
                                ilRcf = ilLoop
                            End If
                        Else
                            If llTstDate < llDate Then
                                llDate = llTstDate
                                ilRcf = ilLoop
                            End If
                        End If
                    End If
                End If
            Next ilLoop
        End If
        If ilRcf <> -1 Then
            Exit For
        End If
    Next ilPass
    If (ilRcf = -1) And (LBound(tgMRcf) = UBound(tgMRcf) - 1) Then
        ilRcf = LBound(tgMRcf)
    End If
    ilUpper = 0
    If ilRcf >= 0 Then
        For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
            If (tgMRcf(ilRcf).iCode = tgMRif(llRif).iRcfCode) And (tgMRif(llRif).iVefCode = ilVefCode) Then
                'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                '    If (tgMRif(ilRif).iRdfcode = tgMRdf(ilRdf).iCode) And (tgMRdf(ilRdf).sState = "A") Then
                    ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                    If ilRdf <> -1 Then
                        If tgMRdf(ilRdf).sState <> "D" Then
                            tmRdfInfo(ilUpper).iRdfIndex = ilRdf
                            slStr = Trim$(Str$(tgMRdf(ilRdf).iSortCode))
                            Do While Len(slStr) < 5
                                slStr = "0" & slStr
                            Loop
                            slStr = slStr & Trim$(tgMRdf(ilRdf).sName)
                            ilTime = 1
                            For ilRow = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
                                    ilTime = ilRow
                                    For ilIndex = 1 To 7 Step 1
                                        If tgMRdf(ilRdf).sWkDays(ilTime, ilIndex - 1) <> "Y" Then
                                            slStr = slStr & "B"
                                        Else
                                            slStr = slStr & "A"
                                        End If
                                    Next ilIndex
                                    Exit For
                                End If
                            Next ilRow
                            llSec = tgMRdf(ilRdf).iStartTime(0, ilTime) \ 256 'Obtain seconds
                            llMin = tgMRdf(ilRdf).iStartTime(1, ilTime) And &HFF 'Obtain Minutes
                            llHour = tgMRdf(ilRdf).iStartTime(1, ilTime) \ 256 'Obtain month
                            llTime = 3600 * llHour + 60 * llMin + llSec
                            slTime = Trim$(Str$(llTime))
                            Do While (Len(slTime) < 5)
                                slTime = "0" & slTime
                            Loop
                            slStr = slStr & slTime
                            ilFound = False
                            For ilLoop = 0 To UBound(tmRdfInfo) - 1 Step 1
                                If tmRdfInfo(ilLoop).iRdfIndex = ilRdf Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                tmRdfInfo(ilUpper).sKey = slStr
                                ilUpper = ilUpper + 1
                                ReDim Preserve tmRdfInfo(0 To ilUpper) As RDFINFO
                            End If
                        End If
                '        Exit For
                    End If
                'Next ilRdf
            End If
        Next llRif
    End If
    If UBound(tmRdfInfo) - 1 > 0 Then
        ArraySortTyp fnAV(tmRdfInfo(), 0), UBound(tmRdfInfo), 0, LenB(tmRdfInfo(0)), 0, LenB(tmRdfInfo(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tmRdfInfo) - 1 Step 1
        ilRdf = tmRdfInfo(ilLoop).iRdfIndex
        If (tgMRdf(ilRdf).iLtfCode(0) <> 0) Or (tgMRdf(ilRdf).iLtfCode(1) <> 0) Or (tgMRdf(ilRdf).iLtfCode(2) <> 0) Then
            cbcDPNames.AddItem Trim$(tgMRdf(ilRdf).sName)
        Else
            For ilTest = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                If (tgMRdf(ilRdf).iStartTime(0, ilTest) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilTest) <> 0) Then
                    gUnpackTime tgMRdf(ilRdf).iStartTime(0, ilTest), tgMRdf(ilRdf).iStartTime(1, ilTest), "A", "1", slSTime
                    gUnpackTime tgMRdf(ilRdf).iEndTime(0, ilTest), tgMRdf(ilRdf).iEndTime(1, ilTest), "A", "1", slETime
                    For ilDay = 1 To 7 Step 1
                        If tgMRdf(ilRdf).sWkDays(ilTest, ilDay - 1) = "Y" Then
                            ilDays(ilDay - 1) = 1
                        Else
                            ilDays(ilDay - 1) = 0
                        End If
                        slDays(ilDay - 1) = "N"
                    Next ilDay
                    slDayName = gDayNames(ilDays(), slDays(), 2, slEDays)
                    cbcDPNames.AddItem Trim$(tgMRdf(ilRdf).sName) & " " & slSTime & "-" & slETime & " " & slDayName
                    Exit For
                End If
            Next ilTest
        End If
    Next ilLoop
    cbcDPNames.AddItem "[By Vehicle]", 0
    cbcDPNames.ListIndex = 0
    Screen.MousePointer = vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gGetRecLength                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the record length from   *
'*                     the database                    *
'*                                                     *
'*******************************************************
Private Function mGetRecLength(slFileName As String) As Integer
'
'   ilRecLen = mGetRecLength(slName)
'   Where:
'       slName (I)- Name of the file
'       ilRecLen (O)- record length within the file
'
    Dim hlFile As Integer
    Dim ilRet As Integer
    hlFile = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mGetRecLength = -ilRet
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Exit Function
    End If
    mGetRecLength = btrRecordLength(hlFile)  'Get and save record length
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
End Function
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
    Dim ilMonth As Integer

    imTerminate = False
    mParseCmmdLine
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    bmResearchSaved = False
    imTestAddStdDemo = True
    imConverting = False
    imFirstFocus = True
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0
    imChgMode = False
    imBSMode = False
    imVehSelectedIndex = -1
    imDaySelectedIndex = -1
    imDPSelectedIndex = -1
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptIMS
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    hmVef = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptIMS
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmDrf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptIMS
    On Error GoTo 0
    imDrfRecLen = Len(tmDrf)
    hmDnf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptIMS
    On Error GoTo 0
    imDnfRecLen = Len(tmDnf)
    'Populate arrays to determine if records exist
    ilRet = mAddStdDemo()
    ilRet = mObtainBookName()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Book Name Error", vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
        imTerminate = True
        Exit Sub
    End If
    mVehPop
    If imTerminate = True Then
        Exit Sub
    End If
    ilRet = gObtainRcfRifRdf()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Rate Card Error", vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
        imTerminate = True
        Exit Sub
    End If
    mDaysPop
    ilRet = mObtainCDemo()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Custom Demo Error", vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
        imTerminate = True
        Exit Sub
    End If
    smNowDate = Format$(Now, "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    gObtainMonthYear 0, smNowDate, ilMonth, imNowYear
    'If tgSpf.sCAudPkg = "Y" Then
    '    'ckcDefault(0).Value = False
    '    ckcDefault(0).Value = vbUnchecked
    '    'ckcDefault(1).Value = True
    '    ckcDefault(1).Value = vbChecked
    'End If

    'smRptTime = Format$(Now, "h:m:s AM/PM")
    'gPackTime smRptTime, tmIcf.iTime(0), tmIcf.iTime(1)
    gCenterStdAlone ImptIMS
    If mTestRecLengths() Then
        Screen.MousePointer = vbDefault
        imTerminate = True
        Exit Sub
    End If
    ilRet = 0
    On Error GoTo mInit1Err:
    'hmFrom = FreeFile
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
mInit1Err:
    ilRet = err.Number
    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainBookName                 *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgCompMnf for          *
'*                     collection                      *
'*                                                     *
'*******************************************************
Private Function mObtainBookName() As Integer
'
'   ilRet = mObtainBookName ()
'   Where:
'       tgCompMnf() (I)- MNFCOMPEXT record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim llNoRec As Long         'Number of records in Mnf
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer

    ReDim tgDnfBook(0 To 0) As DNF
    ilUpperBound = UBound(tgDnfBook)
    'ilRet = btrGetFirst(hmDnf, tgDnfBook(ilUpperBound), imDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'Do While ilRet = BTRV_ERR_NONE
    '    ilUpperBound = ilUpperBound + 1
    '    ReDim Preserve tgDnfBook(1 To ilUpperBound) As DNF
    '    ilRet = btrGetNext(hmDnf, tgDnfBook(ilUpperBound), imDnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'Loop
    ilExtLen = Len(tgDnfBook(0))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hmDnf) 'Obtain number of records
    btrExtClear hmDnf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmDnf, tgDnfBook(0), imDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        mObtainBookName = True
        Exit Function
    End If
    'Call btrExtSetBounds(hmDnf, llNoRec, -1, "UC") 'Set extract limits (all records including first)
    Call btrExtSetBounds(hmDnf, llNoRec, -1, "UC", "DNF", "") 'Set extract limits (all records including first)
    ilOffSet = 0
    ilRet = btrExtAddField(hmDnf, ilOffSet, imDnfRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainBookName = False
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hmDnf)    'Extract record
    ilRet = btrExtGetNext(hmDnf, tgDnfBook(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            mObtainBookName = False
            Exit Function
        End If
        ilExtLen = Len(tgDnfBook(0))  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmDnf, tgDnfBook(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilUpperBound = ilUpperBound + 1
            'ReDim Preserve tgDnfBook(1 To ilUpperBound) As DNF
            ReDim Preserve tgDnfBook(0 To ilUpperBound) As DNF
            ilRet = btrExtGetNext(hmDnf, tgDnfBook(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmDnf, tgDnfBook(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    mObtainBookName = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainCDemo                    *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgMnfCDemo             *
'*                                                     *
'*******************************************************
Private Function mObtainCDemo() As Integer
'
'   ilRet = mObtainCDemo ()
'   Where:
'       tgMnfCDemo() (I)- MNF record structure
'       ilRet (O)- True = populated; False = error
'
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record

    'ReDim tgMnfCDemo(1 To 1) As MNF
    'ilUpperBound = 1
    'ilRecLen = Len(tgMnfCDemo(1)) 'btrRecordLength(hmMnf)  'Get and save record length
    ReDim tgMnfCDemo(0 To 0) As MNF
    ilUpperBound = 0
    ilRecLen = Len(tgMnfCDemo(0)) 'btrRecordLength(hmMnf)  'Get and save record length
    'llNoRec = btrRecords(hmMnf) 'Obtain number of records
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlFile) 'Obtain number of records
    btrExtClear hmMnf   'Clear any previous extend operation
    'ilRet = btrGetFirst(hmMnf, tgMnfCDemo(1), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ilRet = btrGetFirst(hmMnf, tgMnfCDemo(0), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        mObtainCDemo = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mObtainCDemo = False
            Exit Function
        End If
    End If
    'Call btrExtSetBounds(hmMnf, llNoRec, 0, "UC") 'Set extract limits (all records)
    Call btrExtSetBounds(hmMnf, llNoRec, 0, "UC", "MNF", "") 'Set extract limits (all records)
    tlCharTypeBuff.sType = "D"
    ilOffSet = 2 'gFieldOffset("Mnf", "MnfType")
    ilRet = btrExtAddLogicConst(hmMnf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
    ilOffSet = 0 'gFieldOffset("Mnf", "MnfCode")
    ilRet = btrExtAddField(hmMnf, ilOffSet, ilRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainCDemo = False
        Exit Function
    End If
    ilRet = btrExtGetNext(hmMnf, tgMnfCDemo(ilUpperBound), ilRecLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            mObtainCDemo = False
            Exit Function
        End If
        'ilRecLen = Len(tgMnfCDemo(1))  'Extract operation record size
        ilRecLen = Len(tgMnfCDemo(0))  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmMnf, tgMnfCDemo(ilUpperBound), ilRecLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If tgMnfCDemo(ilUpperBound).iGroupNo > 0 Then
                ilUpperBound = ilUpperBound + 1
                'ReDim Preserve tgMnfCDemo(1 To ilUpperBound) As MNF
                ReDim Preserve tgMnfCDemo(0 To ilUpperBound) As MNF
            End If
            ilRet = btrExtGetNext(hmMnf, tgMnfCDemo(ilUpperBound), ilRecLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmMnf, tgMnfCDemo(ilUpperBound), ilRecLen, llRecPos)
            Loop
        Loop
    End If
    mObtainCDemo = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer

    slCommand = sgCommandStr  'Command$
    'slCommand = Command$
'    If StrComp(slCommand, "Debug", 1) = 0 Then
'        igStdAloneMode = True 'Switch from/to stand alone mode
'        sgCallAppName = ""
'        slStr = "Guide"
'        ilTestSystem = False
'        imShowHelpMsg = False
'    Else
'        igStdAloneMode = False  'Switch from/to stand alone mode
        ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
        If Trim$(slStr) = "" Then
            MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
            'End
            imTerminate = True
            Exit Sub
        End If
        ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
        ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
        If StrComp(slTestSystem, "Test", 1) = 0 Then
            ilTestSystem = True
        Else
            ilTestSystem = False
        End If
'        imShowHelpMsg = True
'        ilRet = gParseItem(slStr, 3, "^", slHelpSystem)    'Get application name
'        If (ilRet = CP_MSG_NONE) And (UCase$(slHelpSystem) = "NOHELP") Then
'            imShowHelpMsg = False
'        End If
        ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    'gInitStdAlone ImptIMS, slStr, ilTestSystem
    'ilRet = gParseItem(slCommand, 3, "\", slStr)    'Get call source
    igCmmCallSource = Val(slStr)
    'If igStdAloneMode Then
    '    igCmmCallSource = CALLNONE
    'End If
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
    Dim ilLoop As Integer
    

    Screen.MousePointer = vbDefault

    If bmResearchSaved Then
        If (Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES Then
            ilRet = MsgBox("Please update the Vehicle default ACT1 Lineup codes if required", vbOKOnly + vbInformation, "Warning")
        End If
    End If
    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ImptIMS
    igManUnload = NO
End Sub
Private Function mTestColumnTitle() As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilAvgQH As Integer
    Dim ilPop As Integer
    Dim ilLoop As Integer
    
    mTestColumnTitle = False
    ilRet = 0
    err.Clear
    slLine = ""
    Do
        'On Error GoTo mTestColumnTitle:
        If EOF(hmFrom) Then
            Exit Do
        End If
        Line Input #hmFrom, slLine
        On Error GoTo 0
        ilRet = err.Number
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                gParseCDFields slLine, False, smFieldValues()
                '11/10/16: Move fields from 0 to X to 1 to X +1
                For ilLoop = UBound(smFieldValues) - 1 To LBound(smFieldValues) Step -1
                    smFieldValues(ilLoop + 1) = smFieldValues(ilLoop)
                Next ilLoop
                smFieldValues(0) = ""
                'Determine field Type
                'Header Record
                '   ,,Average,,Average,,,,,,,,,,
                '   ,,Persons,,Rating,,Population,,
                'Demo Record
                '   Males   12-17,,12,,0.1,,"21,847",,,
                '   Females 12-17,,.....
                '   Men     18-24,...
                '   Women   18-24,...
                '
                For ilAvgQH = LBound(smFieldValues) + 1 To UBound(smFieldValues) Step 1
                    If StrComp(Trim$(smFieldValues(ilAvgQH)), "Persons", 1) = 0 Then
                        For ilPop = ilAvgQH + 1 To UBound(smFieldValues) Step 1
                            If StrComp(Trim$(smFieldValues(ilPop)), "Population", 1) = 0 Then
                                mTestColumnTitle = True
                                Exit Function
                            End If
                        Next ilPop
                    End If
                Next ilAvgQH
            End If
        End If
    Loop Until ilEof
    Exit Function
'mTestColumnTitle:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestRecLengths                 *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if record lengths match    *
'*                                                     *
'*******************************************************
Private Function mTestRecLengths() As Integer
    Dim ilSizeError As Integer
    Dim ilSize As Integer
    ilSizeError = False
    ilSize = mGetRecLength("Mnf.Btr")
    If ilSize <> Len(tmMnf) Then
        If ilSize > 0 Then
            MsgBox "Mnf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmMnf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Mnf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Vef.Btr")
    If ilSize <> Len(tmVef) Then
        If ilSize > 0 Then
            MsgBox "Vef size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmVef)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Vef error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Drf.Btr")
    If ilSize <> Len(tmDrf) Then
        If ilSize > 0 Then
            MsgBox "Drf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmDrf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Drf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Dnf.Btr")
    If ilSize <> Len(tmDnf) Then
        If ilSize > 0 Then
            MsgBox "Dnf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmDnf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Dnf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    mTestRecLengths = ilSizeError
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:5/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    'ilRet = gPopUserVehicleBox(Spots, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, cbcVehicle, Traffic!lbcUserVehicle)
    ilRet = gPopUserVehicleBox(ImptIMS, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, cbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ImptIMS
        On Error GoTo 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub plcDefault_Paint()
    plcDefault.CurrentX = 0
    plcDefault.CurrentY = 0
    plcDefault.Print "Set as Vehicle Default"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Import CSV Data"
End Sub
