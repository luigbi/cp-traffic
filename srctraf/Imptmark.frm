VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImptMark 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6330
   ClientLeft      =   2670
   ClientTop       =   1800
   ClientWidth     =   9405
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
   ScaleHeight     =   6330
   ScaleWidth      =   9405
   Begin VB.CommandButton cmcBrowse 
      Caption         =   ".."
      Height          =   285
      Left            =   6840
      TabIndex        =   45
      Top             =   1605
      Width           =   375
   End
   Begin VB.ComboBox cbcBookName 
      Height          =   315
      Index           =   1
      Left            =   4680
      TabIndex        =   21
      Top             =   2595
      Visible         =   0   'False
      Width           =   4230
   End
   Begin VB.ComboBox cbcBookName 
      Height          =   315
      Index           =   0
      Left            =   195
      TabIndex        =   19
      Top             =   2595
      Visible         =   0   'False
      Width           =   4230
   End
   Begin VB.CommandButton cmcMoveToDaypart 
      Appearance      =   0  'Flat
      Caption         =   "Move to &Daypart"
      Enabled         =   0   'False
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
      Left            =   6255
      TabIndex        =   44
      Top             =   5490
      Width           =   2115
   End
   Begin VB.CommandButton cmcMoveToExact 
      Appearance      =   0  'Flat
      Caption         =   "Move to Exact Times"
      Enabled         =   0   'False
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
      Left            =   3945
      TabIndex        =   43
      Top             =   5490
      Width           =   2115
   End
   Begin VB.ListBox lbcDemo 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "Imptmark.frx":0000
      Left            =   8130
      List            =   "Imptmark.frx":0002
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3615
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox plcNew 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   255
      ScaleHeight     =   195
      ScaleWidth      =   4905
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2070
      Width           =   4905
      Begin VB.OptionButton rbcNew 
         Caption         =   "Correct or Extend Existing Book"
         Height          =   195
         Index           =   1
         Left            =   1785
         TabIndex        =   16
         Top             =   0
         Width           =   3045
      End
      Begin VB.OptionButton rbcNew 
         Caption         =   "Add New Book"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   15
         Top             =   0
         Value           =   -1  'True
         Width           =   1620
      End
   End
   Begin VB.PictureBox plcDefault 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   135
      ScaleHeight     =   255
      ScaleWidth      =   6045
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3510
      Width           =   6045
      Begin VB.CheckBox ckcDefault 
         Caption         =   "Rating Book Name"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   2070
         TabIndex        =   26
         Top             =   15
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox ckcDefault 
         Caption         =   "Reallocation Book"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   4020
         TabIndex        =   27
         Top             =   15
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Frame frcBookInfo 
      Height          =   1395
      Left            =   75
      TabIndex        =   13
      Top             =   2055
      Width           =   9045
      Begin VB.TextBox edcBookName 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   120
         MaxLength       =   30
         TabIndex        =   18
         Top             =   540
         Width           =   4230
      End
      Begin VB.TextBox edcBookDate 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1365
         MaxLength       =   10
         TabIndex        =   24
         Top             =   1005
         Width           =   1275
      End
      Begin VB.TextBox edcBookName 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   4605
         MaxLength       =   30
         TabIndex        =   22
         Top             =   540
         Width           =   4230
      End
      Begin VB.Label lacBookName 
         Appearance      =   0  'Flat
         Caption         =   "Book Name: Daypart:"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   300
         Width           =   2250
      End
      Begin VB.Label lacBookDate 
         Appearance      =   0  'Flat
         Caption         =   "Book Date"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   23
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label lacBookName 
         Appearance      =   0  'Flat
         Caption         =   "Exact Times:"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   4605
         TabIndex        =   20
         Top             =   300
         Width           =   2040
      End
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   225
      Left            =   3135
      TabIndex        =   41
      Top             =   4035
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox plcPCForm 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   90
      ScaleHeight     =   240
      ScaleWidth      =   7710
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1305
      Width           =   7710
      Begin VB.OptionButton rbcPCForm 
         Caption         =   "Lineup Analysis(for Stations)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   4920
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   3210
      End
      Begin VB.OptionButton rbcPCForm 
         Caption         =   "Demo Summary(for Networks)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1860
         TabIndex        =   8
         Top             =   0
         Value           =   -1  'True
         Width           =   3315
      End
      Begin VB.OptionButton rbcPCForm 
         Caption         =   "Station with Population by Market"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   5280
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   0
         Width           =   3210
      End
   End
   Begin VB.PictureBox plcImportFrom 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   90
      ScaleHeight     =   240
      ScaleWidth      =   5400
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   5400
      Begin VB.OptionButton rbcImportFrom 
         Caption         =   "On Line System"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3135
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   1680
      End
      Begin VB.OptionButton rbcImportFrom 
         Caption         =   "PC System"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1860
         TabIndex        =   4
         Top             =   0
         Value           =   -1  'True
         Width           =   1230
      End
   End
   Begin VB.ComboBox cbcDPNames 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   4095
      TabIndex        =   31
      Top             =   5895
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.ComboBox cbcDays 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   7500
      TabIndex        =   33
      Top             =   5895
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox cbcVehicle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   795
      TabIndex        =   29
      Top             =   5895
      Visible         =   0   'False
      Width           =   2370
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
      ScaleWidth      =   2580
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   2580
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4695
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6615
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5385
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcErrors 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2040
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4485
      Visible         =   0   'False
      Width           =   5340
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   8100
      Top             =   5145
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
      Left            =   7380
      TabIndex        =   12
      Top             =   1605
      Width           =   1725
   End
   Begin VB.PictureBox plcFrom 
      Height          =   375
      Left            =   915
      ScaleHeight     =   315
      ScaleWidth      =   5760
      TabIndex        =   10
      Top             =   1575
      Width           =   5820
      Begin VB.TextBox edcFrom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   15
         TabIndex        =   11
         Top             =   0
         Width           =   5730
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
      Left            =   840
      TabIndex        =   35
      Top             =   5490
      Width           =   1710
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
      Left            =   2775
      TabIndex        =   36
      Top             =   5490
      Width           =   930
   End
   Begin VB.Label lacMsg 
      Appearance      =   0  'Flat
      Caption         =   $"Imptmark.frx":0004
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   2850
      TabIndex        =   2
      Top             =   0
      Width           =   6405
   End
   Begin VB.Label lacDaypart 
      Appearance      =   0  'Flat
      Caption         =   "Daypart"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3330
      TabIndex        =   30
      Top             =   5940
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lacDays 
      Appearance      =   0  'Flat
      Caption         =   "Days"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6990
      TabIndex        =   32
      Top             =   5940
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lacVehicle 
      Appearance      =   0  'Flat
      Caption         =   "Vehicle"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   28
      Top             =   5925
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   5400
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacFileType 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   735
      TabIndex        =   37
      Top             =   4080
      Width           =   2190
   End
   Begin VB.Label lbcFrom 
      Appearance      =   0  'Flat
      Caption         =   "From File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      TabIndex        =   9
      Top             =   1680
      Width           =   810
   End
End
Attribute VB_Name = "ImptMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Imptmark.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mPCStnConvFile_PopByUSA                                                               *
'******************************************************************************************

' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImptMark.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the import Act1 Data input screen code
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
Dim tmDnfSrchKey As INTKEY0
Dim hmDrf As Integer    'file handle
Dim tmDrfInfo() As DRFINFO
Dim tmDrfPop As DRF
Dim tmDrf As DRF
Dim imDrfRecLen As Integer  'Record length
Dim tmDrfSrchKey As DRFKEY0
Dim tmDrfSrchKey2 As LONGKEY0
Dim hmDpf As Integer    'file handle
Dim tmDpf As DPF
Dim imDpfRecLen As Integer  'Record length
Dim tmDpfSrchKey1 As DPFKEY1
'Dim hmDsf As Integer    'file handle
'Dim tmDsf As DSF
'Dim imDsfRecLen As Integer  'Record length
Dim hmVef As Integer    'file handle
Dim tmVef As VEF
Dim imVefRecLen As Integer  'Record length
Dim tmVefSrchKey As INTKEY0
Dim imVefCodeImpt() As Integer   'Array if vehicles code which match station code
Dim hmMnf As Integer    'file handle
Dim tmMnf As MNF        'Record structure
Dim imMnfRecLen As Integer  'Record length
'7627
'Dim smFieldValues(1 To 1000) As String    '25 fields generated in a record
Dim smFieldValues(0 To 1000) As String    '25 fields generated in a record
'Dim smSvFields(1 To 1000) As String    '25 fields generated in a record
Dim smSvFields(0 To 1000) As String    '25 fields generated in a record
'Dim smFieldValues(1 To 60) As String    '25 fields generated in a record
'Dim smSvFields(1 To 60) As String    '25 fields generated in a record
Dim smDays(0 To 6) As String * 1
Dim tmRdfInfo() As RDFINFO
Dim smDPStamp As String
Dim imMatchRdfCode() As Integer
Dim imTerminate As Integer
Dim imConverting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smNowDate As String
Dim lmNowDate As Long
Dim imNowYear As Integer
Dim smSyncDate As String
Dim smSyncTime As String
Dim imBNSelectedIndex As Integer
Dim imVehSelectedIndex As Integer
Dim imDaySelectedIndex As Integer
Dim imDPSelectedIndex As Integer
Dim imComboBoxIndex As Integer
Dim imBSMode As Integer
Dim imChgMode As Integer
Dim imVefCode As Integer
'Dim smSvLine(1 To 10) As String
Dim smSvLine(0 To 9) As String
Dim smFileNames() As String
Dim smUnfdStations() As String
Dim smBookForm As String
Dim smDataForm As String

Dim bmResearchSaved As Boolean
Dim bmBooksSaved As Boolean

Dim tmNameCode() As SORTCODE
Dim smNameCodeTag As String
'7582
Dim tmExtraDemo() As DPFINFO
Dim lmDpfWithoutPop() As DPFAndColumn
' MsgBox parameters
'Const vbOkOnly = 0                 ' OK button only
'Const vbCritical = 16          ' Critical message
'Const vbApplicationModal = 0
'Const INDEXKEY0 = 0
Private Type DPFAndColumn
    lDpfCode As Long
    iColumn As Integer
End Type

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
    ilRet = gIMoveListBox(ImptMark, lbcDemo, tmNameCode(), smNameCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
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

Private Sub cbcBookName_Change(Index As Integer)
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
                If cbcBookName(Index).ListIndex >= 0 Then
                    cbcBookName(Index).Text = cbcBookName(Index).List(cbcBookName(Index).ListIndex)
                End If
            End If
            ilLoopCount = ilLoopCount + 1
            ' If there are characters in the combobox, look ahead
            '    to see if you can find a match
            If cbcBookName(Index).Text <> "" Then
                gManLookAhead cbcBookName(Index), imBSMode, imComboBoxIndex
            End If
            'imVehSelectedIndex is used to hold the index
            '   because VB has a bug
            imBNSelectedIndex = cbcBookName(Index).ListIndex
            ' this function uses imVehSelectedIndex to find the vehicles
            '      option table vehicle index and returns imVpfIndex
        Loop While imBNSelectedIndex <> cbcBookName(Index).ListIndex
        Screen.MousePointer = vbDefault    'Default
        imChgMode = False
    End If
    mSetCommands Index
End Sub

Private Sub cbcBookName_GotFocus(Index As Integer)
    gCtrlGotFocus cbcBookName(Index)
    imComboBoxIndex = cbcBookName(Index).ListIndex
    imBNSelectedIndex = imComboBoxIndex
End Sub

Private Sub cbcBookName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcBookName_KeyPress(Index As Integer, KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcBookName(Index).SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

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
                            'If tgMRdf(ilRdf).sWkDays(ilRow, ilDay) <> "Y" Then
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
    CMDialogBox.InitDir = sgImportPath
    If rbcImportFrom(0).Value Then
        CMDialogBox.Filter = "CSV (*.csv)|*.csv|All (*.*)|*.*|Blank (*.)|*.|ASC (*.asc)|*.Asc|Text (*.txt)|*.Txt|Print (*.prn)|*.Prn"
        CMDialogBox.DefaultExt = "CSV (*.csv)"
        CMDialogBox.DialogTitle = "Import From CSV File"
    Else
        CMDialogBox.Filter = "All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc|Text (*.txt)|*.Txt|Print (*.prn)|*.Prn"
        CMDialogBox.DefaultExt = "All (*.*)"
        CMDialogBox.DialogTitle = "Import From Text File"
    End If
    CMDialogBox.Action = 1 'Open dialog
    edcFrom.Text = CMDialogBox.fileName
    
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
    If rbcPCForm(2).Value Then
        edcBookName(0).Text = mParseForBookName(edcFrom.Text)
    End If
End Sub

Private Sub cmcCancel_Click()
    If imConverting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcCancel_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
End Sub
Private Sub cmcFileConv_Click()
    Dim slFromName As String
    Dim slDPBookName As String
    Dim slETBookName As String
    Dim slBookDate As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim llTestDate As Long
    Dim llDate As Long
    Dim blIsNewVersion As Boolean
    
    lacFileType.Caption = ""
    cmcCancel.Caption = "&Cancel"
    lbcErrors.Clear
    lbcErrors.Visible = True
    slFromName = Trim$(edcFrom.Text)
    If slFromName = "" Then
        MsgBox "From Name Must be Defined", vbExclamation, "Name Error"
        edcFrom.SetFocus
        Exit Sub
    End If
    'Test if Book Name Exist
    slDPBookName = ""
    slETBookName = ""
    '7582
    If Not rbcPCForm(1).Value Then
        If rbcNew(0).Value Then
            slDPBookName = Trim$(edcBookName(0).Text)
            slETBookName = Trim$(edcBookName(1).Text)
        Else
            If cbcBookName(0).ListIndex >= 0 Then
                slDPBookName = cbcBookName(0).List(cbcBookName(0).ListIndex)
                ilPos = InStrRev(slDPBookName, ": ", -1, vbTextCompare)
                If ilPos > 0 Then
                    slDPBookName = Left$(slDPBookName, ilPos - 1)
                End If
            End If
            If cbcBookName(1).ListIndex >= 0 Then
                slETBookName = cbcBookName(1).List(cbcBookName(1).ListIndex)
                ilPos = InStrRev(slETBookName, ": ", -1, vbTextCompare)
                If ilPos > 0 Then
                    slETBookName = Left$(slETBookName, ilPos - 1)
                End If
            End If
        End If
        If rbcPCForm(0).Value Then
            If (slDPBookName = "") And (slETBookName = "") Then
                MsgBox "Book Name(s) Must be Defined", vbExclamation, "Name Error"
                If rbcNew(0).Value Then
                    edcBookName(0).SetFocus
                Else
                    cbcBookName(0).SetFocus
                End If
                Exit Sub
            End If
        ElseIf rbcPCForm(1).Value Then
            If (slDPBookName = "") Then
                MsgBox "Book Name Must be Defined", vbExclamation, "Name Error"
                If rbcNew(0).Value Then
                    edcBookName(0).SetFocus
                Else
                    cbcBookName(0).SetFocus
                End If
                Exit Sub
            End If
        
        End If
    Else
        'Created from Book Date Column (FAL01) and Market Name (New York, NY)
        slDPBookName = ""
        slETBookName = ""
    End If
    slBookDate = Trim$(edcBookDate.Text)
    If slBookDate = "" Then
        MsgBox "Book Date Must be Defined", vbExclamation, "Name Error"
        edcBookDate.SetFocus
        Exit Sub
    End If
    If Not gValidDate(slBookDate) Then
        MsgBox "Invalid Date", vbExclamation, "Date Error"
        edcBookDate.SetFocus
        Exit Sub
    End If
    '7582 fron not (1) to is (0)
    If rbcPCForm(0).Value Then
        'Check if book name used previously
        ilFound = False
        llTestDate = gDateValue(slBookDate)
        If slDPBookName <> "" Then
            For ilLoop = LBound(tgDnfBook) To UBound(tgDnfBook) - 1 Step 1
                If StrComp(Trim$(slDPBookName), Trim$(tgDnfBook(ilLoop).sBookName), 1) = 0 Then
                    gUnpackDateLong tgDnfBook(ilLoop).iBookDate(0), tgDnfBook(ilLoop).iBookDate(1), llDate
                    If (llDate = llTestDate) Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next ilLoop
            If rbcNew(0).Value Then
                If ilFound Then
                    MsgBox "Book Name/Date Used Previously", vbExclamation, "Name Error"
                    edcBookName(0).SetFocus
                    Exit Sub
                End If
            Else
                If Not ilFound Then
                    MsgBox "Book Name/Date Not Found", vbExclamation, "Name Error"
                    cbcBookName(0).SetFocus
                    Exit Sub
                End If
            End If
        End If
        ilFound = False
        If slETBookName <> "" Then
            For ilLoop = LBound(tgDnfBook) To UBound(tgDnfBook) - 1 Step 1
                If StrComp(Trim$(slETBookName), Trim$(tgDnfBook(ilLoop).sBookName), 1) = 0 Then
                    gUnpackDateLong tgDnfBook(ilLoop).iBookDate(0), tgDnfBook(ilLoop).iBookDate(1), llDate
                    If (llDate = llTestDate) Then
                        ilFound = True
                        Exit For
                    End If
                End If
            Next ilLoop
            If rbcNew(0).Value Then
                If ilFound Then
                    MsgBox "Book Name/Date Used Previously", vbExclamation, "Name Error"
                    edcBookName(0).SetFocus
                    Exit Sub
                End If
            Else
                If Not ilFound Then
                    MsgBox "Book Name/Date Not Found", vbExclamation, "Name Error"
                    cbcBookName(0).SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    'If imVehSelectedIndex < 0 Then
    '    MsgBox "Vehicle Must be Defined", vbExclamation, "Name Error"
    '    cbcVehicle.SetFocus
    '    Exit Sub
    'End If
    'slNameCode = tgUserVehicle(imVehSelectedIndex).sKey    'lbcMster.List(ilLoop)
    'ilRet = gParseItem(slNameCode, 1, "\", slName)
    'ilRet = gParseItem(slName, 3, "|", slName)
    'ilRet = gParseItem(slNameCode, 2, "\", slCode)
    'imVefCode = Val(slCode)
    'If imDaySelectedIndex < 0 Then
    '    MsgBox "Days Must be Defined", vbExclamation, "Name Error"
    '    cbcDays.SetFocus
    '    Exit Sub
    'End If
    'For ilDay = 0 To 6 Step 1
    '    smDays(ilDay) = "N"
    'Next ilDay
    'Select Case imDaySelectedIndex
    '    Case 0  'M-F
    '        ilSY = 0
    '        ilEY = 4
    '    Case 1  'Sa
    '        ilSY = 5
    '        ilEY = 5
    '    Case 2  'Su
    '        ilSY = 6
    '        ilEY = 6
    '    Case 3  'M-Sa
    '        ilSY = 0
    '        ilEY = 5
    '    Case 4  'M-Su
    '        ilSY = 0
    '        ilEY = 6
    '    Case 5  'Sa-Su
    '        ilSY = 5
    '        ilEY = 6
    '    Case 6  'Tu-Su
    '        ilSY = 1
    '        ilEY = 6
    '    Case 7  'Tu-Fr
    '        ilSY = 1
    '        ilEY = 4
    '    Case 8  'We-Su
    '        ilSY = 2
    '        ilEY = 6
    '    Case 9  'Mo
    '        ilSY = 0
    '        ilEY = 0
    '    Case 10  'Tu
    '        ilSY = 1
    '        ilEY = 1
    '    Case 11  'We
    '        ilSY = 2
    '        ilEY = 2
    '    Case 12  'Th
    '        ilSY = 3
    '        ilEY = 3
    '    Case 13  'Fr
    '        ilSY = 4
    '        ilEY = 4
    'End Select
    'For ilDay = ilSY To ilEY Step 1
    '    smDays(ilDay) = "Y"
    'Next ilDay
    Screen.MousePointer = vbHourglass
    
    blIsNewVersion = False
    If mIsNewVersion(slFromName) Then
        blIsNewVersion = True
    End If
    
    'dan 6/11/15 why is this here again?
    slBookDate = Trim$(edcBookDate.Text)
    'Check file names
    gGetSyncDateTime smSyncDate, smSyncTime
    plcGauge.Value = 0
    lmProcessedNoBytes = 0
    'ilIndex = 1
    'ReDim smFileNames(1 To 1) As String
    ilIndex = 0
    ReDim smFileNames(0 To 0) As String
    Do
        ilFound = False
        ilRet = gParseItem(slFromName, ilIndex + 1, "|", slName)
        If ilRet <> CP_MSG_NONE Then
            Exit Do
        End If
        ilFound = True
        smFileNames(ilIndex) = slName
        ilIndex = ilIndex + 1
        'ReDim Preserve smFileNames(1 To ilIndex) As String
        ReDim Preserve smFileNames(0 To ilIndex) As String
    Loop While ilFound
    lmTotalNoBytes = 0
    'For ilIndex = 1 To UBound(smFileNames) - 1 Step 1
    For ilIndex = 0 To UBound(smFileNames) - 1 Step 1
        slFromName = smFileNames(ilIndex)
        ilRet = 0
        'On Error GoTo cmcFileConvErr:
        'hmFrom = FreeFile
        'Open slFromName For Input Access Read As hmFrom
        ilRet = gFileOpen(slFromName, "Input Access Read", hmFrom)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            Close hmFrom
            MsgBox "Unable to find " & slFromName, vbExclamation, "Name Error"
            edcFrom.SetFocus
            Exit Sub
        End If
        lmTotalNoBytes = lmTotalNoBytes + LOF(hmFrom) 'The Loc returns current position \128
        '7582 not (1) to is (0)
        If (Not rbcImportFrom(0).Value) Or (rbcPCForm(0).Value) Then
            If Not mTestColumnTitle(blIsNewVersion) Then
                Screen.MousePointer = vbDefault
                Close hmFrom
                If rbcImportFrom(0).Value Then
                    MsgBox "Column Titles (Demographics, AQH & Population) Missing from " & slFromName, vbExclamation, "Name Error"
                Else
                    MsgBox "Column Titles (AQH AUD & Populations) Missing from " & slFromName, vbExclamation, "Name Error"
                End If
                cmcCancel.SetFocus
                Exit Sub
            End If
        End If
        Close hmFrom
    Next ilIndex
    ilRet = 0
    'hmTo = FreeFile
    'Open sgDBPath & "Messages\" & "ImptAct1.Txt" For Output As hmTo
    ilRet = gFileOpen(sgDBPath & "Messages\" & "ImptAct1.Txt", "Output", hmTo)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Open " & sgDBPath & "Messages\" & "ImptAct1.Txt" & ", Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Print #hmTo, "Import Act1 on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
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
    ReDim smUnfdStations(0 To 0) As String
    'For ilIndex = 1 To UBound(smFileNames) - 1 Step 1
    bmBooksSaved = False
    For ilIndex = 0 To UBound(smFileNames) - 1 Step 1
        slFromName = smFileNames(ilIndex)
        lacFileType.Caption = "Processing " & slFromName
        imConverting = True
        If Not mPrepassDemo(slFromName) Then
            Print #hmTo, "Import Act1 failed because unable to proceess " & slFromName & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
            Close hmTo
            imConverting = False
            'mTerminate
            Exit Sub
        End If
        If rbcImportFrom(0).Value Then
            If rbcPCForm(1).Value Then
                ''Save if add Station Import with Population by USA
                'If Not mPCStnConvFile_POPByUSA(slFromName, slDPBookName, slETBookName, slBookDate) Then
                '    Screen.MousePointer = vbDefault
                '    Print #hmTo, "Import Act1 terminated on " & Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                '    Close hmTo
                '    imConverting = False
                '    'mTerminate
                '    Exit Sub
                'End If
                ''Save if add Station Import with Population by USA
                '7582 replace
                If Not mPCStnConvFile_PopByMarket(slFromName, slBookDate) Then
                    Screen.MousePointer = vbDefault
                    Print #hmTo, "Import Act1 terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
                    Close hmTo
                    imConverting = False
                    'mTerminate
                    Exit Sub
                End If
            ElseIf rbcPCForm(2).Value Then
                If Not mConvertStationVehicles(slFromName, slDPBookName, slBookDate) Then
                    Screen.MousePointer = vbDefault
                    Print #hmTo, "Import Act1 terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
                    Close hmTo
                    imConverting = False
                    'mTerminate
                    Exit Sub
                End If
            Else
                If blIsNewVersion Then
                    If Not mPCNetConvFileNew(slFromName, slDPBookName, slETBookName, slBookDate) Then
                        Screen.MousePointer = vbDefault
                        Print #hmTo, "Import Act1 terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
                        Close hmTo
                        imConverting = False
                        'mTerminate
                        Exit Sub
                    End If
                Else
                    If Not mPCNetConvFile(slFromName, slDPBookName, slETBookName, slBookDate) Then
                        Screen.MousePointer = vbDefault
                        Print #hmTo, "Import Act1 terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
                        Close hmTo
                        imConverting = False
                        'mTerminate
                        Exit Sub
                    End If
                End If
            End If
        Else
            If Not mOnLineConvFile(slFromName, slDPBookName, slETBookName, slBookDate) Then
                Screen.MousePointer = vbDefault
                Print #hmTo, "Import Act1 terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
                Close hmTo
                imConverting = False
                'mTerminate
                Exit Sub
            End If
        End If
    Next ilIndex
    If bmBooksSaved Then
        bmResearchSaved = True
        Print #hmTo, "Import Act1 completed on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    Else
        ilRet = MsgBox("No Book Saved", vbOKOnly + vbInformation, "Warning")
        Print #hmTo, "Import Act1 completed but no book saved, on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    End If
    Close hmTo
    lacFileType.Caption = "Done"

    '11/26/17
    gFileChgdUpdate "vef.btr", False

    ilRet = mObtainBookName()
    imConverting = False
    'cmcFileConv.Enabled = False
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
'cmcFileConvErr:
'    ilRet = Err.Number
'    Resume Next
End Sub
Private Sub cmcFrom_Click()
    lacFileType.Caption = ""
    plcGauge.Value = 0
    'CMDialogBox.DialogTitle = "From File"
    'CMDialogBox.Filter = "Blank|*.|ASC|*.Asc|Text|*.Txt|Print|*.Prn|All|*.*"
    'CMDialogBox.InitDir = Left$(sgImportPath, Len(sgImportPath) - 1)
    'CMDialogBox.Filename = ""
    'CMDialogBox.DefaultExt = ""
    'CMDialogBox.Action = 1 'Open dialog
    'edcFrom.Text = CMDialogBox.Filename
    'ChDir sgCurDir
    
    If rbcImportFrom(0).Value Then
        igBrowserType = 1 + SHIFT8 '1=CSV
    Else
        igBrowserType = 2 + SHIFT8     'Text
    End If
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
    '7582
    If rbcPCForm(2).Value Then
        edcBookName(0).Text = mParseForBookName(edcFrom.Text)
    End If
End Sub

Private Sub cmcMoveToDaypart_Click()
    Dim ilRet As Integer
    Dim ilIndex As Integer
    
    If cbcBookName(1).ListIndex >= 0 Then
        ilIndex = cbcBookName(1).ListIndex
        tmDnfSrchKey.iCode = cbcBookName(1).ItemData(cbcBookName(1).ListIndex)
        ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            tmDnf.sExactTime = "N"
            ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
            Screen.MousePointer = vbHourglass
            ilRet = mObtainBookName()
            cmcMoveToDaypart.Enabled = False
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub cmcMoveToExact_Click()
    Dim ilRet As Integer
    Dim ilIndex As Integer
    
    If cbcBookName(0).ListIndex >= 0 Then
        ilIndex = cbcBookName(0).ListIndex
        tmDnfSrchKey.iCode = cbcBookName(0).ItemData(cbcBookName(0).ListIndex)
        ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            tmDnf.sExactTime = "Y"
            ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
            Screen.MousePointer = vbHourglass
            ilRet = mObtainBookName()
            cmcMoveToExact.Enabled = False
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub


Private Sub edcBookDate_GotFocus()
    lacFileType.Caption = ""
    plcGauge.Value = 0
    gCtrlGotFocus edcBookDate
End Sub
Private Sub edcBookName_GotFocus(Index As Integer)
    lacFileType.Caption = ""
    plcGauge.Value = 0
    gCtrlGotFocus edcBookName(Index)
End Sub
Private Sub edcFrom_GotFocus()
    lacFileType.Caption = ""
    plcGauge.Value = 0
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus edcFrom
End Sub
Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
    If tgSpf.sCAudPkg <> "Y" Then
        ckcDefault(1).Visible = False
    End If
    Me.KeyPreview = True
    Me.Refresh
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        gFunctionKeyBranch KeyCode
        plcPCForm.Visible = False
        plcPCForm.Visible = True
        plcImportFrom.Visible = False
        plcImportFrom.Visible = True
        plcDefault.Visible = False
        plcDefault.Visible = True
        plcFrom.Visible = False
        plcFrom.Visible = True
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
    Erase tgDnfBook
    Erase tgMnfCDemo
    Erase tgMnfSDemo
    Erase imVefCodeImpt
    Erase tmRdfInfo
    Erase tmDrfInfo
    Erase smFileNames
    Erase imMatchRdfCode
    Erase smUnfdStations
    Erase tgDpfInfo
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
'    ilRet = btrClose(hmDsf)
'    btrDestroy hmDsf
    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    ilRet = btrClose(hmDpf)
    btrDestroy hmDpf
    ilRet = btrClose(hmDnf)
    btrDestroy hmDnf
    
    Set ImptMark = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddDemo                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Add non-standard demo name     *
'*                                                     *
'*******************************************************
Private Function mAddDemo(slDemoName As String) As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilRet As Integer
    Dim ilUpper As Integer
    gGetSyncDateTime slSyncDate, slSyncTime
    tmMnf.iCode = 0
    tmMnf.sType = "D"
    tmMnf.sName = slDemoName
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
    If ilRet = BTRV_ERR_NONE Then
        ilUpper = UBound(tgMnfSDemo)
        tgMnfSDemo(ilUpper) = tmMnf
        ilUpper = ilUpper + 1
        'ReDim Preserve tgMnfSDemo(1 To ilUpper) As MNF
        ReDim Preserve tgMnfSDemo(0 To ilUpper) As MNF
    End If
    mAddDemo = True
End Function
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
Private Function mBookNameUsed(slBookName As String, slBookDate As String, ilVefCode As Integer, slDays() As String * 1, slInfoType As String, ilRdfCode As Integer, llStartTime As Long, llEndTime As Long, ilCustomDemo As Integer, ilPrevDnfCode As Integer, llPrevDrfCode As Long) As Integer
'
'   mBookNameUsed(O)- O= No; 1= Name defined, Dates not matching; 2=Name Defined and Dates Match; 3=Name defined and Date Match and DRF record found


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
    Dim llSTime As Long
    Dim llETime As Long
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
                    gUnpackTimeLong tlDrf.iStartTime(0), tlDrf.iStartTime(1), False, llSTime
                    gUnpackTimeLong tlDrf.iEndTime(0), tlDrf.iEndTime(1), False, llETime
                    If (slInfoType = "V") Or ((slInfoType = "D") And (ilRdfCode <> 0)) Or ((slInfoType = "D") And (ilRdfCode = 0) And (llStartTime = llSTime) And (llEndTime = llETime)) Then
                        ilMatch = True
                        If (slInfoType = "D") And (ilRdfCode <> 0) Then
                            If ilRdfCode <> tlDrf.iRdfCode Then
                                ilMatch = False
                            Else
                                'Check days if not all blank
                                For ilDay = 0 To 6 Step 1
                                    If (Trim$(tlDrf.sDay(ilDay)) <> "") And (Asc(tlDrf.sDay(ilDay)) <> 0) Then
                                        If tlDrf.sDay(ilDay) <> slDays(ilDay) Then
                                            ilMatch = False
                                            Exit For
                                        End If
                                    End If
                                Next ilDay
                            End If
                        Else
                            If (slInfoType <> "V") Then
                                For ilDay = 0 To 6 Step 1
                                    If tlDrf.sDay(ilDay) <> slDays(ilDay) Then
                                        ilMatch = False
                                        Exit For
                                    End If
                                Next ilDay
                            End If
                        End If
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
                    End If
                    ilRet = btrGetNext(hmDrf, tlDrf, imDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                Loop
                Exit Function
            End If
        End If
    Next ilLoop
    Exit Function
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
    Dim ilIndex As Integer
    Dim llSec As Long
    Dim llMin As Long
    Dim llHour As Long
    Dim llTime As Long
    Dim llDate As Long
    Dim llTstDate As Long
    Dim ilVefCode As Integer
    Dim ilFound As Integer
    ReDim ilDays(0 To 6) As Integer
    ReDim slDays(0 To 6) As String * 1
    'If imVehSelectedIndex >= 0 Then
    '    slNameCode = tgUserVehicle(imVehSelectedIndex).sKey    'lbcMster.List(ilLoop)
    '    ilRet = gParseItem(slNameCode, 1, "\", slName)
    '    ilRet = gParseItem(slName, 3, "|", slName)
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '    ilVefCode = Val(slCode)
    'Else
    '    ilVefCode = 0
    'End If
    ilVefCode = imVefCode
    If smDPStamp = Trim$(Str$(ilVefCode)) Then
        Exit Sub
    End If
    smDPStamp = Trim$(Str$(ilVefCode))
    Screen.MousePointer = vbHourglass
    ReDim tmRdfInfo(0 To 0) As RDFINFO
    cbcDPNames.Clear
    'Determine Rate Card
    ilRcf = -1
    For ilLoop = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
        If (tgMRcf(ilLoop).iVefCode = ilVefCode) And (tgMRcf(ilLoop).iYear = imNowYear) Then
            If ilRcf = -1 Then
                ilRcf = ilLoop
                gUnpackDateLong tgMRcf(ilLoop).iStartDate(0), tgMRcf(ilLoop).iStartDate(1), llDate
            Else
                gUnpackDateLong tgMRcf(ilLoop).iStartDate(0), tgMRcf(ilLoop).iStartDate(1), llTstDate
                If llTstDate > llDate Then
                    llDate = llTstDate
                    ilRcf = ilLoop
                End If
            End If
        End If
    Next ilLoop
    If ilRcf = -1 Then
        For ilLoop = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
            If (tgMRcf(ilLoop).iVefCode = 0) And (tgMRcf(ilLoop).iYear = imNowYear) Then
                If ilRcf = -1 Then
                    ilRcf = ilLoop
                    gUnpackDateLong tgMRcf(ilLoop).iStartDate(0), tgMRcf(ilLoop).iStartDate(1), llDate
                Else
                    gUnpackDateLong tgMRcf(ilLoop).iStartDate(0), tgMRcf(ilLoop).iStartDate(1), llTstDate
                    If llTstDate > llDate Then
                        llDate = llTstDate
                        ilRcf = ilLoop
                    End If
                End If
            End If
        Next ilLoop
    End If
    If (ilRcf = -1) And (LBound(tgMRcf) = UBound(tgMRcf) - 1) Then
        ilRcf = LBound(tgMRcf)
    End If
    ilUpper = 0
    If ilRcf >= 0 Then
        For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
            If (tgMRcf(ilRcf).iCode = tgMRif(llRif).iRcfCode) And (tgMRif(llRif).iVefCode = ilVefCode) Then
                'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                '    If (tgMRif(llRif).iRdfcode = tgMRdf(ilRdf).iCode) And (tgMRdf(ilRdf).sState = "A") Then
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
                                        'If tgMRdf(ilRdf).sWkDays(ilTime, ilIndex) <> "Y" Then
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
    imFirstActivate = True
    'mParseCmmdLine
    Screen.MousePointer = vbHourglass
    ImptMark.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    bmResearchSaved = False
    imTestAddStdDemo = True
    imConverting = False
    imFirstFocus = True
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0
    imChgMode = False
    imBSMode = False
    imBNSelectedIndex = -1
    imVehSelectedIndex = -1
    imDaySelectedIndex = -1
    imDPSelectedIndex = -1
    hmMnf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptMark
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    hmVef = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptMark
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmDrf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptMark
    On Error GoTo 0
    ReDim tmDrfInfo(0 To 0) As DRFINFO
    imDrfRecLen = Len(tmDrfInfo(0).tDrf)
    hmDpf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptMark
    On Error GoTo 0
    imDpfRecLen = Len(tmDpf)
    hmDnf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptMark
    On Error GoTo 0
    imDnfRecLen = Len(tmDnf)
'    hmDsf = CBtrvTable(TWOHANDLES) 'CBtrvObj
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "Dsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptMark
'    On Error GoTo 0
'    imDsfRecLen = Len(tmDsf)
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
    ilRet = mObtainDemo()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Custom Demo Error", vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
        imTerminate = True
        Exit Sub
    End If
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    gObtainMonthYear 0, smNowDate, ilMonth, imNowYear

    'smRptTime = Format$(Now, "h:m:s AM/PM")
    'gPackTime smRptTime, tmIcf.iTime(0), tmIcf.iTime(1)
    gCenterStdAlone ImptMark
    If mTestRecLengths() Then
        Screen.MousePointer = vbDefault
        imTerminate = True
        Exit Sub
    End If
    ilRet = 0
    On Error GoTo mInit1Err:
    'hmFrom = FreeFile
    Screen.MousePointer = vbDefault
    
    ' TP 10590
    edcBookDate.Text = Format$(gNow(), "m/d/yy")
    
    'imcHelp.Picture = Traffic!imcHelp.Picture
    '7582
    rbcPCForm(1).Visible = False
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
    Dim ilLoop As Integer
    Dim slDate As String
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim ilVehDefined As Integer

    ReDim tgDnfBook(0 To 0) As DNF
    cbcBookName(0).Clear
    cbcBookName(1).Clear
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
    ReDim tlDnfSort(0 To UBound(tgDnfBook)) As DNFSORT
    For ilLoop = LBound(tgDnfBook) To UBound(tgDnfBook) - 1 Step 1
        tlDnfSort(ilLoop).sKey = mCreateSort(tgDnfBook(ilLoop))
        tlDnfSort(ilLoop).tDnf = tgDnfBook(ilLoop)
    Next ilLoop
    If UBound(tlDnfSort) - 1 > 0 Then
        ArraySortTyp fnAV(tlDnfSort(), 0), UBound(tlDnfSort), 0, LenB(tlDnfSort(0)), 0, LenB(tlDnfSort(0).sKey), 0
    End If
    For ilLoop = LBound(tgDnfBook) To UBound(tgDnfBook) - 1 Step 1
        tgDnfBook(ilLoop) = tlDnfSort(ilLoop).tDnf
    Next ilLoop
    
    For ilLoop = LBound(tgDnfBook) To UBound(tgDnfBook) - 1 Step 1
        ilVehDefined = False
'        tmDrfSrchKey.iDnfCode = tgDnfBook(ilLoop).iCode
'        tmDrfSrchKey.sDemoDataType = "D"
'        tmDrfSrchKey.iMnfSocEco = 0
'        tmDrfSrchKey.iVefCode = 0
'        tmDrfSrchKey.sInfoType = "V"
'        tmDrfSrchKey.iRdfcode = 0
'        gUnpackDate tgDnfBook(ilLoop).iBookDate(0), tgDnfBook(ilLoop).iBookDate(1), slDate
'        ilRet = btrGetGreaterOrEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'        Do While (ilRet = BTRV_ERR_NONE) And (tmDrf.iDnfCode = tgDnfBook(ilLoop).iCode) And (tmDrf.sDemoDataType = "D")
'            If (tmDrf.sInfoType = "V") Then
'                ilVehDefined = True
'                Exit Do
'            End If
'            ilRet = btrGetNext(hmDrf, tmDrf, imDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'        Loop
        gUnpackDate tgDnfBook(ilLoop).iBookDate(0), tgDnfBook(ilLoop).iBookDate(1), slDate
        If (Trim$(tgDnfBook(ilLoop).sSource) = "") Or (Trim$(tgDnfBook(ilLoop).sSource) = "A") Or (Trim$(tgDnfBook(ilLoop).sSource) = "M") Then
            If tgDnfBook(ilLoop).sExactTime = "Y" Then
                ilVehDefined = True
            End If
            If ilVehDefined Then
                cbcBookName(1).AddItem Trim$(tgDnfBook(ilLoop).sBookName) & ": " & slDate
                cbcBookName(1).ItemData(cbcBookName(1).NewIndex) = tgDnfBook(ilLoop).iCode
            Else
                cbcBookName(0).AddItem Trim$(tgDnfBook(ilLoop).sBookName) & ": " & slDate
                cbcBookName(0).ItemData(cbcBookName(0).NewIndex) = tgDnfBook(ilLoop).iCode
            End If
        End If
    Next ilLoop
    If cbcBookName(0).ListCount < 12 Then
        gSetComboboxDropdownHeight ImptMark, cbcBookName(0), cbcBookName(0).ListCount
    Else
        gSetComboboxDropdownHeight ImptMark, cbcBookName(0), 12
    End If
    If cbcBookName(1).ListCount < 12 Then
        gSetComboboxDropdownHeight ImptMark, cbcBookName(1), cbcBookName(1).ListCount
    Else
        gSetComboboxDropdownHeight ImptMark, cbcBookName(1), 12
    End If
    mObtainBookName = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainDemo                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgMnfCDemo             *
'*                                                     *
'*******************************************************
Private Function mObtainDemo() As Integer
'
'   ilRet = mObtainDemo ()
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
    'ReDim tgMnfSDemo(1 To 1) As MNF
    'ilUpperBound = 1
    'ilRecLen = Len(tgMnfCDemo(1)) 'btrRecordLength(hmMnf)  'Get and save record length
    ReDim tgMnfCDemo(0 To 0) As MNF
    ReDim tgMnfSDemo(0 To 0) As MNF
    ilUpperBound = 0
    ilRecLen = Len(tgMnfCDemo(0)) 'btrRecordLength(hmMnf)  'Get and save record length
    'llNoRec = btrRecords(hmMnf) 'Obtain number of records
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlFile) 'Obtain number of records
    btrExtClear hmMnf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmMnf, tmMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        mObtainDemo = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mObtainDemo = False
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hmMnf, llNoRec, 0, "UC", "MNF", "") 'Set extract limits (all records)
    tlCharTypeBuff.sType = "D"
    ilOffSet = 2 'gFieldOffset("Mnf", "MnfType")
    ilRet = btrExtAddLogicConst(hmMnf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
    ilOffSet = 0 'gFieldOffset("Mnf", "MnfCode")
    ilRet = btrExtAddField(hmMnf, ilOffSet, ilRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainDemo = False
        Exit Function
    End If
    ilRet = btrExtGetNext(hmMnf, tmMnf, ilRecLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            mObtainDemo = False
            Exit Function
        End If
        'ilRecLen = Len(tgMnfCDemo(1))  'Extract operation record size
        ilRecLen = Len(tgMnfCDemo(0))  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmMnf, tmMnf, ilRecLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If tmMnf.iGroupNo > 0 Then
                ilUpperBound = UBound(tgMnfCDemo)
                tgMnfCDemo(ilUpperBound) = tmMnf
                ilUpperBound = ilUpperBound + 1
                'ReDim Preserve tgMnfCDemo(1 To ilUpperBound) As MNF
                ReDim Preserve tgMnfCDemo(0 To ilUpperBound) As MNF
            Else
                ilUpperBound = UBound(tgMnfSDemo)
                tgMnfSDemo(ilUpperBound) = tmMnf
                ilUpperBound = ilUpperBound + 1
                'ReDim Preserve tgMnfSDemo(1 To ilUpperBound) As MNF
                ReDim Preserve tgMnfSDemo(0 To ilUpperBound) As MNF
            End If
            ilRet = btrExtGetNext(hmMnf, tmMnf, ilRecLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmMnf, tmMnf, ilRecLen, llRecPos)
            Loop
        Loop
    End If
    mObtainDemo = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mOnLineConvFile                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert File                   *
'*                                                     *
'*******************************************************
Private Function mOnLineConvFile(slFromFile As String, slDPBookName As String, slETBookName As String, slBookDate As String) As Integer
    Dim ilRet As Integer
    Dim ilBNRet As Integer
    Dim ilQRet As Integer
    Dim slLine As String
    Dim ilHeaderFd As Integer
    Dim ilPopIndex As Integer
    Dim ilPopCol As Integer
    Dim ilAvgQHIndex As Integer
    Dim ilPopDone As Integer
    Dim ilDemoGender As Integer
    Dim slDemoAge As String
    Dim ilPos As Integer
    Dim slDay As String
    Dim ilDay As Integer
    Dim ilSY As Integer
    Dim ilEY As Integer
    Dim ilLoop As Integer
    Dim ilEof As Integer
    Dim llPercent As Long
    Dim slChar As String
    Dim slTime As String
    Dim slStr As String
    Dim slSexChar As String
    Dim ilIndex As Integer
    Dim ilCustomDemo As Integer
    Dim ilPrevDnfCode As Integer
    Dim llPrevDrfCode As Long
    Dim ilRdf As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim slStationCode As String
    Dim slVehicleName As String
    Dim ilNoDemoFd As Integer
    Dim ilMatch As Integer
    Dim llTime As Long
    Dim ilVef As Integer
    Dim ilVff As Integer
    Dim ilTIndex As Integer
    Dim ilSetBkNm As Integer
    Dim ilSetDnfCode As Integer
    Dim tlDrf As DRF
    ilRet = 0
    'On Error GoTo mOnLineConvFileErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        MsgBox "Open " & slFromFile & ", Error #" & Str$(ilRet), vbExclamation, "Open Error"
        edcFrom.SetFocus
        mOnLineConvFile = False
        Exit Function
    End If
    DoEvents
    If imTerminate Then
        Close hmFrom
        mTerminate
        mOnLineConvFile = False
        Exit Function
    End If
    ilHeaderFd = False
    slLine = ""
    ilPopDone = False
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
    tmDrfPop.sACTLineupCode = ""
    tmDrfPop.sACT1StoredTime = ""
    tmDrfPop.sACT1StoredSpots = ""
    tmDrfPop.sACT1StoreClearPct = ""
    tmDrfPop.sACT1DaypartFilter = ""
    
    ilNoDemoFd = 0
    ilCustomDemo = False
    Do
        err.Clear
        ilRet = 0
        'On Error GoTo mOnLineConvFileErr:
        'Line Input #hmFrom, slLine
        slLine = ""
        Do
            If Not EOF(hmFrom) Then
                slChar = Input(1, #hmFrom)
                If slChar = Chr(13) Then
                    slChar = Input(1, #hmFrom)
                End If
                If slChar = Chr(10) Then
                    Exit Do
                End If
                slLine = slLine & slChar
            Else
                ilEof = True
                Exit Do
            End If
        Loop
        slLine = Trim$(slLine)
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
                    mOnLineConvFile = False
                    Exit Function
                End If
                'Determine field Type
                'Header Record
                '   Continental U.S.  Continental U.S.
                '       AQH AUD   RTG      Populations
                'Demo Record
                '   Males   12-17 12,000 0.1 21,847,000
                '   Females 12-17
                '   Men     18-24
                '   Women   18-24
                '
                If Not ilHeaderFd Then
                    If InStr(1, RTrim$(slLine), "AQH AUD", 1) > 0 Then
                        If InStr(1, RTrim$(slLine), "Populations", 1) > 0 Then
                            'Determine number of Demo columns
                            ilPopCol = InStr(1, RTrim$(slLine), "Populations", 1) - 4
                            ilPos = InStr(1, RTrim$(slLine), "AQH AUD", 1)
                            ReDim tmDrfInfo(0 To 0) As DRFINFO
                            Do While ilPos > 0
                                tmDrfInfo(UBound(tmDrfInfo)).iStartCol = ilPos - 3
                                tmDrfInfo(UBound(tmDrfInfo)).iType = 0  'Vehicle
                                tmDrfInfo(UBound(tmDrfInfo)).iBkNm = 0  'Daypart Book Name
                                tmDrfInfo(UBound(tmDrfInfo)).iSY = -1
                                tmDrfInfo(UBound(tmDrfInfo)).iEY = -1
                                ReDim Preserve tmDrfInfo(0 To UBound(tmDrfInfo) + 1) As DRFINFO
                                ilPos = InStr(ilPos + 7, RTrim$(slLine), "AQH AUD", 1)
                            Loop
                            'Determine Vehicle and Daypart info
                            'For ilLoop = 10 To 1 Step -1
                            For ilLoop = 9 To 0 Step -1
                                If InStr(1, smSvLine(ilLoop), "COVERAGE", 1) > 0 Then
                                    slStationCode = ""
                                    slVehicleName = ""
                                    Exit For
                                End If
                                If Left$(smSvLine(ilLoop), 1) = "(" Then
                                    slStationCode = ""
                                    slVehicleName = ""
                                    ilRow = 2
                                    slChar = Mid$(smSvLine(ilLoop), ilRow, 1)
                                    Do While slChar <> ")"
                                        slStationCode = slStationCode & slChar
                                        ilRow = ilRow + 1
                                        slChar = Mid$(smSvLine(ilLoop), ilRow, 1)
                                    Loop
                                    slStationCode = Trim$(slStationCode)
                                    ilRow = ilRow + 1
                                    ilPos = InStr(1, smSvLine(ilLoop), "Month :", 1)
                                    If ilPos > 0 Then
                                        slVehicleName = Trim$(Mid$(smSvLine(ilLoop), ilRow, ilPos - ilRow))
                                    Else
                                        slVehicleName = Trim$(Mid$(smSvLine(ilLoop), ilRow))
                                    End If
                                    slVehicleName = Trim$(Left$(slVehicleName, Len(slVehicleName) - 8))
                                End If
                                If InStr(1, smSvLine(ilLoop), "Schedule", 1) = 1 Then
                                    'Get Days and Times
                                    slStr = UCase$(Trim$(Mid$(smSvLine(ilLoop), 11)))
                                    'slDay = Mid$(slStr, 1, 2)
                                    'Select Case slDay
                                    '    Case "MO"
                                    '        ilSY = 0
                                    '    Case "TU"
                                    '        ilSY = 1
                                    '    Case "WE"
                                    '        ilSY = 2
                                    '    Case "TH"
                                    '        ilSY = 3
                                    '    Case "FR"
                                    '        ilSY = 4
                                    '    Case "SA"
                                    '        ilSY = 5
                                    '    Case "SU"
                                    '        ilSY = 6
                                    'End Select
                                    'slStr = Mid$(slStr, 3)
                                    'slDay = Mid$(slStr, 1, 2)
                                    'Select Case slDay
                                    '    Case "MO"
                                    '        ilEY = 0
                                    '    Case "TU"
                                    '        ilEY = 1
                                    '    Case "WE"
                                    '        ilEY = 2
                                    '    Case "TH"
                                    '        ilEY = 3
                                    '    Case "FR"
                                    '        ilEY = 4
                                    '    Case "SA"
                                    '        ilEY = 5
                                    '    Case "SU"
                                    '        ilEY = 6
                                    'End Select
                                    mGetDayIndex slStr, ilSY, ilEY
                                    If (ilSY <> -1) And (ilEY <> -1) Then
                                        For ilDay = 0 To 6 Step 1
                                            smDays(ilDay) = "N"
                                        Next ilDay
                                        For ilDay = ilSY To ilEY Step 1
                                            smDays(ilDay) = "Y"
                                        Next ilDay
                                        For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
                                            tmDrfInfo(ilCol).iType = 0  'Vehicle
                                            For ilDay = 0 To 6 Step 1
                                                tmDrfInfo(ilCol).sDays(ilDay) = smDays(ilDay)
                                            Next ilDay
                                        Next ilCol
                                    End If
                                End If
                                For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
                                    slStr = UCase$(Trim$(Mid$(smSvLine(ilLoop), tmDrfInfo(ilCol).iStartCol, 16)))
                                    If InStr(1, slStr, "Stored Sch", 1) > 0 Then
                                        tmDrfInfo(ilCol).iBkNm = 1  'Exact Time
                                        tmDrfInfo(ilCol).iType = 0  'Vehicle
                                        For ilDay = 0 To 6 Step 1
                                            tmDrfInfo(ilCol).sDays(ilDay) = "Y"
                                        Next ilDay
                                    Else
                                        If tmDrfInfo(ilCol).iSY = -1 Then
                                            ilSY = -1
                                            slDay = Mid$(slStr, 1, 2)
                                            'ilTIndex = 3
                                            'Select Case slDay
                                            '    Case "MO"
                                            '        ilSY = 0
                                            '        ilEY = 0
                                            '    Case "TU"
                                            '        ilSY = 1
                                            '        ilEY = 1
                                            '    Case "WE"
                                            '        ilSY = 2
                                            '        ilEY = 2
                                            '    Case "TH"
                                            '        ilSY = 3
                                            '        ilEY = 3
                                            '    Case "FR"
                                            '        ilSY = 4
                                            '        ilEY = 4
                                            '    Case "SA"
                                            '        ilSY = 5
                                            '        ilEY = 5
                                            '        If Mid$(slStr, 3, 2) = "SU" Then
                                            '            ilEY = 6
                                            '            ilTIndex = 5
                                            '        End If
                                            '    Case "SU"
                                            '        ilSY = 6
                                            '        ilEY = 6
                                            '    Case "MF"
                                            '        ilSY = 0
                                            '        ilEY = 4
                                            '    Case "MS"
                                            '        ilSY = 0
                                            '        ilEY = 6
                                            '    Case "SS"
                                            '        ilSY = 5
                                            '        ilEY = 6
                                            'End Select
                                            ilTIndex = InStr(1, slStr, " ", 1) + 1
                                            If ilTIndex > 2 Then
                                                mGetDayIndex Left$(slStr, ilTIndex - 2), ilSY, ilEY
                                            End If
                                            If ilSY <> -1 Then
                                                slChar = Mid$(slStr, ilTIndex, 1)
                                                If (slChar >= "0") And (slChar <= "9") Then
                                                    slStr = Mid$(slStr, ilTIndex)
                                                Else
                                                    slStr = Mid$(slStr, ilTIndex + 1)
                                                End If
                                                slTime = Mid$(slStr, 1, 1)
                                                slStr = Mid$(slStr, 2)
                                                slChar = Mid$(slStr, 1, 1)
                                                If (slChar >= "0") And (slChar <= "9") Then
                                                    slTime = slTime & Mid$(slStr, 1, 2)
                                                    slStr = Mid$(slStr, 3)
                                                Else
                                                    slTime = slTime & Mid$(slStr, 1, 1)
                                                    slStr = Mid$(slStr, 2)
                                                End If
                                                If gValidTime(slTime) Then
                                                    tmDrfInfo(ilCol).lStartTime = gTimeToCurrency(slTime, False)
                                                    slChar = Mid$(slStr, 1, 1)
                                                    If slChar = "-" Then
                                                        slStr = Trim$(Mid$(slStr, 2))
                                                    End If
                                                    slTime = Mid$(slStr, 1, 1)
                                                    slStr = Mid$(slStr, 2)
                                                    slChar = Mid$(slStr, 1, 1)
                                                    If (slChar >= "0") And (slChar <= "9") Then
                                                        slTime = slTime & Mid$(slStr, 1, 2)
                                                        slStr = Mid$(slStr, 3)
                                                    Else
                                                        slTime = slTime & Mid$(slStr, 1, 1)
                                                        slStr = Mid$(slStr, 2)
                                                    End If
                                                    If gValidTime(slTime) Then
                                                        tmDrfInfo(ilCol).iSY = ilSY
                                                        tmDrfInfo(ilCol).iEY = ilEY
                                                        tmDrfInfo(ilCol).lEndTime = gTimeToCurrency(slTime, False)
                                                        For ilDay = 0 To 6 Step 1
                                                            tmDrfInfo(ilCol).sDays(ilDay) = "N"
                                                        Next ilDay
                                                        For ilDay = ilSY To ilEY Step 1
                                                            tmDrfInfo(ilCol).sDays(ilDay) = "Y"
                                                        Next ilDay
                                                        tmDrfInfo(ilCol).iType = 1  'Daypart
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next ilCol
                            Next ilLoop
                            If slVehicleName <> "" Then
                                'Test if vehicle name matches
                                imVehSelectedIndex = -1
                                'For ilLoop = LBound(tgUserVehicle) To UBound(tgUserVehicle) - 1 Step 1
                                '    slNameCode = tgUserVehicle(ilLoop).sKey    'lbcMster.List(ilLoop)
                                '    ilRet = gParseItem(slNameCode, 1, "\", slName)
                                '    ilRet = gParseItem(slName, 3, "|", slName)
                                '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                '    If StrComp(slVehicleName, slName, 1) = 0 Then
                                '        ilHeaderFd = True
                                '        imVefCode = Val(slCode)
                                '        cbcVehicle.ListIndex = ilLoop
                                '        Exit For
                                '    End If
                                'Next ilLoop
                                ReDim imVefCodeImpt(0 To 0) As Integer
                                For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                    If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Or (tgMVef(ilLoop).sType = "R") Or (tgMVef(ilLoop).sType = "V") Then
                                        If ((slStationCode <> "") And (StrComp(UCase(slStationCode), UCase(Trim$(tgMVef(ilLoop).sCodeStn)), 1) = 0)) Or (StrComp(slVehicleName, Trim$(tgMVef(ilLoop).sName), 1) = 0) Then
                                            imVehSelectedIndex = ilLoop
                                            ilHeaderFd = True
                                            imVefCodeImpt(UBound(imVefCodeImpt)) = tgMVef(ilLoop).iCode
                                            ReDim Preserve imVefCodeImpt(0 To UBound(imVefCodeImpt) + 1) As Integer
                                        Else
                                            ilVff = gBinarySearchVff(tgMVef(ilLoop).iCode)
                                            If ilVff <> -1 Then
                                                If ((slStationCode <> "") And (StrComp(UCase(slStationCode), UCase(Trim$(tgVff(ilVff).sACT1LineupCode)), 1) = 0)) Then
                                                    imVehSelectedIndex = ilLoop
                                                    ilHeaderFd = True
                                                    imVefCodeImpt(UBound(imVefCodeImpt)) = tgMVef(ilLoop).iCode
                                                    ReDim Preserve imVefCodeImpt(0 To UBound(imVefCodeImpt) + 1) As Integer
                                                End If
                                            End If
                                        End If
                                    End If
                                Next ilLoop
                                If UBound(imVefCodeImpt) <= 0 Then
                                    Print #hmTo, slVehicleName & " Not Found"
                                    lbcErrors.AddItem slVehicleName & " Not Found"
                                Else
                                    ilNoDemoFd = 0
                                    For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
                                        tmDrfInfo(ilCol).tDrf.lCode = 0
                                        tmDrfInfo(ilCol).tDrf.iDnfCode = tmDnf.iCode
                                        tmDrfInfo(ilCol).tDrf.sDemoDataType = "D"
                                        tmDrfInfo(ilCol).tDrf.iMnfSocEco = 0
                                        If tmDrfInfo(ilCol).iType = 0 Then
                                            tmDrfInfo(ilCol).tDrf.sInfoType = "V"
                                            tmDrfInfo(ilCol).tDrf.iRdfCode = 0
                                            gPackTime "12AM", tmDrfInfo(ilCol).tDrf.iStartTime(0), tmDrfInfo(ilCol).tDrf.iStartTime(1)
                                            gPackTime "12AM", tmDrfInfo(ilCol).tDrf.iEndTime(0), tmDrfInfo(ilCol).tDrf.iEndTime(1)
                                        Else
                                            tmDrfInfo(ilCol).tDrf.sInfoType = "D"
                                            tmDrfInfo(ilCol).tDrf.iRdfCode = 0
                                            gPackTimeLong tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).tDrf.iStartTime(0), tmDrfInfo(ilCol).tDrf.iStartTime(1)
                                            gPackTimeLong tmDrfInfo(ilCol).lEndTime, tmDrfInfo(ilCol).tDrf.iEndTime(0), tmDrfInfo(ilCol).tDrf.iEndTime(1)
                                            For ilDay = 0 To 6 Step 1
                                                tmDrfInfo(ilCol).tDrf.sDay(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
                                            Next ilDay
                                        End If
                                        tmDrfInfo(ilCol).tDrf.sProgCode = ""
                                        tmDrfInfo(ilCol).tDrf.iStartTime2(0) = 1
                                        tmDrfInfo(ilCol).tDrf.iStartTime2(1) = 0
                                        tmDrfInfo(ilCol).tDrf.iEndTime2(0) = 1
                                        tmDrfInfo(ilCol).tDrf.iEndTime2(1) = 0
                                        tmDrfInfo(ilCol).tDrf.iQHIndex = 0
                                        tmDrfInfo(ilCol).tDrf.iCount = 0
                                        tmDrfInfo(ilCol).tDrf.sExStdDP = "N"
                                        tmDrfInfo(ilCol).tDrf.sExRpt = "N"
                                        tmDrfInfo(ilCol).tDrf.sDataType = "A"
                                        For ilLoop = 1 To 16 Step 1
                                            tmDrfInfo(ilCol).tDrf.lDemo(ilLoop - 1) = 0
                                        Next ilLoop
                                    Next ilCol
                                End If
                            End If
                            'For ilLoop = 1 To 10 Step 1
                            For ilLoop = 0 To 9 Step 1
                                smSvLine(ilLoop) = ""
                            Next ilLoop
                        Else
                            'Save Last 10 lines
                            'For ilLoop = 2 To 10 Step 1
                            For ilLoop = 1 To 9 Step 1
                                smSvLine(ilLoop - 1) = smSvLine(ilLoop)
                            Next ilLoop
                            'smSvLine(10) = slLine
                            smSvLine(9) = slLine
                        End If
                    Else
                        'Save Last 10 lines
                        'For ilLoop = 2 To 10 Step 1
                        For ilLoop = 1 To 9 Step 1
                            smSvLine(ilLoop - 1) = smSvLine(ilLoop)
                        Next ilLoop
                        'smSvLine(10) = slLine
                        smSvLine(9) = slLine
                    End If
                Else
                    ilDemoGender = -1
                    ilPos = -1
                    'Test if Custom Demo
                    slLine = RTrim$(slLine)
                    smFieldValues(1) = Trim$(Mid$(slLine, 1, 12))
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
                                smFieldValues(3) = Mid$(slLine, ilPopCol, 16)
                                ilPopIndex = 3
'                                If tgSpf.sSAudData <> "H" Then
'                                    tmDrfPop.lDemo(ilDemoGender) = (CLng(smFieldValues(ilPopIndex)) + 500) \ 1000
'                                Else
'                                    tmDrfPop.lDemo(ilDemoGender) = (CLng(smFieldValues(ilPopIndex)) + 50) \ 100
'                                End If
                                If tgSpf.sSAudData = "H" Then
                                    tmDrfPop.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilPopIndex)) + 50) \ 100
                                ElseIf tgSpf.sSAudData = "N" Then
                                    tmDrfPop.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilPopIndex)) + 5) \ 10
                                ElseIf tgSpf.sSAudData = "U" Then
                                    tmDrfPop.lDemo(ilDemoGender - 1) = CLng(smFieldValues(ilPopIndex))
                                Else
                                    tmDrfPop.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilPopIndex)) + 500) \ 1000
                                End If
                                For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
                                    smFieldValues(2) = Mid$(slLine, tmDrfInfo(ilCol).iStartCol, 10)
                                    ilAvgQHIndex = 2
'                                    If tgSpf.sSAudData <> "H" Then
'                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
'                                    Else
'                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex)) + 50) \ 100
'                                    End If
                                    If tgSpf.sSAudData = "H" Then
                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 50) \ 100
                                    ElseIf tgSpf.sSAudData = "N" Then
                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 5) \ 10
                                    ElseIf tgSpf.sSAudData = "U" Then
                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = CLng(smFieldValues(ilAvgQHIndex))
                                    Else
                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
                                    End If
                                Next ilCol
                                ilNoDemoFd = ilNoDemoFd + 1
                            Else
                                'Print #hmTo, "Unable to find Demo " & smFieldValues(1)
                                'lbcErrors.AddItem "Unable to find Demo " & smFieldValues(1)
                            End If
                        End If
                    Else
                        'tmDrfPop.lDemo(ilDemoGender) = (CLng(smFieldValues(ilPopIndex)) + 500) \ 1000
                        'tmDrf.lDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
                    End If
                    If (ilNoDemoFd >= 16) Or ((ilNoDemoFd > 0) And (InStr(1, slLine, "### Act1", 1) = 1)) Then
                        ilHeaderFd = False
                        '6/9/16: Replaced GoSub
                        'GoSub mWriteRec
                        mWriteRec ilRet, ilCustomDemo, slVehicleName, slDPBookName, slETBookName, slBookDate, ilPrevDnfCode, llPrevDrfCode
                        If ilRet <> 0 Then
                            mOnLineConvFile = False
                            Exit Function
                        End If
                        ilNoDemoFd = 0
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
            If plcGauge.Value <> llPercent Then
                plcGauge.Value = llPercent
            End If
        End If
    Loop Until ilEof
    If (ilHeaderFd) And (ilNoDemoFd > 0) Then
        ilHeaderFd = False
        '6/9/16: Replaced GoSub
        'GoSub mWriteRec
        mWriteRec ilRet, ilCustomDemo, slVehicleName, slDPBookName, slETBookName, slBookDate, ilPrevDnfCode, llPrevDrfCode
        If ilRet <> 0 Then
            mOnLineConvFile = False
            Exit Function
        End If
    End If
    Close hmFrom
    plcGauge.Value = 100
    mOnLineConvFile = True
    MousePointer = vbDefault
    Exit Function
'mOnLineConvFileErr:
'    ilRet = Err.Number
'    Resume Next

'mWriteRec:
'    If ilCustomDemo Then
'        Return
'        'tmDrfPop.sDataType = "B"
'        'tmDrf.sDataType = "B"
'    End If
'    ilRet = 0
'    For ilVef = LBound(imVefCodeImpt) To UBound(imVefCodeImpt) - 1 Step 1
'        imVefCode = imVefCodeImpt(ilVef)
'        'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
'        '    If tgMVef(ilIndex).iCode = imVefCode Then
'            ilIndex = gBinarySearchVef(imVefCode)
'            If ilIndex <> -1 Then
'                slVehicleName = Trim$(tgMVef(ilIndex).sName)
'        '        Exit For
'            End If
'        'Next ilIndex
'        ilSetBkNm = -1
'        For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
'            If ((tmDrfInfo(ilCol).iBkNm = 0) And (Trim$(slDPBookName) <> "")) Or ((tmDrfInfo(ilCol).iBkNm = 1) And (Trim$(slETBookName) <> "")) Then
'
'                tmDrfInfo(ilCol).tDrf.iVefCode = imVefCode
'                If tmDrfInfo(ilCol).tDrf.sInfoType = "D" Then
'                    mDPPop
'                    For ilIndex = LBound(tmRdfInfo) To UBound(tmRdfInfo) - 1 Step 1
'                        ilRdf = tmRdfInfo(ilIndex).iRdfIndex
'                        For ilRow = 1 To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
'                            If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
'                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilRow), tgMRdf(ilRdf).iStartTime(1, ilRow), False, llTime
'                                If llTime = tmDrfInfo(ilCol).lStartTime Then
'                                    gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilRow), tgMRdf(ilRdf).iEndTime(1, ilRow), False, llTime
'                                    If llTime = tmDrfInfo(ilCol).lEndTime Then
'                                        ilMatch = True
'                                        For ilDay = 1 To 7 Step 1
'                                            If tgMRdf(ilRdf).sWkDays(ilRow, ilDay) <> tmDrfInfo(ilCol).sDays(ilDay - 1) Then
'                                                ilMatch = False
'                                                Exit For
'                                            End If
'                                        Next ilDay
'                                        If ilMatch Then
'                                            tmDrfInfo(ilCol).tDrf.iRdfcode = tgMRdf(ilRdf).iCode
'                                        End If
'                                        Exit For
'                                    End If
'                                End If
'                                Exit For
'                            End If
'                        Next ilRow
'                    Next ilIndex
'                End If
'                For ilDay = 0 To 6 Step 1
'                    smDays(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
'                Next ilDay
'                If tmDrfInfo(ilCol).iBkNm = 0 Then
'                    ilBNRet = mBookNameUsed(slDPBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfcode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
'                Else
'                    ilBNRet = mBookNameUsed(slETBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfcode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
'                End If
'                If (ilBNRet = 0) Or (ilBNRet = 1) Then
'                    tmDnf.iCode = 0
'                    If tmDrfInfo(ilCol).iBkNm = 0 Then
'                        tmDnf.sBookName = slDPBookName
'                    Else
'                        tmDnf.sBookName = slETBookName
'                    End If
'                    gPackDate slBookDate, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
'                    gPackDate smNowDate, tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
'                    tmDnf.iUrfCode = tgUrf(0).iCode
'                    tmDnf.sType = "I"
'                    tmDnf.sForm = smDataForm
'                    If tmDrfInfo(ilCol).iBkNm = 0 Then
'                        tmDnf.sExactTime = "N"
'                    Else
'                        tmDnf.sExactTime = "Y"
'                    End If
'                    tmDnf.sSource = "A"
'                    tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
'                    tmDnf.iAutoCode = tmDnf.iCode
'                    ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                            ilRet = csiHandleValue(0, 7)
'                        End If
'                        Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
'                        lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
'                        'mOnLineConvFile = False
'                        'Exit Function
'                        Return
'                    End If
'                    'If tgSpf.sRemoteUsers = "Y" Then
'                        Do
'                            'tmDnfSrchKey.iCode = tmDnf.iCode
'                            'ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'                            tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
'                            tmDnf.iAutoCode = tmDnf.iCode
'                            gPackDate smSyncDate, tmDnf.iSyncDate(0), tmDnf.iSyncDate(1)
'                            gPackTime smSyncTime, tmDnf.iSyncTime(0), tmDnf.iSyncTime(1)
'                            ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
'                        Loop While ilRet = BTRV_ERR_CONFLICT
'                        If ilRet <> BTRV_ERR_NONE Then
'                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                                ilRet = csiHandleValue(0, 7)
'                            End If
'                            Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
'                            lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
'                            'mOnLineConvFile = False
'                            'Exit Function
'                            Return
'                        End If
'                    'End If
'                    ilPrevDnfCode = tmDnf.iCode
'                    ilRet = mObtainBookName()
'                    tmDrfPop.lCode = 0
'                    tmDrfPop.iDnfCode = ilPrevDnfCode
'                    tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
'                    tmDrfPop.lAutoCode = tmDrfPop.lCode
'                    ilRet = btrInsert(hmDrf, tmDrfPop, imDrfRecLen, INDEXKEY2)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                            ilRet = csiHandleValue(0, 7)
'                        End If
'                        Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'                        lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'                        'mOnLineConvFile = False
'                        'Exit Function
'                        Return
'                    End If
'                    'If tgSpf.sRemoteUsers = "Y" Then
'                        Do
'                            'tmDrfSrchKey2.lCode = tmDrfPop.lCode
'                            'ilRet = btrGetEqual(hmDrf, tmDrfPop, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
'                            tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
'                            tmDrfPop.lAutoCode = tmDrfPop.lCode
'                            gPackDate smSyncDate, tmDrfPop.iSyncDate(0), tmDrfPop.iSyncDate(1)
'                            gPackTime smSyncTime, tmDrfPop.iSyncTime(0), tmDrfPop.iSyncTime(1)
'                            ilRet = btrUpdate(hmDrf, tmDrfPop, imDrfRecLen)
'                        Loop While ilRet = BTRV_ERR_CONFLICT
'                        If ilRet <> BTRV_ERR_NONE Then
'                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                                ilRet = csiHandleValue(0, 7)
'                            End If
'                            Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'                            lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'                            'mOnLineConvFile = False
'                            'Exit Function
'                            Return
'                        End If
'                    'End If
'                Else
'                    If StrComp(smBookForm, smDataForm, vbTextCompare) <> 0 Then
'                        If (smDataForm = "8") Or (smBookForm = "8") Then
'                            If tmDrfInfo(ilCol).iBkNm = 0 Then
'                                Print #hmTo, slDPBookName & " for " & slVehicleName & " previously defined with different format then current import data"
'                            Else
'                                Print #hmTo, slETBookName & " for " & slVehicleName & " previously defined with different format then current import data"
'                            End If
'                            lbcErrors.AddItem "Error in Book Forms" & " for " & slVehicleName
'                            ilRet = 0
'                            Return
'                        End If
'                    End If
'                    If ilBNRet = 3 Then
'                        Screen.MousePointer = vbDefault
'                        If tmDrfInfo(ilCol).iBkNm = 0 Then
'                            ilQRet = MsgBox(slDPBookName & " for " & slVehicleName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
'                        Else
'                            ilQRet = MsgBox(slETBookName & " for " & slVehicleName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
'                        End If
'                        If ilQRet = vbNo Then
'                            'mOnLineConvFile = False
'                            'Exit Function
'                            ilRet = 0
'                            Return
'                        End If
'                        Print #hmTo, "Replaced Demo Data (DRF)"
'                    End If
'                End If
'                tmDrfInfo(ilCol).tDrf.iDnfCode = ilPrevDnfCode
'                If llPrevDrfCode = 0 Then
'                    tmDrfInfo(ilCol).tDrf.lCode = 0
'                    tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
'                    tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
'                    ilRet = btrInsert(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen, INDEXKEY2)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                            ilRet = csiHandleValue(0, 7)
'                        End If
'                        Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'                        lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'                        'mOnLineConvFile = False
'                        'Exit Function
'                        Return
'                    End If
'                    'If tgSpf.sRemoteUsers = "Y" Then
'                        Do
'                            'tmDrfSrchKey2.lCode = tmDrf.lCode
'                            'ilRet = btrGetEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
'                            tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
'                            tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
'                            gPackDate smSyncDate, tmDrfInfo(ilCol).tDrf.iSyncDate(0), tmDrfInfo(ilCol).tDrf.iSyncDate(1)
'                            gPackTime smSyncTime, tmDrfInfo(ilCol).tDrf.iSyncTime(0), tmDrfInfo(ilCol).tDrf.iSyncTime(1)
'                            ilRet = btrUpdate(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen)
'                        Loop While ilRet = BTRV_ERR_CONFLICT
'                     'End If
'                Else
'                    Do
'                        tmDrfSrchKey2.lCode = llPrevDrfCode
'                        ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'                        If ilRet <> BTRV_ERR_NONE Then
'                            Exit Do
'                        End If
'                        tmDrfInfo(ilCol).tDrf.lCode = tlDrf.lCode
'                        tmDrfInfo(ilCol).tDrf.iRemoteID = tlDrf.iRemoteID
'                        tmDrfInfo(ilCol).tDrf.lAutoCode = tlDrf.lAutoCode
'                        gPackDate smSyncDate, tmDrfInfo(ilCol).tDrf.iSyncDate(0), tmDrfInfo(ilCol).tDrf.iSyncDate(1)
'                        gPackTime smSyncTime, tmDrfInfo(ilCol).tDrf.iSyncTime(0), tmDrfInfo(ilCol).tDrf.iSyncTime(1)
'                        ilRet = btrUpdate(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen)
'                    Loop While ilRet = BTRV_ERR_CONFLICT
'                End If
'                If ilRet <> BTRV_ERR_NONE Then
'                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                        ilRet = csiHandleValue(0, 7)
'                    End If
'                    Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'                    lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'                    'mOnLineConvFile = False
'                    'Exit Function
'                    Return
'                End If
'                If tmDrfInfo(ilCol).iBkNm = 0 Then
'                    ilSetBkNm = 0
'                    ilSetDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
'                ElseIf (tmDrfInfo(ilCol).iBkNm = 1) And (ilSetBkNm = -1) Then
'                    ilSetBkNm = 1
'                    ilSetDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
'                End If
'            End If
'        Next ilCol
'        If ((ckcDefault(0).Value) Or (ckcDefault(1).Value)) And ((ilSetBkNm = 0) Or (ilSetBkNm = 1)) Then
'            Do
'                tmVefSrchKey.iCode = imVefCode
'                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'                If ilRet <> BTRV_ERR_NONE Then
'                    Exit Do
'                End If
'                If ckcDefault(0).Value = vbChecked Then
'                    tmVef.iDnfCode = ilSetDnfCode
'                End If
'                If ckcDefault(1).Value = vbChecked Then
'                    tmVef.iReallDnfCode = ilSetDnfCode
'                End If
'                'tmVef.iSourceID = tgUrf(0).iRemoteUserID
'                'gPackDate smSyncDate, tmVef.iSyncDate(0), tmVef.iSyncDate(1)
'                'gPackTime smSyncTime, tmVef.iSyncTime(0), tmVef.iSyncTime(1)
'                ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
'            Loop While ilRet = BTRV_ERR_CONFLICT
'            ilRet = gBinarySearchVef(tmVef.iCode)
'            If ilRet <> -1 Then
'                tgMVef(ilRet) = tmVef
'            End If
'        End If
'        Print #hmTo, "Successfully Installed Demo Data" & " for " & Trim$(slVehicleName)
'    Next ilVef
'    For ilVef = LBound(imVefCodeImpt) To UBound(imVefCodeImpt) - 1 Step 1
'        'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
'        '    If tgMVef(ilIndex).iCode = imVefCodeImpt(ilVef) Then
'            ilIndex = gBinarySearchVef(imVefCodeImpt(ilVef))
'            If ilIndex <> -1 Then
'                Print #hmTo, "Successfully Installed Demo Data" & " for " & Trim$(tgMVef(ilIndex).sName)
'        '        Exit For
'            End If
'        'Next ilIndex
'    Next ilVef
'    ilRet = 0
'    Return
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mPCNetConvFile                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert File                   *
'*                                                     *
'*******************************************************
Private Function mPCNetConvFile(slFromFile As String, slDPBookName As String, slETBookName As String, slBookDate As String) As Integer
    Dim ilRet As Integer
    Dim ilBNRet As Integer
    Dim ilQRet As Integer
    Dim slLine As String
    Dim ilHeaderFd As Integer
    Dim ilPopIndex As Integer
    Dim ilPopCol As Integer
    Dim ilAvgQHIndex As Integer
    Dim ilPopDone As Integer
    Dim ilDemoGender As Integer
    Dim slDemoAge As String
    Dim ilPos As Integer
    Dim ilEndPos As Integer
    Dim ilLastNonBlankCol As Integer
    Dim llLineCount As Long
    Dim slDay As String
    Dim ilDay As Integer
    Dim ilSY As Integer
    Dim ilEY As Integer
    Dim ilLoop As Integer
    Dim ilEof As Integer
    Dim llPercent As Long
    Dim slChar As String
    Dim slTime As String
    Dim slStr As String
    Dim slSexChar As String
    Dim ilIndex As Integer
    Dim ilEndIndex As Integer
    Dim ilCustomDemo As Integer
    Dim ilPrevDnfCode As Integer
    Dim llPrevDrfCode As Long
    Dim ilRdf As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim slStationCode As String
    Dim slVehicleName As String
    Dim ilNoDemoFd As Integer
    Dim ilMatch As Integer
    Dim llTime As Long
    Dim ilVef As Integer
    Dim ilVff As Integer
    Dim ilTIndex As Integer
    Dim ilSetBkNm As Integer
    Dim ilSetDnfCode As Integer
    Dim ilDpfUpper As Integer
    Dim ilDpf As Integer
    Dim ilPRet As Integer
    Dim slDemoName As String
    Dim blDontBlankLine As Boolean
    Dim tlDrf As DRF
    ilRet = 0
    'On Error GoTo mPCNetConvFileErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        MsgBox "Open " & slFromFile & ", Error #" & Str$(ilRet), vbExclamation, "Open Error"
        edcFrom.SetFocus
        mPCNetConvFile = False
        Exit Function
    End If
    DoEvents
    If imTerminate Then
        Close hmFrom
        mTerminate
        mPCNetConvFile = False
        Exit Function
    End If
    ilHeaderFd = False
    blDontBlankLine = False
    slLine = ""
    llLineCount = 0
    ilPopDone = False
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
    tmDrfPop.sACTLineupCode = ""
    tmDrfPop.sACT1StoredTime = ""
    tmDrfPop.sACT1StoredSpots = ""
    tmDrfPop.sACT1StoreClearPct = ""
    tmDrfPop.sACT1DaypartFilter = ""
    
    ilNoDemoFd = 0
    ilCustomDemo = False
    Do
        ilRet = 0
        err.Clear
        'On Error GoTo mPCNetConvFileErr:
        'Line Input #hmFrom, slLine
        slLine = ""
        Do
            If Not EOF(hmFrom) Then
                slChar = Input(1, #hmFrom)
                If slChar = Chr(13) Then
                    slChar = Input(1, #hmFrom)
                End If
                If slChar = Chr(10) Then
                    Exit Do
                End If
                slLine = slLine & slChar
            Else
                ilEof = True
                Exit Do
            End If
        Loop
        llLineCount = llLineCount + 1
        slLine = Trim$(slLine)
        On Error GoTo 0
        ilRet = err.Number
        If ilRet = 62 Then
            Exit Do
        End If
        If llLineCount >= 120 Then
            ilRet = ilRet
        End If
        '2/26/18: Filter out lines without research data
        If right(slLine, 1) = "." Then
            slLine = ""
        ElseIf right(slLine, 2) = "." & """" Then
            slLine = ""
        Else
            If (InStr(1, slLine, ",,,,,") > 0) Then
                'If (InStr(1, slLine, ":") > 0) And (InStr(1, slLine, "),") > 0) Then
                If (InStr(1, slLine, ":") <= 0) Or (InStr(1, slLine, "),") <= 0) Then
                    If blDontBlankLine Or (InStr(1, slLine, ":") > 0) And (InStr(1, slLine, "),") <= 0) Or (InStr(1, slLine, ",,,Total") > 0) Or (InStr(1, slLine, "Stored Schedule") > 0) Or (InStr(1, slLine, "Schedule") > 0) Then
                    Else
                        slLine = ""
                    End If
                Else
                    blDontBlankLine = True
                End If
            End If
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                DoEvents
                If imTerminate Then
                    Close hmFrom
                    mTerminate
                    mPCNetConvFile = False
                    Exit Function
                End If
                ilLastNonBlankCol = -1
                gParseCDFields slLine, False, smFieldValues()
                '11/10/16: Move fields from 0 to X to 1 to X +1
                For ilLoop = UBound(smFieldValues) - 1 To LBound(smFieldValues) Step -1
                    smFieldValues(ilLoop + 1) = smFieldValues(ilLoop)
                    If Trim(smFieldValues(ilLoop + 1)) <> "" Then
                        If ilLoop + 1 > ilLastNonBlankCol Then
                            ilLastNonBlankCol = ilLoop + 1
                        End If
                    End If
                Next ilLoop
                smFieldValues(0) = ""
                'Determine field Type
                'Header Record by Daypart
                'RNMF: ReachNet M-F - Sept. '18 (9/04/18),,,,,,,
                ',MF 6a-7p,,,,,Total,
                ',,,,,,United States,
                '   Demographics,AQH,AQH,....,Conv%,Population,,
                '   ,,Rtg.,,.....
                
                'Header Record by Exact time
                'DLHS: D.L. Hughley Show - Sept. '18 (9/17/18),,,,,,,
                ',Stored Schedules,,,,,Total,
                '   Demographics,AQH,AQH,....,Conv%,Population,,
                '   ,,Rtg.,,.....
                 
                 'Demo Record
                '   Males   12-17, 12000, 0.1,..., 21847000
                '   Females 12-17,
                '   Men     18-24,
                '   Women   18-24,
                '
                If Not ilHeaderFd Then
                    If InStr(1, Trim$(slLine), "Demographic", 1) > 0 Then
                        If InStr(1, RTrim$(slLine), "AQH", 1) > 0 Then
                            If InStr(1, RTrim$(slLine), "Population", 1) > 0 Then
                                'Determine number of Demo columns
                                ilPos = 2
                                ReDim tmDrfInfo(0 To 0) As DRFINFO
                                ReDim tgDpfInfo(0 To 0) As DPFINFO
                                Do While ilPos <= UBound(smFieldValues)
                                    If InStr(1, Trim$(smFieldValues(ilPos)), "AQH", 1) > 0 Then
                                        tmDrfInfo(UBound(tmDrfInfo)).iStartCol = ilPos
                                        tmDrfInfo(UBound(tmDrfInfo)).iType = 0  'Vehicle
                                        tmDrfInfo(UBound(tmDrfInfo)).iBkNm = 0  'Daypart Book Name
                                        tmDrfInfo(UBound(tmDrfInfo)).iSY = -1
                                        tmDrfInfo(UBound(tmDrfInfo)).iEY = -1
                                        ReDim Preserve tmDrfInfo(0 To UBound(tmDrfInfo) + 1) As DRFINFO
                                        ilPos = ilPos + 3
                                    ElseIf InStr(1, Trim$(smFieldValues(ilPos)), "Population", 1) > 0 Then
                                        ilPopCol = ilPos
                                        Exit Do
                                    ElseIf InStr(1, Trim$(smFieldValues(ilPos)), "Cov%", 1) > 0 Then
                                        ilPopCol = ilPos + 1
                                        Exit Do
                                    Else
                                        ilPos = ilPos + 1
                                    End If
                                Loop
                                'Determine Vehicle and Daypart info
                                'For ilLoop = 10 To 1 Step -1
                                For ilLoop = 9 To 0 Step -1
                                    'If InStr(1, smSvLine(ilLoop), "COVERAGE", 1) > 0 Then
                                    '    slStationCode = ""
                                    '    slVehicleName = ""
                                    '    Exit For
                                    'End If
                                    If Len(Trim$(smSvLine(ilLoop))) > 0 Then
                                        gParseCDFields smSvLine(ilLoop), False, smSvFields()
                                        '11/10/16: Move fields from 0 to X to 1 to X +1
                                        For ilCol = UBound(smSvFields) - 1 To LBound(smSvFields) Step -1
                                            smSvFields(ilCol + 1) = smSvFields(ilCol)
                                        Next ilCol
                                        smSvFields(0) = ""
                                        ilPos = InStr(1, smSvFields(1), ":", 1)
                                        If ilPos > 0 Then
                                            slStationCode = ""
                                            slVehicleName = ""
                                            slStationCode = Left$(smSvFields(1), ilPos - 1)
                                            If InStr(1, slStationCode, ".", 1) Then
                                                slStationCode = Left(slStationCode, InStr(1, slStationCode, ".", 1) - 1)
                                            End If
                                            ilRow = InStr(1, smSvFields(1), "(", 1)
                                            If ilRow > 0 Then
                                                slVehicleName = Trim$(Mid$(smSvFields(1), ilPos + 1, ilRow - ilPos - 1))
                                            Else
                                                slVehicleName = Trim$(Mid$(smSvFields(1), ilPos + 1))
                                            End If
                                            Exit For
                                        Else
                                            If (InStr(1, smSvLine(ilLoop), "Schedule", 1) >= 1) And (InStr(1, smSvLine(ilLoop), "Stored Schedule:", 1) <= 0) Then
                                                gParseCDFields smSvLine(ilLoop), False, smSvFields()
                                                '-: Move fields from 0 to X to 1 to X +1
                                                For ilCol = UBound(smSvFields) - 1 To LBound(smSvFields) Step -1
                                                    smSvFields(ilCol + 1) = smSvFields(ilCol)
                                                Next ilCol
                                                smSvFields(0) = ""
                                                For ilRow = LBound(smSvFields) + 1 To UBound(smSvFields) - 1 Step 1
                                                    If (InStr(1, smSvFields(ilRow), "Schedule", 1) >= 1) Then
                                                        'ilSY = -1
                                                        'ilEY = -1
                                                        ''Get Days and Times
                                                        'slStr = UCase$(Trim$(smSvFields(ilRow)))
                                                        'slDay = Mid$(slStr, 1, 2)
                                                        'Select Case slDay
                                                        '    Case "MO"
                                                        '        ilSY = 0
                                                        '    Case "TU"
                                                        '        ilSY = 1
                                                        '    Case "WE"
                                                        '        ilSY = 2
                                                        '    Case "TH"
                                                        '        ilSY = 3
                                                        '    Case "FR"
                                                        '        ilSY = 4
                                                        '    Case "SA"
                                                        '        ilSY = 5
                                                        '    Case "SU"
                                                        '        ilSY = 6
                                                        'End Select
                                                        'slStr = Mid$(slStr, 3)
                                                        'slDay = Mid$(slStr, 1, 2)
                                                        'Select Case slDay
                                                        '    Case "MO"
                                                        '        ilEY = 0
                                                        '    Case "TU"
                                                        '        ilEY = 1
                                                        '    Case "WE"
                                                        '        ilEY = 2
                                                        '    Case "TH"
                                                        '        ilEY = 3
                                                        '    Case "FR"
                                                        '        ilEY = 4
                                                        '    Case "SA"
                                                        '        ilEY = 5
                                                        '    Case "SU"
                                                        '        ilEY = 6
                                                        'End Select
                                                        slStr = UCase$(Trim$(smSvFields(ilRow)))
                                                        mGetDayIndex slStr, ilSY, ilEY
                                                        If (ilSY <> -1) And (ilEY <> -1) Then
                                                            For ilDay = 0 To 6 Step 1
                                                                smDays(ilDay) = "N"
                                                            Next ilDay
                                                            For ilDay = ilSY To ilEY Step 1
                                                                smDays(ilDay) = "Y"
                                                            Next ilDay
                                                            For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
                                                                tmDrfInfo(ilCol).iType = 0  'Vehicle
                                                                For ilDay = 0 To 6 Step 1
                                                                    tmDrfInfo(ilCol).sDays(ilDay) = smDays(ilDay)
                                                                Next ilDay
                                                            Next ilCol
                                                        End If
                                                    End If
                                                Next ilRow
                                            End If
                                            For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
                                                slStr = UCase$(Trim$(smSvFields(tmDrfInfo(ilCol).iStartCol)))
                                                If InStr(1, slStr, "Stored Schedule", 1) > 0 Then
                                                    tmDrfInfo(ilCol).iBkNm = 1  'Exact Time
                                                    tmDrfInfo(ilCol).iType = 0  'Vehicle
                                                    For ilDay = 0 To 6 Step 1
                                                        tmDrfInfo(ilCol).sDays(ilDay) = "Y"
                                                    Next ilDay
                                                Else
                                                    If tmDrfInfo(ilCol).iSY = -1 Then
                                                        'slDay = Trim$(Mid$(slStr, 1, 3))
                                                        'ilTIndex = InStr(1, slStr, " ", 1) + 1
                                                        'Select Case slDay
                                                        '    Case "MO", "MON"
                                                        '        ilSY = 0
                                                        '        ilEY = 0
                                                        '    Case "TU", "TUE"
                                                        '        ilSY = 1
                                                        '        ilEY = 1
                                                        '    Case "WE", "WED"
                                                        '        ilSY = 2
                                                        '        ilEY = 2
                                                        '    Case "TH", "THU"
                                                        '        ilSY = 3
                                                        '        ilEY = 3
                                                        '    Case "FR", "FRI"
                                                        '        ilSY = 4
                                                        '        ilEY = 4
                                                        '    Case "SA", "SAT"
                                                        '        ilSY = 5
                                                        '        ilEY = 5
                                                        '        If Mid$(slStr, 3, 2) = "SU" Then
                                                        '            ilEY = 6
                                                        '        End If
                                                        '    Case "SU", "SUN"
                                                        '        ilSY = 6
                                                        '        ilEY = 6
                                                        '    Case "MF"
                                                        '        ilSY = 0
                                                        '        ilEY = 4
                                                        '    Case "MSA"
                                                        '        ilSY = 0
                                                        '        ilEY = 5
                                                        '    Case "MSU"
                                                        '        ilSY = 0
                                                        '        ilEY = 6
                                                        '    Case "MS"
                                                        '        ilSY = 0
                                                        '        ilEY = 6
                                                        '    Case "SS"
                                                        '        ilSY = 5
                                                        '        ilEY = 6
                                                        '    Case "FSU"
                                                        '        ilSY = 4
                                                        '        ilEY = 6
                                                        '    Case "FSA"
                                                        '        ilSY = 4
                                                        '        ilEY = 5
                                                        'End Select
                                                        ilSY = -1
                                                        ilTIndex = InStr(1, slStr, " ", 1) + 1
                                                        If ilTIndex > 2 Then
                                                            mGetDayIndex Left$(slStr, ilTIndex - 2), ilSY, ilEY
                                                        End If
                                                        If ilSY <> -1 Then
                                                            slChar = Mid$(slStr, ilTIndex, 1)
                                                            If (slChar >= "0") And (slChar <= "9") Then
                                                                slStr = Mid$(slStr, ilTIndex)
                                                            Else
                                                                slStr = Mid$(slStr, ilTIndex + 1)
                                                            End If
                                                            slTime = Mid$(slStr, 1, 1)
                                                            slStr = Mid$(slStr, 2)
                                                            slChar = Mid$(slStr, 1, 1)
                                                            If (slChar >= "0") And (slChar <= "9") Then
                                                                slTime = slTime & Mid$(slStr, 1, 2)
                                                                slStr = Mid$(slStr, 3)
                                                            Else
                                                                If Mid$(slStr, 1, 1) = "-" Then
                                                                    slTime = slTime & right$(slStr, 1)
                                                                Else
                                                                    slTime = slTime & Mid$(slStr, 1, 1)
                                                                    slStr = Mid$(slStr, 2)
                                                                End If
                                                            End If
                                                            If gValidTime(slTime) Then
                                                                tmDrfInfo(ilCol).lStartTime = gTimeToCurrency(slTime, False)
                                                                slChar = Mid$(slStr, 1, 1)
                                                                If slChar = "-" Then
                                                                    slStr = Trim$(Mid$(slStr, 2))
                                                                End If
                                                                slTime = Mid$(slStr, 1, 1)
                                                                slStr = Mid$(slStr, 2)
                                                                slChar = Mid$(slStr, 1, 1)
                                                                If (slChar >= "0") And (slChar <= "9") Then
                                                                    slTime = slTime & Mid$(slStr, 1, 2)
                                                                    slStr = Mid$(slStr, 3)
                                                                Else
                                                                    slTime = slTime & Mid$(slStr, 1, 1)
                                                                    slStr = Mid$(slStr, 2)
                                                                End If
                                                                If gValidTime(slTime) Then
                                                                    tmDrfInfo(ilCol).iSY = ilSY
                                                                    tmDrfInfo(ilCol).iEY = ilEY
                                                                    tmDrfInfo(ilCol).lEndTime = gTimeToCurrency(slTime, False)
                                                                    For ilDay = 0 To 6 Step 1
                                                                        tmDrfInfo(ilCol).sDays(ilDay) = "N"
                                                                    Next ilDay
                                                                    For ilDay = ilSY To ilEY Step 1
                                                                        tmDrfInfo(ilCol).sDays(ilDay) = "Y"
                                                                    Next ilDay
                                                                    tmDrfInfo(ilCol).iType = 1  'Daypart
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next ilCol
                                        End If
                                    End If
                                Next ilLoop
                                If slVehicleName <> "" Then
                                    'Test if vehicle name matches
                                    imVehSelectedIndex = -1
                                    'For ilLoop = LBound(tgUserVehicle) To UBound(tgUserVehicle) - 1 Step 1
                                    '    slNameCode = tgUserVehicle(ilLoop).sKey    'lbcMster.List(ilLoop)
                                    '    ilRet = gParseItem(slNameCode, 1, "\", slName)
                                    '    ilRet = gParseItem(slName, 3, "|", slName)
                                    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    '    If StrComp(slVehicleName, slName, 1) = 0 Then
                                    '        ilHeaderFd = True
                                    '        imVefCode = Val(slCode)
                                    '        cbcVehicle.ListIndex = ilLoop
                                    '        Exit For
                                    '    End If
                                    'Next ilLoop
                                    ReDim imVefCodeImpt(0 To 0) As Integer
                                    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                        If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Or (tgMVef(ilLoop).sType = "R") Or (tgMVef(ilLoop).sType = "V") Then
                                            If ((slStationCode <> "") And (StrComp(UCase(slStationCode), UCase(Trim$(tgMVef(ilLoop).sCodeStn)), 1) = 0)) Or (StrComp(UCase(slVehicleName), UCase(Trim$(tgMVef(ilLoop).sName)), 1) = 0) Then
                                                imVehSelectedIndex = ilLoop
                                                ilHeaderFd = True
                                                imVefCodeImpt(UBound(imVefCodeImpt)) = tgMVef(ilLoop).iCode
                                                ReDim Preserve imVefCodeImpt(0 To UBound(imVefCodeImpt) + 1) As Integer
                                            Else
                                                ilVff = gBinarySearchVff(tgMVef(ilLoop).iCode)
                                                If ilVff <> -1 Then
                                                    If ((slStationCode <> "") And (StrComp(UCase(slStationCode), UCase(Trim$(tgVff(ilVff).sACT1LineupCode)), 1) = 0)) Then
                                                        imVehSelectedIndex = ilLoop
                                                        ilHeaderFd = True
                                                        imVefCodeImpt(UBound(imVefCodeImpt)) = tgMVef(ilLoop).iCode
                                                        ReDim Preserve imVefCodeImpt(0 To UBound(imVefCodeImpt) + 1) As Integer
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next ilLoop
                                    If UBound(imVefCodeImpt) <= 0 Then
                                        Print #hmTo, slVehicleName & " Not Found"
                                        lbcErrors.AddItem slVehicleName & " Not Found"
                                    Else
                                        ilNoDemoFd = 0
                                        For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
                                            tmDrfInfo(ilCol).tDrf.lCode = 0
                                            tmDrfInfo(ilCol).tDrf.iDnfCode = tmDnf.iCode
                                            tmDrfInfo(ilCol).tDrf.sDemoDataType = "D"
                                            tmDrfInfo(ilCol).tDrf.iMnfSocEco = 0
                                            If tmDrfInfo(ilCol).iType = 0 Then
                                                tmDrfInfo(ilCol).tDrf.sInfoType = "V"
                                                tmDrfInfo(ilCol).tDrf.iRdfCode = 0
                                                gPackTime "12AM", tmDrfInfo(ilCol).tDrf.iStartTime(0), tmDrfInfo(ilCol).tDrf.iStartTime(1)
                                                gPackTime "12AM", tmDrfInfo(ilCol).tDrf.iEndTime(0), tmDrfInfo(ilCol).tDrf.iEndTime(1)
                                                For ilDay = 0 To 6 Step 1
                                                    tmDrfInfo(ilCol).tDrf.sDay(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
                                                Next ilDay
                                            Else
                                                tmDrfInfo(ilCol).tDrf.sInfoType = "D"
                                                tmDrfInfo(ilCol).tDrf.iRdfCode = 0
                                                gPackTimeLong tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).tDrf.iStartTime(0), tmDrfInfo(ilCol).tDrf.iStartTime(1)
                                                gPackTimeLong tmDrfInfo(ilCol).lEndTime, tmDrfInfo(ilCol).tDrf.iEndTime(0), tmDrfInfo(ilCol).tDrf.iEndTime(1)
                                                For ilDay = 0 To 6 Step 1
                                                    tmDrfInfo(ilCol).tDrf.sDay(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
                                                Next ilDay
                                            End If
                                            tmDrfInfo(ilCol).tDrf.sProgCode = ""
                                            tmDrfInfo(ilCol).tDrf.iStartTime2(0) = 1
                                            tmDrfInfo(ilCol).tDrf.iStartTime2(1) = 0
                                            tmDrfInfo(ilCol).tDrf.iEndTime2(0) = 1
                                            tmDrfInfo(ilCol).tDrf.iEndTime2(1) = 0
                                            tmDrfInfo(ilCol).tDrf.iQHIndex = 0
                                            tmDrfInfo(ilCol).tDrf.iCount = 0
                                            tmDrfInfo(ilCol).tDrf.sExStdDP = "N"
                                            tmDrfInfo(ilCol).tDrf.sExRpt = "N"
                                            tmDrfInfo(ilCol).tDrf.sDataType = "A"
                                            For ilLoop = 1 To 16 Step 1
                                                tmDrfInfo(ilCol).tDrf.lDemo(ilLoop - 1) = 0
                                            Next ilLoop
                                        Next ilCol
                                    End If
                                End If
                                'For ilLoop = 1 To 10 Step 1
                                For ilLoop = 0 To 9 Step 1
                                    smSvLine(ilLoop) = ""
                                Next ilLoop
                            Else
                                'Save Last 10 lines
                                'For ilLoop = 2 To 10 Step 1
                                For ilLoop = 1 To 9 Step 1
                                    smSvLine(ilLoop - 1) = smSvLine(ilLoop)
                                Next ilLoop
                                'smSvLine(10) = slLine
                                smSvLine(9) = slLine
                            End If
                        Else
                            'Save Last 10 lines
                            'For ilLoop = 2 To 10 Step 1
                            For ilLoop = 1 To 9 Step 1
                                smSvLine(ilLoop - 1) = smSvLine(ilLoop)
                            Next ilLoop
                            'smSvLine(10) = slLine
                            smSvLine(9) = slLine
                        End If
                    Else
                        'Save Last 10 lines
                        'For ilLoop = 2 To 10 Step 1
                        For ilLoop = 1 To 9 Step 1
                            smSvLine(ilLoop - 1) = smSvLine(ilLoop)
                        Next ilLoop
                        'smSvLine(10) = slLine
                        smSvLine(9) = slLine
                    End If
                Else
                    ilDemoGender = -1
                    ilPos = -1
                    'Test if Custom Demo
                    'Test if Custom Demo
                    ilRow = InStr(1, smFieldValues(1), ":", 1)
                    If (ilRow <= 0) And (Len(Trim$(smFieldValues(1))) > 0) And (ilLastNonBlankCol >= ilPopCol) Then
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
                            ElseIf (slSexChar = "P") Or (slSexChar = "T") Then
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
                                '2/16/21 - TTP 10080 - found start position, Check for end position...
                                ilEndPos = 0
                                ilEndIndex = Len(smFieldValues(1))
                                Do While ilEndIndex > 1
                                    slChar = Mid$(smFieldValues(1), ilEndIndex, 1)
                                    If ((slChar >= "0") And (slChar <= "9")) Or (slChar = "+") Then
                                        ilEndPos = ilEndIndex
                                        Exit Do
                                    Else
                                        'not a number or special character
                                    End If
                                    ilEndIndex = ilEndIndex - 1
                                Loop
                                'Truncate DemoAge to remove any trailing characters
                                If ilEndPos < Len(smFieldValues(1)) And ilEndPos <> 0 Then
                                    smFieldValues(1) = Trim$(Mid$(smFieldValues(1), 1, ilEndPos))
                                End If
                                
                                If smDataForm <> "8" Then
                                    If (slSexChar = "M") Or (slSexChar = "B") Then
                                        ilDemoGender = 0
                                        slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
                                    ElseIf (slSexChar = "W") Or (slSexChar = "G") Then
                                        ilDemoGender = 8
                                        slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
                                    ElseIf (slSexChar = "T") Or (slSexChar = "P") Then
                                        ilDemoGender = -1
                                        slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
                                    End If
                                    If ilDemoGender >= 0 Then
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
                                    End If
                                Else
                                    If (slSexChar = "M") Or (slSexChar = "B") Then
                                        ilDemoGender = 0
                                        slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
                                    ElseIf (slSexChar = "W") Or (slSexChar = "G") Then
                                        ilDemoGender = 9
                                        slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
                                    ElseIf (slSexChar = "T") Or (slSexChar = "P") Then
                                        ilDemoGender = -1
                                        slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
                                    End If
                                    If ilDemoGender >= 0 Then
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
                                End If  'dataForm not 8
                                If ilDemoGender > 0 Then
                                    ilPopIndex = ilPopCol
'                                    If tgSpf.sSAudData <> "H" Then
'                                        tmDrfPop.lDemo(ilDemoGender) = (CLng(smFieldValues(ilPopIndex)) + 500) \ 1000
'                                    Else
'                                        tmDrfPop.lDemo(ilDemoGender) = (CLng(smFieldValues(ilPopIndex)) + 50) \ 100
'                                    End If
                                    If tgSpf.sSAudData = "H" Then
                                        tmDrfPop.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilPopIndex)) + 50) \ 100
                                    ElseIf tgSpf.sSAudData = "N" Then
                                        tmDrfPop.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilPopIndex)) + 5) \ 10
                                    ElseIf tgSpf.sSAudData = "U" Then
                                        tmDrfPop.lDemo(ilDemoGender - 1) = CLng(smFieldValues(ilPopIndex))
                                    Else
                                        tmDrfPop.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilPopIndex)) + 500) \ 1000
                                    End If
                                    For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
                                        ilAvgQHIndex = tmDrfInfo(ilCol).iStartCol
'                                        If tgSpf.sSAudData <> "H" Then
'                                            tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
'                                        Else
'                                            tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex)) + 50) \ 100
'                                        End If
                                        If tgSpf.sSAudData = "H" Then
                                            tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 50) \ 100
                                        ElseIf tgSpf.sSAudData = "N" Then
                                            tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 5) \ 10
                                        ElseIf tgSpf.sSAudData = "U" Then
                                            tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = CLng(smFieldValues(ilAvgQHIndex))
                                        Else
                                            tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
                                        End If
                                    Next ilCol
                                    ilNoDemoFd = ilNoDemoFd + 1
                                Else
                                    'Print #hmTo, "Unable to find Demo " & smFieldValues(1)
                                    'lbcErrors.AddItem "Unable to find Demo " & smFieldValues(1)
                                    'Plus demo?
                                    ilDpfUpper = -1
                                    If (slSexChar = "M") Or (slSexChar = "B") Then
                                        slDemoName = "M" & slDemoAge
                                    ElseIf (slSexChar = "W") Or (slSexChar = "G") Then
                                        slDemoName = "W" & slDemoAge
                                    ElseIf (slSexChar = "T") Or (slSexChar = "P") Then
                                        If InStr(1, slDemoAge, "12", 1) > 0 Then
                                            slDemoName = "P" & slDemoAge
                                        Else
                                            slDemoName = "A" & slDemoAge
                                        End If
                                    End If
                                    slDemoName = Trim$(slDemoName)
                                    For ilLoop = LBound(tgMnfSDemo) To UBound(tgMnfSDemo) Step 1
                                        If StrComp(Trim$(tgMnfSDemo(ilLoop).sName), slDemoName, 1) = 0 Then
                                            ilDpfUpper = UBound(tgDpfInfo)
                                            Exit For
                                        End If
                                    Next ilLoop
                                    If ilDpfUpper = -1 Then
                                        'Test if valid name like M18-49
                                        If (Len(slDemoName) >= 4) And Len(slDemoName) <= 6 Then
                                            If (Mid$(slDemoName, 4, 1) = "+") Or (Mid$(slDemoName, 4, 1) = "-") Then
                                                ilRet = mAddDemo(slDemoName)
                                            End If
                                        End If
                                    End If
                                    For ilLoop = LBound(tgMnfSDemo) To UBound(tgMnfSDemo) Step 1
                                        If StrComp(Trim$(tgMnfSDemo(ilLoop).sName), slDemoName, 1) = 0 Then
                                            For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
                                                ilDpfUpper = UBound(tgDpfInfo)
                                                tgDpfInfo(ilDpfUpper).iCol = ilCol
                                                tgDpfInfo(ilDpfUpper).tDpf.iMnfDemo = tgMnfSDemo(ilLoop).iCode
'                                                If tgSpf.sSAudData <> "H" Then
'                                                    tgDpfInfo(ilDpfUpper).tDpf.lPop = (CLng(smFieldValues(ilPopCol)) + 500) \ 1000
'                                                Else
'                                                    tgDpfInfo(ilDpfUpper).tDpf.lPop = (CLng(smFieldValues(ilPopCol)) + 50) \ 100
'                                                End If
                                                If tgSpf.sSAudData = "H" Then
                                                    tgDpfInfo(ilDpfUpper).tDpf.lPop = (CLng(smFieldValues(ilPopCol)) + 50) \ 100
                                                ElseIf tgSpf.sSAudData = "N" Then
                                                    tgDpfInfo(ilDpfUpper).tDpf.lPop = (CLng(smFieldValues(ilPopCol)) + 5) \ 10
                                                ElseIf tgSpf.sSAudData = "U" Then
                                                    tgDpfInfo(ilDpfUpper).tDpf.lPop = CLng(smFieldValues(ilPopCol))
                                                Else
                                                    tgDpfInfo(ilDpfUpper).tDpf.lPop = (CLng(smFieldValues(ilPopCol)) + 500) \ 1000
                                                End If
                                                ilAvgQHIndex = tmDrfInfo(ilCol).iStartCol
'                                                If tgSpf.sSAudData <> "H" Then
'                                                    tgDpfInfo(ilDpfUpper).tDpf.lDemo = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
'                                                Else
'                                                    tgDpfInfo(ilDpfUpper).tDpf.lDemo = (CLng(smFieldValues(ilAvgQHIndex)) + 50) \ 100
'                                                End If
                                                If tgSpf.sSAudData = "H" Then
                                                    tgDpfInfo(ilDpfUpper).tDpf.lDemo = (CLng(smFieldValues(ilAvgQHIndex)) + 50) \ 100
                                                ElseIf tgSpf.sSAudData = "N" Then
                                                    tgDpfInfo(ilDpfUpper).tDpf.lDemo = (CLng(smFieldValues(ilAvgQHIndex)) + 5) \ 10
                                                ElseIf tgSpf.sSAudData = "U" Then
                                                    tgDpfInfo(ilDpfUpper).tDpf.lDemo = CLng(smFieldValues(ilAvgQHIndex))
                                                Else
                                                    tgDpfInfo(ilDpfUpper).tDpf.lDemo = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
                                                End If
                                                ReDim Preserve tgDpfInfo(0 To ilDpfUpper + 1) As DPFINFO
                                            Next ilCol
                                            Exit For
                                        End If
                                    Next ilLoop
                                End If
                            End If
                        Else
                            'tmDrfPop.lDemo(ilDemoGender) = (CLng(smFieldValues(ilPopIndex)) + 500) \ 1000
                            'tmDrf.lDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
                        End If
                    End If
                    'If (ilNoDemoFd >= 16) Or ((ilNoDemoFd > 0) And (InStr(1, smFieldValues(1), ":", 1) > 0)) Then
                    If ((ilNoDemoFd > 0) And (InStr(1, smFieldValues(1), ":", 1) > 0)) Then
                        'Save Last 10 lines
                        'For ilLoop = 2 To 10 Step 1
                        For ilLoop = 1 To 9 Step 1
                            smSvLine(ilLoop - 1) = smSvLine(ilLoop)
                        Next ilLoop
                        'smSvLine(10) = slLine
                        smSvLine(9) = slLine
                        ilHeaderFd = False
                        blDontBlankLine = False
                        '6/9/16: Replaced GoSub
                        'GoSub mPCWriteRec
                        mPCWriteRec ilRet, ilCustomDemo, slVehicleName, slDPBookName, slETBookName, slBookDate, ilPrevDnfCode, llPrevDrfCode
                        If ilRet <> 0 Then
                            mPCNetConvFile = False
                            Exit Function
                        End If
                        ilNoDemoFd = 0
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
            If plcGauge.Value <> llPercent Then
                plcGauge.Value = llPercent
            End If
        End If
    Loop Until ilEof
    If (ilHeaderFd) And (ilNoDemoFd > 0) Then
        ilHeaderFd = False
        blDontBlankLine = False
        '6/9/16: Replaced GoSub
        'GoSub mPCWriteRec
        mPCWriteRec ilRet, ilCustomDemo, slVehicleName, slDPBookName, slETBookName, slBookDate, ilPrevDnfCode, llPrevDrfCode
        If ilRet <> 0 Then
            mPCNetConvFile = False
            Exit Function
        End If
    End If
    Close hmFrom
    plcGauge.Value = 100
    mPCNetConvFile = True
    MousePointer = vbDefault
    Exit Function
'mPCNetConvFileErr:
'    ilRet = Err.Number
'    Resume Next

'mPCWriteRec:
'    If ilCustomDemo Then
'        Return
'        'tmDrfPop.sDataType = "B"
'        'tmDrf.sDataType = "B"
'    End If
'    ilRet = 0
'    For ilVef = LBound(imVefCodeImpt) To UBound(imVefCodeImpt) - 1 Step 1
'        imVefCode = imVefCodeImpt(ilVef)
'        'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
'        '    If tgMVef(ilIndex).iCode = imVefCode Then
'            ilIndex = gBinarySearchVef(imVefCode)
'            If ilIndex <> -1 Then
'                slVehicleName = Trim$(tgMVef(ilIndex).sName)
'        '        Exit For
'            End If
'        'Next ilIndex
'        ilSetBkNm = -1
'        For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
'            If ((tmDrfInfo(ilCol).iBkNm = 0) And (Trim$(slDPBookName) <> "")) Or ((tmDrfInfo(ilCol).iBkNm = 1) And (Trim$(slETBookName) <> "")) Then
'
'                tmDrfInfo(ilCol).tDrf.iVefCode = imVefCode
'                If tmDrfInfo(ilCol).tDrf.sInfoType = "D" Then
'                    mDPPop
'                    For ilIndex = LBound(tmRdfInfo) To UBound(tmRdfInfo) - 1 Step 1
'                        ilRdf = tmRdfInfo(ilIndex).iRdfIndex
'                        For ilRow = 1 To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
'                            If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
'                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilRow), tgMRdf(ilRdf).iStartTime(1, ilRow), False, llTime
'                                If llTime = tmDrfInfo(ilCol).lStartTime Then
'                                    gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilRow), tgMRdf(ilRdf).iEndTime(1, ilRow), False, llTime
'                                    If llTime = tmDrfInfo(ilCol).lEndTime Then
'                                        ilMatch = True
'                                        For ilDay = 1 To 7 Step 1
'                                            If tgMRdf(ilRdf).sWkDays(ilRow, ilDay) <> tmDrfInfo(ilCol).sDays(ilDay - 1) Then
'                                                ilMatch = False
'                                                Exit For
'                                            End If
'                                        Next ilDay
'                                        If ilMatch Then
'                                            tmDrfInfo(ilCol).tDrf.iRdfcode = tgMRdf(ilRdf).iCode
'                                        End If
'                                        Exit For
'                                    End If
'                                End If
'                                Exit For
'                            End If
'                        Next ilRow
'                    Next ilIndex
'                End If
'                For ilDay = 0 To 6 Step 1
'                    smDays(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
'                Next ilDay
'                If tmDrfInfo(ilCol).iBkNm = 0 Then
'                    ilBNRet = mBookNameUsed(slDPBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfcode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
'                Else
'                    ilBNRet = mBookNameUsed(slETBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfcode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
'                End If
'                If (ilBNRet = 0) Or (ilBNRet = 1) Then
'                    tmDnf.iCode = 0
'                    If tmDrfInfo(ilCol).iBkNm = 0 Then
'                        tmDnf.sBookName = slDPBookName
'                    Else
'                        tmDnf.sBookName = slETBookName
'                    End If
'                    gPackDate slBookDate, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
'                    gPackDate smNowDate, tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
'                    tmDnf.iUrfCode = tgUrf(0).iCode
'                    tmDnf.sType = "I"
'                    tmDnf.sForm = smDataForm
'                    If tmDrfInfo(ilCol).iBkNm = 0 Then
'                        tmDnf.sExactTime = "N"
'                    Else
'                        tmDnf.sExactTime = "Y"
'                    End If
'                    tmDnf.sSource = "A"
'                    tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
'                    tmDnf.iAutoCode = tmDnf.iCode
'                    ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                            ilRet = csiHandleValue(0, 7)
'                        End If
'                        Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
'                        lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
'                        'mPCNetConvFile = False
'                        'Exit Function
'                        Return
'                    End If
'                    'If tgSpf.sRemoteUsers = "Y" Then
'                        Do
'                            'tmDnfSrchKey.iCode = tmDnf.iCode
'                            'ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'                            tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
'                            tmDnf.iAutoCode = tmDnf.iCode
'                            gPackDate smSyncDate, tmDnf.iSyncDate(0), tmDnf.iSyncDate(1)
'                            gPackTime smSyncTime, tmDnf.iSyncTime(0), tmDnf.iSyncTime(1)
'                            ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
'                        Loop While ilRet = BTRV_ERR_CONFLICT
'                        If ilRet <> BTRV_ERR_NONE Then
'                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                                ilRet = csiHandleValue(0, 7)
'                            End If
'                            Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
'                            lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
'                            'mPCNetConvFile = False
'                            'Exit Function
'                            Return
'                        End If
'                    'End If
'                    ilPrevDnfCode = tmDnf.iCode
'                    ilRet = mObtainBookName()
'                    tmDrfPop.lCode = 0
'                    tmDrfPop.iDnfCode = ilPrevDnfCode
'                    tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
'                    tmDrfPop.lAutoCode = tmDrfPop.lCode
'                    ilRet = btrInsert(hmDrf, tmDrfPop, imDrfRecLen, INDEXKEY2)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                            ilRet = csiHandleValue(0, 7)
'                        End If
'                        Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'                        lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'                        'mPCNetConvFile = False
'                        'Exit Function
'                        Return
'                    End If
'                    'If tgSpf.sRemoteUsers = "Y" Then
'                        Do
'                            'tmDrfSrchKey2.lCode = tmDrfPop.lCode
'                            'ilRet = btrGetEqual(hmDrf, tmDrfPop, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
'                            tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
'                            tmDrfPop.lAutoCode = tmDrfPop.lCode
'                            gPackDate smSyncDate, tmDrfPop.iSyncDate(0), tmDrfPop.iSyncDate(1)
'                            gPackTime smSyncTime, tmDrfPop.iSyncTime(0), tmDrfPop.iSyncTime(1)
'                            ilRet = btrUpdate(hmDrf, tmDrfPop, imDrfRecLen)
'                        Loop While ilRet = BTRV_ERR_CONFLICT
'                        If ilRet <> BTRV_ERR_NONE Then
'                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                                ilRet = csiHandleValue(0, 7)
'                            End If
'                            Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'                            lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'                            'mPCNetConvFile = False
'                            'Exit Function
'                            Return
'                        End If
'                    'End If
'                Else
'                    If StrComp(smBookForm, smDataForm, vbTextCompare) <> 0 Then
'                        If (smDataForm = "8") Or (smBookForm = "8") Then
'                            If tmDrfInfo(ilCol).iBkNm = 0 Then
'                                Print #hmTo, slDPBookName & " for " & slVehicleName & " previously defined with different format then current import data"
'                            Else
'                                Print #hmTo, slETBookName & " for " & slVehicleName & " previously defined with different format then current import data"
'                            End If
'                            lbcErrors.AddItem "Error in Book Forms" & " for " & slVehicleName
'                            ilRet = 0
'                            Return
'                        End If
'                    End If
'                    If ilBNRet = 3 Then
'                        Screen.MousePointer = vbDefault
'                        If tmDrfInfo(ilCol).iBkNm = 0 Then
'                            ilQRet = MsgBox(slDPBookName & " for " & slVehicleName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
'                        Else
'                            ilQRet = MsgBox(slETBookName & " for " & slVehicleName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
'                        End If
'                        If ilQRet = vbNo Then
'                            'mPCNetConvFile = False
'                            'Exit Function
'                            ilRet = 0
'                            Return
'                        End If
'                        Print #hmTo, "Replaced Demo Data (DRF)"
'                    End If
'                End If
'                tmDrfInfo(ilCol).tDrf.iDnfCode = ilPrevDnfCode
'                If llPrevDrfCode = 0 Then
'                    tmDrfInfo(ilCol).tDrf.lCode = 0
'                    tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
'                    tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
'                    ilRet = btrInsert(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen, INDEXKEY2)
'                    If ilRet <> BTRV_ERR_NONE Then
'                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                            ilRet = csiHandleValue(0, 7)
'                        End If
'                        Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'                        lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'                        'mPCNetConvFile = False
'                        'Exit Function
'                        Return
'                    End If
'                    'If tgSpf.sRemoteUsers = "Y" Then
'                        Do
'                            'tmDrfSrchKey2.lCode = tmDrf.lCode
'                            'ilRet = btrGetEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
'                            tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
'                            tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
'                            gPackDate smSyncDate, tmDrfInfo(ilCol).tDrf.iSyncDate(0), tmDrfInfo(ilCol).tDrf.iSyncDate(1)
'                            gPackTime smSyncTime, tmDrfInfo(ilCol).tDrf.iSyncTime(0), tmDrfInfo(ilCol).tDrf.iSyncTime(1)
'                            ilRet = btrUpdate(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen)
'                        Loop While ilRet = BTRV_ERR_CONFLICT
'                     'End If
'                Else
'                    Do
'                        tmDrfSrchKey2.lCode = llPrevDrfCode
'                        ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'                        If ilRet <> BTRV_ERR_NONE Then
'                            Exit Do
'                        End If
'                        'To avoid key being modified
'                        ilRet = btrDelete(hmDrf)
'                        tmDrfInfo(ilCol).tDrf.lCode = tlDrf.lCode
'                        tmDrfInfo(ilCol).tDrf.iRemoteID = tlDrf.iRemoteID
'                        tmDrfInfo(ilCol).tDrf.lAutoCode = tlDrf.lAutoCode
'                        gPackDate smSyncDate, tmDrfInfo(ilCol).tDrf.iSyncDate(0), tmDrfInfo(ilCol).tDrf.iSyncDate(1)
'                        gPackTime smSyncTime, tmDrfInfo(ilCol).tDrf.iSyncTime(0), tmDrfInfo(ilCol).tDrf.iSyncTime(1)
'                        'ilRet = btrUpdate(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen)
'                        ilRet = btrInsert(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen, INDEXKEY2)
'                    Loop While ilRet = BTRV_ERR_CONFLICT
'                End If
'                If ilRet <> BTRV_ERR_NONE Then
'                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                        ilRet = csiHandleValue(0, 7)
'                    End If
'                    Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'                    lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'                    'mPCNetConvFile = False
'                    'Exit Function
'                    Return
'                End If
'                'Insert or add Plus
'                For ilDpf = 0 To UBound(tgDpfInfo) - 1 Step 1
'                    If tgDpfInfo(ilDpf).iCol = ilCol Then
'                        tgDpfInfo(ilDpf).tDpf.iDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
'                        tgDpfInfo(ilDpf).tDpf.lDrfCode = tmDrfInfo(ilCol).tDrf.lCode
'                        tmDpfSrchKey1.lDrfCode = tgDpfInfo(ilDpf).tDpf.lDrfCode
'                        tmDpfSrchKey1.iMnfDemo = tgDpfInfo(ilDpf).tDpf.iMnfDemo
'                        ilPRet = btrGetEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'                        If ilPRet <> BTRV_ERR_NONE Then
'                            tgDpfInfo(ilDpf).tDpf.lCode = 0
'                            ilPRet = btrInsert(hmDpf, tgDpfInfo(ilDpf).tDpf, imDpfRecLen, INDEXKEY0)
'                            If ilPRet <> BTRV_ERR_NONE Then
'                                If (ilPRet = 30000) Or (ilPRet = 30001) Or (ilPRet = 30002) Or (ilPRet = 30003) Then
'                                    ilPRet = csiHandleValue(0, 7)
'                                End If
'                                Print #hmTo, "Warning: Error when Adding Demo Plus Data File (DPF)" & Str$(ilPRet) & " for " & slVehicleName
'                                lbcErrors.AddItem "Error Adding DPF" & " for " & slVehicleName
'                                ''mPCNetConvFile = False
'                                ''Exit Function
'                                'Return
'                            End If
'                        Else
'                            tmDpf.iDnfCode = tgDpfInfo(ilDpf).tDpf.iDnfCode
'                            tmDpf.lPop = tgDpfInfo(ilDpf).tDpf.lPop
'                            tmDpf.lDemo = tgDpfInfo(ilDpf).tDpf.lDemo
'                            ilPRet = btrUpdate(hmDpf, tmDpf, imDpfRecLen)
'                            If ilPRet <> BTRV_ERR_NONE Then
'                                If (ilPRet = 30000) Or (ilPRet = 30001) Or (ilPRet = 30002) Or (ilPRet = 30003) Then
'                                    ilPRet = csiHandleValue(0, 7)
'                                End If
'                                Print #hmTo, "Warning: Error when Updating Demo Plus Data File (DPF)" & Str$(ilPRet) & " for " & slVehicleName
'                                lbcErrors.AddItem "Error Updating DPF" & " for " & slVehicleName
'                                ''mPCNetConvFile = False
'                                ''Exit Function
'                                'Return
'                            End If
'                        End If
'                    End If
'                Next ilDpf
'                If tmDrfInfo(ilCol).iBkNm = 0 Then
'                    ilSetBkNm = 0
'                    ilSetDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
'                ElseIf (tmDrfInfo(ilCol).iBkNm = 1) And (ilSetBkNm = -1) Then
'                    ilSetBkNm = 1
'                    ilSetDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
'                End If
'            End If
'        Next ilCol
'        If ((ckcDefault(0).Value) Or (ckcDefault(1).Value)) And ((ilSetBkNm = 0) Or (ilSetBkNm = 1)) Then
'            Do
'                tmVefSrchKey.iCode = imVefCode
'                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'                If ilRet <> BTRV_ERR_NONE Then
'                    Exit Do
'                End If
'                If ckcDefault(0).Value = vbChecked Then
'                    tmVef.iDnfCode = ilSetDnfCode
'                End If
'                If ckcDefault(1).Value = vbChecked Then
'                    tmVef.iReallDnfCode = ilSetDnfCode
'                End If
'                'tmVef.iSourceID = tgUrf(0).iRemoteUserID
'                'gPackDate smSyncDate, tmVef.iSyncDate(0), tmVef.iSyncDate(1)
'                'gPackTime smSyncTime, tmVef.iSyncTime(0), tmVef.iSyncTime(1)
'                ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
'            Loop While ilRet = BTRV_ERR_CONFLICT
'            ilRet = gBinarySearchVef(tmVef.iCode)
'            If ilRet <> -1 Then
'                tgMVef(ilRet) = tmVef
'            End If
'        End If
'        Print #hmTo, "Successfully Installed Demo Data" & " for " & Trim$(slVehicleName)
'    Next ilVef
'    For ilVef = LBound(imVefCodeImpt) To UBound(imVefCodeImpt) - 1 Step 1
'        'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
'        '    If tgMVef(ilIndex).iCode = imVefCodeImpt(ilVef) Then
'            ilIndex = gBinarySearchVef(imVefCodeImpt(ilVef))
'            If ilIndex <> -1 Then
'                Print #hmTo, "Successfully Installed Demo Data" & " for " & Trim$(tgMVef(ilIndex).sName)
'        '        Exit For
'            End If
'        'Next ilIndex
'    Next ilVef
'    ilRet = 0
'    Return
End Function


Private Sub mPCWriteRec(ilRet As Integer, ilCustomDemo As Integer, slVehicleName As String, slDPBookName As String, slETBookName As String, slBookDate As String, ilPrevDnfCode As Integer, llPrevDrfCode As Long)
    Dim ilVef As Integer
    Dim ilIndex As Integer
    Dim ilRow As Integer
    Dim llTime As Long
    Dim ilMatch As Integer
    Dim ilDay As Integer
    Dim ilCol As Integer
    Dim ilBNRet As Integer
    Dim ilQRet As Integer
    Dim ilDpf As Integer
    Dim ilPRet As Integer
    Dim ilSetBkNm As Integer
    Dim ilSetDnfCode As Integer
    Dim ilRdf As Integer
    Dim tlDrf As DRF
    
    If ilCustomDemo Then
        'Return
        Exit Sub
        'tmDrfPop.sDataType = "B"
        'tmDrf.sDataType = "B"
    End If
    ilRet = 0
    For ilVef = LBound(imVefCodeImpt) To UBound(imVefCodeImpt) - 1 Step 1
        imVefCode = imVefCodeImpt(ilVef)
        'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If tgMVef(ilIndex).iCode = imVefCode Then
            ilIndex = gBinarySearchVef(imVefCode)
            If ilIndex <> -1 Then
                slVehicleName = Trim$(tgMVef(ilIndex).sName)
        '        Exit For
            End If
        'Next ilIndex
        ilSetBkNm = -1
        For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
            If ((tmDrfInfo(ilCol).iBkNm = 0) And (Trim$(slDPBookName) <> "")) Or ((tmDrfInfo(ilCol).iBkNm = 1) And (Trim$(slETBookName) <> "")) Then

                tmDrfInfo(ilCol).tDrf.iVefCode = imVefCode
                If tmDrfInfo(ilCol).tDrf.sInfoType = "D" Then
                    mDPPop
                    For ilIndex = LBound(tmRdfInfo) To UBound(tmRdfInfo) - 1 Step 1
                        ilRdf = tmRdfInfo(ilIndex).iRdfIndex
                        For ilRow = 1 To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                            If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilRow), tgMRdf(ilRdf).iStartTime(1, ilRow), False, llTime
                                If llTime = tmDrfInfo(ilCol).lStartTime Then
                                    gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilRow), tgMRdf(ilRdf).iEndTime(1, ilRow), False, llTime
                                    If llTime = tmDrfInfo(ilCol).lEndTime Then
                                        ilMatch = True
                                        For ilDay = 1 To 7 Step 1
                                            If tgMRdf(ilRdf).sWkDays(ilRow, ilDay - 1) <> tmDrfInfo(ilCol).sDays(ilDay - 1) Then
                                                ilMatch = False
                                                Exit For
                                            End If
                                        Next ilDay
                                        If ilMatch Then
                                            tmDrfInfo(ilCol).tDrf.iRdfCode = tgMRdf(ilRdf).iCode
                                        End If
                                        Exit For
                                    End If
                                End If
                                Exit For
                            End If
                        Next ilRow
                    Next ilIndex
                End If
                For ilDay = 0 To 6 Step 1
                    smDays(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
                Next ilDay
                If tmDrfInfo(ilCol).iBkNm = 0 Then
                    ilBNRet = mBookNameUsed(slDPBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfCode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
                Else
                    ilBNRet = mBookNameUsed(slETBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfCode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
                End If
                If (ilBNRet = 0) Or (ilBNRet = 1) Then
                    tmDnf.iCode = 0
                    If tmDrfInfo(ilCol).iBkNm = 0 Then
                        tmDnf.sBookName = slDPBookName
                    Else
                        tmDnf.sBookName = slETBookName
                    End If
                    gPackDate slBookDate, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
                    gPackDate smNowDate, tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
                    tmDnf.iUrfCode = tgUrf(0).iCode
                    tmDnf.sType = "I"
                    tmDnf.sForm = smDataForm
                    If tmDrfInfo(ilCol).iBkNm = 0 Then
                        tmDnf.sExactTime = "N"
                    Else
                        tmDnf.sExactTime = "Y"
                    End If
                    tmDnf.sSource = "A"
                    tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmDnf.iAutoCode = tmDnf.iCode
                    ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
                    If ilRet <> BTRV_ERR_NONE Then
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
                        lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
                        'mPCNetConvFile = False
                        'Exit Function
                        'Return
                        Exit Sub
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
                        If ilRet <> BTRV_ERR_NONE Then
                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
                            lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
                            'mPCNetConvFile = False
                            'Exit Function
                            'Return
                            Exit Sub
                        End If
                    'End If
                    ilPrevDnfCode = tmDnf.iCode
                    ilRet = mObtainBookName()
                    tmDrfPop.lCode = 0
                    tmDrfPop.iDnfCode = ilPrevDnfCode
                    tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
                    tmDrfPop.lAutoCode = tmDrfPop.lCode
                    ilRet = btrInsert(hmDrf, tmDrfPop, imDrfRecLen, INDEXKEY2)
                    If ilRet <> BTRV_ERR_NONE Then
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slVehicleName
                        lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
                        'mPCNetConvFile = False
                        'Exit Function
                        'Return
                        Exit Sub
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
                        If ilRet <> BTRV_ERR_NONE Then
                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slVehicleName
                            lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
                            'mPCNetConvFile = False
                            'Exit Function
                            'Return
                            Exit Sub
                        End If
                    'End If
                Else
                    If StrComp(smBookForm, smDataForm, vbTextCompare) <> 0 Then
                        If (smDataForm = "8") Or (smBookForm = "8") Then
                            If tmDrfInfo(ilCol).iBkNm = 0 Then
                                Print #hmTo, slDPBookName & " for " & slVehicleName & " previously defined with different format then current import data"
                            Else
                                Print #hmTo, slETBookName & " for " & slVehicleName & " previously defined with different format then current import data"
                            End If
                            lbcErrors.AddItem "Error in Book Forms" & " for " & slVehicleName
                            ilRet = 0
                            'Return
                            Exit Sub
                        End If
                    End If
                    If ilBNRet = 3 Then
                        Screen.MousePointer = vbDefault
                        If tmDrfInfo(ilCol).iBkNm = 0 Then
                            ilQRet = MsgBox(slDPBookName & " for " & slVehicleName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
                        Else
                            ilQRet = MsgBox(slETBookName & " for " & slVehicleName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
                        End If
                        If ilQRet = vbNo Then
                            'mPCNetConvFile = False
                            'Exit Function
                            ilRet = 0
                            'Return
                            Exit Sub
                        End If
                        Print #hmTo, "Replaced Demo Data (DRF)"
                    End If
                End If
                tmDrfInfo(ilCol).tDrf.iDnfCode = ilPrevDnfCode
                If llPrevDrfCode = 0 Then
                    tmDrfInfo(ilCol).tDrf.lCode = 0
                    tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
                    ilRet = btrInsert(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen, INDEXKEY2)
                    If ilRet <> BTRV_ERR_NONE Then
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
                        lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
                        'mPCNetConvFile = False
                        'Exit Function
                        'Return
                        Exit Sub
                    End If
                    'If tgSpf.sRemoteUsers = "Y" Then
                        Do
                            'tmDrfSrchKey2.lCode = tmDrf.lCode
                            'ilRet = btrGetEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                            tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
                            tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
                            gPackDate smSyncDate, tmDrfInfo(ilCol).tDrf.iSyncDate(0), tmDrfInfo(ilCol).tDrf.iSyncDate(1)
                            gPackTime smSyncTime, tmDrfInfo(ilCol).tDrf.iSyncTime(0), tmDrfInfo(ilCol).tDrf.iSyncTime(1)
                            ilRet = btrUpdate(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                     'End If
                Else
                    Do
                        tmDrfSrchKey2.lCode = llPrevDrfCode
                        ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                        If ilRet <> BTRV_ERR_NONE Then
                            Exit Do
                        End If
                        'To avoid key being modified
                        ilRet = btrDelete(hmDrf)
                        tmDrfInfo(ilCol).tDrf.lCode = tlDrf.lCode
                        tmDrfInfo(ilCol).tDrf.iRemoteID = tlDrf.iRemoteID
                        tmDrfInfo(ilCol).tDrf.lAutoCode = tlDrf.lAutoCode
                        gPackDate smSyncDate, tmDrfInfo(ilCol).tDrf.iSyncDate(0), tmDrfInfo(ilCol).tDrf.iSyncDate(1)
                        gPackTime smSyncTime, tmDrfInfo(ilCol).tDrf.iSyncTime(0), tmDrfInfo(ilCol).tDrf.iSyncTime(1)
                        'ilRet = btrUpdate(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen)
                        ilRet = btrInsert(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen, INDEXKEY2)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                End If
                If ilRet <> BTRV_ERR_NONE Then
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
                    lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
                    'mPCNetConvFile = False
                    'Exit Function
                    'Return
                    Exit Sub
                End If
                'Insert or add Plus
                For ilDpf = 0 To UBound(tgDpfInfo) - 1 Step 1
                    If tgDpfInfo(ilDpf).iCol = ilCol Then
                        tgDpfInfo(ilDpf).tDpf.iDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
                        tgDpfInfo(ilDpf).tDpf.lDrfCode = tmDrfInfo(ilCol).tDrf.lCode
                        tmDpfSrchKey1.lDrfCode = tgDpfInfo(ilDpf).tDpf.lDrfCode
                        tmDpfSrchKey1.iMnfDemo = tgDpfInfo(ilDpf).tDpf.iMnfDemo
                        ilPRet = btrGetEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                        If ilPRet <> BTRV_ERR_NONE Then
                            tgDpfInfo(ilDpf).tDpf.lCode = 0
                            ilPRet = btrInsert(hmDpf, tgDpfInfo(ilDpf).tDpf, imDpfRecLen, INDEXKEY0)
                            If ilPRet <> BTRV_ERR_NONE Then
                                If (ilPRet = 30000) Or (ilPRet = 30001) Or (ilPRet = 30002) Or (ilPRet = 30003) Then
                                    ilPRet = csiHandleValue(0, 7)
                                End If
                                Print #hmTo, "Warning: Error when Adding Demo Plus Data File (DPF)" & Str$(ilPRet) & " for " & slVehicleName
                                lbcErrors.AddItem "Error Adding DPF" & " for " & slVehicleName
                                ''mPCNetConvFile = False
                                ''Exit Function
                                'Return
                            End If
                        Else
                            tmDpf.iDnfCode = tgDpfInfo(ilDpf).tDpf.iDnfCode
                            tmDpf.lPop = tgDpfInfo(ilDpf).tDpf.lPop
                            tmDpf.lDemo = tgDpfInfo(ilDpf).tDpf.lDemo
                            ilPRet = btrUpdate(hmDpf, tmDpf, imDpfRecLen)
                            If ilPRet <> BTRV_ERR_NONE Then
                                If (ilPRet = 30000) Or (ilPRet = 30001) Or (ilPRet = 30002) Or (ilPRet = 30003) Then
                                    ilPRet = csiHandleValue(0, 7)
                                End If
                                Print #hmTo, "Warning: Error when Updating Demo Plus Data File (DPF)" & Str$(ilPRet) & " for " & slVehicleName
                                lbcErrors.AddItem "Error Updating DPF" & " for " & slVehicleName
                                ''mPCNetConvFile = False
                                ''Exit Function
                                'Return
                            End If
                        End If
                    End If
                Next ilDpf
                If tmDrfInfo(ilCol).iBkNm = 0 Then
                    ilSetBkNm = 0
                    ilSetDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
                ElseIf (tmDrfInfo(ilCol).iBkNm = 1) And (ilSetBkNm = -1) Then
                    ilSetBkNm = 1
                    ilSetDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
                End If
            End If
        Next ilCol
        If ((ckcDefault(0).Value) Or (ckcDefault(1).Value)) And ((ilSetBkNm = 0) Or (ilSetBkNm = 1)) Then
            Do
                tmVefSrchKey.iCode = imVefCode
                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If ckcDefault(0).Value = vbChecked Then
                    tmVef.iDnfCode = ilSetDnfCode
                End If
                If ckcDefault(1).Value = vbChecked Then
                    tmVef.iReallDnfCode = ilSetDnfCode
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
        End If
        Print #hmTo, "Successfully Installed Demo Data" & " for " & Trim$(slVehicleName)
        bmBooksSaved = True
    Next ilVef
    For ilVef = LBound(imVefCodeImpt) To UBound(imVefCodeImpt) - 1 Step 1
        'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If tgMVef(ilIndex).iCode = imVefCodeImpt(ilVef) Then
            ilIndex = gBinarySearchVef(imVefCodeImpt(ilVef))
            If ilIndex <> -1 Then
                Print #hmTo, "Successfully Installed Demo Data" & " for " & Trim$(tgMVef(ilIndex).sName)
        '        Exit For
            End If
        'Next ilIndex
    Next ilVef
    ilRet = 0
    'Return
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPCStnConvFile_PopByMarket      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert File                   *
'*                                                     *
'*******************************************************
Private Function mPCStnConvFile_PopByMarket(slFromFile As String, slBookDate As String) As Integer
    Dim ilRet As Integer
    Dim ilBNRet As Integer
    Dim slLine As String
    Dim ilHeaderFd As Integer
    Dim ilAvgQHIndex As Integer
    Dim ilPopDone As Integer
    Dim ilDayTimeFd As Integer
    Dim ilDemoGender As Integer
    Dim slDemoAge As String
    Dim ilPos As Integer
    Dim slDay As String
    Dim ilDay As Integer
    Dim ilSY As Integer
    Dim ilEY As Integer
    Dim ilLoop As Integer
    Dim ilEof As Integer
    Dim llPercent As Long
    Dim slChar As String
    Dim slStr As String
    Dim slSexChar As String
    Dim ilIndex As Integer
    Dim ilCustomDemo As Integer
    Dim ilPrevDnfCode As Integer
    Dim llPrevDrfCode As Long
    Dim ilRdf As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim slVehicleName As String
    Dim llTime As Long
    'Gender index 1 to 16
    'ReDim llPopDemo(1 To 16) As Long
    ReDim llPopDemo(0 To 16) As Long
    Dim ilTIndex As Integer
    Dim ilSetBkNm As Integer
    Dim ilMarketCol As Integer
    Dim ilStnCol As Integer
    Dim ilBookCol As Integer
    Dim ilSetDnfCode As Integer
    Dim ilDemoFd As Integer
    Dim slSTime As String
    Dim slETime As String
    Dim ilMatch As Integer
    Dim ilFound As Integer
    Dim slBookName As String
    Dim slBook As String    'Book date if no column defined
    Dim ilSet As Integer
    Dim tlDrf As DRF
    ilRet = 0
    'On Error GoTo mPCStnConvFile_PopByMarketErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        MsgBox "Open " & slFromFile & ", Error #" & Str$(ilRet), vbExclamation, "Open Error"
        edcFrom.SetFocus
        mPCStnConvFile_PopByMarket = False
        Exit Function
    End If
    DoEvents
    If imTerminate Then
        Close hmFrom
        mTerminate
        mPCStnConvFile_PopByMarket = False
        Exit Function
    End If
    ilHeaderFd = False
    ilDayTimeFd = False
    slLine = ""
    ilPopDone = False
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
    tmDrfPop.sACTLineupCode = ""
    tmDrfPop.sACT1StoredTime = ""
    tmDrfPop.sACT1StoredSpots = ""
    tmDrfPop.sACT1StoreClearPct = ""
    tmDrfPop.sACT1DaypartFilter = ""
    
    ilCustomDemo = False
    ilBookCol = -1
    slBook = ""
    Do
        err.Clear
        ilRet = 0
        'On Error GoTo mPCStnConvFile_PopByMarketErr:
        'Dan added 6/9/15
        slLine = ""
        'Line Input #hmFrom, slLine
        Do
            If Not EOF(hmFrom) Then
                slChar = Input(1, #hmFrom)
                If slChar = Chr(13) Then
                    slChar = Input(1, #hmFrom)
                End If
                If slChar = Chr(10) Then
                    Exit Do
                End If
                slLine = slLine & slChar
            Else
                ilEof = True
                Exit Do
            End If
        Loop
        slLine = Trim$(slLine)
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
                    mPCStnConvFile_PopByMarket = False
                    Exit Function
                End If
                gParseCDFields slLine, False, smFieldValues()
                '11/10/16: Move fields from 0 to X to 1 to X +1
                For ilLoop = UBound(smFieldValues) - 1 To LBound(smFieldValues) Step -1
                    smFieldValues(ilLoop + 1) = smFieldValues(ilLoop)
                Next ilLoop
                smFieldValues(0) = ""
                'Determine field Type
                'Demo Record
                '   Schedule: MF 6-10A,,,,, Men 12-17,,,Men 25-35,,,.....
                'Title Record
                '   Met,Market,Book,Station,,AQH,Population,,AQH,Population,,....,
                'Values by station
                '   Rank,"Los Angles, CA", Win01,WXXX-FM,Status, 28900,538300,, 5500,725600,,
                '
                If (Not ilHeaderFd) Or (Not ilDayTimeFd) Then
                    'Remove Book test as two types of import:  One will have the same book name for all station and that name is
                    'specified at top, the other will have two names and will be defined in column title
                    'If (InStr(1, Trim$(slLine), "Market", 1) > 0) And (InStr(1, Trim$(slLine), "Book", 1) > 0) And (InStr(1, Trim$(slLine), "Station", 1) > 0) Then
                    If (InStr(1, Trim$(slLine), "Market", 1) > 0) And (InStr(1, Trim$(slLine), "Station", 1) > 0) Then
                        If InStr(1, RTrim$(slLine), "AQH", 1) > 0 Then
                            'Determine number of Demo columns
                            ilHeaderFd = True
                            ilPos = 2
                            ReDim tmDrfInfo(0 To 0) As DRFINFO
                            Do While ilPos <= UBound(smFieldValues)
                                If InStr(1, Trim$(smFieldValues(ilPos)), "AQH", 1) > 0 Then
                                    tmDrfInfo(UBound(tmDrfInfo)).iStartCol = ilPos
                                    tmDrfInfo(UBound(tmDrfInfo)).iType = 0  'Vehicle
                                    tmDrfInfo(UBound(tmDrfInfo)).iBkNm = 0  'Daypart Book Name
                                    tmDrfInfo(UBound(tmDrfInfo)).iSY = -1
                                    tmDrfInfo(UBound(tmDrfInfo)).iEY = -1
                                    tmDrfInfo(UBound(tmDrfInfo)).iDemoGender = -1
                                    ReDim Preserve tmDrfInfo(0 To UBound(tmDrfInfo) + 1) As DRFINFO
                                    ilPos = ilPos + 1
                                ElseIf InStr(1, Trim$(smFieldValues(ilPos)), "Market", 1) > 0 Then
                                    ilMarketCol = ilPos
                                ElseIf InStr(1, Trim$(smFieldValues(ilPos)), "Book", 1) > 0 Then
                                    ilBookCol = ilPos
                                ElseIf InStr(1, Trim$(smFieldValues(ilPos)), "Station", 1) > 0 Then
                                    ilStnCol = ilPos
                                End If
                                ilPos = ilPos + 1
                            Loop
                            'Get Demo and Dayparts
                            ilDayTimeFd = False
                            ilCustomDemo = False
                            'For ilLoop = 10 To 1 Step -1
                            For ilLoop = 9 To 0 Step -1
                                If Len(Trim$(smSvLine(ilLoop))) > 0 Then
                                    gParseCDFields smSvLine(ilLoop), False, smFieldValues()
                                    If ilBookCol = -1 Then
                                        ilPos = InStr(1, smFieldValues(1), "Summer", 1)
                                        If ilPos > 0 Then
                                            slBook = "Sum" & right$(Trim$(smFieldValues(1)), 2)
                                            ilBookCol = -2  'Found
                                        End If
                                        ilPos = InStr(1, smFieldValues(1), "Fall", 1)
                                        If ilPos > 0 Then
                                            slBook = "Fal" & right$(Trim$(smFieldValues(1)), 2)
                                            ilBookCol = -2  'Found
                                        End If
                                        ilPos = InStr(1, smFieldValues(1), "Winter", 1)
                                        If ilPos > 0 Then
                                            slBook = "Win" & right$(Trim$(smFieldValues(1)), 2)
                                            ilBookCol = -2  'Found
                                        End If
                                        ilPos = InStr(1, smFieldValues(1), "Spring", 1)
                                        If ilPos > 0 Then
                                            slBook = "Spr" & right$(Trim$(smFieldValues(1)), 2)
                                            ilBookCol = -2  'Found
                                        End If
                                        If (ilDayTimeFd) And (ilBookCol = -2) Then
                                            Exit For
                                        End If
                                    End If
                                    ilPos = InStr(1, smFieldValues(1), "Schedule:", 1)
                                    If ilPos > 0 Then
                                        ilCol = 0
                                        smFieldValues(1) = Trim$(Mid$(smFieldValues(1), ilPos + 9))
                                        slStr = UCase$(Trim$(smFieldValues(1)))
                                        ilSY = -1
                                        slDay = Trim$(Mid$(slStr, 1, 3))
                                        ilTIndex = InStr(1, slStr, " ", 1) + 1
                                        'Select Case slDay
                                        '    Case "MO", "MON"
                                        '        ilSY = 0
                                        '        ilEY = 0
                                        '    Case "TU", "TUE"
                                        '        ilSY = 1
                                        '        ilEY = 1
                                        '    Case "WE", "WED"
                                        '        ilSY = 2
                                        '        ilEY = 2
                                        '    Case "TH", "THU"
                                        '        ilSY = 3
                                        '        ilEY = 3
                                        '    Case "FR", "FRI"
                                        '        ilSY = 4
                                        '        ilEY = 4
                                        '    Case "SA", "SAT"
                                        '        ilSY = 5
                                        '        ilEY = 5
                                        '        If Mid$(slStr, 3, 2) = "SU" Then
                                        '            ilEY = 6
                                        '        End If
                                        '    Case "SU", "SUN"
                                        '        ilSY = 6
                                        '        ilEY = 6
                                        '    Case "MF"
                                        '        ilSY = 0
                                        '        ilEY = 4
                                        '    Case "MSA"
                                        '        ilSY = 0
                                        '        ilEY = 5
                                        '    Case "MSU"
                                        '        ilSY = 0
                                        '        ilEY = 6
                                        '    Case "MS"
                                        '        ilSY = 0
                                        '        ilEY = 6
                                        '    Case "SS"
                                        '        ilSY = 5
                                        '        ilEY = 6
                                        '    Case "FSU"
                                        '        ilSY = 4
                                        '        ilEY = 6
                                        '    Case "FSA"
                                        '        ilSY = 4
                                        '        ilEY = 5
                                        'End Select
                                        If ilTIndex > 2 Then
                                            mGetDayIndex Left$(slStr, ilTIndex - 2), ilSY, ilEY
                                        End If
                                        If ilSY <> -1 Then
                                            tmDrfInfo(ilCol).lStartTime2 = -1
                                            slChar = Mid$(slStr, 1, 1)
                                            Do While (slChar < "0") Or (slChar > "9")
                                                slStr = Mid$(slStr, 2)
                                                slChar = Mid$(slStr, 1, 1)
                                            Loop
                                            slSTime = ""
                                            slChar = UCase$(Mid$(slStr, 1, 1))
                                            Do
                                                slSTime = slSTime & slChar
                                                If (slChar = "A") Or (slChar = "P") Then
                                                    Exit Do
                                                End If
                                                slStr = Mid$(slStr, 2)
                                                slChar = UCase$(Mid$(slStr, 1, 1))
                                            Loop While (slChar <> "-")
                                            If gValidTime(slSTime) Then
                                                slChar = Mid$(slStr, 1, 1)
                                                Do While (slChar < "0") Or (slChar > "9")
                                                    slStr = Mid$(slStr, 2)
                                                    slChar = Mid$(slStr, 1, 1)
                                                Loop
                                                slETime = ""
                                                slChar = UCase$(Mid$(slStr, 1, 1))
                                                Do
                                                    slETime = slETime & slChar
                                                    If (slChar = "A") Or (slChar = "P") Then
                                                        Exit Do
                                                    End If
                                                    If Len(slStr) = 0 Then
                                                        Exit Do
                                                    End If
                                                    slStr = Mid$(slStr, 2)
                                                    slChar = UCase$(Mid$(slStr, 1, 1))
                                                Loop While (slChar <> "/")
                                                If gValidTime(slETime) Then
                                                    ilDayTimeFd = True
                                                    If (right$(slSTime, 1) >= "0") And (right$(slSTime, 1) <= "9") Then
                                                        slSTime = slSTime & right$(slETime, 1)
                                                    End If
                                                    tmDrfInfo(ilCol).lStartTime = gTimeToCurrency(slSTime, False)
                                                    tmDrfInfo(ilCol).iSY = ilSY
                                                    tmDrfInfo(ilCol).iEY = ilEY
                                                    tmDrfInfo(ilCol).lEndTime = gTimeToCurrency(slETime, False)
                                                    For ilDay = 0 To 6 Step 1
                                                        tmDrfInfo(ilCol).sDays(ilDay) = "N"
                                                    Next ilDay
                                                    For ilDay = ilSY To ilEY Step 1
                                                        tmDrfInfo(ilCol).sDays(ilDay) = "Y"
                                                    Next ilDay
                                                    tmDrfInfo(ilCol).iType = 1  'Daypart

                                                    If InStr(1, slStr, "/", 1) > 0 Then
                                                        slChar = Mid$(slStr, 1, 1)
                                                        Do While (slChar < "0") Or (slChar > "9")
                                                            slStr = Mid$(slStr, 2)
                                                            slChar = Mid$(slStr, 1, 1)
                                                        Loop
                                                        slSTime = ""
                                                        slChar = UCase$(Mid$(slStr, 1, 1))
                                                        Do
                                                            slSTime = slSTime & slChar
                                                            If (slChar = "A") Or (slChar = "P") Then
                                                                Exit Do
                                                            End If
                                                            slStr = Mid$(slStr, 2)
                                                            slChar = UCase$(Mid$(slStr, 1, 1))
                                                        Loop While (slChar <> "-")
                                                        If gValidTime(slSTime) Then
                                                            slChar = Mid$(slStr, 1, 1)
                                                            Do While (slChar < "0") Or (slChar > "9")
                                                                slStr = Mid$(slStr, 2)
                                                                slChar = Mid$(slStr, 1, 1)
                                                            Loop
                                                            slETime = ""
                                                            slChar = UCase$(Mid$(slStr, 1, 1))
                                                            Do
                                                                slETime = slETime & slChar
                                                                If (slChar = "A") Or (slChar = "P") Then
                                                                    Exit Do
                                                                End If
                                                                If Len(slStr) = 0 Then
                                                                    Exit Do
                                                                End If
                                                                slStr = Mid$(slStr, 2)
                                                                slChar = UCase$(Mid$(slStr, 1, 1))
                                                            Loop While (slChar >= "0") And (slChar <= "9")
                                                            If gValidTime(slETime) Then
                                                                If (right$(slSTime, 1) >= "0") And (right$(slSTime, 1) <= "9") Then
                                                                    slSTime = slSTime & right$(slETime, 1)
                                                                End If
                                                                tmDrfInfo(ilCol).lStartTime2 = gTimeToCurrency(slSTime, False)
                                                                tmDrfInfo(ilCol).lEndTime2 = gTimeToCurrency(slETime, False)
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                        For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
                                            slStr = UCase$(Trim$(smFieldValues(tmDrfInfo(ilCol).iStartCol)))
                                            smFieldValues(1) = slStr
                                            'For ilLoop = LBound(tgMnfCDemo) To UBound(tgMnfCDemo) Step 1
                                            '    If StrComp(Trim$(tgMnfCDemo(ilLoop).sName), smFieldValues(1), 1) = 0 Then
                                            '        ilDemoGender = tgMnfCDemo(ilLoop).iGroupNo
                                            '        ilCustomDemo = True
                                            '        ilDemoFd = True
                                            '        Exit For
                                            '    End If
                                            'Next ilCol
                                            If tmDrfInfo(ilCol).iDemoGender = -1 Then
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
                                                                ilDemoFd = True
                                                            Case "18-24"
                                                                ilDemoGender = ilDemoGender + 2
                                                                ilDemoFd = True
                                                            Case "25-34"
                                                                ilDemoGender = ilDemoGender + 3
                                                                ilDemoFd = True
                                                            Case "35-44"
                                                                ilDemoGender = ilDemoGender + 4
                                                                ilDemoFd = True
                                                            Case "45-49"
                                                                ilDemoGender = ilDemoGender + 5
                                                                ilDemoFd = True
                                                            Case "50-54"
                                                                ilDemoGender = ilDemoGender + 6
                                                                ilDemoFd = True
                                                            Case "55-64"
                                                                ilDemoGender = ilDemoGender + 7
                                                                ilDemoFd = True
                                                            Case "65+"
                                                                ilDemoGender = ilDemoGender + 8
                                                                ilDemoFd = True
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
                                                                ilDemoFd = True
                                                            Case "18-20"
                                                                ilDemoGender = ilDemoGender + 2
                                                                ilDemoFd = True
                                                            Case "21-24"
                                                                ilDemoGender = ilDemoGender + 3
                                                                ilDemoFd = True
                                                            Case "25-34"
                                                                ilDemoGender = ilDemoGender + 4
                                                                ilDemoFd = True
                                                            Case "35-44"
                                                                ilDemoGender = ilDemoGender + 5
                                                                ilDemoFd = True
                                                            Case "45-49"
                                                                ilDemoGender = ilDemoGender + 6
                                                                ilDemoFd = True
                                                            Case "50-54"
                                                                ilDemoGender = ilDemoGender + 7
                                                                ilDemoFd = True
                                                            Case "55-64"
                                                                ilDemoGender = ilDemoGender + 8
                                                                ilDemoFd = True
                                                            Case "65+"
                                                                ilDemoGender = ilDemoGender + 9
                                                                ilDemoFd = True
                                                            Case Else
                                                                ilDemoGender = -1
                                                        End Select
                                                    End If
                                                    tmDrfInfo(ilCol).iDemoGender = ilDemoGender
                                                End If
                                            End If
                                        Next ilCol
                                        If (ilDayTimeFd) And ((ilBookCol >= 0) Or (ilBookCol = -2)) Then
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next ilLoop
                        Else
                            'Save Last 10 lines
                            'For ilLoop = 2 To 10 Step 1
                            For ilLoop = 1 To 9 Step 1
                                smSvLine(ilLoop - 1) = smSvLine(ilLoop)
                            Next ilLoop
                            'smSvLine(10) = slLine
                            smSvLine(9) = slLine
                        End If
                    Else
                        'Save Last 10 lines
                        'For ilLoop = 2 To 10 Step 1
                        For ilLoop = 1 To 9 Step 1
                            smSvLine(ilLoop - 1) = smSvLine(ilLoop)
                        Next ilLoop
                        'smSvLine(10) = slLine
                        smSvLine(9) = slLine
                    End If
                ElseIf Trim$(smFieldValues(ilStnCol + 1)) <> "*" Then
                    If (InStr(1, smFieldValues(1), "Station Notes:", 1) > 0) Or (InStr(1, smFieldValues(1), "Audience Data:", 1) > 0) Or (InStr(1, smFieldValues(1), "Total", 1) > 0) Or (InStr(1, smFieldValues(1), "Coverage Rtg", 1) > 0) Or (InStr(1, smFieldValues(1), "USA Rtg", 1) > 0) Then
                        ilHeaderFd = False
                    ElseIf Trim$(smFieldValues(ilStnCol)) <> "" Then
                        If UBound(tmDrfInfo) > LBound(tmDrfInfo) Then
                            ilCol = 0
                            imVefCode = 0
                            For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                ilPos = InStr(1, tgMVef(ilIndex).sName, smFieldValues(ilStnCol), 1)
                                If ilPos > 0 Then
                                    imVefCode = tgMVef(ilIndex).iCode
                                    Exit For
                                End If
                            Next ilIndex
                            If imVefCode > 0 Then
                                If ilBookCol >= 0 Then
                                    slBookName = Left$(Trim$(smFieldValues(ilBookCol)) & ":" & Trim$(smFieldValues(ilMarketCol)), Len(tmDnf.sBookName))
                                ElseIf ilBookCol = -2 Then
                                    slBookName = Left$(Trim$(slBook) & ":" & Trim$(smFieldValues(ilMarketCol)), Len(tmDnf.sBookName))
                                End If
                                mDPPop
                                For ilLoop = 1 To 16 Step 1
                                    tmDrfInfo(ilCol).tDrf.lDemo(ilLoop - 1) = 0
                                Next ilLoop
                                tmDrfInfo(ilCol).tDrf.lCode = 0
                                tmDrfInfo(ilCol).tDrf.iDnfCode = tmDnf.iCode
                                tmDrfInfo(ilCol).tDrf.sDemoDataType = "D"
                                tmDrfInfo(ilCol).tDrf.iMnfSocEco = 0
                                tmDrfInfo(ilCol).iBkNm = 0  'Book Name
                                tmDrfInfo(ilCol).tDrf.iStartTime2(0) = 1
                                tmDrfInfo(ilCol).tDrf.iStartTime2(1) = 0
                                tmDrfInfo(ilCol).tDrf.iEndTime2(0) = 1
                                tmDrfInfo(ilCol).tDrf.iEndTime2(1) = 0
                                If tmDrfInfo(ilCol).iType = 0 Then
                                    tmDrfInfo(ilCol).tDrf.sInfoType = "V"
                                    tmDrfInfo(ilCol).tDrf.iRdfCode = 0
                                    gPackTime "12AM", tmDrfInfo(ilCol).tDrf.iStartTime(0), tmDrfInfo(ilCol).tDrf.iStartTime(1)
                                    gPackTime "12AM", tmDrfInfo(ilCol).tDrf.iEndTime(0), tmDrfInfo(ilCol).tDrf.iEndTime(1)
                                Else
                                    tmDrfInfo(ilCol).tDrf.sInfoType = "D"
                                    tmDrfInfo(ilCol).tDrf.iRdfCode = 0
                                    gPackTimeLong tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).tDrf.iStartTime(0), tmDrfInfo(ilCol).tDrf.iStartTime(1)
                                    gPackTimeLong tmDrfInfo(ilCol).lEndTime, tmDrfInfo(ilCol).tDrf.iEndTime(0), tmDrfInfo(ilCol).tDrf.iEndTime(1)
                                    For ilDay = 0 To 6 Step 1
                                        tmDrfInfo(ilCol).tDrf.sDay(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
                                    Next ilDay
                                    If tmDrfInfo(ilCol).lStartTime2 <> -1 Then
                                        gPackTimeLong tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).tDrf.iStartTime(0), tmDrfInfo(ilCol).tDrf.iStartTime(1)
                                        gPackTimeLong tmDrfInfo(ilCol).lEndTime, tmDrfInfo(ilCol).tDrf.iEndTime(0), tmDrfInfo(ilCol).tDrf.iEndTime(1)
                                        gPackTimeLong tmDrfInfo(ilCol).lStartTime2, tmDrfInfo(ilCol).tDrf.iStartTime2(0), tmDrfInfo(ilCol).tDrf.iStartTime2(1)
                                        gPackTimeLong tmDrfInfo(ilCol).lEndTime2, tmDrfInfo(ilCol).tDrf.iEndTime2(0), tmDrfInfo(ilCol).tDrf.iEndTime2(1)
                                    End If
                                End If
                                tmDrfInfo(ilCol).tDrf.sProgCode = ""
                                tmDrfInfo(ilCol).tDrf.iQHIndex = 0
                                tmDrfInfo(ilCol).tDrf.iCount = 0
                                tmDrfInfo(ilCol).tDrf.sExStdDP = "N"
                                tmDrfInfo(ilCol).tDrf.sExRpt = "N"
                                tmDrfInfo(ilCol).tDrf.sDataType = "A"
                                For ilIndex = 0 To UBound(tmDrfInfo) - 1 Step 1
                                    ilDemoGender = tmDrfInfo(ilIndex).iDemoGender
                                    ilAvgQHIndex = tmDrfInfo(ilIndex).iStartCol
'                                    If tgSpf.sSAudData <> "H" Then
'                                        llPopDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex + 1)) + 500) \ 1000
'                                    Else
'                                        llPopDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex + 1)) + 50) \ 100
'                                    End If
                                    If tgSpf.sSAudData = "H" Then
                                        llPopDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex + 1)) + 50) \ 100
                                    ElseIf tgSpf.sSAudData = "N" Then
                                        llPopDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex + 1)) + 5) \ 10
                                    ElseIf tgSpf.sSAudData = "U" Then
                                        llPopDemo(ilDemoGender) = CLng(smFieldValues(ilAvgQHIndex + 1))
                                    Else
                                        llPopDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex + 1)) + 500) \ 1000
                                    End If
'                                    If tgSpf.sSAudData <> "H" Then
'                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
'                                    Else
'                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex)) + 50) \ 100
'                                    End If
                                    If tgSpf.sSAudData = "H" Then
                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 50) \ 100
                                    ElseIf tgSpf.sSAudData = "N" Then
                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 5) \ 10
                                    ElseIf tgSpf.sSAudData = "U" Then
                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = CLng(smFieldValues(ilAvgQHIndex))
                                    Else
                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
                                    End If
                                Next ilIndex
                                If tmDrfInfo(ilCol).tDrf.sInfoType = "V" Then
                                    '6/9/16: Replaced GoSub
                                    'GoSub mPCStnWriteRec_Mkt
                                    mPCStnWriteRec_Mkt_1 ilRet, ilCustomDemo, slVehicleName, slBookName, slBookDate, ilPrevDnfCode, llPrevDrfCode, ilDemoGender, llPopDemo()
                                Else
                                    'Determine number of matching dayparts
                                    ReDim imMatchRdfCode(0 To 0) As Integer
                                    For ilIndex = LBound(tmRdfInfo) To UBound(tmRdfInfo) - 1 Step 1
                                        ilMatch = False
                                        ilRdf = tmRdfInfo(ilIndex).iRdfIndex
                                        For ilRow = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                            If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
                                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilRow), tgMRdf(ilRdf).iStartTime(1, ilRow), False, llTime
                                                If llTime = tmDrfInfo(ilCol).lStartTime Then
                                                    gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilRow), tgMRdf(ilRdf).iEndTime(1, ilRow), False, llTime
                                                    If llTime = tmDrfInfo(ilCol).lEndTime Then
                                                        ilMatch = True
                                                        For ilDay = 1 To 7 Step 1
                                                            If tgMRdf(ilRdf).sWkDays(ilRow, ilDay - 1) <> tmDrfInfo(ilCol).sDays(ilDay - 1) Then
                                                                ilMatch = False
                                                                Exit For
                                                            End If
                                                        Next ilDay
                                                        'If ilMatch Then
                                                        '    tmDrfInfo(ilCol).tDrf.iRdfCode = tgMRdf(ilRdf).iCode
                                                        'End If
                                                        If ilMatch Then
                                                            Exit For
                                                        End If
                                                    End If
                                                End If
                                                'Exit For
                                            End If
                                        Next ilRow
                                        If ilMatch And (tmDrfInfo(ilCol).lStartTime2 <> -1) Then
                                            ilMatch = False
                                            For ilRow = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                                If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
                                                    gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilRow), tgMRdf(ilRdf).iStartTime(1, ilRow), False, llTime
                                                    If llTime = tmDrfInfo(ilCol).lStartTime2 Then
                                                        gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilRow), tgMRdf(ilRdf).iEndTime(1, ilRow), False, llTime
                                                        If llTime = tmDrfInfo(ilCol).lEndTime2 Then
                                                            ilMatch = True
                                                            imMatchRdfCode(UBound(imMatchRdfCode)) = tgMRdf(ilRdf).iCode
                                                            ReDim Preserve imMatchRdfCode(0 To UBound(imMatchRdfCode) + 1) As Integer
                                                            Exit For
                                                        End If
                                                    End If
                                                    Exit For
                                                End If
                                            Next ilRow
                                        ElseIf ilMatch Then
                                            imMatchRdfCode(UBound(imMatchRdfCode)) = tgMRdf(ilRdf).iCode
                                            ReDim Preserve imMatchRdfCode(0 To UBound(imMatchRdfCode) + 1) As Integer
                                        End If
                                    Next ilIndex
                                    If UBound(imMatchRdfCode) > LBound(imMatchRdfCode) Then
                                        For ilMatch = LBound(imMatchRdfCode) To UBound(imMatchRdfCode) - 1 Step 1
                                            tmDrfInfo(ilCol).tDrf.lCode = 0
                                            tmDrfInfo(ilCol).tDrf.iRdfCode = imMatchRdfCode(ilMatch)
                                            '6/9/16: Replaced GoSub
                                            'GoSub mPCStnWriteRec_Mkt
                                            mPCStnWriteRec_Mkt_1 ilRet, ilCustomDemo, slVehicleName, slBookName, slBookDate, ilPrevDnfCode, llPrevDrfCode, ilDemoGender, llPopDemo()
                                            If ilRet <> 0 Then
                                                mPCStnConvFile_PopByMarket = False
                                                Exit Function
                                            End If
                                        Next ilMatch
                                    Else
                                        tmDrfInfo(ilCol).tDrf.lCode = 0
                                        tmDrfInfo(ilCol).tDrf.iRdfCode = 0
                                        '6/9/16: Replaced GoSub
                                        'GoSub mPCStnWriteRec_Mkt
                                        mPCStnWriteRec_Mkt_1 ilRet, ilCustomDemo, slVehicleName, slBookName, slBookDate, ilPrevDnfCode, llPrevDrfCode, ilDemoGender, llPopDemo()
                                    End If
                                End If
                                If ilRet <> 0 Then
                                    mPCStnConvFile_PopByMarket = False
                                    Exit Function
                                End If
                            Else
                                ilFound = False
                                For ilLoop = 0 To UBound(smUnfdStations) - 1 Step 1
                                    If StrComp(Trim$(smUnfdStations(ilLoop)), Trim$(smFieldValues(ilStnCol)), 1) = 0 Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilLoop
                                If Not ilFound Then
                                    Print #hmTo, "Unable to find " & smFieldValues(ilStnCol)
                                    lbcErrors.AddItem Trim$(smFieldValues(ilStnCol)) & " not found"
                                    smUnfdStations(UBound(smUnfdStations)) = smFieldValues(ilStnCol)
                                    ReDim Preserve smUnfdStations(0 To UBound(smUnfdStations) + 1) As String
                                End If
                            End If
                        End If
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
            If plcGauge.Value <> llPercent Then
                plcGauge.Value = llPercent
            End If
        End If
    Loop Until ilEof
    Close hmFrom
    plcGauge.Value = 100
    mPCStnConvFile_PopByMarket = True
    MousePointer = vbDefault
    Exit Function
'mPCStnConvFile_PopByMarketErr:
'    ilRet = Err.Number
'    Resume Next

'mPCStnWriteRec_Mkt:
'    If ilCustomDemo Then
'        Return
'        'tmDrfPop.sDataType = "B"
'        'tmDrf.sDataType = "B"
'    End If
'    ilSetBkNm = -1
'    ilRet = 0
'    If ((tmDrfInfo(ilCol).iBkNm = 0) And (Trim$(slBookName) <> "")) Then
'        tmDrfInfo(ilCol).tDrf.iVefCode = imVefCode
'        For ilDay = 0 To 6 Step 1
'            smDays(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
'        Next ilDay
'        'If tmDrfInfo(ilCol).iBkNm = 0 Then
'            ilBNRet = mBookNameUsed(slBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfcode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
'        'Else
'        '    ilBNRet = mBookNameUsed(slETBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfCode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
'        'End If
'        If (ilBNRet = 0) Or (ilBNRet = 1) Then
'            tmDnf.iCode = 0
'            tmDnf.sBookName = slBookName
'            gPackDate slBookDate, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
'            gPackDate smNowDate, tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
'            tmDnf.iUrfCode = tgUrf(0).iCode
'            tmDnf.sType = "I"
'            tmDnf.sForm = smDataForm
'            'If tmDrfInfo(ilCol).iBkNm = 0 Then
'                tmDnf.sExactTime = "N"
'            'Else
'            '    tmDnf.sExactTime = "Y"
'            'End If
'            tmDnf.sSource = "A"
'            tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
'            tmDnf.iAutoCode = tmDnf.iCode
'            ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
'            If ilRet <> BTRV_ERR_NONE Then
'                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                    ilRet = csiHandleValue(0, 7)
'                End If
'                Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
'                lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
'                Return
'            End If
'            Do
'                tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
'                tmDnf.iAutoCode = tmDnf.iCode
'                gPackDate smSyncDate, tmDnf.iSyncDate(0), tmDnf.iSyncDate(1)
'                gPackTime smSyncTime, tmDnf.iSyncTime(0), tmDnf.iSyncTime(1)
'                ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
'            Loop While ilRet = BTRV_ERR_CONFLICT
'            If ilRet <> BTRV_ERR_NONE Then
'                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                    ilRet = csiHandleValue(0, 7)
'                End If
'                Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
'                lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
'                Return
'            End If
'            ilPrevDnfCode = tmDnf.iCode
'            ilRet = mObtainBookName()
'            '6/9/16: Replace GoSub
'            'GoSub lCreatePop_Mkt
'            mCreatePop_Mkt
'            tmDrfPop.lCode = 0
'            tmDrfPop.iDnfCode = ilPrevDnfCode
'            For i= 0 To UBound(tmDrfInfo) - 1 Step 1
'                ilDemoGender = tmDrfInfo(ilSet).iDemoGender
'                tmDrfPop.lDemo(ilDemoGender) = llPopDemo(ilDemoGender)
'            Next ilSet
'            tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
'            tmDrfPop.lAutoCode = tmDrfPop.lCode
'            ilRet = btrInsert(hmDrf, tmDrfPop, imDrfRecLen, INDEXKEY2)
'            If ilRet <> BTRV_ERR_NONE Then
'                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                    ilRet = csiHandleValue(0, 7)
'                End If
'                Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'                lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'                Return
'            End If
'            Do
'                tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
'                tmDrfPop.lAutoCode = tmDrfPop.lCode
'                gPackDate smSyncDate, tmDrfPop.iSyncDate(0), tmDrfPop.iSyncDate(1)
'                gPackTime smSyncTime, tmDrfPop.iSyncTime(0), tmDrfPop.iSyncTime(1)
'                ilRet = btrUpdate(hmDrf, tmDrfPop, imDrfRecLen)
'            Loop While ilRet = BTRV_ERR_CONFLICT
'            If ilRet <> BTRV_ERR_NONE Then
'                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                    ilRet = csiHandleValue(0, 7)
'                End If
'                Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'                lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'                Return
'            End If
'        Else
'            If StrComp(smBookForm, smDataForm, vbTextCompare) <> 0 Then
'                If (smDataForm = "8") Or (smBookForm = "8") Then
'                    Print #hmTo, slBookName & " for " & slVehicleName & " previously defined with different format then current import data"
'                    lbcErrors.AddItem "Error in Book Forms" & " for " & slVehicleName
'                    Return
'                End If
'            End If
'            'If tmDrfPop.lCode = 0 Then
'                tmDrfSrchKey.iDnfCode = tmDnf.iCode
'                tmDrfSrchKey.sDemoDataType = "P"
'                tmDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
'                tmDrfSrchKey.iVefCode = 0
'                tmDrfSrchKey.sInfoType = ""
'                tmDrfSrchKey.iRdfcode = 0
'                ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'                If (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tmDnf.iCode) And (tlDrf.iVefCode = 0) And (tlDrf.sDemoDataType = "P") Then
'                    tmDrfPop.lCode = tlDrf.lCode
'                    Do
'                        tmDrfSrchKey2.lCode = tmDrfPop.lCode
'                        ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'                        If ilRet <> BTRV_ERR_NONE Then
'                            Exit Do
'                        End If
'                        'tlDrf.lDemo(ilDemoGender) = llPopDemo
'                        For i= 0 To UBound(tmDrfInfo) - 1 Step 1
'                            ilDemoGender = tmDrfInfo(ilSet).iDemoGender
'                            tlDrf.lDemo(ilDemoGender) = llPopDemo(ilDemoGender)
'                        Next ilSet
'                        gPackDate smSyncDate, tlDrf.iSyncDate(0), tlDrf.iSyncDate(1)
'                        gPackTime smSyncTime, tlDrf.iSyncTime(0), tlDrf.iSyncTime(1)
'                        ilRet = btrUpdate(hmDrf, tlDrf, imDrfRecLen)
'                    Loop While ilRet = BTRV_ERR_CONFLICT
'                Else
'                    tmDrfPop.lCode = -1
'                End If
'            'End If
'        End If
'        tmDrfInfo(ilCol).tDrf.iDnfCode = ilPrevDnfCode
'        If llPrevDrfCode = 0 Then
'            tmDrfInfo(ilCol).tDrf.lCode = 0
'            tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
'            tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
'            ilRet = btrInsert(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen, INDEXKEY2)
'            If ilRet <> BTRV_ERR_NONE Then
'                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                    ilRet = csiHandleValue(0, 7)
'                End If
'                Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'                lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'                Return
'            End If
'            Do
'                tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
'                tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
'                gPackDate smSyncDate, tmDrfInfo(ilCol).tDrf.iSyncDate(0), tmDrfInfo(ilCol).tDrf.iSyncDate(1)
'                gPackTime smSyncTime, tmDrfInfo(ilCol).tDrf.iSyncTime(0), tmDrfInfo(ilCol).tDrf.iSyncTime(1)
'                ilRet = btrUpdate(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen)
'            Loop While ilRet = BTRV_ERR_CONFLICT
'        Else
'            Do
'                tmDrfSrchKey2.lCode = llPrevDrfCode
'                ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'                If ilRet <> BTRV_ERR_NONE Then
'                    Exit Do
'                End If
'                'tlDrf.lDemo(ilDemoGender) = tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender)
'                For i= 0 To UBound(tmDrfInfo) - 1 Step 1
'                    ilDemoGender = tmDrfInfo(ilSet).iDemoGender
'                    tlDrf.lDemo(ilDemoGender) = tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender)
'                Next ilSet
'                gPackDate smSyncDate, tlDrf.iSyncDate(0), tlDrf.iSyncDate(1)
'                gPackTime smSyncTime, tlDrf.iSyncTime(0), tlDrf.iSyncTime(1)
'                ilRet = btrUpdate(hmDrf, tlDrf, imDrfRecLen)
'            Loop While ilRet = BTRV_ERR_CONFLICT
'        End If
'        If ilRet <> BTRV_ERR_NONE Then
'            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                ilRet = csiHandleValue(0, 7)
'            End If
'            Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'            lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'            'mSvPCStnConvFile = False
'            'Exit Function
'            Return
'        End If
'        If tmDrfInfo(ilCol).iBkNm = 0 Then
'            ilSetBkNm = 0
'            ilSetDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
'        End If
'    End If
'    If ((ckcDefault(0).Value = vbChecked) Or (ckcDefault(1).Value = vbChecked)) And (ilSetBkNm = 0) Then
'        Do
'            tmVefSrchKey.iCode = imVefCode
'            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'            If ilRet <> BTRV_ERR_NONE Then
'                Exit Do
'            End If
'            If ckcDefault(0).Value = vbChecked Then
'                tmVef.iDnfCode = ilSetDnfCode
'            End If
'            If ckcDefault(1).Value = vbChecked Then
'                tmVef.iReallDnfCode = ilSetDnfCode
'            End If
'            'tmVef.iSourceID = tgUrf(0).iRemoteUserID
'            'gPackDate smSyncDate, tmVef.iSyncDate(0), tmVef.iSyncDate(1)
'            'gPackTime smSyncTime, tmVef.iSyncTime(0), tmVef.iSyncTime(1)
'            ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
'        Loop While ilRet = BTRV_ERR_CONFLICT
'        ilRet = gBinarySearchVef(tmVef.iCode)
'        If ilRet <> -1 Then
'            tgMVef(ilRet) = tmVef
'        End If
'    End If
'    ilRet = 0
'    Return
'lCreatePop_Mkt:
'    tmDrfPop.lCode = 0
'    tmDrfPop.iDnfCode = tmDnf.iCode
'    tmDrfPop.sDemoDataType = "P"
'    tmDrfPop.iMnfSocEco = 0
'    tmDrfPop.iVefCode = 0
'    tmDrfPop.sInfoType = ""
'    tmDrfPop.iRdfcode = 0
'    tmDrfPop.sProgCode = ""
'    tmDrfPop.iStartTime(0) = 1
'    tmDrfPop.iStartTime(1) = 0
'    tmDrfPop.iEndTime(0) = 1
'    tmDrfPop.iEndTime(1) = 0
'    tmDrfPop.iStartTime2(0) = 1
'    tmDrfPop.iStartTime2(1) = 0
'    tmDrfPop.iEndTime2(0) = 1
'    tmDrfPop.iEndTime2(1) = 0
'    For ilDay = 0 To 6 Step 1
'        tmDrfPop.sDay(ilDay) = "Y"
'    Next ilDay
'    tmDrfPop.iQHIndex = 0
'    tmDrfPop.iCount = 0
'    tmDrfPop.sExStdDP = "N"
'    tmDrfPop.sExRpt = "N"
'    tmDrfPop.sDataType = "A"
'    For i= 1 To 16 Step 1
'        tmDrfPop.lDemo(ilSet) = 0
'    Next ilSet
'    Return
End Function




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
    Screen.MousePointer = vbDefault

    ' JD 09-13-21  This statement no longer applies.
    'If bmResearchSaved Then
    '    If (Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES Then
    '        ilRet = MsgBox("Please update the Vehicle default ACT1 Lineup codes if required", vbOKOnly + vbInformation, "Warning")
    '    End If
    'End If

    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ImptMark
    igManUnload = NO
End Sub
Private Function mTestColumnTitle(bIsNewVersion As Boolean) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    mTestColumnTitle = False
    ilRet = 0
    slLine = ""
    err.Clear
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
                If rbcImportFrom(0).Value Then
                    'Determine field Type
                    'Header Record
                    '   Demographics,AQH,AQH,....,Conv%,Population,,
                    '   ,,Rtg.,,.....
                    'Demo Record
                    '   Males   12-17, 12000, 0.1,..., 21847000
                    '   Females 12-17,
                    '   Men     18-24,
                    '   Women   18-24,
                    '
                    If bIsNewVersion Then
                        If InStr(1, Trim$(slLine), "CustomData", 1) > 0 Then
                            mTestColumnTitle = True
                            Exit Function
                        End If
                    Else
                        If InStr(1, Trim$(slLine), "Demographic", 1) > 0 Then
                            If InStr(1, Trim$(slLine), "AQH", 1) > 0 Then
                                If InStr(1, Trim$(slLine), "Population", 1) > 0 Then
                                    mTestColumnTitle = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Else
                    'Determine field Type
                    'Header Record
                    '   Continental U.S.  Continental U.S.
                    '       AQH AUD   RTG      Populations
                    'Demo Record
                    '   Males   12-17 12,000 0.1 21,847,000
                    '   Females 12-17
                    '   Men     18-24
                    '   Women   18-24
                    '
                    If InStr(1, Trim$(slLine), "AQH AUD", 1) > 0 Then
                        If InStr(1, Trim$(slLine), "Populations", 1) > 0 Then
                            mTestColumnTitle = True
                            Exit Function
                        End If
                    End If
                End If
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
    If ilSize <> Len(tmDrfInfo(0).tDrf) Then
        If ilSize > 0 Then
            MsgBox "Drf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmDrfInfo(0).tDrf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
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
    If tgSaf(0).sAudByPackage <> "Y" Then
        ilRet = gPopUserVehicleBox(ImptMark, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH, cbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    Else
        ilRet = gPopUserVehicleBox(ImptMark, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHSTDPKG + VEHVIRTUAL + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER + ACTIVEVEH, cbcVehicle, tgUserVehicle(), sgUserVehicleTag)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle)", ImptMark
        On Error GoTo 0
    End If
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub cbcBookName_Click(Index As Integer)
    imComboBoxIndex = cbcBookName(Index).ListIndex
    cbcBookName_Change Index
End Sub

Private Sub rbcImportFrom_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcImportFrom(Index).Value
    'End of coded added
    If Value Then
        If Index = 0 Then
            plcPCForm.Visible = True
        Else
            plcPCForm.Visible = False
        End If
    End If
End Sub
Private Sub rbcImportFrom_GotFocus(Index As Integer)
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
End Sub

Private Sub rbcNew_Click(Index As Integer)
    If rbcNew(Index).Value Then
        If Index = 1 Then
            edcBookName(1).Visible = False
            edcBookName(0).Visible = False
            cbcBookName(0).Visible = True
            '7582
            If Not rbcPCForm(2).Value Then
                cbcBookName(1).Visible = True
            End If
        Else
            cbcBookName(1).Visible = False
            cbcBookName(0).Visible = False
            '7582
            If Not rbcPCForm(2).Value Then
                edcBookName(1).Visible = True
            End If
            edcBookName(0).Visible = True
        End If
    End If
End Sub

Private Sub rbcPCForm_Click(Index As Integer)
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = rbcPCForm(Index).Value
    'End of coded added
    If Value Then
        If Index = 1 Then
            ''Save if add Station Import with Population by USA
            'lacBookName(1).Visible = False
            'edcBookName(1).Visible = False
            'lacBookName(0).Caption = "Book Name"
            ''Save if add Station Import with Population by USA
            lacBookName(1).Visible = False
            edcBookName(1).Visible = False
            edcBookName(0).Visible = False
            cbcBookName(1).Visible = False
            cbcBookName(0).Visible = False
            plcNew.Visible = False
            lacBookName(0).Width = 7200
            lacBookName(0).Caption = "Book Name Created by Combining Survey Period with Market Name"
        '7582
        ElseIf Index = 2 Then
            lacBookName(1).Visible = True
            lacBookName(0).Width = 1890
            plcNew.Visible = False
            edcBookName(1).Visible = False
            cbcBookName(0).Visible = False
            cbcBookName(1).Visible = False
            lacBookName(1).Visible = False
            edcBookName(0).Visible = True
'            If rbcNew(0).Value Then
'                edcBookName(0).Visible = True
'            Else
'                lbcBookName(0).Visible = True
'            End If
            lacBookName(0).Caption = "Book Name:"
            If Len(edcBookName(0).Text) = 0 Then
                edcBookName(0).Text = mParseForBookName(edcFrom.Text)
            End If
        Else
            lacBookName(1).Visible = True
            lacBookName(0).Width = 1890
            plcNew.Visible = True
            'dan 6/12/15
            edcBookName(1).Visible = False
            cbcBookName(0).Visible = False
            cbcBookName(1).Visible = False
            edcBookName(0).Visible = False
            If rbcNew(0).Value Then
                edcBookName(1).Visible = True
                edcBookName(0).Visible = True
            Else
                cbcBookName(1).Visible = True
                cbcBookName(0).Visible = True
            End If
            lacBookName(0).Caption = "Book Name: Daypart"
        End If
    End If
End Sub
Private Sub plcPCForm_Paint()
    plcPCForm.CurrentX = 0
    plcPCForm.CurrentY = 0
    '7582
    'plcPCForm.Print "PC Research Source"
    plcPCForm.Print "Act 1 Report Source"
End Sub
Private Sub plcImportFrom_Paint()
    plcImportFrom.CurrentX = 0
    plcImportFrom.CurrentY = 0
    plcImportFrom.Print "Import Generated on"
End Sub
Private Sub plcDefault_Paint()
    plcDefault.CurrentX = 0
    plcDefault.CurrentY = 0
    plcDefault.Print "Set as Vehicle Default"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Import Act 1 Data"
End Sub

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

    'Remove test 3/10/04 as ACT 1 is not allowed to have 20 or 21 except within a conbination with other demos
    mPrepassDemo = True
    smDataForm = "6"
    Exit Function

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
    slLine = ""
    err.Clear
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

'*******************************************************
'*                                                     *
'*      Procedure Name:mConvertStationVehicles      *
'*                                                     *
'*             Created:6/9/15       By:D. Michaelson      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert File                   *
'*                                                     *
'*******************************************************
Private Function mConvertStationVehicles(slFromFile As String, slBookName As String, slBookDate As String) As Integer
' This was modelled after mPCStnConvFile_PopByMarket.  It would be nice to fix it up.
' we go and find the 'AQH'. that tells us where the category columns will be in the previous lines,
' so we go back through the lines and pick up the categories along with the daypart. That's the header
' now we go through each line between 'Rank' and 'Total' and gather the numbers.  After that, we look for USA Pop line
' for populations.  Then we go ot the next AQH and do it again.
' But really, we only need pop and categories one time.  We currently do pop just once (and 'extra'), but the other is
' done over and over.
' It's also designed to get a new book name for each, but it's the same for the whole file.
    Dim ilRet As Integer
    Dim ilBNRet As Integer
    Dim slLine As String
    Dim ilHeaderFd As Integer
    Dim ilAvgQHIndex As Integer
    Dim ilPopDone As Integer
    Dim ilDayTimeFd As Integer
    Dim ilDemoGender As Integer
    Dim slDemoAge As String
    Dim ilPos As Integer
    Dim slDay As String
    Dim ilDay As Integer
    Dim ilSY As Integer
    Dim ilEY As Integer
    Dim ilLoop As Integer
    Dim ilEof As Integer
    Dim llPercent As Long
    Dim slChar As String
    Dim slStr As String
    Dim slSexChar As String
    Dim ilIndex As Integer
    Dim ilCustomDemo As Integer
    Dim ilPrevDnfCode As Integer
    Dim llPrevDrfCode As Long
    Dim ilRdf As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim slVehicleName As String
    Dim llTime As Long
    'Gender index 1 to 16
    'ReDim llPopDemo(1 To 16) As Long
    ReDim llPopDemo(0 To 16) As Long
    Dim ilTIndex As Integer
    Dim ilSetBkNm As Integer
    Dim ilStnCol As Integer
    Dim ilSetDnfCode As Integer
    Dim ilDemoFd As Integer
    Dim slSTime As String
    Dim slETime As String
    Dim ilMatch As Integer
    Dim ilFound As Integer
    Dim ilSet As Integer
    Dim tlDrf As DRF
    Dim blPopFound As Boolean
    Dim blCategoriesFound As Boolean
    Dim blTempCatFound As Boolean
    Dim ilLoop1 As Integer
    
    'On Error GoTo mConvertStationVehiclesErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        MsgBox "Open " & slFromFile & ", Error #" & Str$(ilRet), vbExclamation, "Open Error"
        edcFrom.SetFocus
        mConvertStationVehicles = False
        GoTo Cleanup
    End If
    DoEvents
    If imTerminate Then
        mTerminate
        mConvertStationVehicles = False
        GoTo Cleanup
    End If
    blPopFound = False
    blCategoriesFound = False
    blTempCatFound = False
    ilHeaderFd = False
    ilDayTimeFd = False
    ReDim tmExtraDemo(0)
    ReDim lmDpfWithoutPop(0)
    ilPopDone = False
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
    tmDrfPop.sACTLineupCode = ""
    tmDrfPop.sACT1StoredTime = ""
    tmDrfPop.sACT1StoredSpots = ""
    tmDrfPop.sACT1StoreClearPct = ""
    tmDrfPop.sACT1DaypartFilter = ""
    
    ilCustomDemo = False
 '   ilBookCol = -1
    Do
        ilRet = 0
        err.Clear
        'On Error GoTo mConvertStationVehiclesErr:
        slLine = ""
        'Line Input #hmFrom, slLine
        Do
            If Not EOF(hmFrom) Then
                slChar = Input(1, #hmFrom)
                If slChar = Chr(13) Then 'Carriage return
                    slChar = Input(1, #hmFrom)
                End If
                If slChar = Chr(10) Then 'line feed
                    Exit Do
                End If
                slLine = slLine & slChar
            Else
                ilEof = True
                Exit Do
            End If
        Loop
        slLine = Trim$(slLine)
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
                    mTerminate
                    mConvertStationVehicles = False
                    GoTo Cleanup
                End If
                gParseCDFields slLine, False, smFieldValues()
                '11/10/16: Move fields from 0 to X to 1 to X +1
                For ilLoop = UBound(smFieldValues) - 1 To LBound(smFieldValues) Step -1
                    smFieldValues(ilLoop + 1) = smFieldValues(ilLoop)
                Next ilLoop
                smFieldValues(0) = ""
            'example of file:
                '"Nielsen Audio DMA Area"
            'book name
                '"Fall Nationwide 2014[¿]"
                '"AFCP: Westwood One FM Connection (5/07/15)"
            'daypart
                'Schedule: MSu 6-10a,,,,,Boys 12-17,,,,Men 18-24
            'get AQH columns here
                'DMA,Market,Station, ,MSL, ,AQH,In-Tab,, ,AQH,
                'Rank,,,,StnID,,,
            'Values by station
                '   25,Raleigh-Durham (Fayette..[PPM+D],WZFX-FM,+,10114, ,1500,
                '..etc
            'End of values:
                'Total
            ''junk'
                'Cov Pop
                'Coverage Rtg
            'population here
                'USA Pop
            ''junk'
                'USA Rtg
                'etc
            'end of junk looks like this:
                'please see the ""ACT 1 Accuracy Expectation"" document.
            'then start again with daypart
                'Schedule: SS 3-7p,,,,,Boys 12-17
                'only run this if still looking for header or dayparts
                If (Not ilHeaderFd) Or (Not ilDayTimeFd) Then
                    'get population here
                    If InStr(1, Trim$(smFieldValues(1)), "USA Pop", vbTextCompare) > 0 Then
                        If Not blPopFound Then
                            For ilIndex = 0 To UBound(tmDrfInfo) - 1 Step 1
                                ilDemoGender = tmDrfInfo(ilIndex).iDemoGender
                                ilAvgQHIndex = tmDrfInfo(ilIndex).iStartCol
                                If ilDemoGender > -1 Then
                                    If tgSpf.sSAudData = "H" Then
                                        llPopDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex)) + 50) \ 100
                                    ElseIf tgSpf.sSAudData = "N" Then
                                        llPopDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex)) + 5) \ 10
                                    ElseIf tgSpf.sSAudData = "U" Then
                                        llPopDemo(ilDemoGender) = CLng(smFieldValues(ilAvgQHIndex))
                                    Else
                                        llPopDemo(ilDemoGender) = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
                                    End If
                                Else
                                    If Not mExtraDemoOrPop(ilAvgQHIndex, False) Then
        'Dan to do
                                    End If
                                End If
                            Next ilIndex
                            blPopFound = True
                        End If
                    ElseIf (InStr(1, Trim$(slLine), "Market", 1) > 0) And (InStr(1, Trim$(slLine), "Station", 1) > 0) Then
                        If InStr(1, RTrim$(slLine), "AQH", 1) > 0 Then
                            'Determine number of Demo columns
                            ilHeaderFd = True
                            ilPos = 2
                            ReDim tmDrfInfo(0 To 0) As DRFINFO
                            Do While ilPos <= UBound(smFieldValues)
                                If InStr(1, Trim$(smFieldValues(ilPos)), "AQH", 1) > 0 Then
                                    tmDrfInfo(UBound(tmDrfInfo)).iStartCol = ilPos
                                    tmDrfInfo(UBound(tmDrfInfo)).iType = 0  'Vehicle
                                    tmDrfInfo(UBound(tmDrfInfo)).iBkNm = 0  'Daypart Book Name
                                    tmDrfInfo(UBound(tmDrfInfo)).iSY = -1
                                    tmDrfInfo(UBound(tmDrfInfo)).iEY = -1
                                    tmDrfInfo(UBound(tmDrfInfo)).iDemoGender = -1
                                    ReDim Preserve tmDrfInfo(0 To UBound(tmDrfInfo) + 1) As DRFINFO
                                    ilPos = ilPos + 1
                                ElseIf InStr(1, Trim$(smFieldValues(ilPos)), "Station", 1) > 0 Then
                                    ilStnCol = ilPos
                                End If
                                ilPos = ilPos + 1
                            Loop
                            'Get Demo and Dayparts
                            ilDayTimeFd = False
                            ilCustomDemo = False
                            'go backwards to get daypart and categories, and book
                            'For ilLoop = 10 To 1 Step -1
                            For ilLoop = 9 To 0 Step -1
                                If Len(Trim$(smSvLine(ilLoop))) > 0 Then
                                    gParseCDFields smSvLine(ilLoop), False, smFieldValues()
                                    '1/23/19: Move fields from 0 to X to 1 to X +1
                                    For ilLoop1 = UBound(smFieldValues) - 1 To LBound(smFieldValues) Step -1
                                        smFieldValues(ilLoop1 + 1) = smFieldValues(ilLoop1)
                                    Next ilLoop1
                                    smFieldValues(0) = ""
                                    
                                    ilPos = InStr(1, smFieldValues(1), "Schedule:", 1)
                                    If ilPos > 0 Then
                                        ilCol = 0
                                        smFieldValues(1) = Trim$(Mid$(smFieldValues(1), ilPos + 9))
                                        slStr = UCase$(Trim$(smFieldValues(1)))
                                        ilSY = -1
                                        slDay = Trim$(Mid$(slStr, 1, 3))
                                        ilTIndex = InStr(1, slStr, " ", 1) + 1
                                        'Select Case slDay
                                        '    Case "MO", "MON"
                                        '        ilSY = 0
                                        '        ilEY = 0
                                        '    Case "TU", "TUE"
                                        '        ilSY = 1
                                        '        ilEY = 1
                                        '    Case "WE", "WED"
                                        '        ilSY = 2
                                        '        ilEY = 2
                                        '    Case "TH", "THU"
                                        '        ilSY = 3
                                        '        ilEY = 3
                                        '    Case "FR", "FRI"
                                        '        ilSY = 4
                                        '        ilEY = 4
                                        '    Case "SA", "SAT"
                                        '        ilSY = 5
                                        '        ilEY = 5
                                        '        If Mid$(slStr, 3, 2) = "SU" Then
                                        '            ilEY = 6
                                        '        End If
                                        '    Case "SU", "SUN"
                                        '        ilSY = 6
                                        '        ilEY = 6
                                        '    Case "MF"
                                        '        ilSY = 0
                                        '        ilEY = 4
                                        '    Case "MSA"
                                        '        ilSY = 0
                                        '        ilEY = 5
                                        '    Case "MSU"
                                        '        ilSY = 0
                                        '        ilEY = 6
                                        '    Case "MS"
                                        '        ilSY = 0
                                        '        ilEY = 6
                                        '    Case "SS"
                                        '        ilSY = 5
                                        '        ilEY = 6
                                        '    Case "FSU"
                                        '        ilSY = 4
                                        '        ilEY = 6
                                        '    Case "FSA"
                                        '        ilSY = 4
                                        '        ilEY = 5
                                        'End Select
                                        If ilTIndex > 2 Then
                                            mGetDayIndex Left$(slStr, ilTIndex - 2), ilSY, ilEY
                                        End If
                                        
                                        If ilSY <> -1 Then
                                            tmDrfInfo(ilCol).lStartTime2 = -1
                                            slChar = Mid$(slStr, 1, 1)
                                            Do While (slChar < "0") Or (slChar > "9")
                                                slStr = Mid$(slStr, 2)
                                                slChar = Mid$(slStr, 1, 1)
                                            Loop
                                            slSTime = ""
                                            slChar = UCase$(Mid$(slStr, 1, 1))
                                            Do
                                                slSTime = slSTime & slChar
                                                If (slChar = "A") Or (slChar = "P") Then
                                                    Exit Do
                                                End If
                                                slStr = Mid$(slStr, 2)
                                                slChar = UCase$(Mid$(slStr, 1, 1))
                                            Loop While (slChar <> "-")
                                            If gValidTime(slSTime) Then
                                                slChar = Mid$(slStr, 1, 1)
                                                Do While (slChar < "0") Or (slChar > "9")
                                                    slStr = Mid$(slStr, 2)
                                                    slChar = Mid$(slStr, 1, 1)
                                                Loop
                                                slETime = ""
                                                slChar = UCase$(Mid$(slStr, 1, 1))
                                                Do
                                                    slETime = slETime & slChar
                                                    If (slChar = "A") Or (slChar = "P") Then
                                                        Exit Do
                                                    End If
                                                    If Len(slStr) = 0 Then
                                                        Exit Do
                                                    End If
                                                    slStr = Mid$(slStr, 2)
                                                    slChar = UCase$(Mid$(slStr, 1, 1))
                                                Loop While (slChar <> "/")
                                                If gValidTime(slETime) Then
                                                    ilDayTimeFd = True
                                                    If (right$(slSTime, 1) >= "0") And (right$(slSTime, 1) <= "9") Then
                                                        slSTime = slSTime & right$(slETime, 1)
                                                    End If
                                                    tmDrfInfo(ilCol).lStartTime = gTimeToCurrency(slSTime, False)
                                                    tmDrfInfo(ilCol).iSY = ilSY
                                                    tmDrfInfo(ilCol).iEY = ilEY
                                                    tmDrfInfo(ilCol).lEndTime = gTimeToCurrency(slETime, False)
                                                    For ilDay = 0 To 6 Step 1
                                                        tmDrfInfo(ilCol).sDays(ilDay) = "N"
                                                    Next ilDay
                                                    For ilDay = ilSY To ilEY Step 1
                                                        tmDrfInfo(ilCol).sDays(ilDay) = "Y"
                                                    Next ilDay
                                                    tmDrfInfo(ilCol).iType = 1  'Daypart

                                                    If InStr(1, slStr, "/", 1) > 0 Then
                                                        slChar = Mid$(slStr, 1, 1)
                                                        Do While (slChar < "0") Or (slChar > "9")
                                                            slStr = Mid$(slStr, 2)
                                                            slChar = Mid$(slStr, 1, 1)
                                                        Loop
                                                        slSTime = ""
                                                        slChar = UCase$(Mid$(slStr, 1, 1))
                                                        Do
                                                            slSTime = slSTime & slChar
                                                            If (slChar = "A") Or (slChar = "P") Then
                                                                Exit Do
                                                            End If
                                                            slStr = Mid$(slStr, 2)
                                                            slChar = UCase$(Mid$(slStr, 1, 1))
                                                        Loop While (slChar <> "-")
                                                        If gValidTime(slSTime) Then
                                                            slChar = Mid$(slStr, 1, 1)
                                                            Do While (slChar < "0") Or (slChar > "9")
                                                                slStr = Mid$(slStr, 2)
                                                                slChar = Mid$(slStr, 1, 1)
                                                            Loop
                                                            slETime = ""
                                                            slChar = UCase$(Mid$(slStr, 1, 1))
                                                            Do
                                                                slETime = slETime & slChar
                                                                If (slChar = "A") Or (slChar = "P") Then
                                                                    Exit Do
                                                                End If
                                                                If Len(slStr) = 0 Then
                                                                    Exit Do
                                                                End If
                                                                slStr = Mid$(slStr, 2)
                                                                slChar = UCase$(Mid$(slStr, 1, 1))
                                                            Loop While (slChar >= "0") And (slChar <= "9")
                                                            If gValidTime(slETime) Then
                                                                If (right$(slSTime, 1) >= "0") And (right$(slSTime, 1) <= "9") Then
                                                                    slSTime = slSTime & right$(slETime, 1)
                                                                End If
                                                                tmDrfInfo(ilCol).lStartTime2 = gTimeToCurrency(slSTime, False)
                                                                tmDrfInfo(ilCol).lEndTime2 = gTimeToCurrency(slETime, False)
                                                            End If ' valid end time again
                                                        End If ' valid start time again
                                                    End If '/ ?
                                                End If 'Valied EndTime
                                            End If 'gValid StartTime
                                        End If 'ilSy
                                        'gather categories
                                        For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
                                            '7852 columns don't line up.  It's one less
                                            slStr = UCase$(Trim$(smFieldValues(tmDrfInfo(ilCol).iStartCol - 1)))
                                            smFieldValues(1) = slStr
                                            If tmDrfInfo(ilCol).iDemoGender = -1 Then
                                                slSexChar = UCase$(Left$(smFieldValues(1), 1))
                                                If (slSexChar = "M") Or (slSexChar = "B") Or (slSexChar = "W") Or (slSexChar = "F") Or (slSexChar = "G") Or (slSexChar = "A") Or (slSexChar = "P") Or (slSexChar = "T") Then
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
                                                    If (slSexChar <> "A") And (slSexChar <> "P") And (slSexChar <> "T") Then
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
                                                                    ilDemoFd = True
                                                                Case "18-24"
                                                                    ilDemoGender = ilDemoGender + 2
                                                                    ilDemoFd = True
                                                                Case "25-34"
                                                                    ilDemoGender = ilDemoGender + 3
                                                                    ilDemoFd = True
                                                                Case "35-44"
                                                                    ilDemoGender = ilDemoGender + 4
                                                                    ilDemoFd = True
                                                                Case "45-49"
                                                                    ilDemoGender = ilDemoGender + 5
                                                                    ilDemoFd = True
                                                                Case "50-54"
                                                                    ilDemoGender = ilDemoGender + 6
                                                                    ilDemoFd = True
                                                                Case "55-64"
                                                                    ilDemoGender = ilDemoGender + 7
                                                                    ilDemoFd = True
                                                                Case "65+"
                                                                    ilDemoGender = ilDemoGender + 8
                                                                    ilDemoFd = True
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
                                                                    ilDemoFd = True
                                                                Case "18-20"
                                                                    ilDemoGender = ilDemoGender + 2
                                                                    ilDemoFd = True
                                                                Case "21-24"
                                                                    ilDemoGender = ilDemoGender + 3
                                                                    ilDemoFd = True
                                                                Case "25-34"
                                                                    ilDemoGender = ilDemoGender + 4
                                                                    ilDemoFd = True
                                                                Case "35-44"
                                                                    ilDemoGender = ilDemoGender + 5
                                                                    ilDemoFd = True
                                                                Case "45-49"
                                                                    ilDemoGender = ilDemoGender + 6
                                                                    ilDemoFd = True
                                                                Case "50-54"
                                                                    ilDemoGender = ilDemoGender + 7
                                                                    ilDemoFd = True
                                                                Case "55-64"
                                                                    ilDemoGender = ilDemoGender + 8
                                                                    ilDemoFd = True
                                                                Case "65+"
                                                                    ilDemoGender = ilDemoGender + 9
                                                                    ilDemoFd = True
                                                                Case Else
                                                                    ilDemoGender = -1
                                                            End Select
                                                        End If 'dataForm not 8
                                                    Else
                                                        slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
                                                        ilDemoGender = -1
                                                    End If
                                                    If ilDemoGender > -1 Then
                                                        tmDrfInfo(ilCol).iDemoGender = ilDemoGender
                                                    ElseIf Not blCategoriesFound Then
                                                        blTempCatFound = True
                                                        If Not mFindExraCategory(slSexChar, slDemoAge, tmDrfInfo(ilCol).iStartCol) Then
                                                            Print #hmTo, "invalid category " & smFieldValues(1)
                                                            lbcErrors.AddItem "invalid category " & smFieldValues(1)
                                                        End If
                                                    End If
                                                End If 'ilPos numbers can be parsed
                                            End If 'ilDemoGender = -1?
                                        Next ilCol
                                        'set this after the whole row is done.
                                        If blTempCatFound Then
                                            blCategoriesFound = True
                                        End If
                                        If (ilDayTimeFd) Then
                                            Exit For
                                        End If
                                    End If 'backward...schedule in line?
                                End If  ' backwards line has length?
                            'go backwards to get categories and daypart
                            Next ilLoop
                        Else
                            'Save Last 10 lines
                            'For ilLoop = 2 To 10 Step 1
                            For ilLoop = 1 To 9 Step 1
                                smSvLine(ilLoop - 1) = smSvLine(ilLoop)
                            Next ilLoop
                            'smSvLine(10) = slLine
                            smSvLine(9) = slLine
                        End If  'AQH in line
                    Else
                        'Save Last 10 lines
                        'For ilLoop = 2 To 10 Step 1
                        For ilLoop = 1 To 9 Step 1
                            smSvLine(ilLoop - 1) = smSvLine(ilLoop)
                        Next ilLoop
                        'smSvLine(10) = slLine
                        smSvLine(9) = slLine
                    End If ' market/station in line
                'found header and dayparts, now let's look for data
                ElseIf InStr(1, smFieldValues(1), "Total", 1) > 0 Then
                    ilHeaderFd = False
                ElseIf Trim$(smFieldValues(ilStnCol)) <> "" Then
                    If UBound(tmDrfInfo) > LBound(tmDrfInfo) Then
                        ilCol = 0
                        imVefCode = 0
                        For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            '9291
                            'ilPos = InStr(1, tgMVef(ilIndex).sName, smFieldValues(ilStnCol), 1)
                            'If ilPos > 0 Then
                            If UCase(Trim$(tgMVef(ilIndex).sName)) = UCase(Trim$(smFieldValues(ilStnCol))) Then
                                imVefCode = tgMVef(ilIndex).iCode
                                Exit For
                            End If
                        Next ilIndex
                        If imVefCode > 0 Then
                            slVehicleName = Trim$(smFieldValues(ilStnCol))
                            mDPPop
                            For ilLoop = 1 To 16 Step 1
                                tmDrfInfo(ilCol).tDrf.lDemo(ilLoop - 1) = 0
                            Next ilLoop
                            tmDrfInfo(ilCol).tDrf.lCode = 0
                            tmDrfInfo(ilCol).tDrf.iDnfCode = tmDnf.iCode
                            tmDrfInfo(ilCol).tDrf.sDemoDataType = "D"
                            tmDrfInfo(ilCol).tDrf.iMnfSocEco = 0
                            tmDrfInfo(ilCol).iBkNm = 0  'Book Name
                            tmDrfInfo(ilCol).tDrf.iStartTime2(0) = 1
                            tmDrfInfo(ilCol).tDrf.iStartTime2(1) = 0
                            tmDrfInfo(ilCol).tDrf.iEndTime2(0) = 1
                            tmDrfInfo(ilCol).tDrf.iEndTime2(1) = 0
                            If tmDrfInfo(ilCol).iType = 0 Then
                                tmDrfInfo(ilCol).tDrf.sInfoType = "V"
                                tmDrfInfo(ilCol).tDrf.iRdfCode = 0
                                gPackTime "12AM", tmDrfInfo(ilCol).tDrf.iStartTime(0), tmDrfInfo(ilCol).tDrf.iStartTime(1)
                                gPackTime "12AM", tmDrfInfo(ilCol).tDrf.iEndTime(0), tmDrfInfo(ilCol).tDrf.iEndTime(1)
                            Else
                                tmDrfInfo(ilCol).tDrf.sInfoType = "D"
                                tmDrfInfo(ilCol).tDrf.iRdfCode = 0
                                gPackTimeLong tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).tDrf.iStartTime(0), tmDrfInfo(ilCol).tDrf.iStartTime(1)
                                gPackTimeLong tmDrfInfo(ilCol).lEndTime, tmDrfInfo(ilCol).tDrf.iEndTime(0), tmDrfInfo(ilCol).tDrf.iEndTime(1)
                                For ilDay = 0 To 6 Step 1
                                    tmDrfInfo(ilCol).tDrf.sDay(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
                                Next ilDay
                                If tmDrfInfo(ilCol).lStartTime2 <> -1 Then
                                    gPackTimeLong tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).tDrf.iStartTime(0), tmDrfInfo(ilCol).tDrf.iStartTime(1)
                                    gPackTimeLong tmDrfInfo(ilCol).lEndTime, tmDrfInfo(ilCol).tDrf.iEndTime(0), tmDrfInfo(ilCol).tDrf.iEndTime(1)
                                    gPackTimeLong tmDrfInfo(ilCol).lStartTime2, tmDrfInfo(ilCol).tDrf.iStartTime2(0), tmDrfInfo(ilCol).tDrf.iStartTime2(1)
                                    gPackTimeLong tmDrfInfo(ilCol).lEndTime2, tmDrfInfo(ilCol).tDrf.iEndTime2(0), tmDrfInfo(ilCol).tDrf.iEndTime2(1)
                                End If
                            End If
                            tmDrfInfo(ilCol).tDrf.sProgCode = ""
                            tmDrfInfo(ilCol).tDrf.iQHIndex = 0
                            tmDrfInfo(ilCol).tDrf.iCount = 0
                            tmDrfInfo(ilCol).tDrf.sExStdDP = "N"
                            tmDrfInfo(ilCol).tDrf.sExRpt = "N"
                            tmDrfInfo(ilCol).tDrf.sDataType = "A"
'handle unfound category--handle better!
On Error GoTo mConvertStationVehiclesErr:
                            For ilIndex = 0 To UBound(tmDrfInfo) - 1 Step 1
                                ilDemoGender = tmDrfInfo(ilIndex).iDemoGender
                                'get the demo data
                                ilAvgQHIndex = tmDrfInfo(ilIndex).iStartCol
                                If ilDemoGender > -1 Then
                                    If tgSpf.sSAudData = "H" Then
                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 50) \ 100
                                    ElseIf tgSpf.sSAudData = "N" Then
                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 5) \ 10
                                    ElseIf tgSpf.sSAudData = "U" Then
                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = CLng(smFieldValues(ilAvgQHIndex))
                                    Else
                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
                                    End If
                                Else
                                    If Not mExtraDemoOrPop(ilAvgQHIndex, True) Then
                                        'Dan anything here?
                                    End If
                                End If
                            Next ilIndex
                            On Error GoTo 0
                            If tmDrfInfo(ilCol).tDrf.sInfoType = "V" Then
                                '6/9/16: Replaced GoSub
                                'GoSub mPCStnWriteRec_Mkt
                                mPCStnWriteRec_Mkt_2 ilRet, ilCustomDemo, slVehicleName, slBookName, slBookDate, ilPrevDnfCode, llPrevDrfCode, ilDemoGender, llPopDemo(), blPopFound
                            Else
                                'Determine number of matching dayparts
                                ReDim imMatchRdfCode(0 To 0) As Integer
                                For ilIndex = LBound(tmRdfInfo) To UBound(tmRdfInfo) - 1 Step 1
                                    ilMatch = False
                                    ilRdf = tmRdfInfo(ilIndex).iRdfIndex
                                    For ilRow = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                        If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
                                            gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilRow), tgMRdf(ilRdf).iStartTime(1, ilRow), False, llTime
                                            If llTime = tmDrfInfo(ilCol).lStartTime Then
                                                gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilRow), tgMRdf(ilRdf).iEndTime(1, ilRow), False, llTime
                                                If llTime = tmDrfInfo(ilCol).lEndTime Then
                                                    ilMatch = True
                                                    For ilDay = 1 To 7 Step 1
                                                        If tgMRdf(ilRdf).sWkDays(ilRow, ilDay - 1) <> tmDrfInfo(ilCol).sDays(ilDay - 1) Then
                                                            ilMatch = False
                                                            Exit For
                                                        End If
                                                    Next ilDay
                                                    If ilMatch Then
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next ilRow
                                    If ilMatch And (tmDrfInfo(ilCol).lStartTime2 <> -1) Then
                                        ilMatch = False
                                        For ilRow = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                            If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
                                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilRow), tgMRdf(ilRdf).iStartTime(1, ilRow), False, llTime
                                                If llTime = tmDrfInfo(ilCol).lStartTime2 Then
                                                    gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilRow), tgMRdf(ilRdf).iEndTime(1, ilRow), False, llTime
                                                    If llTime = tmDrfInfo(ilCol).lEndTime2 Then
                                                        ilMatch = True
                                                        imMatchRdfCode(UBound(imMatchRdfCode)) = tgMRdf(ilRdf).iCode
                                                        ReDim Preserve imMatchRdfCode(0 To UBound(imMatchRdfCode) + 1) As Integer
                                                        Exit For
                                                    End If
                                                End If
                                                Exit For
                                            End If
                                        Next ilRow
                                    ElseIf ilMatch Then
                                        imMatchRdfCode(UBound(imMatchRdfCode)) = tgMRdf(ilRdf).iCode
                                        ReDim Preserve imMatchRdfCode(0 To UBound(imMatchRdfCode) + 1) As Integer
                                    End If
                                Next ilIndex
                                If UBound(imMatchRdfCode) > LBound(imMatchRdfCode) Then
                                    For ilMatch = LBound(imMatchRdfCode) To UBound(imMatchRdfCode) - 1 Step 1
                                        tmDrfInfo(ilCol).tDrf.lCode = 0
                                        tmDrfInfo(ilCol).tDrf.iRdfCode = imMatchRdfCode(ilMatch)
                                        '6/9/16: Replaced GoSub
                                        'GoSub mPCStnWriteRec_Mkt
                                        mPCStnWriteRec_Mkt_2 ilRet, ilCustomDemo, slVehicleName, slBookName, slBookDate, ilPrevDnfCode, llPrevDrfCode, ilDemoGender, llPopDemo(), blPopFound
                                        If ilRet <> 0 Then
                                            mConvertStationVehicles = False
                                            GoTo Cleanup
                                        End If
                                    Next ilMatch
                                Else
                                    tmDrfInfo(ilCol).tDrf.lCode = 0
                                    tmDrfInfo(ilCol).tDrf.iRdfCode = 0
                                    '6/9/16: Replaced GoSub
                                    'GoSub mPCStnWriteRec_Mkt
                                    mPCStnWriteRec_Mkt_2 ilRet, ilCustomDemo, slVehicleName, slBookName, slBookDate, ilPrevDnfCode, llPrevDrfCode, ilDemoGender, llPopDemo(), blPopFound
                                End If
                            End If
                            If ilRet <> 0 Then
                                mConvertStationVehicles = False
                                GoTo Cleanup
                            End If
                        Else
                            ilFound = False
                            For ilLoop = 0 To UBound(smUnfdStations) - 1 Step 1
                                If StrComp(Trim$(smUnfdStations(ilLoop)), Trim$(smFieldValues(ilStnCol)), 1) = 0 Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                            If Not ilFound Then
                                Print #hmTo, "Unable to find " & smFieldValues(ilStnCol)
                                lbcErrors.AddItem Trim$(smFieldValues(ilStnCol)) & " not found"
                                smUnfdStations(UBound(smUnfdStations)) = smFieldValues(ilStnCol)
                                ReDim Preserve smUnfdStations(0 To UBound(smUnfdStations) + 1) As String
                            End If
                        End If 'imVefCode > 0
                    End If 'tmDrf safe
                   ' End If 'data line?
                End If 'header or other?
            End If 'len(slline)
            slVehicleName = ""
            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
            llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If plcGauge.Value <> llPercent Then
                plcGauge.Value = llPercent
            End If
        End If
    Loop Until ilEof
    If blPopFound Then
        If Not mInsertPopulation(slBookName, ilPrevDnfCode, llPopDemo) Then
         '?
        End If
        mUpdatePopulationMissing slBookName
    End If
    mConvertStationVehicles = True
Cleanup:
    Close hmFrom
    plcGauge.Value = 100
    Erase tmExtraDemo
    Erase lmDpfWithoutPop
    MousePointer = vbDefault
    Exit Function
mConvertStationVehiclesErr:
    ilRet = err.Number
    Resume Next

'mPCStnWriteRec_Mkt:
'    If ilCustomDemo Then
'        Return
'    End If
'    ilSetBkNm = -1
'    ilRet = 0
'    If ((tmDrfInfo(ilCol).iBkNm = 0) And (Trim$(slBookName) <> "")) Then
'        tmDrfInfo(ilCol).tDrf.iVefCode = imVefCode
'        For ilDay = 0 To 6 Step 1
'            smDays(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
'        Next ilDay
'        ilBNRet = mBookNameUsed(slBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfcode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
'        If (ilBNRet = 0) Or (ilBNRet = 1) Then
'            tmDnf.iCode = 0
'            tmDnf.sBookName = slBookName
'            gPackDate slBookDate, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
'            gPackDate smNowDate, tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
'            tmDnf.iUrfCode = tgUrf(0).iCode
'            tmDnf.sType = "I"
'            tmDnf.sForm = smDataForm
'            tmDnf.sExactTime = "N"
'            tmDnf.sSource = "A"
'            tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
'            tmDnf.iAutoCode = tmDnf.iCode
'            ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
'            If ilRet <> BTRV_ERR_NONE Then
'                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                    ilRet = csiHandleValue(0, 7)
'                End If
'                Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
'                lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
'                Return
'            End If
'            Do
'                tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
'                tmDnf.iAutoCode = tmDnf.iCode
'                gPackDate smSyncDate, tmDnf.iSyncDate(0), tmDnf.iSyncDate(1)
'                gPackTime smSyncTime, tmDnf.iSyncTime(0), tmDnf.iSyncTime(1)
'                ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
'            Loop While ilRet = BTRV_ERR_CONFLICT
'            If ilRet <> BTRV_ERR_NONE Then
'                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                    ilRet = csiHandleValue(0, 7)
'                End If
'                Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
'                lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
'                Return
'            End If
'            ilPrevDnfCode = tmDnf.iCode
'            ilRet = mObtainBookName()
'        Else
'            If StrComp(smBookForm, smDataForm, vbTextCompare) <> 0 Then
'                If (smDataForm = "8") Or (smBookForm = "8") Then
'                    Print #hmTo, slBookName & " for " & slVehicleName & " previously defined with different format then current import data"
'                    lbcErrors.AddItem "Error in Book Forms" & " for " & slVehicleName
'                    Return
'                End If
'            End If
'            tmDrfSrchKey.iDnfCode = tmDnf.iCode
'            tmDrfSrchKey.sDemoDataType = "P"
'            tmDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
'            tmDrfSrchKey.iVefCode = 0
'            tmDrfSrchKey.sInfoType = ""
'            tmDrfSrchKey.iRdfcode = 0
'            ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'            If (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tmDnf.iCode) And (tlDrf.iVefCode = 0) And (tlDrf.sDemoDataType = "P") Then
'                tmDrfPop.lCode = tlDrf.lCode
'                Do
'                    tmDrfSrchKey2.lCode = tmDrfPop.lCode
'                    ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'                    If ilRet <> BTRV_ERR_NONE Then
'                        Exit Do
'                    End If
'                    For i= 0 To UBound(tmDrfInfo) - 1 Step 1
'                        ilDemoGender = tmDrfInfo(ilSet).iDemoGender
'                        tlDrf.lDemo(ilDemoGender) = llPopDemo(ilDemoGender)
'                    Next ilSet
'                    gPackDate smSyncDate, tlDrf.iSyncDate(0), tlDrf.iSyncDate(1)
'                    gPackTime smSyncTime, tlDrf.iSyncTime(0), tlDrf.iSyncTime(1)
'                    ilRet = btrUpdate(hmDrf, tlDrf, imDrfRecLen)
'                Loop While ilRet = BTRV_ERR_CONFLICT
'            Else
'                tmDrfPop.lCode = -1
'            End If
'        End If
'        tmDrfInfo(ilCol).tDrf.iDnfCode = ilPrevDnfCode
'        If llPrevDrfCode = 0 Then
'            tmDrfInfo(ilCol).tDrf.lCode = 0
'            tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
'            tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
'            ilRet = btrInsert(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen, INDEXKEY2)
'            If ilRet <> BTRV_ERR_NONE Then
'                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                    ilRet = csiHandleValue(0, 7)
'                End If
'                Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'                lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'                Return
'            Else
'                'to pass to extra categories later
'                llPrevDrfCode = tmDrfInfo(ilCol).tDrf.lCode
'            End If
'            Do
'                tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
'                tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
'                gPackDate smSyncDate, tmDrfInfo(ilCol).tDrf.iSyncDate(0), tmDrfInfo(ilCol).tDrf.iSyncDate(1)
'                gPackTime smSyncTime, tmDrfInfo(ilCol).tDrf.iSyncTime(0), tmDrfInfo(ilCol).tDrf.iSyncTime(1)
'                ilRet = btrUpdate(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen)
'            Loop While ilRet = BTRV_ERR_CONFLICT
'        Else
'            Do
'                tmDrfSrchKey2.lCode = llPrevDrfCode
'                ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'                If ilRet <> BTRV_ERR_NONE Then
'                    Exit Do
'                End If
'                'could be -1 for extra category
'                For i= 0 To UBound(tmDrfInfo) - 1 Step 1
'                    ilDemoGender = tmDrfInfo(ilSet).iDemoGender
'                    If ilDemoGender > 0 Then
'                        tlDrf.lDemo(ilDemoGender) = tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender)
'                    End If
'                Next ilSet
'                gPackDate smSyncDate, tlDrf.iSyncDate(0), tlDrf.iSyncDate(1)
'                gPackTime smSyncTime, tlDrf.iSyncTime(0), tlDrf.iSyncTime(1)
'                ilRet = btrUpdate(hmDrf, tlDrf, imDrfRecLen)
'            Loop While ilRet = BTRV_ERR_CONFLICT
'        End If
'        If ilRet <> BTRV_ERR_NONE Then
'            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
'                ilRet = csiHandleValue(0, 7)
'            End If
'            Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
'            lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
'            Return
'        End If
'        If tmDrfInfo(ilCol).iBkNm = 0 Then
'            ilSetBkNm = 0
'            ilSetDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
'        End If
'    End If
'    If ((ckcDefault(0).Value = vbChecked)) And (ilSetBkNm = 0) Then
'        Do
'            tmVefSrchKey.iCode = imVefCode
'            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'            If ilRet <> BTRV_ERR_NONE Then
'                Exit Do
'            End If
'            If ckcDefault(0).Value = vbChecked Then
'                tmVef.iDnfCode = ilSetDnfCode
'            End If
'            If ckcDefault(1).Value = vbChecked Then
'                tmVef.iReallDnfCode = ilSetDnfCode
'            End If
'            ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
'        Loop While ilRet = BTRV_ERR_CONFLICT
'        ilRet = gBinarySearchVef(tmVef.iCode)
'        If ilRet <> -1 Then
'            tgMVef(ilRet) = tmVef
'        End If
'    End If
'    mInsertOrUpdateExraCategory ilPrevDnfCode, llPrevDrfCode, blPopFound, slVehicleName
'    'keeps most values: just reset drf code to 0 and pop,demo values
'    mResetExtraCategories
'    ilRet = 0
'    Return
'lCreatePop_Mkt:
'    tmDrfPop.lCode = 0
'    tmDrfPop.iDnfCode = tmDnf.iCode
'    tmDrfPop.sDemoDataType = "P"
'    tmDrfPop.iMnfSocEco = 0
'    tmDrfPop.iVefCode = 0
'    tmDrfPop.sInfoType = ""
'    tmDrfPop.iRdfcode = 0
'    tmDrfPop.sProgCode = ""
'    tmDrfPop.iStartTime(0) = 1
'    tmDrfPop.iStartTime(1) = 0
'    tmDrfPop.iEndTime(0) = 1
'    tmDrfPop.iEndTime(1) = 0
'    tmDrfPop.iStartTime2(0) = 1
'    tmDrfPop.iStartTime2(1) = 0
'    tmDrfPop.iEndTime2(0) = 1
'    tmDrfPop.iEndTime2(1) = 0
'    For ilDay = 0 To 6 Step 1
'        tmDrfPop.sDay(ilDay) = "Y"
'    Next ilDay
'    tmDrfPop.iQHIndex = 0
'    tmDrfPop.iCount = 0
'    tmDrfPop.sExStdDP = "N"
'    tmDrfPop.sExRpt = "N"
'    tmDrfPop.sDataType = "A"
'    For i= 1 To 16 Step 1
'        tmDrfPop.lDemo(ilSet) = 0
'    Next ilSet
'    Return
End Function
Private Function mInsertPopulation(slBookName As String, ilDnfCode As Integer, llPopDemo() As Long) As Boolean

    Dim blRet As Boolean
    Dim ilSet As Integer
    Dim ilDemoGender As Integer
    Dim ilRet As Integer
    Dim ilDay As Integer
    
    blRet = True
On Error GoTo ERRBOX
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
    For ilSet = 1 To 16 Step 1
        tmDrfPop.lDemo(ilSet - 1) = 0
    Next ilSet
    tmDrfPop.sACTLineupCode = ""
    tmDrfPop.sACT1StoredTime = ""
    tmDrfPop.sACT1StoredSpots = ""
    tmDrfPop.sACT1StoreClearPct = ""
    tmDrfPop.sACT1DaypartFilter = ""
    
    tmDrfPop.lCode = 0
    tmDrfPop.iDnfCode = ilDnfCode
    For ilSet = 0 To UBound(tmDrfInfo) - 1 Step 1
        If tmDrfInfo(ilSet).iDemoGender >= 0 Then
            ilDemoGender = tmDrfInfo(ilSet).iDemoGender
            tmDrfPop.lDemo(ilDemoGender - 1) = llPopDemo(ilDemoGender)
        End If
    Next ilSet
    tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
    tmDrfPop.lAutoCode = tmDrfPop.lCode
    ilRet = btrInsert(hmDrf, tmDrfPop, imDrfRecLen, INDEXKEY2)
    If ilRet <> BTRV_ERR_NONE Then
        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
            ilRet = csiHandleValue(0, 7)
        End If
        Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slBookName
        lbcErrors.AddItem "Error Adding DRF" & " for " & slBookName
        blRet = False
    End If
    Do
        tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
        tmDrfPop.lAutoCode = tmDrfPop.lCode
        gPackDate smSyncDate, tmDrfPop.iSyncDate(0), tmDrfPop.iSyncDate(1)
        gPackTime smSyncTime, tmDrfPop.iSyncTime(0), tmDrfPop.iSyncTime(1)
        ilRet = btrUpdate(hmDrf, tmDrfPop, imDrfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
            ilRet = csiHandleValue(0, 7)
        End If
        Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slBookName
        lbcErrors.AddItem "Error Adding DRF" & " for " & slBookName
        blRet = False
    End If
    mInsertPopulation = blRet
    Exit Function
ERRBOX:
    mInsertPopulation = False
End Function
Private Function mParseForBookName(slFromName As String) As String
    Dim slRet As String
    Dim slLine As String
    Dim slChar As String
    Dim ilEof As Integer
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim ilSelected As Integer
    Dim ilLoop As Integer
    
    Screen.MousePointer = vbHourglass
    slRet = ""
    ilRet = 0
'On Error GoTo MPARSEERR:
    'hmFrom = FreeFile
    slFromName = mLoseLastLetterIf(slFromName, "|")
    If Len(slFromName) = 0 Then
        GoTo Cleanup
    End If
    'Open slFromName For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromName, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        MsgBox "Unable to find " & slFromName, vbExclamation, "Name Error"
        GoTo Cleanup
    End If
    Do
        ilRet = 0
        err.Clear
        slLine = ""
        Do
            If Not EOF(hmFrom) Then
                slChar = Input(1, #hmFrom)
                If slChar = Chr(13) Then 'Carriage return
                    slChar = Input(1, #hmFrom)
                End If
                If slChar = Chr(10) Then 'line feed
                    Exit Do
                End If
                slLine = slLine & slChar
            Else
                ilEof = True
                Exit Do
            End If
        Loop
        slLine = Trim$(slLine)
'On Error GoTo ERRBOX
        ilRet = err.Number
        If ilRet = 62 Then
            Exit Do
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                DoEvents
                gParseCDFields slLine, False, smFieldValues()
                '11/10/16: Move fields from 0 to X to 1 to X +1
                For ilLoop = UBound(smFieldValues) - 1 To LBound(smFieldValues) Step -1
                    smFieldValues(ilLoop + 1) = smFieldValues(ilLoop)
                Next ilLoop
                smFieldValues(0) = ""
                If Len(Trim$(smFieldValues(1))) > 0 Then
                    If InStr(1, smFieldValues(1), "Summer", 1) > 0 Or InStr(1, smFieldValues(1), "Fall", 1) > 0 Or InStr(1, smFieldValues(1), "Winter", 1) > 0 Or InStr(1, smFieldValues(1), "Spring", 1) > 0 Then
                        ilPos = InStr(smFieldValues(1), "[")
                        If ilPos > 0 Then
                            slRet = Mid(smFieldValues(1), 1, ilPos - 1)
                        Else
                            slRet = smFieldValues(1)
                        End If
                        Exit Do
                    End If
                End If
            End If
        End If
    Loop Until ilEof
Cleanup:
    Screen.MousePointer = vbDefault
    Close hmFrom
    mParseForBookName = slRet
    Exit Function
'MPARSEERR:
'    ilRet = Err.Number
'    Resume Next
ERRBOX:
    slRet = ""
    GoTo Cleanup
End Function
Private Function mLoseLastLetterIf(slInput As String, slRemoveThis As String) As String
    Dim llLength As Long
    Dim slNewString As String
    Dim llLastLetter As Long
    
    llLength = Len(slInput)
    llLastLetter = InStrRev(slInput, slRemoveThis)
    If llLength > 0 And llLastLetter = llLength Then
        slNewString = Mid(slInput, 1, llLength - 1)
    Else
        slNewString = slInput
    End If
    mLoseLastLetterIf = slNewString
End Function
Private Function mFindExraCategory(slSexChar As String, slDemoAge As String, ilCol As Integer) As Boolean
    'how to handle extra categories.  tmExtraDemo is an array meant to hold all extra categories for one line of information.
    'Here we fill out the mnf code, and the column # (so we can match demos later)  Dick said we are NOT to add a new category, only to update ones already existing
    Dim slDemoName As String
    Dim ilLoop As Integer
    Dim blRet As Boolean
    Dim ilUpper As Integer
    
On Error GoTo ERRBOX
    blRet = False
    If (slSexChar = "M") Or (slSexChar = "B") Then
        slDemoName = "M" & slDemoAge
    ElseIf (slSexChar = "W") Or (slSexChar = "G") Then
        slDemoName = "W" & slDemoAge
    ElseIf (slSexChar = "T") Or (slSexChar = "P") Then
        If InStr(1, slDemoAge, "12", 1) > 0 Then
            slDemoName = "P" & slDemoAge
        Else
            slDemoName = "A" & slDemoAge
        End If
    End If
    slDemoName = Trim$(slDemoName)
    For ilLoop = LBound(tgMnfSDemo) To UBound(tgMnfSDemo) Step 1
        If StrComp(Trim$(tgMnfSDemo(ilLoop).sName), slDemoName, 1) = 0 Then
            blRet = True
            Exit For
        End If
    Next ilLoop
    If blRet Then
        ilUpper = UBound(tmExtraDemo)
        tmExtraDemo(ilUpper).iCol = ilCol
        tmExtraDemo(ilUpper).tDpf.iMnfDemo = tgMnfSDemo(ilLoop).iCode
        tmExtraDemo(ilUpper).tDpf.iDnfCode = 0
        ReDim Preserve tmExtraDemo(ilUpper + 1)
    End If
    mFindExraCategory = blRet
    Exit Function
ERRBOX:
    mFindExraCategory = False
End Function
Private Function mExtraDemoOrPop(ilCurrentIndex As Integer, blIsDemo As Boolean) As Boolean
    'still filling out our extra category array, which is by line.  This is called for each demo that may have an extra category. Use the column to find the right category.
    Dim ilLoop As Integer
    Dim blRet As Boolean
    Dim slDemoName As String
    Dim llValue As Long
    
On Error GoTo ERRBOX
    blRet = False
    For ilLoop = 0 To UBound(tmExtraDemo) - 1
        If tmExtraDemo(ilLoop).iCol = ilCurrentIndex Then
            blRet = True
            If tgSpf.sSAudData = "H" Then
                llValue = (CLng(smFieldValues(ilCurrentIndex)) + 50) \ 100
            ElseIf tgSpf.sSAudData = "N" Then
                llValue = (CLng(smFieldValues(ilCurrentIndex)) + 5) \ 10
            ElseIf tgSpf.sSAudData = "U" Then
                llValue = CLng(smFieldValues(ilCurrentIndex))
            Else
                llValue = (CLng(smFieldValues(ilCurrentIndex)) + 500) \ 1000
            End If
            If blIsDemo Then
                tmExtraDemo(ilLoop).tDpf.lDemo = llValue
            Else
                tmExtraDemo(ilLoop).tDpf.lPop = llValue
            End If
            Exit For
        End If
    Next ilLoop
    mExtraDemoOrPop = blRet
    Exit Function
ERRBOX:
    mExtraDemoOrPop = False
End Function

Private Function mInsertOrUpdateExraCategory(ilDnfCode As Integer, llDrfCode As Long, blPop As Boolean, slVehicleName As String) As Boolean
    'Now all we are missing for extra categories is the drfCode, the dnfCode, and the population. Fill in (population if we have it yet--won't have for the first)
    'search if the dpf exists by dnf,drf, and mnf.
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim llDpf As Long
    Dim ilUpper As Integer
    
    For ilIndex = 0 To UBound(tmExtraDemo) - 1 Step 1
        tmExtraDemo(ilIndex).tDpf.iDnfCode = ilDnfCode
        tmExtraDemo(ilIndex).tDpf.lDrfCode = llDrfCode
        llDpf = mGetDPF(ilDnfCode, llDrfCode, tmExtraDemo(ilIndex).tDpf.iMnfDemo)
        If llDpf = 0 Then
            tmExtraDemo(ilIndex).tDpf.lCode = 0
            ilRet = btrInsert(hmDpf, tmExtraDemo(ilIndex).tDpf, imDpfRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmTo, "Warning: Error when Adding Demo Plus Data File (DPF)" & Str$(ilRet) & " for " & slVehicleName
                lbcErrors.AddItem "Error Adding DPF" & " for " & slVehicleName
            End If
        Else
            tmExtraDemo(ilIndex).tDpf.lCode = llDpf
            ilRet = btrUpdate(hmDpf, tmExtraDemo(ilIndex).tDpf, imDpfRecLen)
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmTo, "Warning: Error when Updating Demo Plus Data File (DPF)" & Str$(ilRet) & " for " & slVehicleName
                lbcErrors.AddItem "Error Updating DPF" & " for " & slVehicleName
            End If
        End If
        'if population hasn't been done yet, save for later
        If Not blPop Then
            ilUpper = UBound(lmDpfWithoutPop)
            lmDpfWithoutPop(ilUpper).iColumn = tmExtraDemo(ilIndex).iCol
            lmDpfWithoutPop(ilUpper).lDpfCode = tmExtraDemo(ilIndex).tDpf.lCode
            ReDim Preserve lmDpfWithoutPop(ilUpper + 1)
        End If
    Next ilIndex

End Function
Private Function mGetDPF(ilDnfCode As Integer, llDrfCode As Long, ilMnfCode As Integer) As Long
    Dim ilRet As Integer
    Dim llRet As Long
    
    llRet = 0
    tmDpfSrchKey1.lDrfCode = llDrfCode
    tmDpfSrchKey1.iMnfDemo = ilMnfCode
    ilRet = btrGetEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE And tmDpf.iMnfDemo = ilMnfCode And tmDpf.lDrfCode = llDrfCode
        If tmDpf.iDnfCode = ilDnfCode Then
            llRet = tmDpf.lCode
            Exit Do
        End If
        ilRet = btrGetNext(hmDpf, tmDpf, imDpfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    mGetDPF = llRet
End Function
Private Sub mResetExtraCategories()
    Dim ilIndex As Integer
    
    For ilIndex = 0 To UBound(tmExtraDemo) - 1
        tmExtraDemo(ilIndex).tDpf.lDrfCode = 0
        tmExtraDemo(ilIndex).tDpf.lDemo = 0
    Next ilIndex
End Sub
Private Function mUpdatePopulationMissing(slBookName As String) As Boolean
    Dim blRet As Boolean
    Dim ilRet As Integer
    Dim llDpf As Long
    Dim ilIndex As Integer
    Dim ilColumn As Integer
    Dim c As Integer
    
    blRet = True
    For ilIndex = 0 To UBound(lmDpfWithoutPop) - 1
        ilColumn = lmDpfWithoutPop(ilIndex).iColumn
        For c = 0 To UBound(tmExtraDemo) - 1
            If ilColumn = tmExtraDemo(c).iCol Then
                llDpf = lmDpfWithoutPop(ilIndex).lDpfCode
                ilRet = btrGetEqual(hmDpf, tmDpf, imDpfRecLen, llDpf, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_NONE Then
                    tmDpf.lPop = tmExtraDemo(c).tDpf.lPop
                    ilRet = btrUpdate(hmDpf, tmDpf, imDpfRecLen)
                    If ilRet <> BTRV_ERR_NONE Then
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        Print #hmTo, "Error when Adding population to Demo Plus File (DPF)" & Str$(ilRet) & " for " & slBookName
                        lbcErrors.AddItem "Error Adding population to DPF " & " for " & slBookName
                        blRet = False
                    End If
                End If
            End If
        Next c
    Next ilIndex
    mUpdatePopulationMissing = blRet
    Exit Function
ERRBOX:
    mUpdatePopulationMissing = False
End Function

Private Sub mWriteRec(ilRet As Integer, ilCustomDemo As Integer, slVehicleName As String, slDPBookName As String, slETBookName As String, slBookDate As String, ilPrevDnfCode As Integer, llPrevDrfCode As Long)
    Dim ilVef As Integer
    Dim ilCol As Integer
    Dim ilIndex As Integer
    Dim ilSetBkNm As Integer
    Dim ilSetDnfCode As Integer
    Dim ilRdf As Integer
    Dim ilRow As Integer
    Dim llTime As Long
    Dim ilMatch As Integer
    Dim ilDay As Integer
    Dim ilBNRet As Integer
    Dim ilQRet As Integer
    Dim tlDrf As DRF
    
    If ilCustomDemo Then
        'Return
        Exit Sub
        'tmDrfPop.sDataType = "B"
        'tmDrf.sDataType = "B"
    End If
    ilRet = 0
    For ilVef = LBound(imVefCodeImpt) To UBound(imVefCodeImpt) - 1 Step 1
        imVefCode = imVefCodeImpt(ilVef)
        'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If tgMVef(ilIndex).iCode = imVefCode Then
            ilIndex = gBinarySearchVef(imVefCode)
            If ilIndex <> -1 Then
                slVehicleName = Trim$(tgMVef(ilIndex).sName)
        '        Exit For
            End If
        'Next ilIndex
        ilSetBkNm = -1
        For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
            If ((tmDrfInfo(ilCol).iBkNm = 0) And (Trim$(slDPBookName) <> "")) Or ((tmDrfInfo(ilCol).iBkNm = 1) And (Trim$(slETBookName) <> "")) Then

                tmDrfInfo(ilCol).tDrf.iVefCode = imVefCode
                If tmDrfInfo(ilCol).tDrf.sInfoType = "D" Then
                    mDPPop
                    For ilIndex = LBound(tmRdfInfo) To UBound(tmRdfInfo) - 1 Step 1
                        ilRdf = tmRdfInfo(ilIndex).iRdfIndex
                        For ilRow = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                            If (tgMRdf(ilRdf).iStartTime(0, ilRow) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilRow) <> 0) Then
                                gUnpackTimeLong tgMRdf(ilRdf).iStartTime(0, ilRow), tgMRdf(ilRdf).iStartTime(1, ilRow), False, llTime
                                If llTime = tmDrfInfo(ilCol).lStartTime Then
                                    gUnpackTimeLong tgMRdf(ilRdf).iEndTime(0, ilRow), tgMRdf(ilRdf).iEndTime(1, ilRow), False, llTime
                                    If llTime = tmDrfInfo(ilCol).lEndTime Then
                                        ilMatch = True
                                        For ilDay = 1 To 7 Step 1
                                            If tgMRdf(ilRdf).sWkDays(ilRow, ilDay - 1) <> tmDrfInfo(ilCol).sDays(ilDay - 1) Then
                                                ilMatch = False
                                                Exit For
                                            End If
                                        Next ilDay
                                        If ilMatch Then
                                            tmDrfInfo(ilCol).tDrf.iRdfCode = tgMRdf(ilRdf).iCode
                                        End If
                                        Exit For
                                    End If
                                End If
                                Exit For
                            End If
                        Next ilRow
                    Next ilIndex
                End If
                For ilDay = 0 To 6 Step 1
                    smDays(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
                Next ilDay
                If tmDrfInfo(ilCol).iBkNm = 0 Then
                    ilBNRet = mBookNameUsed(slDPBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfCode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
                Else
                    ilBNRet = mBookNameUsed(slETBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfCode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
                End If
                If (ilBNRet = 0) Or (ilBNRet = 1) Then
                    tmDnf.iCode = 0
                    If tmDrfInfo(ilCol).iBkNm = 0 Then
                        tmDnf.sBookName = slDPBookName
                    Else
                        tmDnf.sBookName = slETBookName
                    End If
                    gPackDate slBookDate, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
                    gPackDate smNowDate, tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
                    tmDnf.iUrfCode = tgUrf(0).iCode
                    tmDnf.sType = "I"
                    tmDnf.sForm = smDataForm
                    If tmDrfInfo(ilCol).iBkNm = 0 Then
                        tmDnf.sExactTime = "N"
                    Else
                        tmDnf.sExactTime = "Y"
                    End If
                    tmDnf.sSource = "A"
                    tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmDnf.iAutoCode = tmDnf.iCode
                    ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
                    If ilRet <> BTRV_ERR_NONE Then
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
                        lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
                        'mOnLineConvFile = False
                        'Exit Function
                        'Return
                        Exit Sub
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
                        If ilRet <> BTRV_ERR_NONE Then
                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
                            lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
                            'mOnLineConvFile = False
                            'Exit Function
                            'Return
                            Exit Sub
                        End If
                    'End If
                    ilPrevDnfCode = tmDnf.iCode
                    ilRet = mObtainBookName()
                    tmDrfPop.lCode = 0
                    tmDrfPop.iDnfCode = ilPrevDnfCode
                    tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
                    tmDrfPop.lAutoCode = tmDrfPop.lCode
                    ilRet = btrInsert(hmDrf, tmDrfPop, imDrfRecLen, INDEXKEY2)
                    If ilRet <> BTRV_ERR_NONE Then
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slVehicleName
                        lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
                        'mOnLineConvFile = False
                        'Exit Function
                        'Return
                        Exit Sub
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
                        If ilRet <> BTRV_ERR_NONE Then
                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slVehicleName
                            lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
                            'mOnLineConvFile = False
                            'Exit Function
                            'Return
                            Exit Sub
                        End If
                    'End If
                Else
                    If StrComp(smBookForm, smDataForm, vbTextCompare) <> 0 Then
                        If (smDataForm = "8") Or (smBookForm = "8") Then
                            If tmDrfInfo(ilCol).iBkNm = 0 Then
                                Print #hmTo, slDPBookName & " for " & slVehicleName & " previously defined with different format then current import data"
                            Else
                                Print #hmTo, slETBookName & " for " & slVehicleName & " previously defined with different format then current import data"
                            End If
                            lbcErrors.AddItem "Error in Book Forms" & " for " & slVehicleName
                            ilRet = 0
                            'Return
                            Exit Sub
                        End If
                    End If
                    If ilBNRet = 3 Then
                        Screen.MousePointer = vbDefault
                        If tmDrfInfo(ilCol).iBkNm = 0 Then
                            ilQRet = MsgBox(slDPBookName & " for " & slVehicleName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
                        Else
                            ilQRet = MsgBox(slETBookName & " for " & slVehicleName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
                        End If
                        If ilQRet = vbNo Then
                            'mOnLineConvFile = False
                            'Exit Function
                            ilRet = 0
                            'Return
                            Exit Sub
                        End If
                        Print #hmTo, "Replaced Demo Data (DRF)"
                    End If
                End If
                tmDrfInfo(ilCol).tDrf.iDnfCode = ilPrevDnfCode
                If llPrevDrfCode = 0 Then
                    tmDrfInfo(ilCol).tDrf.lCode = 0
                    tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
                    tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
                    ilRet = btrInsert(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen, INDEXKEY2)
                    If ilRet <> BTRV_ERR_NONE Then
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
                        lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
                        'mOnLineConvFile = False
                        'Exit Function
                        'Return
                        Exit Sub
                    End If
                    'If tgSpf.sRemoteUsers = "Y" Then
                        Do
                            'tmDrfSrchKey2.lCode = tmDrf.lCode
                            'ilRet = btrGetEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                            tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
                            tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
                            gPackDate smSyncDate, tmDrfInfo(ilCol).tDrf.iSyncDate(0), tmDrfInfo(ilCol).tDrf.iSyncDate(1)
                            gPackTime smSyncTime, tmDrfInfo(ilCol).tDrf.iSyncTime(0), tmDrfInfo(ilCol).tDrf.iSyncTime(1)
                            ilRet = btrUpdate(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                     'End If
                Else
                    Do
                        tmDrfSrchKey2.lCode = llPrevDrfCode
                        ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                        If ilRet <> BTRV_ERR_NONE Then
                            Exit Do
                        End If
                        tmDrfInfo(ilCol).tDrf.lCode = tlDrf.lCode
                        tmDrfInfo(ilCol).tDrf.iRemoteID = tlDrf.iRemoteID
                        tmDrfInfo(ilCol).tDrf.lAutoCode = tlDrf.lAutoCode
                        gPackDate smSyncDate, tmDrfInfo(ilCol).tDrf.iSyncDate(0), tmDrfInfo(ilCol).tDrf.iSyncDate(1)
                        gPackTime smSyncTime, tmDrfInfo(ilCol).tDrf.iSyncTime(0), tmDrfInfo(ilCol).tDrf.iSyncTime(1)
                        ilRet = btrUpdate(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                End If
                If ilRet <> BTRV_ERR_NONE Then
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
                    lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
                    'mOnLineConvFile = False
                    'Exit Function
                    'Return
                    Exit Sub
                End If
                If tmDrfInfo(ilCol).iBkNm = 0 Then
                    ilSetBkNm = 0
                    ilSetDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
                ElseIf (tmDrfInfo(ilCol).iBkNm = 1) And (ilSetBkNm = -1) Then
                    ilSetBkNm = 1
                    ilSetDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
                End If
            End If
        Next ilCol
        If ((ckcDefault(0).Value) Or (ckcDefault(1).Value)) And ((ilSetBkNm = 0) Or (ilSetBkNm = 1)) Then
            Do
                tmVefSrchKey.iCode = imVefCode
                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If ckcDefault(0).Value = vbChecked Then
                    tmVef.iDnfCode = ilSetDnfCode
                End If
                If ckcDefault(1).Value = vbChecked Then
                    tmVef.iReallDnfCode = ilSetDnfCode
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
        End If
        Print #hmTo, "Successfully Installed Demo Data" & " for " & Trim$(slVehicleName)
        bmBooksSaved = True
    Next ilVef

    For ilVef = LBound(imVefCodeImpt) To UBound(imVefCodeImpt) - 1 Step 1
        'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        '    If tgMVef(ilIndex).iCode = imVefCodeImpt(ilVef) Then
            ilIndex = gBinarySearchVef(imVefCodeImpt(ilVef))
            If ilIndex <> -1 Then
                Print #hmTo, "Successfully Installed Demo Data" & " for " & Trim$(tgMVef(ilIndex).sName)
        '        Exit For
            End If
        'Next ilIndex
    Next ilVef
    ilRet = 0
End Sub

Private Sub mPCStnWriteRec_Mkt_1(ilRet As Integer, ilCustomDemo As Integer, slVehicleName As String, slBookName As String, slBookDate As String, ilPrevDnfCode As Integer, llPrevDrfCode As Long, ilDemoGender As Integer, llPopDemo() As Long)
    Dim ilSetBkNm As Integer
    Dim ilCol As Integer
    Dim ilDay As Integer
    Dim ilBNRet As Integer
    Dim ilSet As Integer
    Dim ilSetDnfCode As Integer
    Dim tlDrf As DRF
    
    If ilCustomDemo Then
        'Return
        Exit Sub
        'tmDrfPop.sDataType = "B"
        'tmDrf.sDataType = "B"
    End If
    ilSetBkNm = -1
    ilRet = 0
    If ((tmDrfInfo(ilCol).iBkNm = 0) And (Trim$(slBookName) <> "")) Then
        tmDrfInfo(ilCol).tDrf.iVefCode = imVefCode
        For ilDay = 0 To 6 Step 1
            smDays(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
        Next ilDay
        'If tmDrfInfo(ilCol).iBkNm = 0 Then
            ilBNRet = mBookNameUsed(slBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfCode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
        'Else
        '    ilBNRet = mBookNameUsed(slETBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfCode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
        'End If
        If (ilBNRet = 0) Or (ilBNRet = 1) Then
            tmDnf.iCode = 0
            tmDnf.sBookName = slBookName
            gPackDate slBookDate, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
            gPackDate smNowDate, tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
            tmDnf.iUrfCode = tgUrf(0).iCode
            tmDnf.sType = "I"
            tmDnf.sForm = smDataForm
            'If tmDrfInfo(ilCol).iBkNm = 0 Then
                tmDnf.sExactTime = "N"
            'Else
            '    tmDnf.sExactTime = "Y"
            'End If
            tmDnf.sSource = "A"
            tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmDnf.iAutoCode = tmDnf.iCode
            ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
                lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
                'Return
                Exit Sub
            End If
            Do
                tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
                tmDnf.iAutoCode = tmDnf.iCode
                gPackDate smSyncDate, tmDnf.iSyncDate(0), tmDnf.iSyncDate(1)
                gPackTime smSyncTime, tmDnf.iSyncTime(0), tmDnf.iSyncTime(1)
                ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
                lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
                'Return
                Exit Sub
            End If
            ilPrevDnfCode = tmDnf.iCode
            ilRet = mObtainBookName()
            '6/9/16: Replace GoSub
            'GoSub lCreatePop_Mkt
            mCreatePop_Mkt
            tmDrfPop.lCode = 0
            tmDrfPop.iDnfCode = ilPrevDnfCode
            For ilSet = 0 To UBound(tmDrfInfo) - 1 Step 1
                ilDemoGender = tmDrfInfo(ilSet).iDemoGender
                tmDrfPop.lDemo(ilDemoGender - 1) = llPopDemo(ilDemoGender)
            Next ilSet
            tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
            tmDrfPop.lAutoCode = tmDrfPop.lCode
            ilRet = btrInsert(hmDrf, tmDrfPop, imDrfRecLen, INDEXKEY2)
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slVehicleName
                lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
                'Return
                Exit Sub
            End If
            Do
                tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
                tmDrfPop.lAutoCode = tmDrfPop.lCode
                gPackDate smSyncDate, tmDrfPop.iSyncDate(0), tmDrfPop.iSyncDate(1)
                gPackTime smSyncTime, tmDrfPop.iSyncTime(0), tmDrfPop.iSyncTime(1)
                ilRet = btrUpdate(hmDrf, tmDrfPop, imDrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet) & " for " & slVehicleName
                lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
                'Return
                Exit Sub
            End If
        Else
            If StrComp(smBookForm, smDataForm, vbTextCompare) <> 0 Then
                If (smDataForm = "8") Or (smBookForm = "8") Then
                    Print #hmTo, slBookName & " for " & slVehicleName & " previously defined with different format then current import data"
                    lbcErrors.AddItem "Error in Book Forms" & " for " & slVehicleName
                    'Return
                    Exit Sub
                End If
            End If
            'If tmDrfPop.lCode = 0 Then
                tmDrfSrchKey.iDnfCode = tmDnf.iCode
                tmDrfSrchKey.sDemoDataType = "P"
                tmDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
                tmDrfSrchKey.iVefCode = 0
                tmDrfSrchKey.sInfoType = ""
                tmDrfSrchKey.iRdfCode = 0
                ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tmDnf.iCode) And (tlDrf.iVefCode = 0) And (tlDrf.sDemoDataType = "P") Then
                    tmDrfPop.lCode = tlDrf.lCode
                    Do
                        tmDrfSrchKey2.lCode = tmDrfPop.lCode
                        ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                        If ilRet <> BTRV_ERR_NONE Then
                            Exit Do
                        End If
                        'tlDrf.lDemo(ilDemoGender) = llPopDemo
                        For ilSet = 0 To UBound(tmDrfInfo) - 1 Step 1
                            ilDemoGender = tmDrfInfo(ilSet).iDemoGender
                            tlDrf.lDemo(ilDemoGender - 1) = llPopDemo(ilDemoGender)
                        Next ilSet
                        gPackDate smSyncDate, tlDrf.iSyncDate(0), tlDrf.iSyncDate(1)
                        gPackTime smSyncTime, tlDrf.iSyncTime(0), tlDrf.iSyncTime(1)
                        ilRet = btrUpdate(hmDrf, tlDrf, imDrfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                Else
                    tmDrfPop.lCode = -1
                End If
            'End If
        End If
        tmDrfInfo(ilCol).tDrf.iDnfCode = ilPrevDnfCode
        If llPrevDrfCode = 0 Then
            tmDrfInfo(ilCol).tDrf.lCode = 0
            tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
            tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
            ilRet = btrInsert(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen, INDEXKEY2)
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
                lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
                'Return
                Exit Sub
            End If
            Do
                tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
                gPackDate smSyncDate, tmDrfInfo(ilCol).tDrf.iSyncDate(0), tmDrfInfo(ilCol).tDrf.iSyncDate(1)
                gPackTime smSyncTime, tmDrfInfo(ilCol).tDrf.iSyncTime(0), tmDrfInfo(ilCol).tDrf.iSyncTime(1)
                ilRet = btrUpdate(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        Else
            Do
                tmDrfSrchKey2.lCode = llPrevDrfCode
                ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                'tlDrf.lDemo(ilDemoGender) = tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender)
                For ilSet = 0 To UBound(tmDrfInfo) - 1 Step 1
                    ilDemoGender = tmDrfInfo(ilSet).iDemoGender
                    tlDrf.lDemo(ilDemoGender - 1) = tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1)
                Next ilSet
                gPackDate smSyncDate, tlDrf.iSyncDate(0), tlDrf.iSyncDate(1)
                gPackTime smSyncTime, tlDrf.iSyncTime(0), tlDrf.iSyncTime(1)
                ilRet = btrUpdate(hmDrf, tlDrf, imDrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        End If
        If ilRet <> BTRV_ERR_NONE Then
            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                ilRet = csiHandleValue(0, 7)
            End If
            Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
            lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
            ''mSvPCStnConvFile = False
            ''Exit Function
            'Return
            Exit Sub
        End If
        If tmDrfInfo(ilCol).iBkNm = 0 Then
            ilSetBkNm = 0
            ilSetDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
        End If
        bmBooksSaved = True
    End If
    If ((ckcDefault(0).Value = vbChecked) Or (ckcDefault(1).Value = vbChecked)) And (ilSetBkNm = 0) Then
        Do
            tmVefSrchKey.iCode = imVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            If ckcDefault(0).Value = vbChecked Then
                tmVef.iDnfCode = ilSetDnfCode
            End If
            If ckcDefault(1).Value = vbChecked Then
                tmVef.iReallDnfCode = ilSetDnfCode
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
    End If
    ilRet = 0
End Sub

Private Sub mCreatePop_Mkt()
    Dim ilSet As Integer
    Dim ilDay As Integer
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
    For ilSet = 1 To 16 Step 1
        tmDrfPop.lDemo(ilSet - 1) = 0
    Next ilSet
    tmDrfPop.sACTLineupCode = ""
    tmDrfPop.sACT1StoredTime = ""
    tmDrfPop.sACT1StoredSpots = ""
    tmDrfPop.sACT1StoreClearPct = ""
    tmDrfPop.sACT1DaypartFilter = ""
End Sub

Private Sub mPCStnWriteRec_Mkt_2(ilRet As Integer, ilCustomDemo As Integer, slVehicleName As String, slBookName As String, slBookDate As String, ilPrevDnfCode As Integer, llPrevDrfCode As Long, ilDemoGender As Integer, llPopDemo() As Long, blPopFound As Boolean)
    Dim ilSetBkNm As Integer
    Dim ilCol As Integer
    Dim ilDay As Integer
    Dim ilBNRet As Integer
    Dim ilSet As Integer
    Dim ilSetDnfCode As Integer
    Dim tlDrf As DRF
    
    If ilCustomDemo Then
        'Return
        Exit Sub
    End If
    ilSetBkNm = -1
    ilRet = 0
    If ((tmDrfInfo(ilCol).iBkNm = 0) And (Trim$(slBookName) <> "")) Then
        tmDrfInfo(ilCol).tDrf.iVefCode = imVefCode
        For ilDay = 0 To 6 Step 1
            smDays(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
        Next ilDay
        ilBNRet = mBookNameUsed(slBookName, slBookDate, imVefCode, smDays(), tmDrfInfo(ilCol).tDrf.sInfoType, tmDrfInfo(ilCol).tDrf.iRdfCode, tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).lEndTime, ilCustomDemo, ilPrevDnfCode, llPrevDrfCode)
        If (ilBNRet = 0) Or (ilBNRet = 1) Then
            tmDnf.iCode = 0
            tmDnf.sBookName = slBookName
            gPackDate slBookDate, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
            gPackDate smNowDate, tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
            tmDnf.iUrfCode = tgUrf(0).iCode
            tmDnf.sType = "I"
            tmDnf.sForm = smDataForm
            tmDnf.sExactTime = "N"
            tmDnf.sSource = "A"
            tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmDnf.iAutoCode = tmDnf.iCode
            ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
                lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
                'Return
                Exit Sub
            End If
            Do
                tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
                tmDnf.iAutoCode = tmDnf.iCode
                gPackDate smSyncDate, tmDnf.iSyncDate(0), tmDnf.iSyncDate(1)
                gPackTime smSyncTime, tmDnf.iSyncTime(0), tmDnf.iSyncTime(1)
                ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet) & " for " & tmDnf.sBookName
                lbcErrors.AddItem "Error Adding DNF" & " for " & tmDnf.sBookName
                'Return
                Exit Sub
            End If
            ilPrevDnfCode = tmDnf.iCode
            ilRet = mObtainBookName()
        Else
            If StrComp(smBookForm, smDataForm, vbTextCompare) <> 0 Then
                If (smDataForm = "8") Or (smBookForm = "8") Then
                    Print #hmTo, slBookName & " for " & slVehicleName & " previously defined with different format then current import data"
                    lbcErrors.AddItem "Error in Book Forms" & " for " & slVehicleName
                    'Return
                    Exit Sub
                End If
            End If
            tmDrfSrchKey.iDnfCode = tmDnf.iCode
            tmDrfSrchKey.sDemoDataType = "P"
            tmDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
            tmDrfSrchKey.iVefCode = 0
            tmDrfSrchKey.sInfoType = ""
            tmDrfSrchKey.iRdfCode = 0
            ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tmDnf.iCode) And (tlDrf.iVefCode = 0) And (tlDrf.sDemoDataType = "P") Then
                tmDrfPop.lCode = tlDrf.lCode
                Do
                    tmDrfSrchKey2.lCode = tmDrfPop.lCode
                    ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    For ilSet = 0 To UBound(tmDrfInfo) - 1 Step 1
                        ilDemoGender = tmDrfInfo(ilSet).iDemoGender
                        tlDrf.lDemo(ilDemoGender - 1) = llPopDemo(ilDemoGender)
                    Next ilSet
                    gPackDate smSyncDate, tlDrf.iSyncDate(0), tlDrf.iSyncDate(1)
                    gPackTime smSyncTime, tlDrf.iSyncTime(0), tlDrf.iSyncTime(1)
                    ilRet = btrUpdate(hmDrf, tlDrf, imDrfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
            Else
                tmDrfPop.lCode = -1
            End If
        End If
        tmDrfInfo(ilCol).tDrf.iDnfCode = ilPrevDnfCode
        If llPrevDrfCode = 0 Then
            tmDrfInfo(ilCol).tDrf.lCode = 0
            tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
            tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
            ilRet = btrInsert(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen, INDEXKEY2)
            If ilRet <> BTRV_ERR_NONE Then
                If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                    ilRet = csiHandleValue(0, 7)
                End If
                Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
                lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
                'Return
                Exit Sub
            Else
                'to pass to extra categories later
                llPrevDrfCode = tmDrfInfo(ilCol).tDrf.lCode
            End If
            Do
                tmDrfInfo(ilCol).tDrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmDrfInfo(ilCol).tDrf.lAutoCode = tmDrfInfo(ilCol).tDrf.lCode
                gPackDate smSyncDate, tmDrfInfo(ilCol).tDrf.iSyncDate(0), tmDrfInfo(ilCol).tDrf.iSyncDate(1)
                gPackTime smSyncTime, tmDrfInfo(ilCol).tDrf.iSyncTime(0), tmDrfInfo(ilCol).tDrf.iSyncTime(1)
                ilRet = btrUpdate(hmDrf, tmDrfInfo(ilCol).tDrf, imDrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        Else
            Do
                tmDrfSrchKey2.lCode = llPrevDrfCode
                ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                'could be -1 for extra category
                For ilSet = 0 To UBound(tmDrfInfo) - 1 Step 1
                    ilDemoGender = tmDrfInfo(ilSet).iDemoGender
                    If ilDemoGender > 0 Then
                        tlDrf.lDemo(ilDemoGender - 1) = tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1)
                    End If
                Next ilSet
                gPackDate smSyncDate, tlDrf.iSyncDate(0), tlDrf.iSyncDate(1)
                gPackTime smSyncTime, tlDrf.iSyncTime(0), tlDrf.iSyncTime(1)
                ilRet = btrUpdate(hmDrf, tlDrf, imDrfRecLen)
            Loop While ilRet = BTRV_ERR_CONFLICT
        End If
        If ilRet <> BTRV_ERR_NONE Then
            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                ilRet = csiHandleValue(0, 7)
            End If
            Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet) & " for " & slVehicleName
            lbcErrors.AddItem "Error Adding DRF" & " for " & slVehicleName
            'Return
            Exit Sub
        End If
        If tmDrfInfo(ilCol).iBkNm = 0 Then
            ilSetBkNm = 0
            ilSetDnfCode = tmDrfInfo(ilCol).tDrf.iDnfCode
        End If
        bmBooksSaved = True
    End If
    If ((ckcDefault(0).Value = vbChecked)) And (ilSetBkNm = 0) Then
        Do
            tmVefSrchKey.iCode = imVefCode
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            If ckcDefault(0).Value = vbChecked Then
                tmVef.iDnfCode = ilSetDnfCode
            End If
            If ckcDefault(1).Value = vbChecked Then
                tmVef.iReallDnfCode = ilSetDnfCode
            End If
            ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        ilRet = gBinarySearchVef(tmVef.iCode)
        If ilRet <> -1 Then
            tgMVef(ilRet) = tmVef
        End If
    End If
    mInsertOrUpdateExraCategory ilPrevDnfCode, llPrevDrfCode, blPopFound, slVehicleName
    'keeps most values: just reset drf code to 0 and pop,demo values
    mResetExtraCategories
    ilRet = 0
End Sub

Private Function mCreateSort(tlDnf As DNF) As String
    Dim llDate As Long
    Dim slSortDate As String
    Dim slDate As String
    Dim slName As String
    
    gUnpackDateLong tlDnf.iBookDate(0), tlDnf.iBookDate(1), llDate
    llDate = 99999 - llDate
    slSortDate = Trim$(Str$(llDate))
    Do While Len(slSortDate) < 5
        slSortDate = "0" & slSortDate
    Loop
    gUnpackDate tlDnf.iBookDate(0), tlDnf.iBookDate(1), slDate
    slName = Trim$(tlDnf.sBookName) & ": " & slDate
    slName = slSortDate & "|" & slName
    mCreateSort = slName
End Function

Private Sub mSetCommands(Index As Integer)
    Dim ilRet As Integer
       
    If imBNSelectedIndex >= 0 Then
        tmDnfSrchKey.iCode = cbcBookName(Index).ItemData(imBNSelectedIndex)
        ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            If (Trim$(tmDnf.sSource) = "") Or (Trim$(tmDnf.sSource) = "M") Then
                If Index = 0 Then
                    cmcMoveToExact.Enabled = True
                Else
                    cmcMoveToDaypart.Enabled = True
                End If
            Else
                If Index = 0 Then
                    cmcMoveToExact.Enabled = False
                Else
                    cmcMoveToDaypart.Enabled = False
                End If
            End If
        Else
            If Index = 0 Then
                cmcMoveToExact.Enabled = False
            Else
                cmcMoveToDaypart.Enabled = False
            End If
        End If
    Else
        If Index = 0 Then
            cmcMoveToExact.Enabled = False
        Else
            cmcMoveToDaypart.Enabled = False
        End If
    End If
End Sub

Private Sub mGetDayIndex(slInDayNames As String, ilStartDayIndex As Integer, ilEndDayIndex As Integer)
    Dim slDayNames As String
    Dim ilPass As Integer
    Dim ilDay As Integer
    Dim ilPos As Integer
    Dim ilChar As Integer
    ReDim slNames(0 To 6) As String

    slNames(0) = "MON"
    slNames(1) = "TUE"
    slNames(2) = "WED"
    slNames(3) = "THU"
    slNames(4) = "FRI"
    slNames(5) = "SAT"
    slNames(6) = "SUN"
    ilStartDayIndex = -1
    ilEndDayIndex = -1
    slDayNames = Trim$(slInDayNames)
    If slDayNames = "SS" Then
        ilStartDayIndex = 5
        ilEndDayIndex = 6
        Exit Sub
    End If
    For ilPass = 0 To 1 Step 1
        For ilChar = 3 To 1 Step -1
            For ilDay = 0 To 6 Step 1
                ilPos = InStr(1, slDayNames, Left(slNames(ilDay), ilChar))
                If ilPos = 1 Then
                    If ilPass = 0 Then
                        ilStartDayIndex = ilDay
                        If Len(slDayNames) = ilChar Then
                            ilEndDayIndex = ilDay
                            Exit Sub
                        Else
                            slDayNames = Mid$(slDayNames, ilChar + 1)
                        End If
                        Exit For
                    Else
                        ilEndDayIndex = ilDay
                        Exit Sub
                    End If
                End If
                If (ilPass = 0) And (ilStartDayIndex <> -1) Then
                    Exit For
                End If
            Next ilDay
        Next ilChar
    Next ilPass
    Exit Sub
End Sub

Private Function mIsNewVersion(sPFN As String) As Boolean
    Dim ilRet As Integer
    Dim slLine As String
    Dim slChar As String
    Dim ilEof As Integer
    Dim slStrArr() As String
    
    mIsNewVersion = False
    ilRet = gFileOpen(sPFN, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        MsgBox "Open " & sPFN & ", Error #" & Str$(ilRet), vbExclamation, "Open Error"
        Exit Function
    End If
    Do
        ilRet = 0
        err.Clear
        On Error GoTo Err1:
        slLine = ""
        Do
            If Not EOF(hmFrom) Then
                slChar = Input(1, #hmFrom)
                If slChar = Chr(13) Then
                    slChar = Input(1, #hmFrom)
                End If
                If slChar = Chr(10) Then
                    Exit Do
                End If
                slLine = slLine & slChar
            Else
                ilEof = True
                Exit Do
            End If
        Loop
        slLine = Trim$(slLine)
        If InStr(1, RTrim$(slLine), "ROWTYPE", 1) > 0 Then
            slStrArr = Split(slLine, ",")
            If UBound(slStrArr) > 4 Then
                If Trim(slStrArr(1)) = "VALUE" _
                    And Trim(slStrArr(2)) = "IMP" _
                    And Trim(slStrArr(3)) = "RTG" _
                    And Trim(slStrArr(4)) = "COVPCT" Then
                    mIsNewVersion = True
                End If
            End If
            Close hmFrom
            Exit Function
        End If
    Loop Until ilEof
    
    Close hmFrom
    Exit Function
    
Err1:
End Function

'
' JD 09-1-21 : Created
'
' Added support for new ACT1 format. The import process was written in C#. This code creates
' the parameter file to pass the information to the Act1Import.exe program.
'
Private Function mPCNetConvFileNew(slFromFile As String, slDPBookName As String, slETBookName As String, slBookDate As String) As Integer
    Dim ilRet As Integer
    Dim slParamFile As String
    Dim slDSN As String
    Dim slImportProgram As String
    Dim llErrorCode As Long
    Dim slLine As String
    Dim slResults() As String
    Dim hlTo As Integer
    Dim hlFrom As Integer
    Dim slLocation As String

    On Error GoTo Err1
    If igTestSystem Then
        slLocation = "TestLocations"
    Else
        slLocation = "Locations"
    End If
    
    bmBooksSaved = False
    If Not gLoadOptionTrafficThenAffiliate(slLocation, "Name", slDSN) Then
        gMsgBox "Unable to load database name from ini file."
        Exit Function
    End If

    ' Create the file that the Act1Import.exe program needs to process the import file.
    slParamFile = sgExportPath & "ImportResearch_Parms.txt"
    ilRet = gFileOpen(slParamFile, "Output Access Write", hlTo)
    Print #hlTo, "DBCONN, " & slDSN
    Print #hlTo, "USERCODE, " & tgUrf(0).iCode
    Print #hlTo, "IMPORTFILE, " & slFromFile
    Print #hlTo, "BOOKNAME_DAYPART, " & slDPBookName
    Print #hlTo, "BOOKNAME_EXACTTIMES, " & slETBookName
    Print #hlTo, "BOOKDATE, " & slBookDate
    If rbcPCForm(0).Value Then
       Print #hlTo, "DEMOSUMMARY, Y"
       Print #hlTo, "LINUPANALYSIS, N"
    Else
       Print #hlTo, "DEMOSUMMARY, N"
       Print #hlTo, "LINUPANALYSIS, Y"
    End If

    If rbcNew(0).Value Then
        Print #hlTo, "NEW_BOOK, Y"
    Else
        Print #hlTo, "NEW_BOOK, N"
    End If
    If ckcDefault(0).Value Then
       Print #hlTo, "SETVEHICLEASDEFAULT, Y"
    Else
       Print #hlTo, "SETVEHICLEASDEFAULT, N"
    End If
    Print #hlTo, "CSIDATE, " & gNow()
    
    Close hlTo

    slImportProgram = sgExePath & "Act1Import.exe"
    If gFileExist(slImportProgram) = 1 Then
        gMsgBox "The program " & slImportProgram & " does not exist. Import cannot run."
        Exit Function
    End If

    ' Add the full path to the parameter file  as a command line argument.
    slImportProgram = slImportProgram & " " & slParamFile
    
    cmcFileConv.Enabled = False
    lbcErrors.AddItem "Import Started"

    gShellAndWait ImptMark, slImportProgram, vbHide, True

    cmcFileConv.Enabled = True

    ' Get the results of the import
    ilRet = gFileOpen(slParamFile, "Input Access Read", hlFrom)
    While Not EOF(hlFrom)
        Line Input #hlFrom, slLine
        gLogMsg slLine, "Act1Import.txt", False
        lbcErrors.AddItem slLine
    Wend
    Close hlFrom

    bmBooksSaved = True
    mPCNetConvFileNew = True
    MousePointer = vbDefault
    plcGauge.Value = 100
    Exit Function
Err1:
    llErrorCode = err.Number
    gMsg = "A general error has occured in ImptMark.frm-mPCNetConvFileNew"
    gLogMsg gMsg & err.Description & " Error #" & err.Number, "Act1Import.txt", False
    gMsgBox gMsg
    Exit Function
End Function

    
' First go at it
'Private Function mPCNetConvFileNew(slFromFile As String, slDPBookName As String, slETBookName As String, slBookDate As String) As Integer
'    Dim ilBNRet As Integer
'    Dim ilQRet As Integer
'    Dim slLine As String
'    Dim ilHeaderFd As Integer
'    Dim ilPopIndex As Integer
'    Dim ilPopCol As Integer
'    Dim ilAvgQHIndex As Integer
'    Dim ilPopDone As Integer
'    Dim ilDemoGender As Integer
'    Dim slDemoAge As String
'    Dim ilPos As Integer
'    Dim ilEndPos As Integer
'    Dim ilLastNonBlankCol As Integer
'    Dim llLineCount As Long
'    Dim slDay As String
'    Dim ilDay As Integer
'    Dim ilSY As Integer
'    Dim ilEY As Integer
'    Dim ilLoop As Integer
'    Dim ilEof As Integer
'    Dim llPercent As Long
'    Dim slChar As String
'    Dim slTime As String
'    Dim slStr As String
'    Dim slSexChar As String
'    Dim ilIndex As Integer
'    Dim ilEndIndex As Integer
'    Dim ilCustomDemo As Integer
'    Dim ilPrevDnfCode As Integer
'    Dim llPrevDrfCode As Long
'    Dim ilRdf As Integer
'    Dim ilRow As Integer
'    Dim ilCol As Integer
'    Dim slStationCode As String
'    Dim slAct1LineUpCode As String
'    Dim slACT1code As String
'    Dim ACT1stored As String
'    Dim slVehicleName As String
'    Dim slStrArr() As String
'    Dim ilNoDemoFd As Integer
'    Dim ilMatch As Integer
'    Dim llTime As Long
'    Dim ilVef As Integer
'    Dim ilVff As Integer
'    Dim ilTIndex As Integer
'    Dim ilSetBkNm As Integer
'    Dim ilSetDnfCode As Integer
'    Dim ilDpfUpper As Integer
'    Dim ilDpf As Integer
'    Dim ilPRet As Integer
'    Dim slDemoName As String
'    Dim blDontBlankLine As Boolean
'    Dim tlDrf As DRF
'    ilRet = 0
'    'On Error GoTo mPCNetConvFileErr:
'    'hmFrom = FreeFile
'    'Open slFromFile For Input Access Read As hmFrom
'    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
'    If ilRet <> 0 Then
'        Close hmFrom
'        MsgBox "Open " & slFromFile & ", Error #" & Str$(ilRet), vbExclamation, "Open Error"
'        edcFrom.SetFocus
'        mPCNetConvFileNew = False
'        Exit Function
'    End If
'    DoEvents
'    If imTerminate Then
'        Close hmFrom
'        mTerminate
'        mPCNetConvFileNew = False
'        Exit Function
'    End If
'    ilHeaderFd = False
'    blDontBlankLine = False
'    slLine = ""
'    llLineCount = 0
'    ilPopDone = False
'    tmDrfPop.lCode = 0
'    tmDrfPop.iDnfCode = tmDnf.iCode
'    tmDrfPop.sDemoDataType = "P"
'    tmDrfPop.iMnfSocEco = 0
'    tmDrfPop.iVefCode = 0
'    tmDrfPop.sInfoType = ""
'    tmDrfPop.iRdfCode = 0
'    tmDrfPop.sProgCode = ""
'    tmDrfPop.iStartTime(0) = 1
'    tmDrfPop.iStartTime(1) = 0
'    tmDrfPop.iEndTime(0) = 1
'    tmDrfPop.iEndTime(1) = 0
'    tmDrfPop.iStartTime2(0) = 1
'    tmDrfPop.iStartTime2(1) = 0
'    tmDrfPop.iEndTime2(0) = 1
'    tmDrfPop.iEndTime2(1) = 0
'    For ilDay = 0 To 6 Step 1
'        tmDrfPop.sDay(ilDay) = "Y"
'    Next ilDay
'    tmDrfPop.iQHIndex = 0
'    tmDrfPop.iCount = 0
'    tmDrfPop.sExStdDP = "N"
'    tmDrfPop.sExRpt = "N"
'    tmDrfPop.sDataType = "A"
'    For ilLoop = 1 To 16 Step 1
'        tmDrfPop.lDemo(ilLoop - 1) = 0
'    Next ilLoop
'    tmDrfPop.sACTLineupCode = ""
'    tmDrfPop.sACT1StoredTime = ""
'    tmDrfPop.sACT1StoredSpots = ""
'    tmDrfPop.sACT1StoreClearPct = ""
'    tmDrfPop.sACT1DaypartFilter = ""
'
'    ilNoDemoFd = 0
'    ilCustomDemo = False
'    Do
'        ilRet = 0
'        err.Clear
'        'On Error GoTo mPCNetConvFileErr:
'        'Line Input #hmFrom, slLine
'        slLine = ""
'        Do
'            If Not EOF(hmFrom) Then
'                slChar = Input(1, #hmFrom)
'                If slChar = Chr(13) Then
'                    slChar = Input(1, #hmFrom)
'                End If
'                If slChar = Chr(10) Then
'                    Exit Do
'                End If
'                slLine = slLine & slChar
'            Else
'                ilEof = True
'                Exit Do
'            End If
'        Loop
'        llLineCount = llLineCount + 1
'        slLine = Trim$(slLine)
'        On Error GoTo 0
'        ilRet = err.Number
'        If ilRet = 62 Then
'            Exit Do
'        End If
'        If Len(slLine) > 0 Then
'            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
'                ilEof = True
'            Else
'                DoEvents
'                If imTerminate Then
'                    Close hmFrom
'                    mTerminate
'                    mPCNetConvFileNew = False
'                    Exit Function
'                End If
'                ilLastNonBlankCol = -1
'                gParseCDFields slLine, False, smFieldValues()
'                '11/10/16: Move fields from 0 to X to 1 to X +1
'                For ilLoop = UBound(smFieldValues) - 1 To LBound(smFieldValues) Step -1
'                    smFieldValues(ilLoop + 1) = smFieldValues(ilLoop)
'                    If Trim(smFieldValues(ilLoop + 1)) <> "" Then
'                        If ilLoop + 1 > ilLastNonBlankCol Then
'                            ilLastNonBlankCol = ilLoop + 1
'                        End If
'                    End If
'                Next ilLoop
'                smFieldValues(0) = ""
'
'                If InStr(1, Trim$(slLine), "CustomData", 1) > 0 Then
'                    ' CustomData,Phil Collins
'                    slVehicleName = ""
'                    slStrArr = Split(slLine, ",")
'                    If UBound(slStrArr) > 0 Then
'                        slVehicleName = slStrArr(1)
'                    End If
'
'                    ReDim imVefCodeImpt(0 To 0) As Integer
'                    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
'                        If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Or (tgMVef(ilLoop).sType = "R") Or (tgMVef(ilLoop).sType = "V") Then
'                            If ((slStationCode <> "") And (StrComp(UCase(slStationCode), UCase(Trim$(tgMVef(ilLoop).sCodeStn)), 1) = 0)) Or (StrComp(UCase(slVehicleName), UCase(Trim$(tgMVef(ilLoop).sName)), 1) = 0) Then
'                                imVehSelectedIndex = ilLoop
'                                ilHeaderFd = True
'                                imVefCodeImpt(UBound(imVefCodeImpt)) = tgMVef(ilLoop).iCode
'                                ReDim Preserve imVefCodeImpt(0 To UBound(imVefCodeImpt) + 1) As Integer
'                            Else
'                                ilVff = gBinarySearchVff(tgMVef(ilLoop).iCode)
'                                If ilVff <> -1 Then
'                                    If ((slStationCode <> "") And (StrComp(UCase(slStationCode), UCase(Trim$(tgVff(ilVff).sACT1LineupCode)), 1) = 0)) Then
'                                        imVehSelectedIndex = ilLoop
'                                        ilHeaderFd = True
'                                        imVefCodeImpt(UBound(imVefCodeImpt)) = tgMVef(ilLoop).iCode
'                                        ReDim Preserve imVefCodeImpt(0 To UBound(imVefCodeImpt) + 1) As Integer
'                                    End If
'                                End If
'                            End If
'                        End If
'                    Next ilLoop
'                    If UBound(imVefCodeImpt) <= 0 Then
'                        Print #hmTo, slVehicleName & " Not Found"
'                        lbcErrors.AddItem slVehicleName & " Not Found"
'                    Else
'                        ilNoDemoFd = 0
'                        For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
'                            tmDrfInfo(ilCol).tDrf.lCode = 0
'                            tmDrfInfo(ilCol).tDrf.iDnfCode = tmDnf.iCode
'                            tmDrfInfo(ilCol).tDrf.sDemoDataType = "D"
'                            tmDrfInfo(ilCol).tDrf.iMnfSocEco = 0
'                            If tmDrfInfo(ilCol).iType = 0 Then
'                                tmDrfInfo(ilCol).tDrf.sInfoType = "V"
'                                tmDrfInfo(ilCol).tDrf.iRdfCode = 0
'                                gPackTime "12AM", tmDrfInfo(ilCol).tDrf.iStartTime(0), tmDrfInfo(ilCol).tDrf.iStartTime(1)
'                                gPackTime "12AM", tmDrfInfo(ilCol).tDrf.iEndTime(0), tmDrfInfo(ilCol).tDrf.iEndTime(1)
'                                For ilDay = 0 To 6 Step 1
'                                    tmDrfInfo(ilCol).tDrf.sDay(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
'                                Next ilDay
'                            Else
'                                tmDrfInfo(ilCol).tDrf.sInfoType = "D"
'                                tmDrfInfo(ilCol).tDrf.iRdfCode = 0
'                                gPackTimeLong tmDrfInfo(ilCol).lStartTime, tmDrfInfo(ilCol).tDrf.iStartTime(0), tmDrfInfo(ilCol).tDrf.iStartTime(1)
'                                gPackTimeLong tmDrfInfo(ilCol).lEndTime, tmDrfInfo(ilCol).tDrf.iEndTime(0), tmDrfInfo(ilCol).tDrf.iEndTime(1)
'                                For ilDay = 0 To 6 Step 1
'                                    tmDrfInfo(ilCol).tDrf.sDay(ilDay) = tmDrfInfo(ilCol).sDays(ilDay)
'                                Next ilDay
'                            End If
'                            tmDrfInfo(ilCol).tDrf.sProgCode = ""
'                            tmDrfInfo(ilCol).tDrf.iStartTime2(0) = 1
'                            tmDrfInfo(ilCol).tDrf.iStartTime2(1) = 0
'                            tmDrfInfo(ilCol).tDrf.iEndTime2(0) = 1
'                            tmDrfInfo(ilCol).tDrf.iEndTime2(1) = 0
'                            tmDrfInfo(ilCol).tDrf.iQHIndex = 0
'                            tmDrfInfo(ilCol).tDrf.iCount = 0
'                            tmDrfInfo(ilCol).tDrf.sExStdDP = "N"
'                            tmDrfInfo(ilCol).tDrf.sExRpt = "N"
'                            tmDrfInfo(ilCol).tDrf.sDataType = "A"
'                            For ilLoop = 1 To 16 Step 1
'                                tmDrfInfo(ilCol).tDrf.lDemo(ilLoop - 1) = 0
'                            Next ilLoop
'                        Next ilCol
'                    End If
'
'                ElseIf InStr(1, Trim$(slLine), "ACT1code", 1) > 0 Then
'                    ' ACT1code,PCOLN
'                    slStrArr = Split(slLine, ",")
'                    If UBound(slStrArr) > 0 Then
'                        slACT1code = slStrArr(1)
'                    End If
'
'                ElseIf InStr(1, Trim$(slLine), "LineupCode", 1) > 0 Then
'                    slStrArr = Split(slLine, ",")
'                    If UBound(slStrArr) > 0 Then
'                        slAct1LineUpCode = slStrArr(1)
'                    End If
'
'                ElseIf InStr(1, Trim$(slLine), "ACT1stored", 1) > 0 Then
'                    ' ACT1stored,T
'                    ' T = Stored Times
'                    ' S = Stored Spots
'                    ' C = Store Clear%
'                    ' F = Daypart filter
'
'                    ReDim tmDrfInfo(0 To 0) As DRFINFO
'                    ReDim tgDpfInfo(0 To 0) As DPFINFO
'                    tmDrfInfo(UBound(tmDrfInfo)).iStartCol = 2
'                    tmDrfInfo(UBound(tmDrfInfo)).iType = 0  'Vehicle
'                    tmDrfInfo(UBound(tmDrfInfo)).iSY = -1
'                    tmDrfInfo(UBound(tmDrfInfo)).iEY = -1
'                    ReDim Preserve tmDrfInfo(0 To UBound(tmDrfInfo) + 1) As DRFINFO
'                    ilCol = UBound(tmDrfInfo)
'
'                    tmDrfInfo(ilCol).tDrf.sACTLineupCode = slAct1LineUpCode
'                    tmDrfPop.sACTLineupCode = slAct1LineUpCode
'
'                    slStrArr = Split(slLine, ",")
'                    If UBound(slStrArr) > 0 Then
'                        If InStr(1, slStrArr(1), "T", 1) > 0 Then
'                            tmDrfInfo(ilCol).tDrf.sACT1StoredTime = "T"
'                            tmDrfPop.sACT1StoredTime = "T"
'                        End If
'                        If InStr(1, slStrArr(1), "S", 1) > 0 Then
'                            tmDrfInfo(ilCol).tDrf.sACT1StoredSpots = "S"
'                            tmDrfPop.sACT1StoredSpots = "S"
'                        End If
'                        If InStr(1, slStrArr(1), "C", 1) > 0 Then
'                            tmDrfInfo(ilCol).tDrf.sACT1StoreClearPct = "C"
'                            tmDrfPop.sACT1StoreClearPct = "C"
'                        End If
'                        If InStr(1, slStrArr(1), "F", 1) > 0 Then
'                            tmDrfInfo(ilCol).tDrf.sACT1DaypartFilter = "F"
'                            tmDrfPop.sACT1DaypartFilter = "F"
'                        End If
'                    End If
'
'                ElseIf InStr(1, Trim$(slLine), "Daypart", 1) > 0 Then
'                    ' Daypart,Stored Schedules
'                    ' Daypart,MF 6-10a
'                    slStrArr = Split(slLine, ",")
'                    If UBound(slStrArr) > 0 Then
'                        slStr = slStrArr(1)
'                        If InStr(1, slStr, "Stored Schedule", 1) > 0 Then
'                            tmDrfInfo(ilCol).iBkNm = 1  'Exact Time
'                            tmDrfInfo(ilCol).iType = 0  'Vehicle
'                            For ilDay = 0 To 6 Step 1
'                                tmDrfInfo(ilCol).sDays(ilDay) = "Y"
'                            Next ilDay
'                        Else
'                            For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
'                                If tmDrfInfo(ilCol).iSY = -1 Then
'                                    ilSY = -1
'                                    ilTIndex = InStr(1, slStr, " ", 1) + 1
'                                    If ilTIndex > 2 Then
'                                        mGetDayIndex Left$(slStr, ilTIndex - 2), ilSY, ilEY
'                                    End If
'                                    If ilSY <> -1 Then
'                                        slChar = Mid$(slStr, ilTIndex, 1)
'                                        If (slChar >= "0") And (slChar <= "9") Then
'                                            slStr = Mid$(slStr, ilTIndex)
'                                        Else
'                                            slStr = Mid$(slStr, ilTIndex + 1)
'                                        End If
'                                        slTime = Mid$(slStr, 1, 1)
'                                        slStr = Mid$(slStr, 2)
'                                        slChar = Mid$(slStr, 1, 1)
'                                        If (slChar >= "0") And (slChar <= "9") Then
'                                            slTime = slTime & Mid$(slStr, 1, 2)
'                                            slStr = Mid$(slStr, 3)
'                                        Else
'                                            If Mid$(slStr, 1, 1) = "-" Then
'                                                slTime = slTime & right$(slStr, 1)
'                                            Else
'                                                slTime = slTime & Mid$(slStr, 1, 1)
'                                                slStr = Mid$(slStr, 2)
'                                            End If
'                                        End If
'                                        If gValidTime(slTime) Then
'                                            tmDrfInfo(ilCol).lStartTime = gTimeToCurrency(slTime, False)
'                                            slChar = Mid$(slStr, 1, 1)
'                                            If slChar = "-" Then
'                                                slStr = Trim$(Mid$(slStr, 2))
'                                            End If
'                                            slTime = Mid$(slStr, 1, 1)
'                                            slStr = Mid$(slStr, 2)
'                                            slChar = Mid$(slStr, 1, 1)
'                                            If (slChar >= "0") And (slChar <= "9") Then
'                                                slTime = slTime & Mid$(slStr, 1, 2)
'                                                slStr = Mid$(slStr, 3)
'                                            Else
'                                                slTime = slTime & Mid$(slStr, 1, 1)
'                                                slStr = Mid$(slStr, 2)
'                                            End If
'                                            If gValidTime(slTime) Then
'                                                tmDrfInfo(ilCol).iSY = ilSY
'                                                tmDrfInfo(ilCol).iEY = ilEY
'                                                tmDrfInfo(ilCol).lEndTime = gTimeToCurrency(slTime, False)
'                                                For ilDay = 0 To 6 Step 1
'                                                    tmDrfInfo(ilCol).sDays(ilDay) = "N"
'                                                Next ilDay
'                                                For ilDay = ilSY To ilEY Step 1
'                                                    tmDrfInfo(ilCol).sDays(ilDay) = "Y"
'                                                Next ilDay
'                                                tmDrfInfo(ilCol).iType = 1  'Daypart
'                                            End If
'                                        End If
'                                    End If
'                                End If
'                            Next ilCol
'                        End If
'                    End If
'
'                ElseIf InStr(1, Trim$(slLine), "DEMO", 1) > 0 Then
'                    ' DEMO,Persons 12+,227400,0.08,79.8,277786500
'                    slDemoName = ""
'                    slStrArr = Split(slLine, ",")
'                    If UBound(slStrArr) > 0 Then
'                        slDemoName = slStrArr(1)
'                    End If
'
'                    ' Stuff these values in here so we can reuse the rest of the code.
'                    For ilLoop = LBound(slStrArr) To UBound(slStrArr) Step 1
'                        smFieldValues(ilLoop) = slStrArr(ilLoop)
'                    Next
'
'                    ilDemoGender = -1
'                    ilPopCol = 6
'                    For ilLoop = LBound(tgMnfCDemo) To UBound(tgMnfCDemo) Step 1
'                        If StrComp(Trim$(tgMnfCDemo(ilLoop).sName), slDemoName, 1) = 0 Then
'                            ilDemoGender = tgMnfCDemo(ilLoop).iGroupNo
'                            ilCustomDemo = True
'                            Exit For
'                        End If
'                    Next ilLoop
'                    If ilDemoGender = -1 Then
'                        slSexChar = UCase$(Left$(slDemoName, 1))
'                        If (slSexChar = "M") Or (slSexChar = "B") Or (slSexChar = "W") Or (slSexChar = "F") Or (slSexChar = "G") Then
'                            'Scan for xx-yy
'                            ilIndex = 2
'                            Do While ilIndex < Len(slDemoName)
'                                slChar = Mid$(slDemoName, ilIndex, 1)
'                                If (slChar >= "0") And (slChar <= "9") Then
'                                    ilPos = ilIndex
'                                    Exit Do
'                                End If
'                                ilIndex = ilIndex + 1
'                            Loop
'                        ElseIf (slSexChar = "P") Or (slSexChar = "T") Then
'                            ilIndex = 2
'                            Do While ilIndex < Len(slDemoName)
'                                slChar = Mid$(slDemoName, ilIndex, 1)
'                                If (slChar >= "0") And (slChar <= "9") Then
'                                    ilPos = ilIndex
'                                    Exit Do
'                                End If
'                                ilIndex = ilIndex + 1
'                            Loop
'                        End If
'                        If ilPos > 0 Then
'                            '2/16/21 - TTP 10080 - found start position, Check for end position...
'                            ilEndPos = 0
'                            ilEndIndex = Len(slDemoName)
'                            Do While ilEndIndex > 1
'                                slChar = Mid$(slDemoName, ilEndIndex, 1)
'                                If ((slChar >= "0") And (slChar <= "9")) Or (slChar = "+") Then
'                                    ilEndPos = ilEndIndex
'                                    Exit Do
'                                Else
'                                    'not a number or special character
'                                End If
'                                ilEndIndex = ilEndIndex - 1
'                            Loop
'                            'Truncate DemoAge to remove any trailing characters
'                            If ilEndPos < Len(slDemoName) And ilEndPos <> 0 Then
'                                slDemoName = Trim$(Mid$(slDemoName, 1, ilEndPos))
'                            End If
'
'                            If smDataForm <> "8" Then
'                                If (slSexChar = "M") Or (slSexChar = "B") Then
'                                    ilDemoGender = 0
'                                    slDemoAge = Trim$(Mid$(slDemoName, ilPos))
'                                ElseIf (slSexChar = "W") Or (slSexChar = "G") Then
'                                    ilDemoGender = 8
'                                    slDemoAge = Trim$(Mid$(slDemoName, ilPos))
'                                ElseIf (slSexChar = "T") Or (slSexChar = "P") Then
'                                    ilDemoGender = -1
'                                    slDemoAge = Trim$(Mid$(slDemoName, ilPos))
'                                End If
'                                If ilDemoGender >= 0 Then
'                                    Select Case slDemoAge
'                                        Case "12-17"
'                                            ilDemoGender = ilDemoGender + 1
'                                        Case "18-24"
'                                            ilDemoGender = ilDemoGender + 2
'                                        Case "25-34"
'                                            ilDemoGender = ilDemoGender + 3
'                                        Case "35-44"
'                                            ilDemoGender = ilDemoGender + 4
'                                        Case "45-49"
'                                            ilDemoGender = ilDemoGender + 5
'                                        Case "50-54"
'                                            ilDemoGender = ilDemoGender + 6
'                                        Case "55-64"
'                                            ilDemoGender = ilDemoGender + 7
'                                        Case "65+"
'                                            ilDemoGender = ilDemoGender + 8
'                                        Case Else
'                                            ilDemoGender = -1
'                                    End Select
'                                End If
'                            Else
'                                If (slSexChar = "M") Or (slSexChar = "B") Then
'                                    ilDemoGender = 0
'                                    slDemoAge = Trim$(Mid$(slDemoName, ilPos))
'                                ElseIf (slSexChar = "W") Or (slSexChar = "G") Then
'                                    ilDemoGender = 9
'                                    slDemoAge = Trim$(Mid$(slDemoName, ilPos))
'                                ElseIf (slSexChar = "T") Or (slSexChar = "P") Then
'                                    ilDemoGender = -1
'                                    slDemoAge = Trim$(Mid$(slDemoName, ilPos))
'                                End If
'                                If ilDemoGender >= 0 Then
'                                    Select Case slDemoAge
'                                        Case "12-17"
'                                            ilDemoGender = ilDemoGender + 1
'                                        Case "18-20"
'                                            ilDemoGender = ilDemoGender + 2
'                                        Case "21-24"
'                                            ilDemoGender = ilDemoGender + 3
'                                        Case "25-34"
'                                            ilDemoGender = ilDemoGender + 4
'                                        Case "35-44"
'                                            ilDemoGender = ilDemoGender + 5
'                                        Case "45-49"
'                                            ilDemoGender = ilDemoGender + 6
'                                        Case "50-54"
'                                            ilDemoGender = ilDemoGender + 7
'                                        Case "55-64"
'                                            ilDemoGender = ilDemoGender + 8
'                                        Case "65+"
'                                            ilDemoGender = ilDemoGender + 9
'                                        Case Else
'                                            ilDemoGender = -1
'                                    End Select
'                                End If
'                            End If  'dataForm not 8
'                            If ilDemoGender > 0 Then
'                                ilPopIndex = ilPopCol
'                                If tgSpf.sSAudData = "H" Then
'                                    tmDrfPop.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilPopIndex)) + 50) \ 100
'                                ElseIf tgSpf.sSAudData = "N" Then
'                                    tmDrfPop.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilPopIndex)) + 5) \ 10
'                                ElseIf tgSpf.sSAudData = "U" Then
'                                    tmDrfPop.lDemo(ilDemoGender - 1) = CLng(smFieldValues(ilPopIndex))
'                                Else
'                                    tmDrfPop.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilPopIndex)) + 500) \ 1000
'                                End If
'                                For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
'                                    ilAvgQHIndex = tmDrfInfo(ilCol).iStartCol
'                                    If tgSpf.sSAudData = "H" Then
'                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 50) \ 100
'                                    ElseIf tgSpf.sSAudData = "N" Then
'                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 5) \ 10
'                                    ElseIf tgSpf.sSAudData = "U" Then
'                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = CLng(smFieldValues(ilAvgQHIndex))
'                                    Else
'                                        tmDrfInfo(ilCol).tDrf.lDemo(ilDemoGender - 1) = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
'                                    End If
'                                Next ilCol
'                                ilNoDemoFd = ilNoDemoFd + 1
'                            Else
'                                ilDpfUpper = -1
'                                If (slSexChar = "M") Or (slSexChar = "B") Then
'                                    slDemoName = "M" & slDemoAge
'                                ElseIf (slSexChar = "W") Or (slSexChar = "G") Then
'                                    slDemoName = "W" & slDemoAge
'                                ElseIf (slSexChar = "T") Or (slSexChar = "P") Then
'                                    If InStr(1, slDemoAge, "12", 1) > 0 Then
'                                        slDemoName = "P" & slDemoAge
'                                    Else
'                                        slDemoName = "A" & slDemoAge
'                                    End If
'                                End If
'                                slDemoName = Trim$(slDemoName)
'                                For ilLoop = LBound(tgMnfSDemo) To UBound(tgMnfSDemo) Step 1
'                                    If StrComp(Trim$(tgMnfSDemo(ilLoop).sName), slDemoName, 1) = 0 Then
'                                        ilDpfUpper = UBound(tgDpfInfo)
'                                        Exit For
'                                    End If
'                                Next ilLoop
'                                If ilDpfUpper = -1 Then
'                                    'Test if valid name like M18-49
'                                    If (Len(slDemoName) >= 4) And Len(slDemoName) <= 6 Then
'                                        If (Mid$(slDemoName, 4, 1) = "+") Or (Mid$(slDemoName, 4, 1) = "-") Then
'                                            ilRet = mAddDemo(slDemoName)
'                                        End If
'                                    End If
'                                End If
'                                For ilLoop = LBound(tgMnfSDemo) To UBound(tgMnfSDemo) Step 1
'                                    If StrComp(Trim$(tgMnfSDemo(ilLoop).sName), slDemoName, 1) = 0 Then
'                                        For ilCol = 0 To UBound(tmDrfInfo) - 1 Step 1
'                                            ilDpfUpper = UBound(tgDpfInfo)
'                                            tgDpfInfo(ilDpfUpper).iCol = ilCol
'                                            tgDpfInfo(ilDpfUpper).tDpf.iMnfDemo = tgMnfSDemo(ilLoop).iCode
'                                            If tgSpf.sSAudData = "H" Then
'                                                tgDpfInfo(ilDpfUpper).tDpf.lPop = (CLng(smFieldValues(ilPopCol)) + 50) \ 100
'                                            ElseIf tgSpf.sSAudData = "N" Then
'                                                tgDpfInfo(ilDpfUpper).tDpf.lPop = (CLng(smFieldValues(ilPopCol)) + 5) \ 10
'                                            ElseIf tgSpf.sSAudData = "U" Then
'                                                tgDpfInfo(ilDpfUpper).tDpf.lPop = CLng(smFieldValues(ilPopCol))
'                                            Else
'                                                tgDpfInfo(ilDpfUpper).tDpf.lPop = (CLng(smFieldValues(ilPopCol)) + 500) \ 1000
'                                            End If
'                                            ilAvgQHIndex = tmDrfInfo(ilCol).iStartCol
'                                            If tgSpf.sSAudData = "H" Then
'                                                tgDpfInfo(ilDpfUpper).tDpf.lDemo = (CLng(smFieldValues(ilAvgQHIndex)) + 50) \ 100
'                                            ElseIf tgSpf.sSAudData = "N" Then
'                                                tgDpfInfo(ilDpfUpper).tDpf.lDemo = (CLng(smFieldValues(ilAvgQHIndex)) + 5) \ 10
'                                            ElseIf tgSpf.sSAudData = "U" Then
'                                                tgDpfInfo(ilDpfUpper).tDpf.lDemo = CLng(smFieldValues(ilAvgQHIndex))
'                                            Else
'                                                tgDpfInfo(ilDpfUpper).tDpf.lDemo = (CLng(smFieldValues(ilAvgQHIndex)) + 500) \ 1000
'                                            End If
'                                            ReDim Preserve tgDpfInfo(0 To ilDpfUpper + 1) As DPFINFO
'                                        Next ilCol
'                                        Exit For
'                                    End If
'                                Next ilLoop
'                            End If
'                        End If
'                    End If
'
'                End If
'
'            End If
'            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
'            llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
'            If llPercent >= 100 Then
'                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
'                    llPercent = 99
'                Else
'                    llPercent = 100
'                End If
'            End If
'            If plcGauge.Value <> llPercent Then
'                plcGauge.Value = llPercent
'            End If
'        End If
'    Loop Until ilEof
'    If (ilHeaderFd) And (ilNoDemoFd > 0) Then
'        ilHeaderFd = False
'        blDontBlankLine = False
'        '6/9/16: Replaced GoSub
'        'GoSub mPCWriteRec
'        mPCWriteRec ilRet, ilCustomDemo, slVehicleName, slDPBookName, slETBookName, slBookDate, ilPrevDnfCode, llPrevDrfCode
'        If ilRet <> 0 Then
'            mPCNetConvFileNew = False
'            Exit Function
'        End If
'    End If
'    Close hmFrom
'    plcGauge.Value = 100
'    mPCNetConvFileNew = True
'    MousePointer = vbDefault
'    Exit Function
'End Function

