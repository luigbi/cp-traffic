VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImptVeh 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5700
   ClientLeft      =   1185
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
      Left            =   6600
      TabIndex        =   19
      Top             =   1320
      Width           =   375
   End
   Begin VB.ListBox lbcDemo 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "ImptVeh.frx":0000
      Left            =   8295
      List            =   "ImptVeh.frx":0002
      TabIndex        =   18
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
      TabIndex        =   10
      Top             =   2505
      Width           =   6075
      Begin VB.CheckBox ckcDefault 
         Caption         =   "Reallocation Book"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   4020
         TabIndex        =   12
         Top             =   15
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
   Begin VB.TextBox edcBookDate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      MaxLength       =   10
      TabIndex        =   9
      Top             =   2025
      Width           =   1275
   End
   Begin VB.TextBox edcBookName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      MaxLength       =   30
      TabIndex        =   7
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
      Left            =   7095
      TabIndex        =   5
      Top             =   1290
      Width           =   1605
   End
   Begin VB.PictureBox plcFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1140
      ScaleHeight     =   345
      ScaleWidth      =   5355
      TabIndex        =   3
      Top             =   1260
      Width           =   5385
      Begin VB.TextBox edcFrom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   60
         TabIndex        =   4
         Top             =   60
         Width           =   5265
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
      TabIndex        =   17
      Top             =   2985
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label lacBookDate 
      Appearance      =   0  'Flat
      Caption         =   "Book Date"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   75
      TabIndex        =   8
      Top             =   2070
      Width           =   1095
   End
   Begin VB.Label lacBookName 
      Appearance      =   0  'Flat
      Caption         =   "Book Name"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   75
      TabIndex        =   6
      Top             =   1725
      Width           =   1095
   End
   Begin VB.Label lacMsg 
      Appearance      =   0  'Flat
      Caption         =   $"ImptVeh.frx":0004
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
      TabIndex        =   16
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
Attribute VB_Name = "ImptVeh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ImptVeh.FRM on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImptVeh.Frm
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
Dim hmDpf As Integer    'file handle
Dim tmDpf As DPF
Dim tmPlusDpf() As DPF
Dim imDpfRecLen As Integer  'Record length
Dim tmDpfSrchKey As LONGKEY0
Dim tmDpfSrchKey1 As DPFKEY1
Dim hmVef As Integer    'file handle
Dim tmVef As VEF
Dim imVefRecLen As Integer  'Record length
Dim tmVefSrchKey As INTKEY0
Dim imVefCodeInDnf() As Integer   'Array if vehicles code which are vehicles with data in dnf
Dim hmMnf As Integer    'file handle
Dim tmMnf As MNF        'Record structure
Dim imMnfRecLen As Integer  'Record length
'Dim smFieldValues(1 To 30) As String    '25 fields generated in a record
Dim smFieldValues(0 To 35) As String    '25 fields generated in a record
'Private imDemoMap(1 To 30) As Integer
Private imDemoMap(0 To 35) As Integer
'Private lmPopValues(1 To 30) As Long
Private lmPopValues(0 To 35) As Long
Dim imTerminate As Integer
Dim imConverting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smNowDate As String
Dim lmNowDate As Long
Dim imNowYear As Integer
Dim smSyncDate As String
Dim smSyncTime As String
Dim imComboBoxIndex As Integer
Dim imBSMode As Integer
Dim imChgMode As Integer
Dim imVefCode As Integer
Dim smDataForm As String
Dim smBookForm As String
Dim imMaxPlusValueM As Integer
Dim imMaxPlusFieldM As Integer
Dim imMaxPlusIndexM As Integer
Dim imMaxPlusValueW As Integer
Dim imMaxPlusFieldW As Integer
Dim imMaxPlusIndexW As Integer
Dim bm65PlusExist As Boolean
Dim bmResearchSaved As Boolean

Private Type STDDEMOINFO
    sName As String * 10
    iMnfCode As Integer
End Type

Dim tmStdDemoInfo() As STDDEMOINFO

Dim tmNameCode() As SORTCODE
Dim smNameCodeTag As String

Const SOURCEINDEX = 1
Const STATIONCODEINDEX = 2
Const VEHICLEINDEX = 3
Const DEMOSTARTINDEX = 4




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
    ilRet = gIMoveListBox(ImptVeh, lbcDemo, tmNameCode(), smNameCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
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


Private Sub cmcBrowse_Click()
    'TTP 10339 - Automation Import - use windows File Browser (to see Local Drives and mapped drives from RDP)
    CMDialogBox.DialogTitle = "Import From File"
    CMDialogBox.Filter = "All (*.*)|*.*|CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc|Text (*.txt)|*.Txt|Print (*.prn)|*.Prn"
    CMDialogBox.InitDir = sgImportPath
    CMDialogBox.DefaultExt = "All (*.*)"
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
    Dim ilLoop As Integer
    Dim blChkSourceExist As Boolean
    
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
    blChkSourceExist = False
    slBookName = Trim$(edcBookName.Text)
    If slBookName = "" Then
        'MsgBox "Book Name Must be Defined", vbExclamation, "Name Error"
        'edcBookName.SetFocus
        'Exit Sub
        blChkSourceExist = True
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
    slBookDate = Trim$(edcBookDate.Text)
    'Check file names
    gGetSyncDateTime smSyncDate, smSyncTime
    For ilLoop = LBound(imDemoMap) To UBound(imDemoMap) Step 1
        imDemoMap(ilLoop) = 0
        lmPopValues(ilLoop) = 0
    Next ilLoop
    plcGauge.Value = 0
    lmProcessedNoBytes = 0
    ilRet = 0
    If Not mPrepassDemo(slFromName) Then
        MsgBox "Unable to process " & slFromName, vbExclamation, "Name Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
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
    If Not mTestColumnTitle(slFromName, blChkSourceExist) Then
        Close hmFrom
        cmcCancel.SetFocus
        Exit Sub
    End If
    Close hmFrom
    'hmTo = FreeFile
    'Open sgDBPath & "Messages\" & "ImptVeh.Txt" For Output As hmTo
    ilRet = gFileOpen(sgDBPath & "Messages\" & "ImptVeh.Txt", "Output", hmTo)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Open " & sgDBPath & "Messages\" & "ImptVeh.Txt" & ", Error #" & Str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Print #hmTo, "Import Vehicle Research on " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmTo, ""
    lacFileType.Caption = "Processing " & slFromName
    imConverting = True
    If Not mConvFile(slFromName, slBookName, slBookDate, slName) Then
        Print #hmTo, "Import Vehicle Research terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Close hmTo
        imConverting = False
        mTerminate
        Exit Sub
    End If
    Print #hmTo, "Import Vehicle Research successfully completed on " & Format$(gNow(), "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Close hmTo
    lacFileType.Caption = "Done"
    cmcCancel.Caption = "&Done"
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
    
    'TTP 10339 - Automation Import - use windows File Browser (to see Local Drives and mapped drives from RDP)
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
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    ilRet = btrClose(hmDnf)
    btrDestroy hmDnf
    ilRet = btrClose(hmDpf)
    btrDestroy hmDpf
    
    Set ImptVeh = Nothing   'Remove data segment

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
Private Function mBookNameUsed(slBookName As String, slBookDate As String, ilVefCode As Integer, ilPrevDnfCode As Integer, llPrevDrfCode As Long) As Integer
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
                tmDrfSrchKey.sInfoType = "V"
                tmDrfSrchKey.iRdfCode = 0
                ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tgDnfBook(ilLoop).iCode) And (tlDrf.sDemoDataType = "D") And (tlDrf.iVefCode = ilVefCode) And (tlDrf.sInfoType = "V") And (tlDrf.iRdfCode = 0)
                    mBookNameUsed = 3
                    llPrevDrfCode = tlDrf.lCode
                    Exit Function
                Loop
                Exit Function
            End If
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
Private Function mConvFile(slFromFile As String, slInBookName As String, slBookDate As String, slName As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilHeaderFd As Integer
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
    Dim ilPrevDnfCode As Integer
    Dim llPrevDrfCode As Long
    Dim ilRdf As Integer
    Dim ilRow As Integer
    Dim tlDrf As DRF
    Dim ilDpf As Integer
    Dim ilField As Integer
    Dim llDemo As Long
    Dim llSum As Long
    Dim llPop As Long
    Dim ilSum1 As Integer
    Dim ilSum2 As Integer
    Dim slDemoPlus As String
    Dim ilStd As Integer
    Dim ilDpfMaxM As Integer
    Dim ilDpfMaxW As Integer
    Dim ilDpfMax  As Integer
    Dim ilStartIndex  As Integer
    Dim ilEndIndex  As Integer
    Dim ilAdj As Integer
    Dim ilAdjInc As Integer
    Dim ilDpf12M As Integer
    Dim ilDpf12W As Integer
    Dim slBookName As String
        
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
    For ilLoop = 1 To 18 Step 1
        tmDrfPop.lDemo(ilLoop - 1) = 0
    Next ilLoop
    tmDrf.lCode = 0
    tmDrf.iDnfCode = tmDnf.iCode
    tmDrf.sDemoDataType = "D"
    tmDrf.iMnfSocEco = 0
    tmDrf.iVefCode = imVefCode
    tmDrf.sInfoType = "V"
    tmDrf.iRdfCode = 0
    gPackTime "12AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
    gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    tmDrf.sProgCode = ""
    tmDrf.iStartTime2(0) = 1
    tmDrf.iStartTime2(1) = 0
    tmDrf.iEndTime2(0) = 1
    tmDrf.iEndTime2(1) = 0
    For ilDay = 0 To 6 Step 1
        tmDrf.sDay(ilDay) = "Y"
    Next ilDay
    tmDrf.iQHIndex = 0
    tmDrf.iCount = 0
    tmDrf.sExStdDP = "N"
    tmDrf.sExRpt = "N"
    tmDrf.sDataType = "A"
    For ilLoop = 1 To 18 Step 1
        tmDrf.lDemo(ilLoop - 1) = 0
    Next ilLoop
    tmDpf.lCode = 0
    tmDpf.iDnfCode = tmDnf.iCode
    tmDpf.iMnfDemo = 0
    tmDpf.lDrfCode = 0
    tmDpf.lDemo = 0
    tmDpf.lPop = 0
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
                'Header Record (record 1)
                '   Source,Code,Podcast,M12-17,M18-24,M25-34,M35-44,M45-49,M50-54,M55+,W12-17,W18-24,w25-34,w35-44,w45-49,w50-54,w55+
                'Population record (record 2, used with all imported books that follow)
                '   ,,Population,12877000,16173000,21015000,19968000,10620000,10858000,38769000,12298000,15263000,20592000,20167000,10848000,11273000,45872000
                'Demo Record (record 3 thru n)
                '   BookName1,DP,Dan Patrick Show,98,3539,12191,32344,39423,7275,3638,0,0,0,0,0,0,0
                '   BookName1,Drew , Dr.Drew, 0, 2380, 15503, 7752, 4182, 2380, 1802, 0, 3041, 19811, 9906, 5344, 3041, 2303
                '   BookName2,DP,Dan Patrick Show,98,3539,12191,32344,39423,7275,3638,0,0,0,0,0,0,0
                '   BookName2,Bill , Bill Crystal, 0, 2380, 15503, 7752, 4182, 2380, 1802, 0, 3041, 19811, 9906, 5344, 3041, 2303
                
                If Not ilHeaderFd Then
                    If (StrComp(UCase(Trim$(smFieldValues(SOURCEINDEX))), "SOURCE", 1) = 0) Then
                        If (StrComp(UCase(Trim$(smFieldValues(STATIONCODEINDEX))), "CODE", 1) = 0) Then
                            If (StrComp(UCase(Trim$(smFieldValues(VEHICLEINDEX))), "VEHICLE", 1) = 0) Then
                                ilHeaderFd = True
                            End If
                        End If
                    End If
                Else
                    ilDemoGender = -1
                    ilPos = -1
                    smFieldValues(VEHICLEINDEX) = Trim$(smFieldValues(VEHICLEINDEX))
                    If UCase(Trim$(smFieldValues(VEHICLEINDEX))) = "POPULATION" Then
                        'Population record- record values imported in the mTestColumnTitles
                    Else
                        imVefCode = -1
                        For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            If (StrComp(UCase(Trim$(tgMVef(ilLoop).sCodeStn)), UCase(Trim$(smFieldValues(STATIONCODEINDEX))), vbTextCompare) = 0) Then
                                imVefCode = tgMVef(ilLoop).iCode
                                Exit For
                            End If
                        Next ilLoop
                        If imVefCode <> -1 Then
                            ilDpfMaxM = -1
                            ilDpfMaxW = -1
                            ReDim tmPlusDpf(0 To 0) As DPF
                            For ilLoop = 1 To 18 Step 1
                                tmDrf.lDemo(ilLoop - 1) = 0
                            Next ilLoop
                            For ilField = DEMOSTARTINDEX To UBound(smFieldValues) Step 1
                                If tgSpf.sSAudData = "H" Then
                                    llDemo = (Val(smFieldValues(ilField)) + 50) \ 100
                                ElseIf tgSpf.sSAudData = "N" Then
                                    llDemo = (Val(smFieldValues(ilField)) + 5) \ 10
                                ElseIf tgSpf.sSAudData = "U" Then
                                    llDemo = Val(smFieldValues(ilField))
                                Else
                                    llDemo = (Val(smFieldValues(ilField)) + 500) \ 1000
                                End If
                                If imDemoMap(ilField) >= 1 Then
                                    tmDrf.lDemo(imDemoMap(ilField) - 1) = llDemo
                                ElseIf imDemoMap(ilField) < 0 Then
                                    ilDpf = UBound(tmPlusDpf)
                                    tmPlusDpf(ilDpf).lDemo = llDemo
                                    tmPlusDpf(ilDpf).iMnfDemo = tmStdDemoInfo(-imDemoMap(ilField)).iMnfCode
                                    tmPlusDpf(ilDpf).lPop = lmPopValues(ilField)
                                    tmPlusDpf(ilDpf).lDrfCode = 0
                                    tmPlusDpf(ilDpf).iDnfCode = 0
                                    
                                    '6/17/14
                                    If ilField = imMaxPlusFieldM Then
                                        ilDpfMaxM = ilDpf
                                    ElseIf ilField = imMaxPlusFieldW Then
                                        ilDpfMaxW = ilDpf
                                    End If
                                    
                                    ReDim Preserve tmPlusDpf(0 To ilDpf + 1) As DPF
                                End If
                            Next ilField
                            'Write out DRF, then DPF if required
                            slBookName = Trim$(Trim$(slInBookName) & " " & Trim(smFieldValues(SOURCEINDEX)))
                            ilRet = mAddResearch(slBookName, slBookDate)
                            
                            '6/17/14
                            If (Not bm65PlusExist) Then
                                If smDataForm <> "8" Then
                                    ilAdjInc = 8
                                Else
                                    ilAdjInc = 9
                                End If
                                For ilAdj = 0 To ilAdjInc Step ilAdjInc
                                    ilStartIndex = -1
                                    If (ilAdj = 0) And (imMaxPlusValueM > 0) And (ilDpfMaxM <> -1) Then
                                        ilEndIndex = imMaxPlusIndexM - 1
                                        ilDpfMax = ilDpfMaxM
                                        ilStartIndex = 1
                                    ElseIf (ilAdj = ilAdjInc) And (imMaxPlusValueW > 0) And (ilDpfMaxW <> -1) Then
                                        ilEndIndex = imMaxPlusIndexW - 1
                                        ilDpfMax = ilDpfMaxW
                                        ilStartIndex = ilAdjInc + 1 '9
                                    End If
                                    If ilStartIndex <> -1 Then
                                        'Sum from 12+ to imMaxValue (55+), sum 18+to 55+,..., sum 50+ to 55+
                                        For ilSum1 = ilStartIndex To ilEndIndex Step 1
                                            llSum = 0
                                            llPop = 0
                                            For ilSum2 = ilSum1 To ilEndIndex Step 1
                                                llSum = llSum + tmDrf.lDemo(ilSum2 - 1)
                                                llPop = llPop + tmDrfPop.lDemo(ilSum2 - 1)
                                            Next ilSum2
                                            llSum = llSum + tmPlusDpf(ilDpfMax).lDemo
                                            llPop = llPop + tmPlusDpf(ilDpfMax).lPop
                                            'Write it out
                                            ilDpf = UBound(tmPlusDpf)
                                            tmPlusDpf(ilDpf).lDemo = llSum
                                            'Get Demo from ilSum1, use switch to get value, than find it in the standard
                                            If smDataForm <> "8" Then
                                                slDemoPlus = Trim(Str(Switch(ilSum1 = 1 + ilAdj, 12, ilSum1 = 2 + ilAdj, 18, ilSum1 = 3 + ilAdj, 25, ilSum1 = 4 + ilAdj, 35, ilSum1 = 5 + ilAdj, 45, ilSum1 = 6 + ilAdj, 50))) & "+"
                                            Else
                                                slDemoPlus = Trim(Str(Switch(ilSum1 = 1 + ilAdj, 12, ilSum1 = 2 + ilAdj, 18, ilSum1 = 3 + ilAdj, 21, ilSum1 = 4 + ilAdj, 25, ilSum1 = 5 + ilAdj, 35, ilSum1 = 6 + ilAdj, 45, ilSum1 = 7 + ilAdj, 50))) & "+"
                                            End If
                                            If ilAdj = 0 Then
                                                slDemoPlus = "M" & slDemoPlus
                                            Else
                                                slDemoPlus = "W" & slDemoPlus
                                            End If
                                            For ilStd = LBound(tmStdDemoInfo) To UBound(tmStdDemoInfo) Step 1
                                                If Trim$(tmStdDemoInfo(ilStd).sName) = slDemoPlus Then
                                                    tmPlusDpf(ilDpf).iMnfDemo = tmStdDemoInfo(ilStd).iMnfCode
                                                    Exit For
                                                End If
                                            Next ilStd
                                            tmPlusDpf(ilDpf).lPop = llPop
                                            tmPlusDpf(ilDpf).lDrfCode = 0
                                            tmPlusDpf(ilDpf).iDnfCode = 0
                                            ReDim Preserve tmPlusDpf(0 To ilDpf + 1) As DPF
                                            If ilSum1 = 1 Then
                                                ilDpf12M = ilDpf
                                            ElseIf ilSum1 = 9 Then
                                                ilDpf12W = ilDpf
                                            End If
                                        Next ilSum1
                                    End If
                                Next ilAdj
                                'Create P12+ and A18+, A21+; A35+,....
                                ilStartIndex = 1
                                ilEndIndex = imMaxPlusIndexM - 1
                                For ilSum1 = ilStartIndex To ilEndIndex Step 1
                                    llSum = tmPlusDpf(ilDpf12M + ilSum1 - 1).lDemo + tmPlusDpf(ilDpf12W + ilSum1 - 1).lDemo
                                    llPop = tmPlusDpf(ilDpf12M + ilSum1 - 1).lPop + tmPlusDpf(ilDpf12W + ilSum1 - 1).lPop
                                    ilDpf = UBound(tmPlusDpf)
                                    tmPlusDpf(ilDpf).lDemo = llSum
                                    'Get Demo from ilSum1, use switch to get value, than find it in the standard
                                    If smDataForm <> "8" Then
                                        slDemoPlus = Trim(Str(Switch(ilSum1 = 1, 12, ilSum1 = 2, 18, ilSum1 = 3, 25, ilSum1 = 4, 35, ilSum1 = 5, 45, ilSum1 = 6, 50))) & "+"
                                    Else
                                        slDemoPlus = Trim(Str(Switch(ilSum1 = 1, 12, ilSum1 = 2, 18, ilSum1 = 4, 21, ilSum1 = 4, 25, ilSum1 = 5, 35, ilSum1 = 6, 45, ilSum1 = 7, 50))) & "+"
                                    End If
                                    If ilSum1 = ilStartIndex Then
                                        slDemoPlus = "P" & slDemoPlus
                                    Else
                                        slDemoPlus = "A" & slDemoPlus
                                    End If
                                    For ilStd = LBound(tmStdDemoInfo) To UBound(tmStdDemoInfo) Step 1
                                        If Trim$(tmStdDemoInfo(ilStd).sName) = slDemoPlus Then
                                            tmPlusDpf(ilDpf).iMnfDemo = tmStdDemoInfo(ilStd).iMnfCode
                                            Exit For
                                        End If
                                    Next ilStd
                                    tmPlusDpf(ilDpf).lPop = llPop
                                    tmPlusDpf(ilDpf).lDrfCode = 0
                                    tmPlusDpf(ilDpf).iDnfCode = 0
                                    ReDim Preserve tmPlusDpf(0 To ilDpf + 1) As DPF
                                Next ilSum1
                            End If
                            For ilDpf = 0 To UBound(tmPlusDpf) - 1 Step 1
                                tmDpf = tmPlusDpf(ilDpf)
                                tmDpf.lDrfCode = tmDrf.lCode
                                tmDpf.iDnfCode = tmDrf.iDnfCode
                                ilRet = mAddDpf()
                            Next ilDpf
                            ilRet = mUpdateVehicle()
                        Else
                            If Trim$(smFieldValues(2)) <> "" Then
                                Print #hmTo, "Unable to find Matching Vehicle Station Code " & Trim$(smFieldValues(1)) & " " & Trim$(smFieldValues(2))
                                lbcErrors.AddItem "Unable to find Matching Vehicle Station Code " & Trim$(smFieldValues(1)) & " " & Trim$(smFieldValues(2))
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
    mConvFile = True
    MousePointer = vbDefault
    Exit Function
'mConvFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function


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
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptVeh
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    hmVef = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptVeh
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmDrf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptVeh
    On Error GoTo 0
    imDrfRecLen = Len(tmDrf)
    hmDnf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptVeh
    On Error GoTo 0
    imDnfRecLen = Len(tmDnf)
    hmDpf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptVeh
    On Error GoTo 0
    imDpfRecLen = Len(tmDpf)
    'Populate arrays to determine if records exist
    ilRet = mAddStdDemo()
    ilRet = mObtainBookName()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Book Name Error", vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
        imTerminate = True
        Exit Sub
    End If
    mBuildDemoArray
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
    gCenterStdAlone ImptVeh
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
    'gInitStdAlone ImptVeh, slStr, ilTestSystem
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
    
    Screen.MousePointer = vbDefault
    If bmResearchSaved Then
        If (Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES Then
            ilRet = MsgBox("Please update the Vehicle default ACT1 Lineup codes if required", vbOKOnly + vbInformation, "Warning")
        End If
    End If
    
    sgDoneMsg = ""
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ImptVeh
    igManUnload = NO
End Sub
Private Function mTestColumnTitle(slFromName As String, blChkSourceExist As Boolean) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilAvgQH As Integer
    Dim ilPop As Integer
    Dim ilAdj As Integer
    Dim ilDemo As Integer
    Dim slGender As String
    Dim slDemoAge As String
    Dim ilLoop As Integer
    Dim blTitleFound As Boolean
    Dim blPopFound As Boolean
    Dim llPop As Long
    Dim ilField As Integer
    Dim slStr As String
    
    mTestColumnTitle = False
    bm65PlusExist = False
    imMaxPlusValueM = 0
    imMaxPlusValueW = 0
    blTitleFound = False
    blPopFound = False
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
                'Header Record (record 1)
                '   Source,Code,Podcast,M12-17,M18-24,M25-34,M35-44,M45-49,M50-54,M55+,W12-17,W18-24,w25-34,w35-44,w45-49,w50-54,w55+
                'Population record (record 2, used with all imported books that follow)
                '   ,,Population,12877000,16173000,21015000,19968000,10620000,10858000,38769000,12298000,15263000,20592000,20167000,10848000,11273000,45872000
                'Demo Record (record 3 thru n)
                '   BookName1,DP,Dan Patrick Show,98,3539,12191,32344,39423,7275,3638,0,0,0,0,0,0,0
                '   BookName1,Drew , Dr.Drew, 0, 2380, 15503, 7752, 4182, 2380, 1802, 0, 3041, 19811, 9906, 5344, 3041, 2303
                '   BookName2,DP,Dan Patrick Show,98,3539,12191,32344,39423,7275,3638,0,0,0,0,0,0,0
                '   BookName2,Bill , Bill Crystal, 0, 2380, 15503, 7752, 4182, 2380, 1802, 0, 3041, 19811, 9906, 5344, 3041, 2303
                If (StrComp(UCase(Trim$(smFieldValues(SOURCEINDEX))), "SOURCE", 1) = 0) Then
                    If (StrComp(UCase(Trim$(smFieldValues(STATIONCODEINDEX))), "CODE", 1) = 0) Then
                        If (StrComp(UCase(Trim$(smFieldValues(VEHICLEINDEX))), "VEHICLE", 1) = 0) Then
                            blTitleFound = True
                            'Test Demo names
                            For ilDemo = DEMOSTARTINDEX To UBound(smFieldValues) Step 1
                                If Trim$(smFieldValues(ilDemo)) <> "" Then
                                    slGender = UCase(Left(smFieldValues(ilDemo), 1))
                                    slDemoAge = Mid(smFieldValues(ilDemo), 2)
                                    If (slGender = "M") Then
                                        ilAdj = 0
                                    ElseIf (slGender = "W") Then
                                        If smDataForm <> "8" Then
                                            ilAdj = 8
                                        Else
                                            ilAdj = 9
                                        End If
                                    Else
                                        MsgBox "Column Gender Missing from demo names" & " " & smFieldValues(ilDemo), vbExclamation, "Name Error"
                                        mTestColumnTitle = False
                                        Exit Function
                                    End If
                                    If smDataForm <> "8" Then
                                        Select Case slDemoAge
                                            Case "12-17"
                                                imDemoMap(ilDemo) = 1 + ilAdj
                                            Case "18-24"
                                                imDemoMap(ilDemo) = 2 + ilAdj
                                            Case "25-34"
                                                imDemoMap(ilDemo) = 3 + ilAdj
                                            Case "35-44"
                                                imDemoMap(ilDemo) = 4 + ilAdj
                                            Case "45-49"
                                                imDemoMap(ilDemo) = 5 + ilAdj
                                            Case "50-54"
                                                imDemoMap(ilDemo) = 6 + ilAdj
                                            Case "55-64"
                                                imDemoMap(ilDemo) = 7 + ilAdj
                                            Case "65+"
                                                imDemoMap(ilDemo) = 8 + ilAdj
                                                bm65PlusExist = True
                                            Case Else
                                                imDemoMap(ilDemo) = 0
                                                For ilLoop = 0 To UBound(tmStdDemoInfo) - 1 Step 1
                                                    If UCase(Trim(smFieldValues(ilDemo))) = UCase(Trim(tmStdDemoInfo(ilLoop).sName)) Then
                                                        imDemoMap(ilDemo) = -ilLoop
                                                        slStr = Mid(smFieldValues(ilDemo), 2, 2)
                                                        If ilAdj = 0 Then
                                                            If (Val(slStr) > imMaxPlusValueM) And (InStr(1, smFieldValues(ilDemo), "+", vbBinaryCompare) > 0) Then
                                                                imMaxPlusValueM = Val(slStr)
                                                                imMaxPlusFieldM = ilDemo
                                                                imMaxPlusIndexM = Switch(slStr = "12", 1, slStr = "18", 2, slStr = "25", 3, slStr = "35", 4, slStr = "45", 5, slStr = "50", 6, slStr = "55", 7)
                                                            End If
                                                        Else
                                                            If (Val(slStr) > imMaxPlusValueW) And (InStr(1, smFieldValues(ilDemo), "+", vbBinaryCompare) > 0) Then
                                                                imMaxPlusValueW = Val(slStr)
                                                                imMaxPlusFieldW = ilDemo
                                                                imMaxPlusIndexW = Switch(slStr = "12", 1, slStr = "18", 2, slStr = "25", 3, slStr = "35", 4, slStr = "45", 5, slStr = "50", 6, slStr = "55", 7) + ilAdj
                                                            End If
                                                        End If
                                                        Exit For
                                                    End If
                                                Next ilLoop
                                                If imDemoMap(ilDemo) = 0 Then
                                                    MsgBox "Column Demo Name Missing from standard demo names" & " " & smFieldValues(ilDemo), vbExclamation, "Name Error"
                                                    mTestColumnTitle = False
                                                    Exit Function
                                                End If
                                        End Select
                                    Else
                                        Select Case slDemoAge
                                            Case "12-17"
                                                imDemoMap(ilDemo) = 1 + ilAdj
                                            Case "18-20"
                                                imDemoMap(ilDemo) = 2 + ilAdj
                                            Case "21-24"
                                                imDemoMap(ilDemo) = 3 + ilAdj
                                            Case "25-34"
                                                imDemoMap(ilDemo) = 4 + ilAdj
                                            Case "35-44"
                                                imDemoMap(ilDemo) = 5 + ilAdj
                                            Case "45-49"
                                                imDemoMap(ilDemo) = 6 + ilAdj
                                            Case "50-54"
                                                imDemoMap(ilDemo) = 7 + ilAdj
                                            Case "55-64"
                                                imDemoMap(ilDemo) = 8 + ilAdj
                                            Case "65+"
                                                imDemoMap(ilDemo) = 9 + ilAdj
                                                bm65PlusExist = True
                                            Case Else
                                                imDemoMap(ilDemo) = 0
                                                For ilLoop = 0 To UBound(tmStdDemoInfo) - 1 Step 1
                                                    If UCase(Trim(smFieldValues(ilDemo))) = UCase(Trim(tmStdDemoInfo(ilLoop).sName)) Then
                                                        imDemoMap(ilDemo) = -ilLoop
                                                        slStr = Mid(smFieldValues(ilDemo), 2, 2)
                                                        If ilAdj = 0 Then
                                                            If (Val(slStr) > imMaxPlusValueM) And (InStr(1, smFieldValues(ilDemo), "+", vbBinaryCompare) > 0) Then
                                                                imMaxPlusValueM = Val(slStr)
                                                                imMaxPlusFieldM = ilDemo
                                                                imMaxPlusIndexM = Switch(slStr = "12", 1, slStr = "18", 2, slStr = "21", 3, slStr = "25", 4, slStr = "35", 5, slStr = "45", 6, slStr = "50", 7, slStr = "55", 8)
                                                            End If
                                                        Else
                                                            If (Val(slStr) > imMaxPlusValueW) And (InStr(1, smFieldValues(ilDemo), "+", vbBinaryCompare) > 0) Then
                                                                imMaxPlusValueW = Val(slStr)
                                                                imMaxPlusFieldW = ilDemo
                                                                imMaxPlusIndexW = Switch(slStr = "12", 1, slStr = "18", 2, slStr = "21", 3, slStr = "25", 4, slStr = "35", 5, slStr = "45", 6, slStr = "50", 7, slStr = "55", 8) + ilAdj
                                                            End If
                                                        End If
                                                        Exit For
                                                    End If
                                                Next ilLoop
                                                If imDemoMap(ilDemo) = 0 Then
                                                    MsgBox "Column Demo Name Missing from standard demo names" & " " & smFieldValues(ilDemo), vbExclamation, "Name Error"
                                                    mTestColumnTitle = False
                                                    Exit Function
                                                End If
                                        End Select
                                    End If
                                End If
                            Next ilDemo
                        End If
                    End If
                Else
                    If blTitleFound Then
                        If Not blPopFound Then
                            If UCase(Trim$(smFieldValues(VEHICLEINDEX))) = "POPULATION" Then
                                For ilField = DEMOSTARTINDEX To UBound(smFieldValues) Step 1
                                    If Trim$(smFieldValues(ilField)) <> "" Then
                                        If tgSpf.sSAudData = "H" Then
                                            llPop = (Val(smFieldValues(ilField)) + 50) \ 100
                                        ElseIf tgSpf.sSAudData = "N" Then
                                            llPop = (Val(smFieldValues(ilField)) + 5) \ 10
                                        ElseIf tgSpf.sSAudData = "U" Then
                                            llPop = Val(smFieldValues(ilField))
                                        Else
                                            llPop = (Val(smFieldValues(ilField)) + 500) \ 1000
                                        End If
                                        lmPopValues(ilField) = llPop
                                    End If
                                Next ilField
                                If Not blChkSourceExist Then
                                    mTestColumnTitle = True
                                    Exit Function
                                End If
                                blPopFound = True
                            End If
                        Else
                            If Trim$(smFieldValues(SOURCEINDEX)) = "" Then
                                MsgBox "Book Name must be defined on the Import screen or in the Source column on the Audience data records", vbExclamation, "Name Error"
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Loop Until ilEof
    If Not blTitleFound Then
        MsgBox "Column Titles record (Source, Code, Vehicle, Demo Names) Missing from " & slFromName, vbExclamation, "Name Error"
    ElseIf Not blPopFound Then
        MsgBox "Population record (,,Population, Population Demo Data) Missing from " & slFromName, vbExclamation, "Name Error"
    Else
        mTestColumnTitle = True
    End If
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

Private Sub plcDefault_Paint()
    plcDefault.CurrentX = 0
    plcDefault.CurrentY = 0
    plcDefault.Print "Set as Vehicle Default"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Import Vehicle Data"
End Sub

Private Sub mBuildDemoArray()
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim slName As String
    Dim slCode As String
    
    ReDim tmStdDemoInfo(0 To 0) As STDDEMOINFO
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
    smNameCodeTag = ""
    ilRet = gIMoveListBox(ImptVeh, lbcDemo, tmNameCode(), smNameCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
    For ilIndex = LBound(tmNameCode) To UBound(tmNameCode) - 1 Step 1
        ilRet = gParseItem(tmNameCode(ilIndex).sKey, 1, "\", slName)    'Get application name
        ilRet = gParseItem(tmNameCode(ilIndex).sKey, 2, "\", slCode)    'Get application name
        tmStdDemoInfo(UBound(tmStdDemoInfo)).sName = slName
        tmStdDemoInfo(UBound(tmStdDemoInfo)).iMnfCode = Val(slCode)
        ReDim Preserve tmStdDemoInfo(0 To UBound(tmStdDemoInfo) + 1) As STDDEMOINFO
    Next ilIndex
End Sub
Private Function mAddDpf() As Integer
    Dim ilRet As Integer
    Dim ilUpdate As Integer
    Dim llDpfCode As Long
    Dim llDpf As Long
    Dim tlDpf As DPF

    ilUpdate = False
    llDpfCode = -1
'    If UBound(tmPrevDpf) > LBound(tmPrevDpf) Then
'        For llDpf = LBound(tmPrevDpf) To UBound(tmPrevDpf) - 1 Step 1
'            If (tmDpf.lDrfCode = tmPrevDpf(llDpf).lDrfCode) And (tmDpf.iMnfDemo = tmPrevDpf(llDpf).iMnfDemo) Then
'                llDpfCode = tmPrevDpf(llDpf).lCode
'                ilUpdate = True
'                Exit For
'            End If
'        Next llDpf
'    End If
    
    tmDpfSrchKey1.lDrfCode = tmDpf.lDrfCode
    tmDpfSrchKey1.iMnfDemo = tmDpf.iMnfDemo
    ilRet = btrGetEqual(hmDpf, tlDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        llDpfCode = tlDpf.lCode
        ilUpdate = True
    End If
    If Not ilUpdate Then
        tmDpf.lCode = 0
        ilRet = btrInsert(hmDpf, tmDpf, imDpfRecLen, INDEXKEY0)
    Else
        Do
            tmDpfSrchKey.lCode = llDpfCode
            ilRet = btrGetEqual(hmDpf, tlDpf, imDpfRecLen, tmDpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tmDpf.lCode = llDpfCode
                ilRet = btrUpdate(hmDpf, tmDpf, imDpfRecLen)
            Else
                Exit Do
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
    End If
    mAddDpf = ilRet
End Function

Private Function mAddResearch(slBookName As String, slBookDate As String) As Integer
    Dim ilRet As Integer
    Dim ilPrevDnfCode As Integer
    Dim llPrevDrfCode As Long
    Dim ilField As Integer
    Dim ilVef As Integer
    Dim tlDrf As DRF
    
    ilRet = mBookNameUsed(slBookName, slBookDate, imVefCode, ilPrevDnfCode, llPrevDrfCode)
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
            mAddResearch = False
            Exit Function
        End If
        Do
            tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmDnf.iAutoCode = tmDnf.iCode
            gPackDate smSyncDate, tmDnf.iSyncDate(0), tmDnf.iSyncDate(1)
            gPackTime smSyncTime, tmDnf.iSyncTime(0), tmDnf.iSyncTime(1)
            ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        
        tgDnfBook(UBound(tgDnfBook)) = tmDnf
        'ReDim Preserve tgDnfBook(1 To UBound(tgDnfBook) + 1) As DNF
        ReDim Preserve tgDnfBook(0 To UBound(tgDnfBook) + 1) As DNF

        ilPrevDnfCode = tmDnf.iCode
        tmDrfPop.lCode = 0
        tmDrfPop.iDnfCode = ilPrevDnfCode
        tmDrfPop.sForm = smDataForm
        For ilField = 3 To UBound(imDemoMap) Step 1
            If imDemoMap(ilField) > 0 Then
                tmDrfPop.lDemo(imDemoMap(ilField) - 1) = lmPopValues(ilField)
            End If
        Next ilField
        tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
        tmDrfPop.lAutoCode = tmDrfPop.lCode
        ilRet = btrInsert(hmDrf, tmDrfPop, imDrfRecLen, INDEXKEY0)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            Print #hmTo, "Error when Adding Population File (DRF)" & Str$(ilRet)
            lbcErrors.AddItem "Error Adding DRF" & Str$(ilRet)
            mAddResearch = False
            Exit Function
        End If
        Do
            tmDrfPop.iRemoteID = tgUrf(0).iRemoteUserID
            tmDrfPop.lAutoCode = tmDrfPop.lCode
            gPackDate smSyncDate, tmDrfPop.iSyncDate(0), tmDrfPop.iSyncDate(1)
            gPackTime smSyncTime, tmDrfPop.iSyncTime(0), tmDrfPop.iSyncTime(1)
            ilRet = btrUpdate(hmDrf, tmDrfPop, imDrfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
    Else
        If ilRet = 3 Then
            Screen.MousePointer = vbDefault
            ilVef = gBinarySearchVef(imVefCode)
            If ilVef <> -1 Then
                ilRet = MsgBox(slBookName & " for " & Trim$(tgMVef(ilVef).sName) & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
            Else
                ilRet = MsgBox(slBookName & " for VefCode" & Trim$(Str(imVefCode)) & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
            End If
            If ilRet = vbNo Then
                mAddResearch = False
                Exit Function
            End If
            Print #hmTo, "Replaced Demo Data (DRF)"
        End If
    End If
    tmDrf.iDnfCode = ilPrevDnfCode
    If llPrevDrfCode = 0 Then
        tmDrf.lCode = 0
        tmDrf.sForm = smDataForm
        tmDrf.iVefCode = imVefCode
        tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
        tmDrf.lAutoCode = tmDrf.lCode
        ilRet = btrInsert(hmDrf, tmDrf, imDrfRecLen, INDEXKEY2)
        If ilRet <> BTRV_ERR_NONE Then
            If ilRet >= 30000 Then
                ilRet = csiHandleValue(0, 7)
            End If
            Print #hmTo, "Error when Adding Demo Data File (DRF)" & Str$(ilRet)
            lbcErrors.AddItem "Error Adding DRF" & Str$(ilRet)
            mAddResearch = False
            Exit Function
        End If
        Do
            tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
            tmDrf.lAutoCode = tmDrf.lCode
            gPackDate smSyncDate, tmDrf.iSyncDate(0), tmDrf.iSyncDate(1)
            gPackTime smSyncTime, tmDrf.iSyncTime(0), tmDrf.iSyncTime(1)
            ilRet = btrUpdate(hmDrf, tmDrf, imDrfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
    Else
        Do
            tmDrfSrchKey2.lCode = llPrevDrfCode
            ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                Exit Do
            End If
            tmDrf.lCode = tlDrf.lCode
            tmDrf.sForm = smDataForm
            tmDrf.iVefCode = imVefCode
            tmDrf.lAutoCode = tlDrf.lAutoCode
            tmDrf.iRemoteID = tlDrf.iRemoteID
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
        mAddResearch = False
        Exit Function
    End If
    mAddResearch = True
    Exit Function
End Function

Private Function mUpdateVehicle() As Integer
    Dim ilRet As Integer
    
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
            'If ckcDefault(1).Value Then
            '    tmVef.iReallDnfCode = tmDnf.iCode
            'End If
            ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        ilRet = gBinarySearchVef(tmVef.iCode)
        If ilRet <> -1 Then
            tgMVef(ilRet) = tmVef
        End If
        '11/26/17
        gFileChgdUpdate "vef.btr", False
        
    End If
    mUpdateVehicle = True
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
