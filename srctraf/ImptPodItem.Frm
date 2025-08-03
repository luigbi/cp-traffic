VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImptPodItem 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1500
   ClientLeft      =   795
   ClientTop       =   1485
   ClientWidth     =   7755
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
   ScaleHeight     =   1500
   ScaleWidth      =   7755
   Begin VB.CommandButton cmcBrowse 
      Caption         =   ".."
      Height          =   285
      Left            =   5580
      TabIndex        =   11
      Top             =   375
      Width           =   375
   End
   Begin VB.ListBox lbcInv 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "ImptPodItem.frx":0000
      Left            =   6780
      List            =   "ImptPodItem.frx":0002
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lbcAdvtProd 
      Height          =   840
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   210
      Left            =   1500
      TabIndex        =   6
      Top             =   3075
      Visible         =   0   'False
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
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
      Left            =   3360
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   2820
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
      Left            =   6030
      TabIndex        =   4
      Top             =   375
      Width           =   1605
   End
   Begin VB.PictureBox plcFrom 
      Height          =   375
      Left            =   915
      ScaleHeight     =   315
      ScaleWidth      =   4485
      TabIndex        =   2
      Top             =   345
      Width           =   4545
      Begin VB.TextBox edcFrom 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   4425
      End
   End
   Begin VB.CommandButton cmcFileConv 
      Appearance      =   0  'Flat
      Caption         =   "Import"
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
      Left            =   2100
      TabIndex        =   8
      Top             =   1010
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
      Left            =   4245
      TabIndex        =   9
      Top             =   1010
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Label lacScreen 
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   7260
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   5220
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lbcFrom 
      Appearance      =   0  'Flat
      Caption         =   "Import File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      TabIndex        =   1
      Top             =   435
      Width           =   810
   End
End
Attribute VB_Name = "ImptPodItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImptPodItem.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the import contract conversion input screen code
Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim lmPercent As Long
Dim imShowMsg As Integer
Dim lmLen As Long
Dim hmFrom As Integer   'From file hanle
Dim hmTo As Integer   'From file hanle
Dim hmCif As Integer    'file handle
Dim tmCif As CIF
Dim imCifRecLen As Integer  'Record length
Dim tmCifSrchKey As LONGKEY0
Dim imCifIndex As Integer
Dim hmCpf As Integer    'file handle
Dim tmCpf As CPF
Dim imCpfRecLen As Integer  'Record length
Dim tmCpfSrchKey0 As LONGKEY0
Dim tmCpfSrchKey1 As CPFKEY1
Dim imTerminate As Integer
Dim imConverting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smNowDate As String
Dim lmNowDate As Long
Dim smSyncDate As String
Dim smSyncTime As String
Dim smDataForm As String
Dim imSortCart As Integer
Dim smFieldValues(0 To 9) As String

Dim tmInvNameCode() As SORTCODE
Dim smInvNameCodeTag As String

Dim tmNameCode() As SORTCODE
Dim smNameCodeTag As String





Private Sub cmcBrowse_Click()

End Sub

' MsgBox parameters
'Const vbOkOnly = 0                 ' OK button only
'Const vbCritical = 16          ' Critical message
'Const vbApplicationModal = 0
'Const INDEXKEY0 = 0
Private Sub cmcCancel_Click()
    If imConverting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcFileConv_Click()
    Exit Sub
    Dim slFromName As String
    Dim slBookName As String
    Dim slBookDate As String
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim slPopulationName As String
    Dim slTimeName As String
    Dim slDayPartName As String
    Dim slGroupName As String
    Dim slTimeNameDay As String
    Dim ilLoop As Integer
    Dim ilPrevDnfCode As Integer
    Dim slLine As String
    Dim slChar As String

    lbcErrors.Clear
    lbcErrors.Visible = True
    imShowMsg = True
    lmLen = 0
    slFromName = Trim$(edcFrom.Text)
    If slFromName = "" Then
        MsgBox "From Name Must be Defined", vbExclamation, "Name Error"
        edcFrom.SetFocus
        Exit Sub
    End If
    'Check file names
    If (InStr(slFromName, ":") = 0) And (Left$(slFromName, 2) <> "\\") Then
        slFromName = sgImportPath & slFromName
    End If
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

    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    ilRet = 0
    'hmTo = FreeFile
    'Open sgDBPath & "Messages\" & "ImptPodItem.Txt" For Output As hmTo
    ilRet = gFileOpen(sgDBPath & "Messages\" & "ImptPodItem.Txt", "Output", hmTo)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Open " & sgDBPath & "Messages\" & "ImptPodItem.Txt" & " Error #" & str(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mInvPop
    gGetSyncDateTime smSyncDate, smSyncTime
    plcGauge.Value = 0
    lmPercent = 0
    Print #hmTo, "Import Pod Items on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    Print #hmTo, ""
    imConverting = True
    Print #hmTo, "** Processing: " & slFromName & " **"
    If Not mConvCopyInventory() Then
        Close hmFrom
        Print #hmTo, "Import Pod Items terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        Close hmTo
        imConverting = False
        mTerminate
        Exit Sub
    End If
    Print #hmTo, "Import Pod Items successfully completed on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    Close hmTo
    plcGauge.Value = 100
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    imConverting = False
    Screen.MousePointer = vbDefault
    Exit Sub
cmcFileConvErr:
    ilRet = err.Number
    Resume Next
End Sub
Private Sub cmcFrom_Click()
    igBrowserType = 1  'CSV
    Browser.Show vbModal
    If igBrowserReturn = 1 Then
        edcFrom.Text = Trim$(sgBrowserFile)
    End If
    DoEvents
    edcFrom.SetFocus
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
End Sub

Private Sub edcFrom_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus ActiveControl
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
    
    Erase tmInvNameCode
    Erase tmNameCode
    
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf

    Set ImptPodItem = Nothing   'Remove data segment
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
    Dim slLine As String
    Dim slChar As String
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slAdvtName As String
    Dim ilAdvtIndex As Integer
    Dim ilPos As Integer
    
    imTerminate = False
    imFirstActivate = True
    'mParseCmmdLine
    Screen.MousePointer = vbHourglass
    imConverting = False
    imFirstFocus = True
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0
    imSortCart = 0
    hmCif = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptPodItem
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmCpf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptPodItem
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)

    gCenterStdAlone ImptPodItem
    If mTestRecLengths() Then
        Screen.MousePointer = vbDefault
        imTerminate = True
        Exit Sub
    End If
    If igCopyInvAdfCode > 0 Then
        lacScreen.Caption = "Import Pod Items: "
        ilAdvtIndex = gBinarySearchAdf(igCopyInvAdfCode)
        If ilAdvtIndex <> -1 Then
            If (tgCommAdf(ilAdvtIndex).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdvtIndex).sAddrID) <> "") Then
                slAdvtName = Trim$(tgCommAdf(ilAdvtIndex).sName) & ", " & Trim$(tgCommAdf(ilAdvtIndex).sAddrID)
            Else
                slAdvtName = Trim$(tgCommAdf(ilAdvtIndex).sName)
            End If
            ilPos = InStr(slAdvtName, "&")
            If ilPos > 0 Then
                lacScreen.Caption = lacScreen.Caption & Left$(Trim$(slAdvtName), ilPos - 1) & "&&" & Mid$(Trim$(slAdvtName), ilPos + 1)
            Else
                lacScreen.Caption = lacScreen.Caption & Trim$(slAdvtName)
            End If
        Else
            lacScreen.Caption = "Import Pod Items"
        End If
    Else
        lacScreen.Caption = "Import Pod Items"
    End If
    lbcAdvtProd.Clear 'Force list box to be populated
    mProdPop
    If imTerminate Then
        Exit Sub
    End If
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
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ImptPodItem
    igManUnload = NO
End Sub
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
    ilSize = mGetRecLength("Cif.Btr")
    If ilSize <> Len(tmCif) Then
        If ilSize > 0 Then
            MsgBox "Cif size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmCif)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Cif error: " & str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    ilSize = mGetRecLength("Cpf.Btr")
    If ilSize <> Len(tmCpf) Then
        If ilSize > 0 Then
            MsgBox "Cpf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmCpf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Cpf error: " & str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
    mTestRecLengths = ilSizeError
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mConvPop                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert CHF                    *
'*                                                     *
'*******************************************************
Private Function mConvCopyInventory() As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim llPercent As Long
    Dim ilEof As Integer
    Dim slLen As String
    Dim slISCI As String
    Dim slCreativeTitle As String
    Dim slProduct As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLen As Integer
    Dim ilLoop As Integer
    Dim slChar As String
    Dim tlCif As CIF
    
    If lbcAdvtProd.ListIndex <= 0 Then
        slProduct = ""
    Else
        slProduct = Trim$(lbcAdvtProd.List(lbcAdvtProd.ListIndex))
    End If
    imCifIndex = 0
    slLine = ""
    err.Clear
    'On Error GoTo mConvCopyInventoryErr:
    Do
        ilRet = 0
        'On Error GoTo mConvCopyInventoryErr:
        ''Line Input #hmFrom, slLine
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
                    mConvCopyInventory = False
                    Exit Function
                End If
                gParseCDFields slLine, False, smFieldValues()
                'Determine field Type
                'Spot Length, ISCI Code, Creative Title, Type, Type Name, Type Number, Flight Start Date, Flight End Date, Rotation %
                'Where
                '  Type     Type Name         Type Number
                '  D        DMA Market name   DMA Market rank
                '  M        MSA Market name   MSA Market rank
                '  S        State (2 char)    Blank
                '  Z        Zone (2 char)     Blank
                '  F        Format            Blank
                '  A        Call Letters-Band Blank
                '
                '
                For ilLoop = LBound(smFieldValues) To UBound(smFieldValues) Step 1
                    smFieldValues(ilLoop) = gRemoveIllegalChars(smFieldValues(ilLoop))
                Next ilLoop
                slLen = Trim$(smFieldValues(0))
                slISCI = Trim$(smFieldValues(1))
                '12/22/11: Limit # of ISCI characters
                If ((Asc(tgSpf.sUsingFeatures9) And LIMITISCI) = LIMITISCI) Then
                    slISCI = Left(slISCI, 15)
                Else
                    slISCI = Left(slISCI, 20)
                End If
                slCreativeTitle = Trim$(smFieldValues(2))
                
                tmCpfSrchKey1.sISCI = slISCI
                ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                If (ilRet <> BTRV_ERR_NONE) Then
                    tmCpf.lCode = 0
                    tmCpf.sName = slProduct
                    tmCpf.sISCI = slISCI
                    tmCpf.sCreative = slCreativeTitle
                    tmCpf.iRotEndDate(0) = 0
                    tmCpf.iRotEndDate(1) = 0
                    ilRet = btrInsert(hmCpf, tmCpf, imCpfRecLen, INDEXKEY0)
                    If ilRet <> BTRV_ERR_NONE Then
                        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        Print #hmTo, "Unable to insert the Copy Inventory record, error " & ilRet
                        lbcErrors.AddItem "Unable to insert the Copy Inventory record, error " & ilRet
                        mConvCopyInventory = False
                        Exit Function
                    End If
                    If sgUseCartNo <> "N" Then
                        'Determine Purge cart number to be used
                        slNameCode = tmInvNameCode(imCifIndex).sKey   'lbcInvCode.List(imCifIndex)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        slCode = Trim$(slCode)
                        tmCifSrchKey.lCode = CLng(slCode)
                        ilRet = btrGetEqual(hmCif, tlCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                        If ilRet <> BTRV_ERR_NONE Then
                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            Print #hmTo, "Unable to Retrieve Copy Inventory record, error " & ilRet
                            lbcErrors.AddItem "Unable to Retrieve Copy Inventory record, error " & ilRet
                            mConvCopyInventory = False
                            Exit Function
                        End If
                        ilLen = tlCif.iLen
                        If tlCif.sPurged = "P" Then
                            mInitCif
                            tmCif.sName = tlCif.sName
                            tmCif.sCut = tlCif.sCut
                            tmCif.lcpfCode = tmCpf.lCode
                            tmCif.iLen = CInt(slLen)
                            tmCif.iNoTimesAir = tlCif.iNoTimesAir
                            If ilLen > 0 Then  'Used previously, therefore retain
                                tlCif.sPurged = "H"
                                gPackDate smNowDate, tlCif.iPurgeDate(0), tlCif.iPurgeDate(1)
                                ilRet = btrUpdate(hmCif, tlCif, imCifRecLen)
                                If ilRet <> BTRV_ERR_NONE Then
                                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                        ilRet = csiHandleValue(0, 7)
                                    End If
                                    Print #hmTo, "Unable to Update Copy Inventory record, error " & ilRet
                                    lbcErrors.AddItem "Unable to Update Copy Inventory record, error " & ilRet
                                    mConvCopyInventory = False
                                    Exit Function
                                End If
                                tmCif.lCode = 0
                                ilRet = btrInsert(hmCif, tmCif, imCifRecLen, INDEXKEY0)
                                If ilRet <> BTRV_ERR_NONE Then
                                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                        ilRet = csiHandleValue(0, 7)
                                    End If
                                    Print #hmTo, "Unable to Insert Copy Inventory record, error " & ilRet
                                    lbcErrors.AddItem "Unable to Insert Copy Inventory record, error " & ilRet
                                    mConvCopyInventory = False
                                    Exit Function
                                End If
                            Else
                                'Replace record
                                ilRet = btrDelete(hmCif)
                                If ilRet <> BTRV_ERR_NONE Then
                                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                        ilRet = csiHandleValue(0, 7)
                                    End If
                                    Print #hmTo, "Unable to Delete Copy Inventory record, error " & ilRet
                                    lbcErrors.AddItem "Unable to Update Copy Inventory record, error " & ilRet
                                    mConvCopyInventory = False
                                    Exit Function
                                End If
                                tmCif.lCode = tlCif.lCode
                                ilRet = btrInsert(hmCif, tmCif, imCifRecLen, INDEXKEY0)
                                If ilRet <> BTRV_ERR_NONE Then
                                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                        ilRet = csiHandleValue(0, 7)
                                    End If
                                    Print #hmTo, "Unable to Update Copy Inventory record, error " & ilRet
                                    lbcErrors.AddItem "Unable to Update Copy Inventory record, error " & ilRet
                                    mConvCopyInventory = False
                                    Exit Function
                                End If
                            End If
                            imCifIndex = imCifIndex + 1
                        Else
                            Print #hmTo, "No purge Copy Inventory records to be used"
                            lbcErrors.AddItem "No purge Copy Inventory records to be used"
                            mConvCopyInventory = False
                            Exit Function
                        End If
                    Else
                        mInitCif
                        tmCif.lcpfCode = tmCpf.lCode
                        tmCif.iLen = CInt(slLen)
                        ilRet = btrInsert(hmCif, tmCif, imCifRecLen, INDEXKEY0)
                        If ilRet <> BTRV_ERR_NONE Then
                            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            Print #hmTo, "Unable to Insert Copy Inventory record, error " & ilRet
                            lbcErrors.AddItem "Unable to Insert Copy Inventory record, error " & ilRet
                            mConvCopyInventory = False
                            Exit Function
                        End If
                    End If
                Else
                    Print #hmTo, slISCI & " ISCI previously defined"
                    lbcErrors.AddItem slISCI & " ISCI previously defined, Import record bypassed"
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
    mConvCopyInventory = True
    Exit Function
'mConvCopyInventoryErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mProdPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate advertiser product    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mProdPop()
'
'   mProdPop
'   Where:
'       igCopyInvAdfCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    If igCopyInvAdfCode <= 0 Then
        If lbcAdvtProd.ListCount <= 0 Then
            lbcAdvtProd.AddItem "[None]", 0  'Force as first item on list
        End If
        Exit Sub
    End If
    ilIndex = lbcAdvtProd.ListIndex
    If ilIndex > 0 Then
        slName = lbcAdvtProd.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtProdBox(CopyInv, igCopyInvAdfCode, lbcProd, lbcProdCode)
    If tgSpf.sUseProdSptScr = "P" Then
        ilRet = gPopShortTitleBox(CopyInv, igCopyInvAdfCode, lbcAdvtProd, tmNameCode(), smNameCodeTag)
    Else
        ilRet = gPopAdvtProdBox(CopyInv, igCopyInvAdfCode, lbcAdvtProd, tmNameCode(), smNameCodeTag)
    End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mProdPopErr
        gCPErrorMsg ilRet, "mProdPop (gPopAdvtProdBox)", CopyInv
        On Error GoTo 0
        lbcAdvtProd.AddItem "[None]", 0  'Force as first item on list
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcAdvtProd
            If gLastFound(lbcAdvtProd) > 0 Then
                lbcAdvtProd.ListIndex = gLastFound(lbcAdvtProd)
            Else
                lbcAdvtProd.ListIndex = -1
            End If
        Else
            lbcAdvtProd.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mProdPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

Private Sub mInitCif()
    tmCif.lCode = 0
    tmCif.iMcfCode = igCopyInvMcfCode
    tmCif.sName = ""
    tmCif.sCut = ""
    tmCif.sReel = ""
    tmCif.iLen = 0
    tmCif.iEtfCode = igCopyInvEtfCode
    tmCif.iEnfCode = igCopyInvEnfCode
    tmCif.iAdfCode = igCopyInvAdfCode 'From rotation input screen
    tmCif.lcpfCode = 0
    tmCif.iMnfComp(0) = 0
    tmCif.iMnfComp(1) = 0
    tmCif.iMnfAnn = 0
    tmCif.sHouse = "N"
    tmCif.sCleared = "N"
    tmCif.lCsfCode = 0
    tmCif.iNoTimesAir = 0
    tmCif.sCartDisp = "N"
    tmCif.sTapeDisp = "N"
    tmCif.sPurged = "A"
    gPackDate "", tmCif.iPurgeDate(0), tmCif.iPurgeDate(1)
    gPackDate smNowDate, tmCif.iDateEntrd(0), tmCif.iDateEntrd(1)
    gPackDate "", tmCif.iUsedDate(0), tmCif.iUsedDate(1)
    gPackDate "", tmCif.iRotStartDate(0), tmCif.iRotStartDate(1)
    gPackDate "", tmCif.iRotEndDate(0), tmCif.iRotEndDate(1)
    tmCif.iUrfCode = tgUrf(0).iCode
    tmCif.sPrint = "N"
    tmCif.iLangMnfCode = 0
    tmCif.sUnused = ""
End Sub

Private Sub mInvPop()
    Dim ilRet As Integer
    
    lbcInv.Clear
    smInvNameCodeTag = ""
    If igCopyInvMcfCode > 0 Then
        ilRet = gPopCopyForMediaBox(ImptPodItem, igCopyInvMcfCode, smNowDate, True, True, imSortCart, lbcInv, tmInvNameCode(), smInvNameCodeTag)    'lbcInvCode)
    End If
End Sub

