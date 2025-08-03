VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImptRCPrice 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4770
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
   ScaleHeight     =   4770
   ScaleWidth      =   7755
   Begin VB.CommandButton cmcBrowse 
      Caption         =   ".."
      Height          =   285
      Left            =   5560
      TabIndex        =   12
      Top             =   1260
      Width           =   375
   End
   Begin VB.CheckBox ckcResetStdRate 
      Caption         =   "Recompute Package Rates from Hidden"
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   1725
      Width           =   3960
   End
   Begin VB.ListBox lbcInv 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "ImptRCPrice.frx":0000
      Left            =   6690
      List            =   "ImptRCPrice.frx":0002
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4245
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   195
      Left            =   1740
      TabIndex        =   5
      Top             =   2100
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   344
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
      Height          =   1710
      Left            =   900
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2445
      Visible         =   0   'False
      Width           =   5820
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
      Top             =   1260
      Width           =   1605
   End
   Begin VB.PictureBox plcFrom 
      Height          =   375
      Left            =   915
      ScaleHeight     =   315
      ScaleWidth      =   4485
      TabIndex        =   2
      Top             =   1215
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
   Begin VB.CommandButton cmcImport 
      Appearance      =   0  'Flat
      Caption         =   "Import Rates"
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
      TabIndex        =   7
      Top             =   4305
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
      TabIndex        =   8
      Top             =   4305
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   240
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Label lacMsg 
      Appearance      =   0  'Flat
      Caption         =   $"ImptRCPrice.frx":0004
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   75
      TabIndex        =   10
      Top             =   300
      Width           =   7515
   End
   Begin VB.Label lacScreen 
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   4605
   End
   Begin VB.Label lbcFrom 
      Appearance      =   0  'Flat
      Caption         =   "From File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      TabIndex        =   1
      Top             =   1320
      Width           =   810
   End
End
Attribute VB_Name = "ImptRCPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImptRCPrice.Frm
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
Dim imTerminate As Integer
Dim imConverting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smNowDate As String
Dim lmNowDate As Long
Dim smSyncDate As String
Dim smSyncTime As String
Dim smDataForm As String
Dim smFieldValues(0 To 4) As String






Private Sub cmcBrowse_Click()

    'TTP 10339 - Automation Import - use windows File Browser (to see Local Drives and mapped drives from RDP)
    CMDialogBox.DialogTitle = "Import From CSV File"
    CMDialogBox.Filter = "CSV (*.csv)|*.csv|Blank (*.)|*.|ASC (*.asc)|*.Asc|Text (*.txt)|*.Txt|All (*.*)|*.*"
    CMDialogBox.InitDir = sgImportPath
    CMDialogBox.DefaultExt = "CSV (*.csv)"
    CMDialogBox.Action = 1 'Open dialog
    edcFrom.Text = CMDialogBox.fileName
    
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
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
Private Sub cmcImport_Click()
    Dim ilRet As Integer
    Dim ilPos As Integer
    Dim slLine As String
    Dim slChar As String
    Dim slFromName As String

    ilRet = MsgBox("Vehicle/Daypart Rates will be updated from the Input file, OK to Continue?", vbYesNo + vbQuestion, "Rate Card")
    If ilRet = vbNo Then
        Exit Sub
    End If
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
    'Open sgDBPath & "Messages\" & "ImptRCPrice.Txt" For Output As hmTo
    ilRet = gFileOpen(sgDBPath & "Messages\" & "ImptRCRate.Txt", "Output", hmTo)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Open " & sgDBPath & "Messages\" & "ImptRCRate.Txt" & " Error #" & str(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    gGetSyncDateTime smSyncDate, smSyncTime
    plcGauge.Value = 0
    lmPercent = 0
    Print #hmTo, "Import Rate Card Rates on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    Print #hmTo, ""
    imConverting = True
    Print #hmTo, "** Processing: " & slFromName & " **"
    If Not mImportRateCardPrice() Then
        Close hmFrom
        Print #hmTo, "Import Rate Card Rates terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        Close hmTo
        imConverting = False
        mTerminate
        Exit Sub
    End If
    Print #hmTo, "Import Rate Card Rates successfully completed on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
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
    

    Set ImptRCPrice = Nothing   'Remove data segment
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
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)

    gCenterStdAlone ImptRCPrice
    lacScreen.Caption = "Import Rate Card Rates for " & Trim$(tgRcfI.sName)
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
    Unload ImptRCPrice
    igManUnload = NO
End Sub


Private Function mImportRateCardPrice() As Integer
    Dim slLine As String
    Dim ilRet As Integer
    Dim slChar As String
    Dim ilEof As Integer
    Dim llPercent As Long
    Dim ilLoop As Integer
    Dim slVehicle As String
    Dim slDayPart As String
    Dim slDollars As String
    
    slLine = ""
    err.Clear
    Do
        ilRet = 0
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
                    mImportRateCardPrice = False
                    Exit Function
                End If
                gParseCDFields slLine, False, smFieldValues()
                'Vehicle Name, Daypart, Dollar
                '
                For ilLoop = LBound(smFieldValues) To UBound(smFieldValues) Step 1
                    smFieldValues(ilLoop) = gRemoveIllegalChars(smFieldValues(ilLoop))
                Next ilLoop
                slVehicle = UCase$(Trim$(smFieldValues(0)))
                slDayPart = UCase$(Trim$(smFieldValues(1)))
                slDollars = Trim$(smFieldValues(2))
                If slVehicle <> "VEHICLE" Then
                    ilRet = RateCard.mSetImportPrice("D", slVehicle, slDayPart, slDollars)
                    If ilRet = 1 Then
                        Print #hmTo, "Vehicle " & Trim$(smFieldValues(0)) & " not found. Line " & slLine & " bypassed"
                        lbcErrors.AddItem "Vehicle " & Trim$(smFieldValues(0)) & " not found. Line " & slLine & " bypassed"
                    ElseIf ilRet = 2 Then
                        Print #hmTo, "Daypart " & Trim$(smFieldValues(1)) & " not found. Line " & slLine & " bypassed"
                        lbcErrors.AddItem "Daypart " & Trim$(smFieldValues(1)) & " not found. Line " & slLine & " bypassed"
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
    If ckcResetStdRate.Value = vbChecked Then
        ilRet = RateCard.mSetImportPrice("R", "", "", "")
    End If
    ilRet = RateCard.mSetImportPrice("C", "", "", "")
    mImportRateCardPrice = True
    Exit Function


End Function
