VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImptRad 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5700
   ClientLeft      =   795
   ClientTop       =   1485
   ClientWidth     =   7980
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
   ScaleWidth      =   7980
   Begin VB.CommandButton cmcBrowse 
      Caption         =   ".."
      Height          =   285
      Left            =   5640
      TabIndex        =   22
      Top             =   1620
      Width           =   375
   End
   Begin VB.ListBox lbcDemo 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "Imptrad.frx":0000
      Left            =   6975
      List            =   "Imptrad.frx":0002
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2265
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   210
      Left            =   2475
      TabIndex        =   20
      Top             =   3075
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox plcDefault 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      ScaleHeight     =   255
      ScaleWidth      =   6000
      TabIndex        =   10
      Top             =   2745
      Width           =   6000
      Begin VB.CheckBox ckcDefault 
         Caption         =   "Rating Book Name"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   2070
         TabIndex        =   11
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
         TabIndex        =   12
         Top             =   15
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.TextBox edcBookDate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1155
      MaxLength       =   10
      TabIndex        =   9
      Top             =   2400
      Width           =   1275
   End
   Begin VB.TextBox edcBookName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1155
      MaxLength       =   30
      TabIndex        =   7
      Top             =   2040
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
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7230
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4905
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7155
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4170
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7110
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4605
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
      Height          =   1710
      Left            =   1020
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3450
      Visible         =   0   'False
      Width           =   5340
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
      Left            =   6150
      TabIndex        =   5
      Top             =   1620
      Width           =   1725
   End
   Begin VB.PictureBox plcFrom 
      Height          =   375
      Left            =   1155
      ScaleHeight     =   315
      ScaleWidth      =   4365
      TabIndex        =   3
      Top             =   1590
      Width           =   4425
      Begin VB.TextBox edcFrom 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   4305
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
      Left            =   2100
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
      Left            =   4245
      TabIndex        =   13
      Top             =   5310
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   6480
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.Label lacBookDate 
      Appearance      =   0  'Flat
      Caption         =   "Book Date"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      TabIndex        =   8
      Top             =   2445
      Width           =   1095
   End
   Begin VB.Label lacBookName 
      Appearance      =   0  'Flat
      Caption         =   "Book Name"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      TabIndex        =   6
      Top             =   2085
      Width           =   1095
   End
   Begin VB.Label lacMsg 
      Appearance      =   0  'Flat
      Caption         =   $"Imptrad.frx":0004
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   120
      TabIndex        =   1
      Top             =   375
      Width           =   7500
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   5220
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacFileType 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   60
      TabIndex        =   15
      Top             =   3045
      Width           =   2190
   End
   Begin VB.Label lbcFrom 
      Appearance      =   0  'Flat
      Caption         =   "From File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      TabIndex        =   2
      Top             =   1695
      Width           =   810
   End
End
Attribute VB_Name = "ImptRad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Imptrad.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ImptRad.Frm
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
Dim imTestAddStdDemo As Integer
Dim lmPercent As Long
Dim imShowMsg As Integer
Dim lmLen As Long
Dim hmFrom As Integer   'From file hanle
Dim hmTo As Integer   'From file hanle
Dim hmDnf As Integer    'file handle
Dim tmDnf As DNF
Dim imDnfRecLen As Integer  'Record length
Dim tmDnfSrchKey As INTKEY0
Dim hmDrf As Integer    'file handle
Dim tmDrf As DRF
Dim imDrfRecLen As Integer  'Record length
Dim tmDrfSrchKey As DRFKEY0
Dim hmRdf As Integer    'file handle
Dim tmRdf As RDF
Dim imRdfRecLen As Integer  'Record length
Dim hmVef As Integer    'file handle
Dim tmVef As VEF
Dim imVefRecLen As Integer  'Record length
Dim tmVefSrchKey As INTKEY0
Dim imVefCodeInDnf() As Integer   'Array if vehicles code which are vehicles with data in dnf
Dim smVehNotFound() As String
Dim hmMnf As Integer    'file handle
Dim tmMnf As MNF        'Record structure
Dim imMnfRecLen As Integer  'Record length
'Dim hmDsf As Integer 'Delete Stamp file handle
'Dim tmDsf As DSF        'DSF record image
'Dim tmDsfSrchKey As LONGKEY0    'DSF key record image
'Dim imDsfRecLen As Integer        'DSF record length
'Dim tmRec As LPOPREC
Dim imTerminate As Integer
Dim imConverting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smNowDate As String
Dim lmNowDate As Long
Dim smSyncDate As String
Dim smSyncTime As String
Dim smDataForm As String
Dim imBaseLen As Integer
Dim imNoBuckets As Integer
Dim bmResearchSaved As Boolean

Dim tmNameCode() As SORTCODE
Dim smNameCodeTag As String


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
    ilRet = gIMoveListBox(ImptRad, lbcDemo, tmNameCode(), smNameCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffSet())
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

    lacFileType.Caption = ""
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
    ilRet = mBookNameUsed(slBookName, slBookDate, ilPrevDnfCode)
    If ilRet = 1 Then
        MsgBox "Book Name Previously Used", vbExclamation, "Name Error"
        edcBookName.SetFocus
        Exit Sub
    End If
    If ilRet = 2 Then
        Screen.MousePointer = vbDefault
        ilRet = MsgBox(slBookName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
        If ilRet = vbNo Then
            cmcCancel.SetFocus
            Exit Sub
        End If
        'Remove records associated with book previously imported
        Screen.MousePointer = vbHourglass
        gGetSyncDateTime smSyncDate, smSyncTime
        ilRet = mRemovePrevDnf(ilPrevDnfCode)
        If Not ilRet Then
            cmcCancel.SetFocus
            Exit Sub
        End If
    End If
    'Check file names
    If (InStr(slFromName, ":") = 0) And (Left$(slFromName, 2) <> "\\") Then
        slFromName = sgImportPath & slFromName
    End If
    ilPos = InStr(slFromName, ".")
    If ilPos = 0 Then
        slPopulationName = Left$(slFromName, Len(slFromName) - 1) & "1"
        slTimeName = Left$(slFromName, Len(slFromName) - 1) & "2"
        slDayPartName = Left$(slFromName, Len(slFromName) - 1) & "3"
        slGroupName = Left$(slFromName, Len(slFromName) - 1) & "4"
        slTimeNameDay = Left$(slFromName, Len(slFromName) - 1) & "5"
    Else
        slPopulationName = slFromName
        Mid$(slPopulationName, ilPos - 1, 1) = "1"
        slTimeName = slFromName
        Mid$(slTimeName, ilPos - 1, 1) = "2"
        slDayPartName = slFromName
        Mid$(slDayPartName, ilPos - 1, 1) = "3"
        slGroupName = slFromName
        Mid$(slGroupName, ilPos - 1, 1) = "4"
        slTimeNameDay = slFromName
        Mid$(slTimeNameDay, ilPos - 1, 1) = "5"
    End If
    lmProcessedNoBytes = 0
    ilRet = 0
    ReDim smVehNotFound(0 To 0) As String
    'On Error GoTo cmcFileConvErr:
    'hmFrom = FreeFile
    'Open slPopulationName For Input Access Read As hmFrom
    ilRet = gFileOpen(slPopulationName, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        Close hmFrom
        MsgBox "Unable to find " & slPopulationName, vbExclamation, "Name Error"
        edcFrom.SetFocus
        Exit Sub
    End If
    'Test which demo data format (16 or 18 buckets.  If 16, then record is 124 bytes followed by new line character (hex 0A).
    'If 18, then record is 134 btyes follwed by new line character
    slLine = Input(134, #hmFrom)    'Remove this line to read each character
    slChar = Input(1, #hmFrom)
    If (slChar = Chr(13)) Or (slChar = Chr(10)) Then
        smDataForm = "8"
        imBaseLen = 134
        imNoBuckets = 18
    Else
        smDataForm = "6"
        imBaseLen = 124
        imNoBuckets = 16
    End If

    lmTotalNoBytes = LOF(hmFrom) 'The Loc returns current position \128
    Close hmFrom
    'On Error GoTo cmcFileConvErr:
    'hmFrom = FreeFile
    'Open slTimeName For Input Access Read As hmFrom
    ilRet = gFileOpen(slTimeName, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        Close hmFrom
        MsgBox "Unable to find " & slTimeName, vbExclamation, "Name Error"
        edcFrom.SetFocus
        Exit Sub
    End If
    lmTotalNoBytes = lmTotalNoBytes + LOF(hmFrom) 'The Loc returns current position \128
    Close hmFrom
    'On Error GoTo cmcFileConvErr:
    'hmFrom = FreeFile
    'Open slDayPartName For Input Access Read As hmFrom
    ilRet = gFileOpen(slDayPartName, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        Close hmFrom
        MsgBox "Unable to find " & slDayPartName, vbExclamation, "Name Error"
        edcFrom.SetFocus
        Exit Sub
    End If
    lmTotalNoBytes = lmTotalNoBytes + LOF(hmFrom) 'The Loc returns current position \128
    Close hmFrom
    'hmFrom = FreeFile
    'Open slGroupName For Input Access Read As hmFrom
    ilRet = gFileOpen(slGroupName, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        Close hmFrom
        MsgBox "Unable to find " & slGroupName, vbExclamation, "Name Error"
        edcFrom.SetFocus
        Exit Sub
    End If
    lmTotalNoBytes = lmTotalNoBytes + LOF(hmFrom) 'The Loc returns current position \128
    Close hmFrom
    ilRet = 0
    'On Error GoTo cmcFileConvErr:
    'hmFrom = FreeFile
    'Open slTimeNameDay For Input Access Read As hmFrom
    ilRet = gFileOpen(slTimeNameDay, "Input Access Read", hmFrom)
    If ilRet = 0 Then
        lmTotalNoBytes = lmTotalNoBytes + LOF(hmFrom) 'The Loc returns current position \128
    End If
    Close hmFrom
    ilRet = 0
    'hmTo = FreeFile
    'Open sgDBPath & "Messages\" & "ImptRad.Txt" For Output As hmTo
    ilRet = gFileOpen(sgDBPath & "Messages\" & "ImptRad.Txt", "Output", hmTo)
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Open " & sgDBPath & "Messages\" & "ImptRad.Txt" & " Error #" & Str(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
        cmcCancel.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    gGetSyncDateTime smSyncDate, smSyncTime
    plcGauge.Value = 0
    lmPercent = 0
    Print #hmTo, "Import RADAR on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    Print #hmTo, ""
    'ReDim imVefCodeInDnf(1 To 1) As Integer
    ReDim imVefCodeInDnf(0 To 0) As Integer
    tmDnf.iCode = 0
    tmDnf.sBookName = slBookName
    gPackDate slBookDate, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
    gPackDate smNowDate, tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
    tmDnf.iUrfCode = tgUrf(0).iCode
    tmDnf.sType = "I"
    tmDnf.sForm = smDataForm
    tmDnf.sExactTime = "N"
    tmDnf.sSource = "R"
    tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
    tmDnf.iAutoCode = tmDnf.iCode
    ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        Print #hmTo, "Error when Adding Demo Name File (DNF)" & Str$(ilRet)
        Close hmTo
        If gOkAddStrToListBox("Error Adding DNF", lmLen, imShowMsg) Then
            lbcErrors.AddItem "Error Adding DNF"
        Else
            imShowMsg = False
        End If
        imConverting = False
        mTerminate
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
    'End If
    'lmCount = 0
    lacFileType.Caption = "Processing Populations"
    imConverting = True
    Print #hmTo, "** Processing Populations: " & slPopulationName & " **"
    If Not mConvPopulation(slPopulationName) Then
        Print #hmTo, "Import RADAR terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        Close hmTo
        imConverting = False
        mTerminate
        Exit Sub
    End If
    'Process program file
    lacFileType.Caption = "Processing Times"
    Print #hmTo, "** Processing Times: " & slTimeName & " **"
    If Not mConvTime(slTimeName) Then
        Print #hmTo, "Import RADAR terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        Close hmTo
        imConverting = False
        mTerminate
        Exit Sub
    End If
    'Process daypart file
    lacFileType.Caption = "Processing Dayparts"
    Print #hmTo, "** Processing Dayparts: " & slDayPartName & " **"
    If Not mConvDaypart(slDayPartName) Then
        Print #hmTo, "Import RADAR terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        Close hmTo
        imConverting = False
        mTerminate
        Exit Sub
    End If
    'Process Group file
    lacFileType.Caption = "Processing Groups"
    Print #hmTo, "** Processing Groups: " & slGroupName & " **"
    If Not mConvGroup(slGroupName) Then
        Print #hmTo, "Import RADAR terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        Close hmTo
        imConverting = False
        mTerminate
        Exit Sub
    End If
    'Process program file # 5
    lacFileType.Caption = "Processing Times by Day"
    Print #hmTo, "** Processing Times by Day: " & slTimeNameDay & " **"
    If Not mConvTimeDay(slTimeNameDay) Then
        Print #hmTo, "Import RADAR terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        Close hmTo
        imConverting = False
        mTerminate
        Exit Sub
    End If
    If ckcDefault(0).Value = vbChecked Or ckcDefault(1).Value = vbChecked Then
        For ilLoop = LBound(imVefCodeInDnf) To UBound(imVefCodeInDnf) - 1 Step 1
            Do
                tmVefSrchKey.iCode = imVefCodeInDnf(ilLoop)
                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
                If ckcDefault(0).Value = vbChecked Then
                    tmVef.iDnfCode = tmDnf.iCode
                End If
                If ckcDefault(1).Value = vbChecked Then
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
        Next ilLoop
        '11/26/17
        gFileChgdUpdate "vef.btr", False
        
    End If
    Print #hmTo, "Import RADAR successfully completed on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    Close hmTo
    ilRet = mObtainBookName()
    lacFileType.Caption = "Done"
    plcGauge.Value = 100
    bmResearchSaved = True
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    imConverting = False
    Screen.MousePointer = vbDefault
    Exit Sub
'cmcFileConvErr:
'    ilRet = Err.Number
'    Resume Next
End Sub
Private Sub cmcFrom_Click()
    lacFileType.Caption = ""
    'CMDialogBox.DialogTitle = "From File"
    'CMDialogBox.Filter = "Blank|*.|ASC|*.Asc|Text|*.Txt|Print|*.Prn|All|*.*"
    'CMDialogBox.InitDir = Left$(sgImportPath, Len(sgImportPath) - 1)
    'CMDialogBox.Filename = "ABC0"
    'CMDialogBox.DefaultExt = ""
    'CMDialogBox.Action = 1 'Open dialog
    'edcFrom.Text = CMDialogBox.Filename
    'ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
    'ChDir sgCurDir
    'DoEvents
    'edcFrom.SetFocus
    
    igBrowserType = 7 'blank ext
    sgBrowseMaskFile = "*"
    Browser.Show vbModal
    If igBrowserReturn = 1 Then
        edcFrom.Text = sgBrowserFile
    End If
    
    DoEvents
    If edcBookName.Text = "" Then
        mSetBookName
    End If
    edcFrom.SetFocus
    If InStr(1, sgCurDir, ":") > 0 Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
        ChDir sgCurDir
    End If
End Sub

Private Sub edcBookDate_GotFocus()
    lacFileType.Caption = ""
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcBookName_GotFocus()
    lacFileType.Caption = ""
    If edcBookName.Text = "" Then
        mSetBookName
    End If
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcFrom_Change()
    edcBookName.Text = ""
End Sub

Private Sub edcFrom_GotFocus()
    lacFileType.Caption = ""
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    gCtrlGotFocus ActiveControl
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
    Erase tgVefRad
    Erase tgDnfBook
    Erase tgMnfSocEcoRad
    Erase tgVehMerge
    'Erase tgRdf
    Erase imVefCodeInDnf
    Erase smVehNotFound
    ilRet = btrClose(hmRdf)
    btrDestroy hmRdf
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    ilRet = btrClose(hmDnf)
    btrDestroy hmDnf
    
    Set ImptRad = Nothing   'Remove data segment

End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAddGroup Name                  *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgMnfSocEcoRad            *
'*                                                     *
'*******************************************************
Private Sub mAddGroupName(slName As String, slGroup As String)
    Dim ilRet As Integer
    Dim ilUpperBound As Integer
    ilUpperBound = UBound(tgMnfSocEcoRad)
    tmMnf.iCode = 0
    tmMnf.sType = "F"
    tmMnf.sName = slName
    tmMnf.sRPU = ""
    tmMnf.sUnitType = slGroup
    tmMnf.iMerge = 0
    tmMnf.iGroupNo = 0
    tmMnf.sCodeStn = ""
    tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
    tmMnf.iAutoCode = tmMnf.iCode
    ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
    'If tgSpf.sRemoteUsers = "Y" Then
        Do
            'tmMnfSrchKey.iCode = tmMnf.iCode
            'ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
            tmMnf.iAutoCode = tmMnf.iCode
            gPackDate smSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
            gPackTime smSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
            ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
    'End If
    tgMnfSocEcoRad(ilUpperBound) = tmMnf
    ilUpperBound = ilUpperBound + 1
    'ReDim Preserve tgMnfSocEcoRad(1 To ilUpperBound) As MNF
    ReDim Preserve tgMnfSocEcoRad(0 To ilUpperBound) As MNF
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
Private Function mBookNameUsed(slBookName As String, slBookDate As String, ilPrevDnfCode As Integer) As Integer
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
    ilPrevDnfCode = 0
    For ilLoop = LBound(tgDnfBook) To UBound(tgDnfBook) - 1 Step 1
        If StrComp(Trim$(slBookName), Trim$(tgDnfBook(ilLoop).sBookName), 1) = 0 Then
            mBookNameUsed = 1
            gUnpackDateLong tgDnfBook(ilLoop).iBookDate(0), tgDnfBook(ilLoop).iBookDate(1), llDate
            If (llDate = llTestDate) Then
                mBookNameUsed = 2
                ilPrevDnfCode = tgDnfBook(ilLoop).iCode
            End If
            Exit Function
        End If
    Next ilLoop
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mConvDaypart                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert CHF                    *
'*                                                     *
'*******************************************************
Private Function mConvDaypart(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llPercent As Long
    Dim slChar As String
    Dim ilValue As Integer
    Dim ilAddFlag As Integer
    Dim slName As String
    Dim ilDay As Integer
    Dim ilSDemo As Integer
    Dim ilIndex As Integer
    Dim ilDayIndex As Integer
    Dim ilOk As Integer
    Dim ilNoTimes As Integer
    Dim ilTimeCount As Integer
    Dim ilFound As Integer
    Dim slStrDay As String
    Dim slStrTime As String
    Dim llRif As Long
    Dim ilRdf As Integer

    ilRet = 0
    'On Error GoTo mConvDaypartErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        MsgBox "Open " & slFromFile & ", Error #" & Str$(ilRet), vbExclamation, "Open Error"
        edcFrom.SetFocus
        mConvDaypart = False
        Exit Function
    End If
    DoEvents
    If imTerminate Then
        Close hmFrom
        mTerminate
        mConvDaypart = False
        Exit Function
    End If
    ilRet = 0
    err.Clear
    'On Error GoTo mConvDaypartErr:
    Do While Not EOF(hmFrom)
        ilRet = err.Number
        If ilRet <> 0 Then
            Close hmFrom
            MsgBox "Input Error #" & Str$(ilRet) & " when reading Daypart File", vbExclamation, "Read Error"
            mTerminate
            mConvDaypart = False
            Exit Function
        End If
        slLine = Input(imBaseLen, #hmFrom)    'Remove this line to read each character
        slChar = Input(1, #hmFrom)
        DoEvents
        If imTerminate Then
            Close hmFrom
            mTerminate
            mConvDaypart = False
            Exit Function
        End If
        Do While slChar = Chr(9)
            slChar = Input(1, #hmFrom)
        Loop
        If slChar = Chr(13) Then
            slChar = Input(1, #hmFrom)
        End If
        If slChar = Chr(10) Then
            'Process line
            ilAddFlag = True
            tmDrf.iDnfCode = tmDnf.iCode
            tmDrf.sDemoDataType = "D"
            tmDrf.iMnfSocEco = 0
            slStr = Mid$(slLine, 1, 2) 'Demographic Vehicle
            'slStr = UCase$(slStr)
            'If slStr = "AX" Then
            '    slName = "Excel"
            'ElseIf slStr = "AG" Then
            '    slName = "Galaxy"
            'ElseIf slStr = "AN" Then
            '    slName = "Genesis"
            'ElseIf slStr = "AL" Then
            '    slName = "Platinum"
            'ElseIf slStr = "AP" Then
            '    slName = "Prime"
            'ElseIf slStr = "AA" Then
            '    slName = "Advantage Net"
            'Else
            slName = mGetVehName(slStr)
            If Len(slName) = 0 Then
                ilFound = False
                For ilLoop = LBound(smVehNotFound) To UBound(smVehNotFound) - 1 Step 1
                    If StrComp(slStr, smVehNotFound(ilLoop), 1) = 0 Then
                        ilFound = True
                    End If
                Next ilLoop
                If Not ilFound Then
                    smVehNotFound(UBound(smVehNotFound)) = slStr
                    ReDim Preserve smVehNotFound(0 To UBound(smVehNotFound) + 1) As String
                    Print #hmTo, "Unable to Find Vehicle Code " & slStr & " record not added"
                    If gOkAddStrToListBox("Unable to Find Vehicle Code " & slStr, lmLen, imShowMsg) Then
                        lbcErrors.AddItem "Unable to Find Vehicle Code " & slStr
                    Else
                        imShowMsg = False
                    End If
                End If
                slName = " "
                ilAddFlag = False
            End If
            tmDrf.iVefCode = 0
            If Trim$(slName) <> "" Then
                For ilLoop = LBound(tgVefRad) To UBound(tgVefRad) - 1 Step 1
                    If StrComp(slName, Trim$(tgVefRad(ilLoop).sName), 1) = 0 Then
                        tmDrf.iVefCode = tgVefRad(ilLoop).iCode
                        Exit For
                    End If
                Next ilLoop
                If tmDrf.iVefCode = 0 Then
                    ilAddFlag = False
                    ilFound = False
                    For ilLoop = LBound(smVehNotFound) To UBound(smVehNotFound) - 1 Step 1
                        If StrComp(slName, smVehNotFound(ilLoop), 1) = 0 Then
                            ilFound = True
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        smVehNotFound(UBound(smVehNotFound)) = slName
                        ReDim Preserve smVehNotFound(0 To UBound(smVehNotFound) + 1) As String
                        Print #hmTo, "Unable to Find Vehicle " & slName & " record not added"
                        If gOkAddStrToListBox("Unable to Find Vehicle " & slName, lmLen, imShowMsg) Then
                            lbcErrors.AddItem "Unable to Find Vehicle " & slName
                        Else
                            imShowMsg = False
                        End If
                    End If
                End If
            End If
            ilFound = False
            For ilLoop = LBound(imVefCodeInDnf) To UBound(imVefCodeInDnf) - 1 Step 1
                If imVefCodeInDnf(ilLoop) = tmDrf.iVefCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                imVefCodeInDnf(UBound(imVefCodeInDnf)) = tmDrf.iVefCode
                'ReDim Preserve imVefCodeInDnf(1 To UBound(imVefCodeInDnf) + 1) As Integer
                ReDim Preserve imVefCodeInDnf(0 To UBound(imVefCodeInDnf) + 1) As Integer
            End If
            tmDrf.sInfoType = "D"
            tmDrf.iRdfCode = 0
            slStr = Mid$(slLine, 8, 3) 'Day Code
            For ilDay = 0 To 6 Step 1
                tmDrf.sDay(ilDay) = "N"
            Next ilDay
            slStr = UCase$(Trim$(slStr))
            slStrDay = slStr
            'If slStr = "M-S" Then
            If (slStr = "M-S") Or (slStr = "MS*") Then
                For ilDay = 0 To 6 Step 1
                    tmDrf.sDay(ilDay) = "Y"
                Next ilDay
'           ElseIf slStr = "M-F" Then
            ElseIf (slStr = "M-F") Or (slStr = "MF*") Then
                For ilDay = 0 To 4 Step 1
                    tmDrf.sDay(ilDay) = "Y"
                Next ilDay
            ElseIf slStr = "SAT" Then
                tmDrf.sDay(5) = "Y"
            ElseIf slStr = "SUN" Then
                tmDrf.sDay(6) = "Y"
            ElseIf slStr = "ALL" Then
                For ilDay = 0 To 6 Step 1
                    tmDrf.sDay(ilDay) = "Y"
                Next ilDay
            ElseIf slStr = "MSA" Then
                For ilDay = 0 To 5 Step 1
                    tmDrf.sDay(ilDay) = "Y"
                Next ilDay
            ElseIf slStr = "S-S" Then
                For ilDay = 5 To 6 Step 1
                    tmDrf.sDay(ilDay) = "Y"
                Next ilDay
            Else
                ilAddFlag = False
                Print #hmTo, "Unable to Find Days " & slStrDay & " on " & slName & " record not added"
                If gOkAddStrToListBox("Unable to Find Days " & slStrDay, lmLen, imShowMsg) Then
                    lbcErrors.AddItem "Unable to Find Days " & slStrDay
                Else
                    imShowMsg = False
                End If
            End If
            tmDrf.iStartTime2(0) = 1
            tmDrf.iStartTime2(1) = 0
            tmDrf.iEndTime2(0) = 1
            tmDrf.iEndTime2(1) = 0
            ilNoTimes = 1
            slStr = Mid$(slLine, 16, 8) 'Daypart Labels
            slStr = UCase$(Trim$(slStr))
            slStrTime = slStr
'            If slStr = "12M-12M" Then
'                gPackTime "12AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "6A-12M" Then
'                gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "6A-7P" Then
'                gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "7PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "12M-6A" Then
'                gPackTime "12AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "6AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "6A-10A" Then
'                gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "10AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "10A-3P" Then
'                gPackTime "10AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "3PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "3P-7P" Then
'                gPackTime "3PM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "7PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "7P-12M" Then
'                gPackTime "7PM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "6-10+3-7" Then
'                gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "10AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'                gPackTime "3PM", tmDrf.iStartTime2(0), tmDrf.iStartTime2(1)
'                gPackTime "7PM", tmDrf.iEndTime2(0), tmDrf.iEndTime2(1)
'                ilNoTimes = 2
'            Else
'                ilAddFlag = False
'                Print #hmTo, "Unable to Find Daypart Time " & slStrTime & " on " & slName & " record not added"
'                If gOkAddStrToListBox("Unable to Find Daypart Time " & slStrTime, lmLen, imShowMsg) Then
'                    lbcErrors.AddItem "Unable to Find Daypart Time " & slStrTime
'                Else
'                    imShowMsg = False
'                End If
'            End If
            '11/29/08: Arbitron added 5a and 8p times
            mConvertDayparts slStr, slStrTime, slName, ilAddFlag, ilNoTimes

            'Scan if daypart matches a daypart definition
            If ilAddFlag Then
                For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                    If (tgMRif(llRif).iVefCode = tmDrf.iVefCode) Then
                        'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                        '    If tgMRif(llRif).iRdfcode = tgMRdf(ilRdf).iCode Then
                            ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                            If ilRdf <> -1 Then
                                'Ignore Dormant Dayparts- Jim request 6/23/04 because of demo to XM
                                If tgMRdf(ilRdf).sState <> "D" Then
                                    If (tgMRdf(ilRdf).iLtfCode(0) = 0) And (tgMRdf(ilRdf).iLtfCode(1) = 0) And (tgMRdf(ilRdf).iLtfCode(2) = 0) Then
                                        ilTimeCount = 0
                                        For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                                ilTimeCount = ilTimeCount + 1
                                            End If
                                        Next ilIndex
                                        If ilTimeCount = ilNoTimes Then
                                            ilTimeCount = 0
                                            For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                                If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                                    ilOk = False
                                                    If ((tgMRdf(ilRdf).iStartTime(0, ilIndex) = tmDrf.iStartTime(0)) And (tgMRdf(ilRdf).iStartTime(1, ilIndex) = tmDrf.iStartTime(1))) Then
                                                        If (tgMRdf(ilRdf).iEndTime(0, ilIndex) = tmDrf.iEndTime(0)) And (tgMRdf(ilRdf).iEndTime(1, ilIndex) = tmDrf.iEndTime(1)) Then
                                                            ilOk = True
                                                        End If
                                                    End If
                                                    If ((tgMRdf(ilRdf).iStartTime(0, ilIndex) = tmDrf.iStartTime2(0)) And (tgMRdf(ilRdf).iStartTime(1, ilIndex) = tmDrf.iStartTime2(1))) Then
                                                        If (tgMRdf(ilRdf).iEndTime(0, ilIndex) = tmDrf.iEndTime2(0)) And (tgMRdf(ilRdf).iEndTime(1, ilIndex) = tmDrf.iEndTime2(1)) Then
                                                            ilOk = True
                                                        End If
                                                    End If
                                                    If ilOk Then
                                                        'Exact time match- check days
                                                        ilOk = True
                                                        For ilDayIndex = 0 To 6 Step 1
                                                            'If (tmDrf.sDay(ilDayIndex) = "Y") And (tgMRdf(ilRdf).sWkDays(ilIndex, ilDayIndex + 1) <> "Y") Then
                                                            If (tmDrf.sDay(ilDayIndex) = "Y") And (tgMRdf(ilRdf).sWkDays(ilIndex, ilDayIndex) <> "Y") Then
                                                                ilOk = False
                                                                Exit For
                                                            'ElseIf (tmDrf.sDay(ilDayIndex) = "N") And (tgMRdf(ilRdf).sWkDays(ilIndex, ilDayIndex + 1) <> "N") Then
                                                            ElseIf (tmDrf.sDay(ilDayIndex) = "N") And (tgMRdf(ilRdf).sWkDays(ilIndex, ilDayIndex) <> "N") Then
                                                                ilOk = False
                                                                Exit For
                                                            End If
                                                        Next ilDayIndex
                                                        If ilOk Then
                                                            ilTimeCount = ilTimeCount + 1
                                                            If ilTimeCount = ilNoTimes Then
                                                                tmDrf.iRdfCode = tgMRdf(ilRdf).iCode
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next ilIndex
                                        End If
                                    End If
                                End If
                        '        Exit For
                            End If
                        'Next ilRdf
                        If tmDrf.iRdfCode > 0 Then
                            Exit For
                        End If
                    End If
                Next llRif
                If tmDrf.iRdfCode = 0 Then
                    'Print #hmTo, "Unable to Find Matching Daypart, using Daypart Times/Days: " & slStrTime & " " & slStrDay & " on " & slName
                    'If gOkAddStrToListBox("Unable to Find Matching Daypart", lmLen, imShowMsg) Then
                    '    lbcErrors.AddItem "Unable to Find Matching Daypart"
                    'Else
                    '    imShowMsg = False
                    'End If
                End If
            End If
            tmDrf.sProgCode = ""
            tmDrf.iQHIndex = 0
            slStr = Mid$(slLine, 34, 4) 'Avg # of Broadcast
            ilValue = Val(slStr)
            tmDrf.iCount = ilValue
            tmDrf.sExStdDP = "N"
            slStr = Mid$(slLine, 5, 1) 'Avg # of Stations
            If UCase$(slStr) = "Z" Then
                tmDrf.sExRpt = "Y"
            Else
                tmDrf.sExRpt = "N"
            End If
            slStr = Mid$(slLine, 29, 1) 'Type of data
            If UCase$(slStr) = "A" Then
                tmDrf.sDataType = "A"
            Else
                ilAddFlag = False
                tmDrf.sDataType = "C"
            End If
            For ilLoop = 1 To 18 Step 1
                tmDrf.lDemo(ilLoop - 1) = 0
            Next ilLoop
            ilSDemo = 45
            For ilLoop = 1 To imNoBuckets Step 1
                slStr = Mid$(slLine, ilSDemo, 5)
                If Trim$(slStr) <> "" Then
                    tmDrf.lDemo(ilLoop - 1) = Val(slStr)
                Else
                    tmDrf.lDemo(ilLoop - 1) = 0
                End If
                ilSDemo = ilSDemo + 5
            Next ilLoop
            If ilAddFlag Then
                tmDrf.sForm = smDataForm
                tmDrf.lCode = 0
                tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmDrf.lAutoCode = tmDrf.lCode
                If tgSpf.sSAudData = "H" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 10 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If tgSpf.sSAudData = "N" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 100 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If tgSpf.sSAudData = "U" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 1000 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If (slStrDay = "M-S") Or (slStrDay = "MS*") Or (slStrDay = "M-F") Or (slStrDay = "MF*") Then
                    If mDrfDefined() Then
                        ilAddFlag = False
                    End If
                End If
                If ilAddFlag Then
                    ilRet = btrInsert(hmDrf, tmDrf, imDrfRecLen, INDEXKEY2)
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
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + 125
            llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmPercent <> llPercent Then
                plcGauge.Value = llPercent
                lmPercent = llPercent
            End If
            slLine = ""
            slChar = ""
        Else
            slLine = slLine & slChar
        End If
    Loop
    Close hmFrom
    mConvDaypart = True
    Exit Function
'mConvDaypartErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mConvGroup                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert CHF                    *
'*                                                     *
'*******************************************************
Private Function mConvGroup(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llPercent As Long
    Dim slChar As String
    Dim ilValue As Integer
    Dim ilAddFlag As Integer
    Dim slName As String
    Dim ilDay As Integer
    Dim ilSDemo As Integer
    Dim ilIndex As Integer
    Dim ilDayIndex As Integer
    Dim ilOk As Integer
    Dim ilNoTimes As Integer
    Dim ilTimeCount As Integer
    Dim ilFound As Integer
    Dim slStrDay As String
    Dim slStrTime As String
    Dim llRif As Long
    Dim ilRdf As Integer
    ilRet = 0
    'On Error GoTo mConvGroupErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        MsgBox "Open " & slFromFile & ", Error #" & Str$(ilRet), vbExclamation, "Open Error"
        edcFrom.SetFocus
        mConvGroup = False
        Exit Function
    End If
    DoEvents
    If imTerminate Then
        Close hmFrom
        mTerminate
        mConvGroup = False
        Exit Function
    End If
    ilRet = 0
    err.Clear
    'On Error GoTo mConvGroupErr:
    Do While Not EOF(hmFrom)
        ilRet = err.Number
        If ilRet <> 0 Then
            Close hmFrom
            MsgBox "Input Error #" & Str$(ilRet) & " when reading Group File", vbExclamation, "Read Error"
            mTerminate
            mConvGroup = False
            Exit Function
        End If
        slLine = Input(imBaseLen, #hmFrom)    'Remove this line to read each character
        slChar = Input(1, #hmFrom)
        DoEvents
        If imTerminate Then
            Close hmFrom
            mTerminate
            mConvGroup = False
            Exit Function
        End If
        Do While slChar = Chr(9)
            slChar = Input(1, #hmFrom)
        Loop
        If slChar = Chr(13) Then
            slChar = Input(1, #hmFrom)
        End If
        If slChar = Chr(10) Then
            'Process line
            'Process line
            ilAddFlag = True
            tmDrf.iDnfCode = tmDnf.iCode
            tmDrf.sDemoDataType = "D"
            tmDrf.iMnfSocEco = 0
            '8/28/07:  Changed from 26-27 to 25-27
            'slStr = Mid$(slLine, 26, 2) 'Demographic Group (T0 or A1-L3)
            slStr = Trim$(Mid$(slLine, 25, 3)) 'Demographic Group (T0 or A1-L3)
            For ilLoop = LBound(tgMnfSocEcoRad) To UBound(tgMnfSocEcoRad) - 1 Step 1
                If StrComp(slStr, Trim$(tgMnfSocEcoRad(ilLoop).sUnitType), 1) = 0 Then
                    tmDrf.iMnfSocEco = tgMnfSocEcoRad(ilLoop).iCode
                    Exit For
                End If
            Next ilLoop
            If tmDrf.iMnfSocEco = 0 Then
                'ilAddFlag = False
                Print #hmTo, "Population Group " & slStr & " added"
                If gOkAddStrToListBox("Population Group " & slStr & " added", lmLen, imShowMsg) Then
                    lbcErrors.AddItem "Population Group " & slStr & " added"
                Else
                    imShowMsg = False
                End If
                mAddGroupName slStr & " Definition Missing", slStr
                tmDrf.iMnfSocEco = tgMnfSocEcoRad(UBound(tgMnfSocEcoRad) - 1).iCode
            End If
            slStr = Mid$(slLine, 1, 2) 'Demographic Vehicle
            'slStr = UCase$(slStr)
            'If slStr = "AX" Then
            '    slName = "Excel"
            'ElseIf slStr = "AG" Then
            '    slName = "Galaxy"
            'ElseIf slStr = "AN" Then
            '    slName = "Genesis"
            'ElseIf slStr = "AL" Then
            '    slName = "Platinum"
            'ElseIf slStr = "AP" Then
            '    slName = "Prime"
            'ElseIf slStr = "AA" Then
            '    slName = "Advantage Net"
            'Else
            slName = mGetVehName(slStr)
            If Len(slName) = 0 Then
                ilFound = False
                For ilLoop = LBound(smVehNotFound) To UBound(smVehNotFound) - 1 Step 1
                    If StrComp(slStr, smVehNotFound(ilLoop), 1) = 0 Then
                        ilFound = True
                    End If
                Next ilLoop
                If Not ilFound Then
                    smVehNotFound(UBound(smVehNotFound)) = slStr
                    ReDim Preserve smVehNotFound(0 To UBound(smVehNotFound) + 1) As String
                    If ilAddFlag Then
                        Print #hmTo, "Unable to Find Vehicle Code " & slStr & " record not added"
                        If gOkAddStrToListBox("Unable to Find Vehicle Code " & slStr, lmLen, imShowMsg) Then
                            lbcErrors.AddItem "Unable to Find Vehicle Code " & slStr
                        Else
                            imShowMsg = False
                        End If
                    End If
                End If
                slName = " "
                ilAddFlag = False
            End If
            tmDrf.iVefCode = 0
            If Trim$(slName) <> "" Then
                For ilLoop = LBound(tgVefRad) To UBound(tgVefRad) - 1 Step 1
                    If StrComp(slName, Trim$(tgVefRad(ilLoop).sName), 1) = 0 Then
                        tmDrf.iVefCode = tgVefRad(ilLoop).iCode
                        Exit For
                    End If
                Next ilLoop
                If tmDrf.iVefCode = 0 Then
                    ilFound = False
                    For ilLoop = LBound(smVehNotFound) To UBound(smVehNotFound) - 1 Step 1
                        If StrComp(slName, smVehNotFound(ilLoop), 1) = 0 Then
                            ilFound = True
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        smVehNotFound(UBound(smVehNotFound)) = slName
                        ReDim Preserve smVehNotFound(0 To UBound(smVehNotFound) + 1) As String
                        If ilAddFlag Then
                            Print #hmTo, "Unable to Find Vehicle " & slName & " record not added"
                            If gOkAddStrToListBox("Unable to Find Vehicle " & slName, lmLen, imShowMsg) Then
                                lbcErrors.AddItem "Unable to Find Vehicle " & slName
                            Else
                                imShowMsg = False
                            End If
                        End If
                    End If
                    ilAddFlag = False
                End If
            End If
            ilFound = False
            For ilLoop = LBound(imVefCodeInDnf) To UBound(imVefCodeInDnf) - 1 Step 1
                If imVefCodeInDnf(ilLoop) = tmDrf.iVefCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                imVefCodeInDnf(UBound(imVefCodeInDnf)) = tmDrf.iVefCode
                'ReDim Preserve imVefCodeInDnf(1 To UBound(imVefCodeInDnf) + 1) As Integer
                ReDim Preserve imVefCodeInDnf(0 To UBound(imVefCodeInDnf) + 1) As Integer
            End If
            tmDrf.sInfoType = "D"
            tmDrf.iRdfCode = 0
            slStr = Mid$(slLine, 8, 3) 'Day Code
            For ilDay = 0 To 6 Step 1
                tmDrf.sDay(ilDay) = "N"
            Next ilDay
            slStr = UCase$(Trim$(slStr))
            slStrDay = slStr
            'If slStr = "M-S" Then
            If (slStr = "M-S") Or (slStr = "MS*") Then
                For ilDay = 0 To 6 Step 1
                    tmDrf.sDay(ilDay) = "Y"
                Next ilDay
            'ElseIf slStr = "M-F" Then
            ElseIf (slStr = "M-F") Or (slStr = "MF*") Then
                For ilDay = 0 To 4 Step 1
                    tmDrf.sDay(ilDay) = "Y"
                Next ilDay
            ElseIf slStr = "SAT" Then
                tmDrf.sDay(5) = "Y"
            ElseIf slStr = "SUN" Then
                tmDrf.sDay(6) = "Y"
            ElseIf slStr = "ALL" Then
                For ilDay = 0 To 6 Step 1
                    tmDrf.sDay(ilDay) = "Y"
                Next ilDay
            ElseIf slStr = "MSA" Then
                For ilDay = 0 To 5 Step 1
                    tmDrf.sDay(ilDay) = "Y"
                Next ilDay
            ElseIf slStr = "S-S" Then
                For ilDay = 5 To 6 Step 1
                    tmDrf.sDay(ilDay) = "Y"
                Next ilDay
            Else
                If ilAddFlag Then
                    ilAddFlag = False
                    Print #hmTo, "Unable to Find Days " & slStrDay & " on " & slName & " record not added"
                    If gOkAddStrToListBox("Unable to Find Days " & slStrDay, lmLen, imShowMsg) Then
                        lbcErrors.AddItem "Unable to Find Days " & slStrDay
                    Else
                        imShowMsg = False
                    End If
                End If
            End If
            tmDrf.iStartTime2(0) = 1
            tmDrf.iStartTime2(1) = 0
            tmDrf.iEndTime2(0) = 1
            tmDrf.iEndTime2(1) = 0
            ilNoTimes = 1
            slStr = Mid$(slLine, 16, 8) 'Daypart Labels
            slStr = UCase$(Trim$(slStr))
            slStrTime = slStr
'            If slStr = "12M-12M" Then
'                gPackTime "12AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "6A-12M" Then
'                gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "6A-7P" Then
'                gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "7PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "12M-6A" Then
'                gPackTime "12AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "6AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "6A-10A" Then
'                gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "10AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "10A-3P" Then
'                gPackTime "10AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "3PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "3P-7P" Then
'                gPackTime "3PM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "7PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "7P-12M" Then
'                gPackTime "7PM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            ElseIf slStr = "6-10+3-7" Then
'                gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'                gPackTime "10AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'                gPackTime "3PM", tmDrf.iStartTime2(0), tmDrf.iStartTime2(1)
'                gPackTime "7PM", tmDrf.iEndTime2(0), tmDrf.iEndTime2(1)
'                ilNoTimes = 2
'            Else
'                If ilAddFlag Then
'                    ilAddFlag = False
'                    Print #hmTo, "Unable to Find Daypart Time " & slStrTime & " on " & slName & " record not added"
'                    If gOkAddStrToListBox("Unable to Find Daypart Time " & slStrTime, lmLen, imShowMsg) Then
'                        lbcErrors.AddItem "Unable to Find Daypart Time " & slStrTime
'                    Else
'                        imShowMsg = False
'                    End If
'                End If
'            End If
            '11/29/08: Arbitron added 5a and 8p times
            mConvertDayparts slStr, slStrTime, slName, ilAddFlag, ilNoTimes
            'Scan if daypart matches a daypart definition
            If ilAddFlag Then
                For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                    If (tgMRif(llRif).iVefCode = tmDrf.iVefCode) Then
                        'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                        '    If tgMRif(llRif).iRdfcode = tgMRdf(ilRdf).iCode Then
                            ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                            If ilRdf <> -1 Then
                                'Ignore Dormant Dayparts- Jim request 6/23/04 because of demo to XM
                                If tgMRdf(ilRdf).sState <> "D" Then
                                    If (tgMRdf(ilRdf).iLtfCode(0) = 0) And (tgMRdf(ilRdf).iLtfCode(1) = 0) And (tgMRdf(ilRdf).iLtfCode(2) = 0) Then
                                        ilTimeCount = 0
                                        For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                                ilTimeCount = ilTimeCount + 1
                                            End If
                                        Next ilIndex
                                        If ilTimeCount = ilNoTimes Then
                                            ilTimeCount = 0
                                            For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                                If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
                                                    ilOk = False
                                                    If ((tgMRdf(ilRdf).iStartTime(0, ilIndex) = tmDrf.iStartTime(0)) And (tgMRdf(ilRdf).iStartTime(1, ilIndex) = tmDrf.iStartTime(1))) Then
                                                        If (tgMRdf(ilRdf).iEndTime(0, ilIndex) = tmDrf.iEndTime(0)) And (tgMRdf(ilRdf).iEndTime(1, ilIndex) = tmDrf.iEndTime(1)) Then
                                                            ilOk = True
                                                        End If
                                                    End If
                                                    If ((tgMRdf(ilRdf).iStartTime(0, ilIndex) = tmDrf.iStartTime2(0)) And (tgMRdf(ilRdf).iStartTime(1, ilIndex) = tmDrf.iStartTime2(1))) Then
                                                        If (tgMRdf(ilRdf).iEndTime(0, ilIndex) = tmDrf.iEndTime2(0)) And (tgMRdf(ilRdf).iEndTime(1, ilIndex) = tmDrf.iEndTime2(1)) Then
                                                            ilOk = True
                                                        End If
                                                    End If
                                                    If ilOk Then
                                                        'Exact time match- check days
                                                        ilOk = True
                                                        For ilDayIndex = 0 To 6 Step 1
                                                            'If (tmDrf.sDay(ilDayIndex) = "Y") And (tgMRdf(ilRdf).sWkDays(ilIndex, ilDayIndex + 1) <> "Y") Then
                                                            If (tmDrf.sDay(ilDayIndex) = "Y") And (tgMRdf(ilRdf).sWkDays(ilIndex, ilDayIndex) <> "Y") Then
                                                                ilOk = False
                                                                Exit For
                                                            'ElseIf (tmDrf.sDay(ilDayIndex) = "N") And (tgMRdf(ilRdf).sWkDays(ilIndex, ilDayIndex + 1) <> "N") Then
                                                            ElseIf (tmDrf.sDay(ilDayIndex) = "N") And (tgMRdf(ilRdf).sWkDays(ilIndex, ilDayIndex) <> "N") Then
                                                                ilOk = False
                                                                Exit For
                                                            End If
                                                        Next ilDayIndex
                                                        If ilOk Then
                                                            ilTimeCount = ilTimeCount + 1
                                                            If ilTimeCount = ilNoTimes Then
                                                                tmDrf.iRdfCode = tgMRdf(ilRdf).iCode
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next ilIndex
                                        End If
                                    End If
                                End If
                        '        Exit For
                            End If
                        'Next ilRdf
                        If tmDrf.iRdfCode > 0 Then
                            Exit For
                        End If
                    End If
                Next llRif
                If tmDrf.iRdfCode = 0 Then
                    'Print #hmTo, "Unable to Find Matching Daypart, using Daypart Times/Days: " & slStrTime & " " & slStrDay & " on " & slName
                    ''lbcErrors.AddItem "Unable to Find Matching Daypart"
                End If
            End If
            tmDrf.sProgCode = ""
            tmDrf.iQHIndex = 0
            slStr = Mid$(slLine, 34, 4) 'Avg # of Broadcast
            ilValue = Val(slStr)
            tmDrf.iCount = ilValue
            tmDrf.sExStdDP = "N"
            slStr = Mid$(slLine, 5, 1) 'Exclude Standard report
            If UCase$(slStr) = "Z" Then
                tmDrf.sExRpt = "Y"
            Else
                tmDrf.sExRpt = "N"
            End If
            slStr = Mid$(slLine, 29, 1) 'Type of data
            If UCase$(slStr) = "A" Then
                tmDrf.sDataType = "A"
            Else
                ilAddFlag = False
                tmDrf.sDataType = "C"
            End If
            For ilLoop = 1 To 18 Step 1
                tmDrf.lDemo(ilLoop - 1) = 0
            Next ilLoop
            ilSDemo = 45
            For ilLoop = 1 To imNoBuckets Step 1
                slStr = Mid$(slLine, ilSDemo, 5)
                If Trim$(slStr) <> "" Then
                    tmDrf.lDemo(ilLoop - 1) = Val(slStr)
                Else
                    tmDrf.lDemo(ilLoop - 1) = 0
                End If
                ilSDemo = ilSDemo + 5
            Next ilLoop
            If ilAddFlag Then
                tmDrf.sForm = smDataForm
                tmDrf.lCode = 0
                tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmDrf.lAutoCode = tmDrf.lCode
                If tgSpf.sSAudData = "H" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 10 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If tgSpf.sSAudData = "N" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 100 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If tgSpf.sSAudData = "U" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 1000 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If (slStrDay = "M-S") Or (slStrDay = "MS*") Or (slStrDay = "M-F") Or (slStrDay = "MF*") Then
                    If mDrfDefined() Then
                        ilAddFlag = False
                    End If
                End If
                If ilAddFlag Then
                    ilRet = btrInsert(hmDrf, tmDrf, imDrfRecLen, INDEXKEY2)
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
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + 125
            llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmPercent <> llPercent Then
                plcGauge.Value = llPercent
                lmPercent = llPercent
            End If
            slLine = ""
            slChar = ""
        Else
            slLine = slLine & slChar
        End If
    Loop
    Close hmFrom
    mConvGroup = True
    Exit Function
'mConvGroupErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mConvPopulation                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert CHF                    *
'*                                                     *
'*******************************************************
Private Function mConvPopulation(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llPercent As Long
    Dim slChar As String
    Dim ilDay As Integer
    Dim ilSDemo As Integer
    Dim ilAddFlag As Integer
    ilRet = 0
    'On Error GoTo mConvPopulationErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        MsgBox "Open " & slFromFile & ", Error #" & Str$(ilRet), vbExclamation, "Open Error"
        edcFrom.SetFocus
        mConvPopulation = False
        Exit Function
    End If
    DoEvents
    If imTerminate Then
        Close hmFrom
        mTerminate
        mConvPopulation = False
        Exit Function
    End If
    ilRet = 0
    err.Clear
    slLine = ""
    'On Error GoTo mConvPopulationErr:
    Do While Not EOF(hmFrom)
        ilRet = err.Number
        If ilRet <> 0 Then
            Close hmFrom
            MsgBox "Input Error #" & Str$(ilRet) & " when reading Population File", vbExclamation, "Read Error"
            mTerminate
            mConvPopulation = False
            Exit Function
        End If
        slLine = Input(imBaseLen, #hmFrom)    'Remove this line to read each character
        slChar = Input(1, #hmFrom)
        DoEvents
        If imTerminate Then
            Close hmFrom
            mTerminate
            mConvPopulation = False
            Exit Function
        End If
        Do While slChar = Chr(9)
            slChar = Input(1, #hmFrom)
        Loop
        If slChar = Chr(13) Then
            slChar = Input(1, #hmFrom)
        End If
        If slChar = Chr(10) Then
            ilAddFlag = True
            'Process line
            tmDrf.iDnfCode = tmDnf.iCode
            tmDrf.sDemoDataType = "P"
            '8/28/07:  Changed from 26-27 to 25-27
            'slStr = Mid$(slLine, 26, 2) 'Demographic Group (T0 or A1-L3)
            slStr = Trim$(Mid$(slLine, 25, 3)) 'Demographic Group (T0 or A1-L3)
            If (StrComp(slStr, "T0", 1) = 0) Or (StrComp(slStr, "Z0", 1) = 0) Then
                tmDrf.iMnfSocEco = 0
            Else
                tmDrf.iMnfSocEco = 0
                For ilLoop = LBound(tgMnfSocEcoRad) To UBound(tgMnfSocEcoRad) - 1 Step 1
                    If StrComp(slStr, Trim$(tgMnfSocEcoRad(ilLoop).sUnitType), 1) = 0 Then
                        tmDrf.iMnfSocEco = tgMnfSocEcoRad(ilLoop).iCode
                        Exit For
                    End If
                Next ilLoop
                If tmDrf.iMnfSocEco = 0 Then
                    'ilAddFlag = False
                    Print #hmTo, "Population Group " & slStr & " added"
                    If gOkAddStrToListBox("Population Group " & slStr & " added", lmLen, imShowMsg) Then
                        lbcErrors.AddItem "Population Group " & slStr & " added"
                    Else
                        imShowMsg = False
                    End If
                    mAddGroupName slStr & " Name Missing", slStr
                    tmDrf.iMnfSocEco = tgMnfSocEcoRad(UBound(tgMnfSocEcoRad) - 1).iCode
                End If
            End If
            tmDrf.iVefCode = 0
            tmDrf.sInfoType = ""
            tmDrf.iRdfCode = 0
            tmDrf.sProgCode = ""
            tmDrf.iStartTime(0) = 1
            tmDrf.iStartTime(1) = 0
            tmDrf.iEndTime(0) = 1
            tmDrf.iEndTime(1) = 0
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
            ilSDemo = 45
            For ilLoop = 1 To imNoBuckets Step 1
                slStr = Mid$(slLine, ilSDemo, 5)
                If Trim$(slStr) <> "" Then
                    tmDrf.lDemo(ilLoop - 1) = Val(slStr)
                Else
                    tmDrf.lDemo(ilLoop - 1) = 0
                End If
                ilSDemo = ilSDemo + 5
            Next ilLoop
            If ilAddFlag Then
                tmDrf.sForm = smDataForm
                tmDrf.lCode = 0
                tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmDrf.lAutoCode = tmDrf.lCode
                gPackDate smSyncDate, tmDrf.iSyncDate(0), tmDrf.iSyncDate(1)
                gPackTime smSyncTime, tmDrf.iSyncTime(0), tmDrf.iSyncTime(1)
                If tgSpf.sSAudData = "H" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 10 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If tgSpf.sSAudData = "N" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 100 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If tgSpf.sSAudData = "U" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 1000 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                ilRet = btrInsert(hmDrf, tmDrf, imDrfRecLen, INDEXKEY2)
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
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + 125
            llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmPercent <> llPercent Then
                plcGauge.Value = llPercent
                lmPercent = llPercent
            End If
            slLine = ""
            slChar = ""
        Else
            slLine = slLine & slChar
        End If
    Loop
    Close hmFrom
    mConvPopulation = True
    Exit Function
'mConvPopulationErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mConvTime                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert CHF                    *
'*                                                     *
'*******************************************************
Private Function mConvTime(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilPos As Integer
    Dim llPercent As Long
    Dim slChar As String
    Dim ilValue As Integer
    Dim ilHour As Integer
    Dim ilMin As Integer
    Dim slMin As String
    Dim ilAddFlag As Integer
    Dim slName As String
    Dim ilDay As Integer
    Dim ilSDemo As Integer
    Dim slTime As String
    Dim ilFound As Integer
    Dim ilDayValue As Integer
    ilRet = 0
    ReDim tgVehMerge(0 To 0) As VEHMERGE
    'On Error GoTo mConvTimeErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        MsgBox "Open " & slFromFile & ", Error #" & Str$(ilRet), vbExclamation, "Open Error"
        edcFrom.SetFocus
        mConvTime = False
        Exit Function
    End If
    DoEvents
    If imTerminate Then
        Close hmFrom
        mTerminate
        mConvTime = False
        Exit Function
    End If
    ilRet = 0
    err.Clear
    'On Error GoTo mConvTimeErr:
    Do While Not EOF(hmFrom)
        ilRet = err.Number
        If ilRet <> 0 Then
            Close hmFrom
            MsgBox "Input Error #" & Str$(ilRet) & " when reading Program File", vbExclamation, "Read Error"
            mTerminate
            mConvTime = False
            Exit Function
        End If
        slLine = Input(imBaseLen, #hmFrom)    'Remove this line to read each character
        slChar = Input(1, #hmFrom)
        DoEvents
        If imTerminate Then
            Close hmFrom
            mTerminate
            mConvTime = False
            Exit Function
        End If
        Do While slChar = Chr(9)
            slChar = Input(1, #hmFrom)
        Loop
        If slChar = Chr(13) Then
            slChar = Input(1, #hmFrom)
        End If
        If slChar = Chr(10) Then
            'Process line
            ilAddFlag = True
            tmDrf.iDnfCode = tmDnf.iCode
            tmDrf.sDemoDataType = "D"
            tmDrf.iMnfSocEco = 0
            slStr = Mid$(slLine, 1, 2) 'Demographic Vehicle
            'slStr = UCase$(slStr)
            'If slStr = "AX" Then
            '    slName = "Excel"
            'ElseIf slStr = "AG" Then
            '    slName = "Galaxy"
            'ElseIf slStr = "AN" Then
            '    slName = "Genesis"
            'ElseIf slStr = "AL" Then
            '    slName = "Platinum"
            'ElseIf slStr = "AP" Then
            '    slName = "Prime"
            'ElseIf slStr = "AA" Then
            '    slName = "Advantage Net"
            'Else
            slName = mGetVehName(slStr)
            If Len(slName) = 0 Then
                ilFound = False
                For ilLoop = LBound(smVehNotFound) To UBound(smVehNotFound) - 1 Step 1
                    If StrComp(slStr, smVehNotFound(ilLoop), 1) = 0 Then
                        ilFound = True
                    End If
                Next ilLoop
                If Not ilFound Then
                    smVehNotFound(UBound(smVehNotFound)) = slStr
                    ReDim Preserve smVehNotFound(0 To UBound(smVehNotFound) + 1) As String
                    Print #hmTo, "Unable to Find Vehicle Code " & slStr & " record not added"
                    If gOkAddStrToListBox("Unable to Find Vehicle Code " & slStr, lmLen, imShowMsg) Then
                        lbcErrors.AddItem "Unable to Find Vehicle Code " & slStr
                    Else
                        imShowMsg = False
                    End If
                End If
                slName = " "
                ilAddFlag = False
            End If
            tmDrf.iVefCode = 0
            If Trim$(slName) <> "" Then
                For ilLoop = LBound(tgVefRad) To UBound(tgVefRad) - 1 Step 1
                    If StrComp(slName, Trim$(tgVefRad(ilLoop).sName), 1) = 0 Then
                        tmDrf.iVefCode = tgVefRad(ilLoop).iCode
                        Exit For
                    End If
                Next ilLoop
                If tmDrf.iVefCode = 0 Then
                    ilAddFlag = False
                    ilFound = False
                    For ilLoop = LBound(smVehNotFound) To UBound(smVehNotFound) - 1 Step 1
                        If StrComp(slName, smVehNotFound(ilLoop), 1) = 0 Then
                            ilFound = True
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        smVehNotFound(UBound(smVehNotFound)) = slName
                        ReDim Preserve smVehNotFound(0 To UBound(smVehNotFound) + 1) As String
                        Print #hmTo, "Unable to Find Vehicle " & slName & " record not added"
                        If gOkAddStrToListBox("Unable to Find Vehicle " & slName, lmLen, imShowMsg) Then
                            lbcErrors.AddItem "Unable to Find Vehicle " & slName
                        Else
                            imShowMsg = False
                        End If
                    End If
                End If
            End If
            ilFound = False
            For ilLoop = LBound(imVefCodeInDnf) To UBound(imVefCodeInDnf) - 1 Step 1
                If imVefCodeInDnf(ilLoop) = tmDrf.iVefCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                imVefCodeInDnf(UBound(imVefCodeInDnf)) = tmDrf.iVefCode
                'ReDim Preserve imVefCodeInDnf(1 To UBound(imVefCodeInDnf) + 1) As Integer
                ReDim Preserve imVefCodeInDnf(0 To UBound(imVefCodeInDnf) + 1) As Integer
            End If
            tmDrf.sInfoType = "T"
            tmDrf.iRdfCode = 0
            slStr = Mid$(slLine, 31, 2) 'Day Code
            ilDayValue = Val(slStr)
            slStr = Mid$(slLine, 3, 4)  '5) 'Demographic Prog Code
            Do While Len(slStr) < Len(tmDrf.sProgCode)
                slStr = "0" & slStr
            Loop
            tmDrf.sProgCode = slStr
            If (ilDayValue <> 12) And (ilDayValue <> 13) Then
                slStr = Mid$(slLine, 7, 5) 'Demographic Prog Name- first 5 characters contains time
                slStr = Left$(slStr, 2) & ":" & right(slStr, 3)
                If Not gValidTime(slStr) Then
                    slTime = slStr
                    'Obtain the time from the quarter hour index
                    slStr = Mid$(slLine, 35, 2) 'Quarter Hour Code
                    ilValue = Val(slStr)
                    ilHour = ilValue \ 4
                    ilMin = (ilValue - 1) Mod 4
                    Select Case ilMin
                        Case 0
                            slMin = "00"
                        Case 1
                            slMin = "15"
                        Case 2
                            slMin = "30"
                        Case 3
                            slMin = "45"
                    End Select
                    If ilHour = 0 Then
                        slStr = "12" & ":" & slMin & "AM"
                    ElseIf ilHour < 12 Then
                        slStr = Trim$(Str$(ilHour)) & ":" & slMin & "AM"
                    ElseIf ilHour = 12 Then
                        slStr = "12" & ":" & slMin & "PM"
                    Else
                        slStr = Trim$(Str$(ilHour - 12)) & ":" & slMin & "AM"
                    End If
                    Print #hmTo, "Invalid Time " & slTime & ", using Quarter Hour Time" & slStr
                    If gOkAddStrToListBox("Invalid Time, using Quarter Hour " & slStr, lmLen, imShowMsg) Then
                        lbcErrors.AddItem "Invalid Time, using Quarter Hour " & slStr
                    Else
                        imShowMsg = False
                    End If
                End If
                gPackTime slStr, tmDrf.iStartTime(0), tmDrf.iStartTime(1)
                gPackTime slStr, tmDrf.iEndTime(0), tmDrf.iEndTime(1)
            Else
                gPackTime "12AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
                gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
                tmDrf.sInfoType = "V"
                tmDrf.iRdfCode = 0
            End If
            tmDrf.iStartTime2(0) = 1
            tmDrf.iStartTime2(1) = 0
            tmDrf.iEndTime2(0) = 1
            tmDrf.iEndTime2(1) = 0
            For ilDay = 0 To 6 Step 1
                tmDrf.sDay(ilDay) = "N"
            Next ilDay
            Select Case ilDayValue
                Case 1  'M-F
                    For ilDay = 0 To 4 Step 1
                        tmDrf.sDay(ilDay) = "Y"
                    Next ilDay
                Case 2  'Sat
                    tmDrf.sDay(5) = "Y"
                Case 3  'Sun
                    tmDrf.sDay(6) = "Y"
                Case 4  'T-S
                    For ilDay = 1 To 6 Step 1
                        tmDrf.sDay(ilDay) = "Y"
                    Next ilDay
                Case 5  'T-F
                    For ilDay = 1 To 4 Step 1
                        tmDrf.sDay(ilDay) = "Y"
                    Next ilDay
                Case 6  'W-S
                    For ilDay = 2 To 6 Step 1
                        tmDrf.sDay(ilDay) = "Y"
                    Next ilDay
                Case 7  'Mon
                    tmDrf.sDay(0) = "Y"
                Case 8  'Tue
                    tmDrf.sDay(1) = "Y"
                Case 9  'Wed
                    tmDrf.sDay(2) = "Y"
                Case 10 'Thu
                    tmDrf.sDay(3) = "Y"
                Case 11 'Fri
                    tmDrf.sDay(4) = "Y"
                Case 12 'MF ROS
                    For ilDay = 0 To 4 Step 1
                        tmDrf.sDay(ilDay) = "Y"
                    Next ilDay
                Case 13 'MS ROS
                    For ilDay = 0 To 6 Step 1
                        tmDrf.sDay(ilDay) = "Y"
                    Next ilDay
                Case Else
                    ilAddFlag = False
                    Print #hmTo, "Invalid Day Code " & Str$(ilDayValue) & " on " & slName & " record not added"
                    If gOkAddStrToListBox("Invalid Day Code" & Str$(ilDayValue), lmLen, imShowMsg) Then
                        lbcErrors.AddItem "Invalid Day Code" & Str$(ilDayValue)
                    Else
                        imShowMsg = False
                    End If
            End Select
            slStr = Mid$(slLine, 35, 2) 'Quarter Hour Code
            ilValue = Val(slStr)
            tmDrf.iQHIndex = ilValue
            slStr = Mid$(slLine, 37, 4) 'Avg # of Stations
            ilValue = Val(slStr)
            tmDrf.iCount = ilValue
            slStr = Mid$(slLine, 41, 1) 'Avg # of Stations
            If UCase$(slStr) = "X" Then
                tmDrf.sExStdDP = "Y"
            Else
                tmDrf.sExStdDP = "N"
            End If
            slStr = Mid$(slLine, 42, 1) 'Avg # of Stations
            If UCase$(slStr) = "Y" Then
                tmDrf.sExRpt = "Y"
            Else
                tmDrf.sExRpt = "N"
            End If
            slStr = Mid$(slLine, 43, 1) 'Type of data
            If UCase$(slStr) = "A" Then
                tmDrf.sDataType = "A"
            Else
                ilAddFlag = False
                tmDrf.sDataType = "C"
            End If
            For ilLoop = 1 To 18 Step 1
                tmDrf.lDemo(ilLoop - 1) = 0
            Next ilLoop
            ilSDemo = 45
            For ilLoop = 1 To imNoBuckets Step 1
                slStr = Mid$(slLine, ilSDemo, 5)
                If Trim$(slStr) <> "" Then
                    tmDrf.lDemo(ilLoop - 1) = Val(slStr)
                Else
                    tmDrf.lDemo(ilLoop - 1) = 0
                End If
                ilSDemo = ilSDemo + 5
            Next ilLoop
            If ilAddFlag Then
                tmDrf.sForm = smDataForm
                tmDrf.lCode = 0
                tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmDrf.lAutoCode = tmDrf.lCode
                If tgSpf.sSAudData = "H" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 10 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If tgSpf.sSAudData = "N" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 100 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If tgSpf.sSAudData = "U" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 1000 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If (ilDayValue = 12) Or (ilDayValue = 13) Then
                    If mDrfDefined() Then
                        ilAddFlag = False
                    End If
                End If
                If (ilAddFlag) And (tmDrf.sInfoType = "V") Then
                    ilAddFlag = False
                    ilFound = False
                    For ilLoop = LBound(tgVehMerge) To UBound(tgVehMerge) - 1 Step 1
                        If tgVehMerge(ilLoop).tDrf.iVefCode = tmDrf.iVefCode Then
                            ilFound = True
                            For ilPos = 1 To imNoBuckets Step 1
                                tgVehMerge(ilLoop).tDrf.lDemo(ilPos - 1) = tgVehMerge(ilLoop).tDrf.lDemo(ilPos - 1) + tmDrf.lDemo(ilPos - 1)
                            Next ilPos
                            tgVehMerge(ilLoop).iCount = tgVehMerge(ilLoop).iCount + 1
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        tgVehMerge(UBound(tgVehMerge)).tDrf = tmDrf
                        tgVehMerge(UBound(tgVehMerge)).iCount = 1
                        ReDim Preserve tgVehMerge(0 To UBound(tgVehMerge) + 1) As VEHMERGE
                    End If
                End If
                If ilAddFlag Then
                    ilRet = btrInsert(hmDrf, tmDrf, imDrfRecLen, INDEXKEY2)
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
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + 125
            llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmPercent <> llPercent Then
                plcGauge.Value = llPercent
                lmPercent = llPercent
            End If
            slLine = ""
            slChar = ""
        Else
            slLine = slLine & slChar
        End If
    Loop
    For ilLoop = LBound(tgVehMerge) To UBound(tgVehMerge) - 1 Step 1
        For ilPos = 1 To imNoBuckets Step 1
            tgVehMerge(ilLoop).tDrf.lDemo(ilPos - 1) = tgVehMerge(ilLoop).tDrf.lDemo(ilPos - 1) / tgVehMerge(ilLoop).iCount
        Next ilPos
        LSet tmDrf = tgVehMerge(ilLoop)
        slStr = "0"
        Do While Len(slStr) < Len(tmDrf.sProgCode)
            slStr = "0" & slStr
        Loop
        tmDrf.sForm = smDataForm
        tmDrf.sProgCode = slStr
        ilRet = btrInsert(hmDrf, tmDrf, imDrfRecLen, INDEXKEY2)
        Do
            'tmDrfSrchKey2.lCode = tmDrf.lCode
            'ilRet = btrGetEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
            tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
            tmDrf.lAutoCode = tmDrf.lCode
            gPackDate smSyncDate, tmDrf.iSyncDate(0), tmDrf.iSyncDate(1)
            gPackTime smSyncTime, tmDrf.iSyncTime(0), tmDrf.iSyncTime(1)
            ilRet = btrUpdate(hmDrf, tmDrf, imDrfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
    Next ilLoop
    Close hmFrom
    mConvTime = True
    Exit Function
'mConvTimeErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mConvTimeDay                    *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Convert CHF                    *
'*                                                     *
'*******************************************************
Private Function mConvTimeDay(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim llPercent As Long
    Dim slChar As String
    Dim ilValue As Integer
    Dim ilHour As Integer
    Dim ilMin As Integer
    Dim slMin As String
    Dim ilAddFlag As Integer
    Dim slName As String
    Dim ilDay As Integer
    Dim ilSDemo As Integer
    Dim slTime As String
    Dim ilFound As Integer
    Dim ilDayValue As Integer
    ilRet = 0
    'On Error GoTo mConvTimeDayErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        'MsgBox "Open " & slFromFile, vbExclamation, "Open Error"
        'edcFrom.SetFocus
        lbcErrors.AddItem "Warning: No Time by Day file found " & slFromFile
        Print #hmTo, "Warning: No Time by Day file found " & slFromFile
        mConvTimeDay = True
        Exit Function
    End If
    DoEvents
    If imTerminate Then
        Close hmFrom
        mTerminate
        mConvTimeDay = False
        Exit Function
    End If
    ilRet = 0
    err.Clear
    'On Error GoTo mConvTimeDayErr:
    Do While Not EOF(hmFrom)
        ilRet = err.Number
        If ilRet <> 0 Then
            Close hmFrom
            MsgBox "Input Error #" & Str$(ilRet) & " when reading Program File", vbExclamation, "Read Error"
            mTerminate
            mConvTimeDay = False
            Exit Function
        End If
        slLine = Input(imBaseLen, #hmFrom)    'Remove this line to read each character
        slChar = Input(1, #hmFrom)
        DoEvents
        If imTerminate Then
            Close hmFrom
            mTerminate
            mConvTimeDay = False
            Exit Function
        End If
        Do While slChar = Chr(9)
            slChar = Input(1, #hmFrom)
        Loop
        If slChar = Chr(13) Then
            slChar = Input(1, #hmFrom)
        End If
        If slChar = Chr(10) Then
            'Process line
            ilAddFlag = True
            tmDrf.iDnfCode = tmDnf.iCode
            tmDrf.sDemoDataType = "D"
            tmDrf.iMnfSocEco = 0
            slStr = Mid$(slLine, 1, 2) 'Demographic Vehicle
            'slStr = UCase$(slStr)
            'If slStr = "AX" Then
            '    slName = "Excel"
            'ElseIf slStr = "AG" Then
            '    slName = "Galaxy"
            'ElseIf slStr = "AN" Then
            '    slName = "Genesis"
            'ElseIf slStr = "AL" Then
            '    slName = "Platinum"
            'ElseIf slStr = "AP" Then
            '    slName = "Prime"
            'ElseIf slStr = "AA" Then
            '    slName = "Advantage Net"
            'Else
            slName = mGetVehName(slStr)
            If Len(slName) = 0 Then
                ilFound = False
                For ilLoop = LBound(smVehNotFound) To UBound(smVehNotFound) - 1 Step 1
                    If StrComp(slStr, smVehNotFound(ilLoop), 1) = 0 Then
                        ilFound = True
                    End If
                Next ilLoop
                If Not ilFound Then
                    smVehNotFound(UBound(smVehNotFound)) = slStr
                    ReDim Preserve smVehNotFound(0 To UBound(smVehNotFound) + 1) As String
                    Print #hmTo, "Unable to Find Vehicle Code " & slStr & " record not added"
                    If gOkAddStrToListBox("Unable to Find Vehicle Code " & slStr, lmLen, imShowMsg) Then
                        lbcErrors.AddItem "Unable to Find Vehicle Code " & slStr
                    Else
                        imShowMsg = False
                    End If
                End If
                slName = " "
                ilAddFlag = False
            End If
            tmDrf.iVefCode = 0
            If Trim$(slName) <> "" Then
                For ilLoop = LBound(tgVefRad) To UBound(tgVefRad) - 1 Step 1
                    If StrComp(slName, Trim$(tgVefRad(ilLoop).sName), 1) = 0 Then
                        tmDrf.iVefCode = tgVefRad(ilLoop).iCode
                        Exit For
                    End If
                Next ilLoop
                If tmDrf.iVefCode = 0 Then
                    ilAddFlag = False
                    ilFound = False
                    For ilLoop = LBound(smVehNotFound) To UBound(smVehNotFound) - 1 Step 1
                        If StrComp(slName, smVehNotFound(ilLoop), 1) = 0 Then
                            ilFound = True
                        End If
                    Next ilLoop
                    If Not ilFound Then
                        smVehNotFound(UBound(smVehNotFound)) = slName
                        ReDim Preserve smVehNotFound(0 To UBound(smVehNotFound) + 1) As String
                        Print #hmTo, "Unable to Find Vehicle " & slName & " record not added"
                        If gOkAddStrToListBox("Unable to Find Vehicle " & slName, lmLen, imShowMsg) Then
                            lbcErrors.AddItem "Unable to Find Vehicle " & slName
                        Else
                            imShowMsg = False
                        End If
                    End If
                End If
            End If
            ilFound = False
            For ilLoop = LBound(imVefCodeInDnf) To UBound(imVefCodeInDnf) - 1 Step 1
                If imVefCodeInDnf(ilLoop) = tmDrf.iVefCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                imVefCodeInDnf(UBound(imVefCodeInDnf)) = tmDrf.iVefCode
                'ReDim Preserve imVefCodeInDnf(1 To UBound(imVefCodeInDnf) + 1) As Integer
                ReDim Preserve imVefCodeInDnf(0 To UBound(imVefCodeInDnf) + 1) As Integer
            End If
            tmDrf.sInfoType = "T"
            tmDrf.iRdfCode = 0
            slStr = Mid$(slLine, 31, 2) 'Day Code
            ilDayValue = Val(slStr)
            slStr = Mid$(slLine, 3, 4)  '5) 'Demographic Prog Code
            Do While Len(slStr) < Len(tmDrf.sProgCode)
                slStr = "0" & slStr
            Loop
            tmDrf.sProgCode = slStr
            If (ilDayValue <> 12) And (ilDayValue <> 13) Then
                slStr = Mid$(slLine, 7, 5) 'Demographic Prog Name- first 5 characters contains time
                slStr = Left$(slStr, 2) & ":" & right(slStr, 3)
                If Not gValidTime(slStr) Then
                    slTime = slStr
                    'Obtain the time from the quarter hour index
                    slStr = Mid$(slLine, 35, 2) 'Quarter Hour Code
                    ilValue = Val(slStr)
                    ilHour = ilValue \ 4
                    ilMin = (ilValue - 1) Mod 4
                    Select Case ilMin
                        Case 0
                            slMin = "00"
                        Case 1
                            slMin = "15"
                        Case 2
                            slMin = "30"
                        Case 3
                            slMin = "45"
                    End Select
                    If ilHour = 0 Then
                        slStr = "12" & ":" & slMin & "AM"
                    ElseIf ilHour < 12 Then
                        slStr = Trim$(Str$(ilHour)) & ":" & slMin & "AM"
                    ElseIf ilHour = 12 Then
                        slStr = "12" & ":" & slMin & "PM"
                    Else
                        slStr = Trim$(Str$(ilHour - 12)) & ":" & slMin & "AM"
                    End If
                    Print #hmTo, "Invalid Time " & slTime & ", using Quarter Hour Time" & slStr
                    If gOkAddStrToListBox("Invalid Time, using Quarter Hour " & slStr, lmLen, imShowMsg) Then
                        lbcErrors.AddItem "Invalid Time, using Quarter Hour " & slStr
                    Else
                        imShowMsg = False
                    End If
                End If
                gPackTime slStr, tmDrf.iStartTime(0), tmDrf.iStartTime(1)
                gPackTime slStr, tmDrf.iEndTime(0), tmDrf.iEndTime(1)
            Else
                gPackTime "12AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
                gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
                tmDrf.sInfoType = "V"
                tmDrf.iRdfCode = 0
            End If
            tmDrf.iStartTime2(0) = 1
            tmDrf.iStartTime2(1) = 0
            tmDrf.iEndTime2(0) = 1
            tmDrf.iEndTime2(1) = 0
            For ilDay = 0 To 6 Step 1
                tmDrf.sDay(ilDay) = "N"
            Next ilDay
            Select Case ilDayValue
                'Case 1  'M-F
                '    For ilDay = 0 To 4 Step 1
                '        tmDrf.sDay(ilDay) = "Y"
                '    Next ilDay
                Case 2  'Sat
                    tmDrf.sDay(5) = "Y"
                Case 3  'Sun
                    tmDrf.sDay(6) = "Y"
                'Case 4  'T-S
                '    For ilDay = 1 To 6 Step 1
                '        tmDrf.sDay(ilDay) = "Y"
                '    Next ilDay
                'Case 5  'T-F
                '    For ilDay = 1 To 4 Step 1
                '        tmDrf.sDay(ilDay) = "Y"
                '    Next ilDay
                'Case 6  'W-S
                '    For ilDay = 2 To 6 Step 1
                '        tmDrf.sDay(ilDay) = "Y"
                '    Next ilDay
                Case 7  'Mon
                    tmDrf.sDay(0) = "Y"
                Case 8  'Tue
                    tmDrf.sDay(1) = "Y"
                Case 9  'Wed
                    tmDrf.sDay(2) = "Y"
                Case 10 'Thu
                    tmDrf.sDay(3) = "Y"
                Case 11 'Fri
                    tmDrf.sDay(4) = "Y"
                Case 12 'MF ROS
                    For ilDay = 0 To 4 Step 1
                        tmDrf.sDay(ilDay) = "Y"
                    Next ilDay
                Case 13 'MS ROS
                    For ilDay = 0 To 6 Step 1
                        tmDrf.sDay(ilDay) = "Y"
                    Next ilDay
                Case Else
                    ilAddFlag = False
                    Print #hmTo, "Invalid Day Code " & Str$(ilDayValue) & " on " & slName & " record not added"
                    If gOkAddStrToListBox("Invalid Day Code" & Str$(ilDayValue), lmLen, imShowMsg) Then
                        lbcErrors.AddItem "Invalid Day Code" & Str$(ilDayValue)
                    Else
                        imShowMsg = False
                    End If
            End Select
            slStr = Mid$(slLine, 35, 2) 'Quarter Hour Code
            ilValue = Val(slStr)
            tmDrf.iQHIndex = ilValue
            tmDrf.iCount = 0
            slStr = Mid$(slLine, 41, 1) 'Avg # of Stations
            If UCase$(slStr) = "X" Then
                tmDrf.sExStdDP = "Y"
            Else
                tmDrf.sExStdDP = "N"
            End If
            slStr = Mid$(slLine, 42, 1) 'Avg # of Stations
            If UCase$(slStr) = "Y" Then
                tmDrf.sExRpt = "Y"
            Else
                tmDrf.sExRpt = "N"
            End If
            tmDrf.sDataType = "A"
            For ilLoop = 1 To 18 Step 1
                tmDrf.lDemo(ilLoop - 1) = 0
            Next ilLoop
            ilSDemo = 45
            For ilLoop = 1 To imNoBuckets Step 1
                slStr = Mid$(slLine, ilSDemo, 5)
                If Trim$(slStr) <> "" Then
                    tmDrf.lDemo(ilLoop - 1) = Val(slStr)
                Else
                    tmDrf.lDemo(ilLoop - 1) = 0
                End If
                ilSDemo = ilSDemo + 5
            Next ilLoop
            If ilAddFlag Then
                tmDrf.sForm = smDataForm
                tmDrf.lCode = 0
                tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
                tmDrf.lAutoCode = tmDrf.lCode
                If tgSpf.sSAudData = "H" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 10 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If tgSpf.sSAudData = "N" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 100 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If tgSpf.sSAudData = "U" Then
                    For ilLoop = 1 To imNoBuckets Step 1
                        tmDrf.lDemo(ilLoop - 1) = 1000 * tmDrf.lDemo(ilLoop - 1)
                    Next ilLoop
                End If
                If mDrfDefined() Then
                    ilAddFlag = False
                End If
                If ilAddFlag Then
                    ilRet = btrInsert(hmDrf, tmDrf, imDrfRecLen, INDEXKEY2)
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
                End If
            End If
            lmProcessedNoBytes = lmProcessedNoBytes + 125
            llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
            If llPercent >= 100 Then
                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
                    llPercent = 99
                Else
                    llPercent = 100
                End If
            End If
            If lmPercent <> llPercent Then
                plcGauge.Value = llPercent
                lmPercent = llPercent
            End If
            slLine = ""
            slChar = ""
        Else
            slLine = slLine & slChar
        End If
    Loop
    Close hmFrom
    mConvTimeDay = True
    Exit Function
'mConvTimeDayErr:
'    ilRet = Err.Number
'    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mDrfDefined                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Test if record previously      *
'*                      defined                        *
'*                                                     *
'*******************************************************
Private Function mDrfDefined()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilDayMatch As Integer
    Dim ilDemoMatch As Integer
    Dim tlDrf As DRF

    tmDrfSrchKey.iDnfCode = tmDrf.iDnfCode
    tmDrfSrchKey.sDemoDataType = tmDrf.sDemoDataType
    tmDrfSrchKey.iMnfSocEco = tmDrf.iMnfSocEco
    tmDrfSrchKey.iVefCode = tmDrf.iVefCode
    tmDrfSrchKey.sInfoType = tmDrf.sInfoType
    tmDrfSrchKey.iRdfCode = tmDrf.iRdfCode
    ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmDrf.iDnfCode = tlDrf.iDnfCode) And (tmDrf.sDemoDataType = tlDrf.sDemoDataType) And (tmDrf.iMnfSocEco = tlDrf.iMnfSocEco) And (tmDrf.iVefCode = tlDrf.iVefCode) And (tmDrf.sInfoType = tlDrf.sInfoType) And (tmDrf.iRdfCode = tlDrf.iRdfCode)
        ilDayMatch = True
        For ilLoop = 0 To 6 Step 1
            If StrComp(tmDrf.sDay(ilLoop), tlDrf.sDay(ilLoop), 1) <> 0 Then
                ilDayMatch = False
                Exit For
            End If
        Next ilLoop
        If ilDayMatch Then
            If (tmDrf.iStartTime(0) = tlDrf.iStartTime(0)) And (tmDrf.iStartTime(1) = tlDrf.iStartTime(1)) And (tmDrf.iEndTime(0) = tlDrf.iEndTime(0)) And (tmDrf.iEndTime(1) = tlDrf.iEndTime(1)) Then
                If (tmDrf.iStartTime2(0) = tlDrf.iStartTime2(0)) And (tmDrf.iStartTime2(1) = tlDrf.iStartTime2(1)) And (tmDrf.iEndTime2(0) = tlDrf.iEndTime2(0)) And (tmDrf.iEndTime2(1) = tlDrf.iEndTime2(1)) Then
                    If (tmDrf.sProgCode = tlDrf.sProgCode) And (tmDrf.iQHIndex = tlDrf.iQHIndex) And (tmDrf.sDataType = tlDrf.sDataType) Then
                        ilDemoMatch = True
                        For ilLoop = 1 To imNoBuckets Step 1
                            If tmDrf.lDemo(ilLoop - 1) <> tlDrf.lDemo(ilLoop - 1) Then
                                ilDemoMatch = False
                                Exit For
                            End If
                        Next ilLoop
                        If ilDemoMatch Then
                            mDrfDefined = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        ilRet = btrGetNext(hmDrf, tlDrf, imDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mDrfDefined = False
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
'*      Procedure Name:mGetVehName                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get Vehicle Name from Network   *
'*                     code                            *
'*                                                     *
'*******************************************************
Private Function mGetVehName(slNetCode As String) As String
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVef                                                                                 *
'******************************************************************************************

    Dim slName As String
    Dim ilLoop As Integer
    Dim ilVpf As Integer

    slName = ""
    slNetCode = UCase$(slNetCode)
''    If slNetCode = "AX" Then
''        slName = "Excel"
''    ElseIf slNetCode = "AG" Then
'    If slNetCode = "AG" Then
'        slName = "Galaxy"
'    ElseIf slNetCode = "AN" Then
'        slName = "Genesis"
'    ElseIf slNetCode = "AL" Then
'        slName = "Platinum"
'    ElseIf slNetCode = "AP" Then
'        slName = "Prime"
'    ElseIf slNetCode = "AA" Then
'        slName = "Advantage Net"
'    Else
'        For ilLoop = LBound(tgVefRad) To UBound(tgVefRad) - 1 Step 1
'            If StrComp(slNetCode, Trim$(tgVefRad(ilLoop).sCodeStn), 1) = 0 Then
'                slName = Trim$(tgVefRad(ilLoop).sName)
'                Exit For
'            End If
'        Next ilLoop
'    End If
    '12/14/15: Was that testing the last record
    'For ilVpf = LBound(tgVpf) To UBound(tgVpf) - 1 Step 1
    For ilVpf = LBound(tgVpf) To UBound(tgVpf) Step 1
        If StrComp(slNetCode, Trim$(tgVpf(ilVpf).sRadarCode), vbTextCompare) = 0 Then
            For ilLoop = LBound(tgVefRad) To UBound(tgVefRad) - 1 Step 1
                If tgVefRad(ilLoop).iCode = tgVpf(ilVpf).iVefKCode Then
                    slName = Trim$(tgVefRad(ilLoop).sName)
                    Exit For
                End If
            Next ilLoop
            If slName <> "" Then
                Exit For
            End If
        End If
    Next ilVpf
    mGetVehName = slName
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
    imTerminate = False
    imFirstActivate = True
    'mParseCmmdLine
    Screen.MousePointer = vbHourglass
    bmResearchSaved = False
    imTestAddStdDemo = True
    imConverting = False
    imFirstFocus = True
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptRad
    On Error GoTo 0
    imRdfRecLen = Len(tmRdf)
    hmMnf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptRad
    On Error GoTo 0
    imMnfRecLen = Len(tmMnf)
    hmVef = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptRad
    On Error GoTo 0
    imVefRecLen = Len(tmVef)
    hmDrf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptRad
    On Error GoTo 0
    imDrfRecLen = Len(tmDrf)
    hmDnf = CBtrvTable(TWOHANDLES) 'CBtrvObj
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptRad
    On Error GoTo 0
    imDnfRecLen = Len(tmDnf)
'    hmDsf = CBtrvTable(TWOHANDLES) 'CBtrvObj
'    ilRet = btrOpen(hmDsf, "", sgDBPath & "Dsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptRad
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
    ilRet = gObtainRcfRifRdf()
    'ilRet = mObtainDaypart()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Daypart Error", vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
        imTerminate = True
        Exit Sub
    End If
    ilRet = mObtainVehicles()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Vehicle Error", vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
        imTerminate = True
        Exit Sub
    End If
    ilRet = mObtainSocEco()
    If ilRet = False Then
        Screen.MousePointer = vbDefault
        MsgBox "Obtain Soc Eco Groups Error", vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
        imTerminate = True
        Exit Sub
    End If
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)

    'smRptTime = Format$(Now, "h:m:s AM/PM")
    'gPackTime smRptTime, tmIcf.iTime(0), tmIcf.iTime(1)
    gCenterStdAlone ImptRad
    If mTestRecLengths() Then
        Screen.MousePointer = vbDefault
        imTerminate = True
        Exit Sub
    End If
    'Test if Radar file exist
    ilRet = 0
    'On Error GoTo mInit1Err:
    'hmFrom = FreeFile
    'Open sgImportPath & "ABC0" For Input Access Read As hmFrom
    ilRet = gFileOpen(sgImportPath & "ABC0", "Input Access Read", hmFrom)
    If ilRet = 0 Then
        'On Error GoTo mInit1Err:
        'Line Input #hmFrom, slLine
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            Do While slChar = Chr(9)
                slChar = Input(1, #hmFrom)
            Loop
            If slChar = Chr(13) Then
                slChar = Input(1, #hmFrom)
            End If
            If slChar = Chr(10) Then
                Exit Do
            End If
            slLine = slLine & slChar
        Loop
        slLine = Trim$(slLine)
        Close hmFrom
        'Test if Book Name Exist
        ilFound = False
        For ilLoop = LBound(tgDnfBook) To UBound(tgDnfBook) - 1 Step 1
            If StrComp(slLine, Trim$(tgDnfBook(ilLoop).sBookName), 1) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            edcFrom.Text = sgImportPath & "ABC0"
            edcBookName.Text = slLine
        End If
    End If
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = Traffic!imcHelp.Picture
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
'mInit1Err:
'    ilRet = Err.Number
'    Resume Next
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
'*      Procedure Name:mObtainSocEco                   *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgMnfSocEcoRad            *
'*                                                     *
'*******************************************************
Private Function mObtainSocEco() As Integer
'
'   ilRet = mObtainSocEco ()
'   Where:
'       tgMnfSocEcoRad() (I)- MNF record structure
'       ilRet (O)- True = populated; False = error
'
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim slName As String
    Dim slGroup As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record

    'ReDim tgMnfSocEcoRad(1 To 1) As MNF
    ReDim tgMnfSocEcoRad(0 To 0) As MNF
    ilRecLen = Len(tmMnf) 'btrRecordLength(hmMnf)  'Get and save record length
    'llNoRec = btrRecords(hmMnf) 'Obtain number of records
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlFile) 'Obtain number of records
    btrExtClear hmMnf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmMnf, tmMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        mObtainSocEco = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mObtainSocEco = False
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hmMnf, llNoRec, 0, "UC", "MNF", "") 'Set extract limits (all records)
    tlCharTypeBuff.sType = "F"
    ilOffSet = 2 'gFieldOffset("Mnf", "MnfType")
    ilRet = btrExtAddLogicConst(hmMnf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
    ilOffSet = 0 'gFieldOffset("Mnf", "MnfCode")
    ilRet = btrExtAddField(hmMnf, ilOffSet, ilRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainSocEco = False
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hmMnf)    'Extract record
    'If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
    '    If ilRet <> BTRV_ERR_NONE Then
    '        mObtainSocEco = False
    '        Exit Function
    '    End If
    'End If
    'ilUpperBound = UBound(tgMnfSocEcoRad)
    'ilExtLen = Len(tgMnfSocEcoRad(ilUpperBound))  'Extract operation record size
    'ilRet = btrExtGetFirst(hmMnf, tgMnfSocEcoRad(ilUpperBound), ilExtLen, llRecPos)
    'Do While ilRet = BTRV_ERR_NONE
    '    ilUpperBound = ilUpperBound + 1
    '    ReDim Preserve tgMnfSocEcoRad(1 To ilUpperBound) As MNF
    '    ilRet = btrExtGetNext(hmMnf, tgMnfSocEcoRad(ilUpperBound), ilExtLen, llRecPos)
    'Loop
    ilUpperBound = UBound(tgMnfSocEcoRad)
    ilExtLen = Len(tgMnfSocEcoRad(ilUpperBound))  'Extract operation record size
    ilRet = btrExtGetNext(hmMnf, tgMnfSocEcoRad(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            mObtainSocEco = False
            Exit Function
        End If
        ilExtLen = Len(tgMnfSocEcoRad(ilUpperBound))  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmMnf, tgMnfSocEcoRad(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilUpperBound = ilUpperBound + 1
            'ReDim Preserve tgMnfSocEcoRad(1 To ilUpperBound) As MNF
            ReDim Preserve tgMnfSocEcoRad(0 To ilUpperBound) As MNF
            ilRet = btrExtGetNext(hmMnf, tgMnfSocEcoRad(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmMnf, tgMnfSocEcoRad(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    mObtainSocEco = True
    If ilUpperBound = LBound(tgMnfSocEcoRad) Then
        For ilLoop = 1 To 135 Step 1
            'tmMnf.iCode = 0
            'tmMnf.sType = "F"
            gGetGroupName ilLoop, slName, slGroup
            'tmMnf.sName = slName
            'tmMnf.sRPU = ""
            'tmMnf.sUnitType = slGroup
            'tmMnf.iMerge = 0
            'tmMnf.iGroupNo = 0
            'tmMnf.sCodeStn = ""
            'ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
            'tgMnfSocEcoRad(ilUpperBound) = tmMnf
            'ilUpperBound = ilUpperBound + 1
            'ReDim Preserve tgMnfSocEcoRad(1 To ilUpperBound) As MNF
            mAddGroupName slName, slGroup
        Next ilLoop
    End If
    Exit Function
End Function
Private Function mObtainVehicles() As Integer
'
'   ilRet = mObtainVehicles ()
'   Where:
'       tgVeh() (I)- Vef record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim ilRet As Integer
    Dim ilUpperBound As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilOffSet As Integer
    Dim llRecPos As Long

    'ReDim tgVefRad(1 To 1) As VEF
    ReDim tgVefRad(0 To 0) As VEF
    ilUpperBound = UBound(tgVefRad)
    'ilExtLen = Len(tgVefRad(1)) 'btrRecordLength(hmAgf)  'Get and save record length
    ilExtLen = Len(tgVefRad(0)) 'btrRecordLength(hmAgf)  'Get and save record length
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hmVef) 'Obtain number of records
    btrExtClear hmVef   'Clear any previous extend operation
    'ilRet = btrGetFirst(hmVef, tgVefRad(1), imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ilRet = btrGetFirst(hmVef, tgVefRad(0), imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        mObtainVehicles = True
        Exit Function
    End If
    'Do While ilRet = BTRV_ERR_NONE
    '    tgVefRad(ilUpperBound) = tmVef
    '    ilUpperBound = ilUpperBound + 1
    '    ReDim Preserve tgVefRad(1 To ilUpperBound) As VEF
    '    ilRet = btrGetNext(hmVef, tmVef, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    'Loop
    Call btrExtSetBounds(hmVef, llNoRec, -1, "UC", "VEF", "") 'Set extract limits (all records including first)
    ilOffSet = 0
    ilRet = btrExtAddField(hmVef, ilOffSet, imVefRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainVehicles = False
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hmVef)    'Extract record
    ilRet = btrExtGetNext(hmVef, tgVefRad(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            mObtainVehicles = False
            Exit Function
        End If
        'ilExtLen = Len(tgVefRad(1))  'Extract operation record size
        ilExtLen = Len(tgVefRad(0))  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmVef, tgVefRad(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilUpperBound = ilUpperBound + 1
            'ReDim Preserve tgVefRad(1 To ilUpperBound) As VEF
            ReDim Preserve tgVefRad(0 To ilUpperBound) As VEF
            ilRet = btrExtGetNext(hmVef, tgVefRad(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmVef, tgVefRad(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    mObtainVehicles = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mRemovePrevDnf                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove record of previouly     *
'*                      imported book                  *
'*                                                     *
'*******************************************************
Private Function mRemovePrevDnf(ilPrevDnfCode As Integer) As Integer
    Dim ilRet As Integer
    Dim tlDrf As DRF
    Dim tlDnf As DNF
    Do
        tmDrfSrchKey.iDnfCode = ilPrevDnfCode
        tmDrfSrchKey.sDemoDataType = ""
        tmDrfSrchKey.iMnfSocEco = 0
        tmDrfSrchKey.iVefCode = 0
        tmDrfSrchKey.sInfoType = ""
        tmDrfSrchKey.iRdfCode = 0
        ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            Exit Do
        End If
        If tlDrf.iDnfCode <> ilPrevDnfCode Then
            Exit Do
        End If
        'tmRec = tlDrf
        'ilRet = gGetByKeyForUpdate("DRF", hmDrf, tmRec)
        'tlDrf = tmRec
        'If ilRet <> BTRV_ERR_NONE Then
        '    mRemovePrevDnf = False
        '    ilRet = MsgBox("Remove Not Completed, Try Later", vbOkOnly + vbExclamation, "Remove")
        '    Exit Function
        'End If
        ilRet = btrDelete(hmDrf)
    Loop
    ilRet = BTRV_ERR_NONE
    Do
        tmDnfSrchKey.iCode = ilPrevDnfCode
        ilRet = btrGetEqual(hmDnf, tlDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            mRemovePrevDnf = False
            ilRet = MsgBox("Remove Not Completed, Try Later", vbOKOnly + vbExclamation, "Remove")
            Exit Function
        End If
        ilRet = btrDelete(hmDnf)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        mRemovePrevDnf = False
        ilRet = MsgBox("Remove Not Completed, Try Later", vbOKOnly + vbExclamation, "Remove")
        Exit Function
    End If
'    If tgSpf.sRemoteUsers = "Y" Then
'        tmDsf.lCode = 0
'        tmDsf.sFileName = "DNF"
'        gPackDate smSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
'        gPackTime smSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
'        tmDsf.iRemoteID = tlDnf.iRemoteID
'        tmDsf.lAutoCode = tlDnf.iAutoCode
'        tmDsf.iSourceID = tgUrf(0).iRemoteUserID
'        tmDsf.lCntrNo = 0
'        ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
'    End If
    mRemovePrevDnf = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetBookName                    *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set Book Name                   *
'*                                                     *
'*******************************************************
Private Sub mSetBookName()
    Dim slFromName As String
    Dim slLine As String
    Dim ilRet As Integer
    Dim slChar As String
    slFromName = Trim$(edcFrom.Text)
    If slFromName = "" Then
        edcBookName.Text = ""
        Exit Sub
    End If
    ilRet = 0
    'On Error GoTo mSetBookNameErr:
    'hmFrom = FreeFile
    If InStr(slFromName, ":") > 0 Then
        'Open slFromName For Input Access Read As hmFrom
        ilRet = gFileOpen(slFromName, "Input Access Read", hmFrom)
    Else
        'Open sgImportPath & slFromName For Input Access Read As hmFrom
        ilRet = gFileOpen(sgImportPath & slFromName, "Input Access Read", hmFrom)
    End If
    If ilRet = 0 Then
        'On Error GoTo mSetBookNameErr:
        'Line Input #hmFrom, slLine
        Do While Not EOF(hmFrom)
            slChar = Input(1, #hmFrom)
            Do While slChar = Chr(9)
                slChar = Input(1, #hmFrom)
            Loop
            If slChar = Chr(13) Then
                slChar = Input(1, #hmFrom)
            End If
            If slChar = Chr(10) Then
                Exit Do
            End If
            slLine = slLine & slChar
        Loop
        slLine = Trim$(slLine)
        Close hmFrom
        edcBookName.Text = slLine
    Else
        edcBookName.Text = ""
    End If
    Exit Sub
'mSetBookNameErr:
'    ilRet = Err.Number
'    Resume Next
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

'    ilRet = btrClose(hmDsf)
'    btrDestroy hmDsf
    Dim ilRet As Integer
    
    Screen.MousePointer = vbDefault
    
    If bmResearchSaved Then
        If (Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES Then
            ilRet = MsgBox("Please update the Vehicle default ACT1 Lineup codes if required", vbOKOnly + vbInformation, "Warning")
        End If
    End If
    
    igManUnload = YES
    Unload ImptRad
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
    ilSize = mGetRecLength("Rdf.Btr")
    If ilSize <> Len(tmRdf) Then
        If ilSize > 0 Then
            MsgBox "Rdf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmRdf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            ilSizeError = True
        Else
            MsgBox "Rdf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Initialize Error"
            ilSizeError = True
        End If
    End If
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
    plcScreen.Print "Import RADAR Data"
End Sub

Private Sub mConvertDayparts(slStr As String, slStrTime As String, slName As String, ilAddFlag As Integer, ilNoTimes As Integer)
    Dim ilPos As Integer
    Dim ilFound As Integer
    Dim slStartTime As String
    Dim slEndTime As String

    If slStr = "12M-12M" Then
        gPackTime "12AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "5A-12M" Then
        gPackTime "5AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "5A-8P" Then
        gPackTime "5AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "8PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "5A-7P" Then
        gPackTime "5AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "7PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "6A-12M" Then
        gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "6A-8P" Then
        gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "8PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "6A-7P" Then
        gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "7PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "12M-6A" Then
        gPackTime "12AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "6AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "12M-5A" Then
        gPackTime "12AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "5AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "5A-10A" Then
        gPackTime "5AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "10AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "6A-10A" Then
        gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "10AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "10A-3P" Then
        gPackTime "10AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "3PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "3P-8P" Then
        gPackTime "3PM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "8PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "3P-7P" Then
        gPackTime "3PM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "7PM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "7P-12M" Then
        gPackTime "7PM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "8P-12M" Then
        gPackTime "8PM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "12AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
    ElseIf slStr = "6-10+3-7" Then
        gPackTime "6AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "10AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
        gPackTime "3PM", tmDrf.iStartTime2(0), tmDrf.iStartTime2(1)
        gPackTime "7PM", tmDrf.iEndTime2(0), tmDrf.iEndTime2(1)
        ilNoTimes = 2
    ElseIf slStr = "5-10+3-8" Then
        gPackTime "5AM", tmDrf.iStartTime(0), tmDrf.iStartTime(1)
        gPackTime "10AM", tmDrf.iEndTime(0), tmDrf.iEndTime(1)
        gPackTime "3PM", tmDrf.iStartTime2(0), tmDrf.iStartTime2(1)
        gPackTime "8PM", tmDrf.iEndTime2(0), tmDrf.iEndTime2(1)
        ilNoTimes = 2
    Else
        ilFound = False
        ilPos = InStr(1, slStr, "-", vbTextCompare)
        If ilPos > 0 Then
            slStartTime = Left(slStr, ilPos - 1)
            slEndTime = Mid(slStr, ilPos + 1)
            If (gValidTime(slStartTime)) And (gValidTime(slEndTime)) Then
                gPackTime slStartTime, tmDrf.iStartTime(0), tmDrf.iStartTime(1)
                gPackTime slEndTime, tmDrf.iEndTime(0), tmDrf.iEndTime(1)
                ilFound = True
            End If
        End If
        If Not ilFound Then
            ilAddFlag = False
            Print #hmTo, "Unable to Find Daypart Time " & slStrTime & " on " & slName & " record not added"
            If gOkAddStrToListBox("Unable to Find Daypart Time " & slStrTime, lmLen, imShowMsg) Then
                lbcErrors.AddItem "Unable to Find Daypart Time " & slStrTime
            Else
                imShowMsg = False
            End If
        End If
    End If

End Sub
