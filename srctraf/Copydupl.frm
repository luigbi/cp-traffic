VERSION 5.00
Begin VB.Form CopyDupl 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4545
   ClientLeft      =   2115
   ClientTop       =   2160
   ClientWidth     =   4155
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4545
   ScaleWidth      =   4155
   Begin VB.CommandButton cmcSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
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
      TabIndex        =   6
      Top             =   4140
      Width           =   1050
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3585
      Width           =   75
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
      Left            =   1485
      TabIndex        =   5
      Top             =   4140
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
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
      Left            =   210
      TabIndex        =   4
      Top             =   4140
      Width           =   1050
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
      Height          =   495
      Left            =   30
      ScaleHeight     =   495
      ScaleWidth      =   3990
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   3990
   End
   Begin VB.PictureBox plcLine 
      ForeColor       =   &H00000000&
      Height          =   3330
      Left            =   180
      ScaleHeight     =   3270
      ScaleWidth      =   3720
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   540
      Width           =   3780
      Begin VB.ListBox lbcMedia 
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
         Height          =   3180
         Left            =   30
         TabIndex        =   2
         Top             =   45
         Width           =   1020
      End
      Begin VB.ListBox lbcInv 
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
         Height          =   3180
         Left            =   1080
         TabIndex        =   3
         Top             =   45
         Width           =   2625
      End
   End
End
Attribute VB_Name = "CopyDupl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Copydupl.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CopyDupl.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Price Grid Calculate input screen code
Option Explicit
Option Compare Text
'Media Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim tmMediaCode() As SORTCODE
Dim smMediaCodeTag As String
Dim smScreenCaption As String
'Copy inventory file
Dim hmCif As Integer 'Copy inventory file handle
Dim tmCif As CIF        'CIF record image
Dim tmCifSrchKey As LONGKEY0    'CIF key record image
Dim imCifRecLen As Integer        'CIF record length
Dim imCifIndex As Integer
'Copy product/ISCI file
Dim hmCpf As Integer 'Copy product/ISCI file handle
Dim tmCpf As CPF        'CPF record image
Dim tmCpfSrchKey As LONGKEY0    'CPF key record image
Dim tmCpfSrchKey1 As CPFKEY1    'CPF key record image
Dim imCpfRecLen As Integer        'CPF record length
'Media code file
Dim hmMcf As Integer 'Media file handle
Dim tmMcf As MCF        'MCF record image
Dim tmMcfSrchKey As INTKEY0    'MCF key record image
Dim imMcfRecLen As Integer        'MCF record length
Dim imMcfIndex As Integer
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcDone_Click()
    Dim ilRes As Integer
    Dim slMess As String

    If (lbcMedia.ListIndex >= 0) And (lbcInv.ListIndex >= 0) Then
        slMess = "Duplicate to " & Trim$(lbcMedia.List(lbcMedia.ListIndex)) & lbcInv.List(lbcInv.ListIndex)
        ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
        If ilRes = vbCancel Then
            Exit Sub
        End If
        If ilRes = vbYes Then
            ilRes = mSaveRec()
        End If
    End If
    mTerminate
End Sub
Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slName As String
    ilRet = mSaveRec()
    If ilRet Then
        'Remove used media code
        lbcMedia.Clear
        lbcInv.Clear
        For ilIndex = imMcfIndex To UBound(tmMediaCode) - 1 Step 1
            tmMediaCode(ilIndex) = tmMediaCode(ilIndex + 1)
        Next ilIndex
        ReDim Preserve tmMediaCode(LBound(tmMediaCode) To UBound(tmMediaCode) - 1) As SORTCODE
        For ilLoop = 0 To UBound(tmMediaCode) - 1 Step 1
            slNameCode = tmMediaCode(ilLoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            If ilRet = CP_MSG_NONE Then
                slName = Trim$(slName)
                lbcMedia.AddItem slName  'Add ID to list box
            End If
        Next ilLoop
    End If
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

Private Sub Form_Load()
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    
    Erase tmMediaCode
    Erase tgPayableCode
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf

    Set CopyDupl = Nothing   'Remove data segment

End Sub

Private Sub lbcInv_Click()
    imCifIndex = lbcInv.ListIndex
End Sub
Private Sub lbcMedia_Click()
    imMcfIndex = lbcMedia.ListIndex
    lbcInv.Clear
    Screen.MousePointer = vbHourglass
    mReadMcf
    mInvPop
    Screen.MousePointer = vbDefault
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
    Screen.MousePointer = vbHourglass
    imFirstActivate = True
    imTerminate = False
    CopyDupl.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    'gCenterModalForm CopyDupl
    smScreenCaption = "Duplicate Copy for " & sgISCI & " " & sgCreativeTitle
    mMediaPop
    hmCif = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "CIF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CIF.Btr)", CopyDupl
    On Error GoTo 0
    imCifRecLen = Len(tmCif)  'Get and save CIF record length
    hmCpf = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "CPF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: CPF.Btr)", CopyDupl
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)  'Get and save CPF record length
    hmMcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "MCF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: MCF.Btr)", CopyDupl
    On Error GoTo 0
    imMcfRecLen = Len(tmMcf)  'Get and save MCF record length
    If tgCif.lcpfCode > 0 Then
        tmCpfSrchKey.lCode = tgCif.lcpfCode
        ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmCpf.sName = ""
            tmCpf.sISCI = ""
            tmCpf.sCreative = ""
        End If
    Else
        tmCpf.lCode = 0
        tmCpf.sName = ""
        tmCpf.sISCI = ""
        tmCpf.sCreative = ""
    End If
    plcScreen_Paint
    gCenterStdAlone CopyDupl
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInvPop                         *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the inventory         *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mInvPop()
'
'   mInvPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slNowDate As String
    If lbcMedia.ListIndex < 0 Then
        Exit Sub
    End If
    slNowDate = Format$(gNow(), "m/d/yy")
    'If rbcPurged(0).Value Then
        ilRet = gPopCopyForMediaBox(CopyDupl, tmMcf.iCode, slNowDate, True, True, igSortCart, lbcInv, tgPayableCode(), sgPayableCodeTag)    'lbcInvCode)
    'Else
    '    ilRet = gPopCopyForMediaBox(CopyDupl, tmMcf.iCode, slNowDate, True, False, lbcInv, tgPayableCode(), sgPayableCodeTag)  'lbcInvCode)
    'End If
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mInvPopErr
        gCPErrorMsg ilRet, "mInvPop (gIMoveListBox)", CopyDupl
        On Error GoTo 0
    End If
    Exit Sub
mInvPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestISCI                       *
'*                                                     *
'*             Created:4/21/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test is ISCI is unique          *
'*                                                     *
'*******************************************************
Private Function mISCIOk(sISCI As String) As Integer
    Dim tlCpf As CPF
    Dim ilRet As Integer
    Dim hlCif As Integer        'Cif handle
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilCifRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlCif As CIF
    Dim ilOffSet As Integer
    Dim ilTest As Integer
    Dim tlLTypeBuff As POPLCODE   'Type field record
    mISCIOk = True
    If Trim$(sISCI) = "" Then
        Exit Function
    End If
    hlCif = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        On Error GoTo mISCIOkErr
        gBtrvErrorMsg ilRet, "mISCIOk (btrOpen):" & "Cif.Btr", CopyDupl
        On Error GoTo 0
        Exit Function
    End If
    ilCifRecLen = Len(tlCif) 'btrRecordLength(hlAdf)  'Get and save record length
    tmCpfSrchKey1.sISCI = sISCI 'smSave(6)
    ilRet = btrGetEqual(hmCpf, tlCpf, imCpfRecLen, tmCpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (Trim$(tlCpf.sISCI) = Trim$(sISCI))
        'If tgCif.lCpfCode <> tlCpf.lCode Then
            'Test if inventory that is referencing this CPF is Purged or
            'History- if so, then ISCI Ok
            ilExtLen = Len(tlCif)  'Extract operation record size
            llNoRec = gExtNoRec(ilExtLen)  'btrRecords(hlCif) 'Obtain number of records
            btrExtClear hlCif   'Clear any previous extend operation
            ilRet = btrGetFirst(hlCif, tlCif, ilCifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_END_OF_FILE Then
                ilRet = btrClose(hlCif)
                btrDestroy hlCif
                Exit Function
            End If
            If ilRet <> BTRV_ERR_NONE Then
                On Error GoTo mISCIOkErr
                gBtrvErrorMsg ilRet, "mISCIOk (btrGetFirst):" & "Cif.Btr", CopyDupl
                On Error GoTo 0
                Exit Function
            End If
            Call btrExtSetBounds(hlCif, llNoRec, -1, "UC", "CIF", "") 'Set extract limits (all records including first)
            tlLTypeBuff.lCode = tlCpf.lCode
            ilOffSet = gFieldOffset("Cif", "CIFCPFCODE")
            ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlLTypeBuff, 4)
            ilOffSet = 0
            ilRet = btrExtAddField(hlCif, ilOffSet, Len(tlCif))  'Extract the whole record
            ilRet = btrExtGetNext(hlCif, tlCif, ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                    On Error GoTo mISCIOkErr
                    gBtrvErrorMsg ilRet, "mISCIOk (btrExtGetNext):" & "Cif.Btr", CopyDupl
                    On Error GoTo 0
                    Exit Function
                End If
                ilExtLen = Len(tlCif)  'Extract operation record size
                'ilRet = btrExtGetFirst(hlCif, tlCifExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlCif, tlCif, ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    ilTest = True
'                    If (tgSpf.sUseCartNo <> "N") And (tlCif.iMcfCode <> 0) Then
                        If tmMcf.iCode <> tlCif.iMcfCode Then
                            If tgCif.iAdfCode = tlCif.iAdfCode Then
                                ilTest = False
                            End If
                        End If
'                    End If
                    If (tlCif.sPurged = "A") And (ilTest) Then
                        ilRet = btrClose(hlCif)
                        btrDestroy hlCif
                        mISCIOk = False
                        Exit Function
                    End If
                    ilRet = btrExtGetNext(hlCif, tlCif, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlCif, tlCif, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        'End If
        ilRet = btrGetNext(hmCpf, tlCpf, imCpfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilRet = btrClose(hlCif)
    btrDestroy hlCif
    mISCIOk = True
    Exit Function
mISCIOkErr:
    ilRet = btrClose(hlCif)
    btrDestroy hlCif
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMediaPop                       *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the media combo       *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mMediaPop()
'
'   mMediaPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilfilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoop As Integer
    ilfilter(0) = NOFILTER
    slFilter(0) = ""
    ilOffSet(0) = 0
    lbcMedia.Clear
    'ilRet = gIMoveListBox(CopyInv, cbcMedia, lbcMediaCode, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilFilter(), slFilter(), ilOffset())
    ilRet = gIPopListBox(CopyDupl, tmMediaCode(), smMediaCodeTag, "Mcf.btr", gFieldOffset("Mcf", "McfName"), 6, ilfilter(), slFilter(), ilOffSet())    'Repopulate Traffic list box if required
    If ilRet Then
        For ilLoop = 0 To UBound(tmMediaCode) - 1 Step 1  'lbcMediaCode.ListCount - 1 Step 1
            slNameCode = tmMediaCode(ilLoop).sKey    'lbcMediaCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            If tgCif.iMcfCode = Val(slCode) Then
                For ilIndex = ilLoop To UBound(tmMediaCode) - 1 Step 1
                    tmMediaCode(ilIndex) = tmMediaCode(ilIndex + 1)
                Next ilIndex
                ReDim Preserve tmMediaCode(LBound(tmMediaCode) To UBound(tmMediaCode) - 1) As SORTCODE
                Exit For
            End If
        Next ilLoop
        For ilLoop = 0 To UBound(tmMediaCode) - 1 Step 1
            slNameCode = tmMediaCode(ilLoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            If ilRet = CP_MSG_NONE Then
                slName = Trim$(slName)
                lbcMedia.AddItem slName  'Add ID to list box
            End If
        Next ilLoop
    End If
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadMcf                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Sub mReadMcf()
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    If imMcfIndex < 0 Then
        Exit Sub
    End If
    slNameCode = tmMediaCode(imMcfIndex).sKey    'lbcMediaCode.List(imMcfIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadMcfErr
    gCPErrorMsg ilRet, "mReadMcfErr (gParseItem field 2)", CopyDupl
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmMcfSrchKey.iCode = CInt(slCode)
    ilRet = btrGetEqual(hmMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo mReadMcfErr
    gBtrvErrorMsg ilRet, "mReadMcfErr (btrGetEqual: Media Code)", CopyDupl
    On Error GoTo 0
    Exit Sub
mReadMcfErr:
    On Error GoTo 0
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:5/14/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim slNameCode As String
    Dim slCode As String
    Dim tlCif As CIF
    If lbcMedia.ListIndex < 0 Then
        mSaveRec = False
        Exit Function
    End If
    If lbcInv.ListIndex < 0 Then
        mSaveRec = False
        Exit Function
    End If
    If Not mISCIOk(tmCpf.sISCI) Then
        ilRet = MsgBox("ISCI Previously Used", vbOKOnly + vbExclamation, "Incomplete")
        mSaveRec = False
        Exit Function
    End If
    slNameCode = tgPayableCode(imCifIndex).sKey   'lbcInvCode.List(imCifIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    slCode = Trim$(slCode)
    tmCifSrchKey.lCode = CLng(slCode)
    ilRet = btrGetEqual(hmCif, tlCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
    tmCif = tgCif
    tmCif.lCode = tlCif.lCode
    tmCif.iMcfCode = tlCif.iMcfCode
    tmCif.sName = tlCif.sName
    tmCif.sCut = tlCif.sCut
    tmCif.iNoTimesAir = tlCif.iNoTimesAir
    If tmCif.sPurged = "P" Then
        slDate = Format$(gNow(), "m/d/yy")
        gPackDate slDate, tmCif.iPurgeDate(0), tmCif.iPurgeDate(1)
    ElseIf tmCif.sPurged = "H" Then
        slDate = Format$(gNow(), "m/d/yy")
        gPackDate slDate, tmCif.iPurgeDate(0), tmCif.iPurgeDate(1)
    Else
        tmCif.iPurgeDate(0) = 0
        tmCif.iPurgeDate(1) = 0
    End If
    slDate = Format$(gNow(), "m/d/yy")
    gPackDate slDate, tmCif.iDateEntrd(0), tmCif.iDateEntrd(1)
    tmCif.iUsedDate(0) = 0
    tmCif.iUsedDate(1) = 0
    tmCif.iRotStartDate(0) = 0
    tmCif.iRotStartDate(1) = 0
    tmCif.iRotEndDate(0) = 0
    tmCif.iRotEndDate(1) = 0
    tmCif.iUrfCode = tgUrf(0).iCode
    ilRet = btrUpdate(hmCif, tmCif, imCifRecLen)
    On Error GoTo mSaveRecErr
    gBtrvErrorMsg ilRet, "mSaveRec (btrInsert: Cif)", CopyDupl
    On Error GoTo 0
    mSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
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
    igManUnload = YES
    Unload CopyDupl
    igManUnload = NO
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub
