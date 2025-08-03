VERSION 5.00
Begin VB.Form FdCartNo 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3450
   ClientLeft      =   765
   ClientTop       =   2175
   ClientWidth     =   8220
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
   ScaleHeight     =   3450
   ScaleWidth      =   8220
   Begin VB.PictureBox plcAdvt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   210
      ScaleHeight     =   1590
      ScaleWidth      =   7695
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1275
      Width           =   7755
      Begin VB.PictureBox pbcLbcAdvt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   75
         ScaleHeight     =   1050
         ScaleWidth      =   7560
         TabIndex        =   12
         Top             =   300
         Width           =   7560
      End
      Begin VB.ListBox lbcAdvt 
         Appearance      =   0  'Flat
         Height          =   1080
         Left            =   60
         TabIndex        =   11
         Top             =   285
         Width           =   7590
      End
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   255
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2505
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   -15
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   285
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   525
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
      Left            =   15
      ScaleHeight     =   165
      ScaleWidth      =   75
      TabIndex        =   6
      Top             =   1770
      Width           =   75
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   30
      ScaleHeight     =   240
      ScaleWidth      =   1290
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -30
      Width           =   1290
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   3585
      TabIndex        =   5
      Top             =   3060
      Width           =   945
   End
   Begin VB.PictureBox plcInvNo 
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
      Height          =   915
      Left            =   1965
      ScaleHeight     =   855
      ScaleWidth      =   4200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   4260
      Begin VB.CommandButton cmcFind 
         Appearance      =   0  'Flat
         Caption         =   "&Find"
         Height          =   285
         Left            =   1650
         TabIndex        =   4
         Top             =   510
         Width           =   945
      End
      Begin VB.TextBox edcFindInv 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2550
         TabIndex        =   3
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label lacFind 
         Appearance      =   0  'Flat
         Caption         =   "Find Advertiser for Copy #"
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
         Height          =   240
         Left            =   105
         TabIndex        =   2
         Top             =   150
         Width           =   2415
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   135
      Top             =   2970
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "FdCartNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Fdcartno.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: FdCartNo.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Cart number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim tmCif As CIF        'Cif record image
Dim tmCifSrchKey As CIFKEY1    'Cif key record image
Dim hmCif As Integer    'Copy inventory file handle
Dim imCifRecLen As Integer        'Cif record length
Dim tmCpf As CPF        'Cpf record image
Dim tmCpfSrchKey As LONGKEY0    'Cpf key record image
Dim hmCpf As Integer    'Copy product file handle
Dim imCpfRecLen As Integer        'Cpf record length
Dim tmAdf As ADF        'Adf record image
Dim tmAdfSrchKey As INTKEY0    'Rvf key record image
Dim hmAdf As Integer    'Advertiser file handle
Dim imAdfRecLen As Integer        'ADF record length
Dim tmMcf() As MCF        'Mcf record image
Dim hmMcf As Integer    'Media code file handle
Dim imMcfRecLen As Integer        'McF record length
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
'Dim imListField(1 To 5) As Integer
Dim imListField(0 To 5) As Integer  'index zero ignored
Dim imLBCtrls As Integer

'Dim imListFieldChar(1 To 4) As Integer
Dim tmFdCartNoInfo() As FDCRTNOINFO

Private Sub cmcDone_Click()
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcFind_Click()
    Dim ilRet As Integer
    Dim slInvNo As String
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilLen As Integer
    Dim ilMcfCode As Integer
    Dim slCut As String
    Dim ilPos As Integer
    Dim slStatus As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slProduct As String
    Dim slDateRange As String
    Dim slAdvtName As String

    lbcAdvt.Clear
    ReDim tmFdCartNoInfo(0 To 0) As FDCRTNOINFO
    slInvNo = edcFindInv.Text
    'First find media code
    ilFound = False
    For ilLoop = 0 To UBound(tmMcf) - 1 Step 1
        slStr = Trim$(tmMcf(ilLoop).sName)
        ilLen = Len(slStr)
        If StrComp(Left$(slInvNo, ilLen), slStr, 1) = 0 Then
            ilFound = True
            ilMcfCode = tmMcf(ilLoop).iCode
            slInvNo = right$(slInvNo, Len(slInvNo) - ilLen)
            Exit For
        End If
    Next ilLoop
    If Not ilFound Then
        cmcDone.SetFocus
        Exit Sub
    End If
    slCut = ""
    ilPos = InStr(slInvNo, "-")
    If ilPos > 0 Then
        slCut = Mid$(slInvNo, ilPos + 1)
        slInvNo = Left$(slInvNo, ilPos - 1)
    End If
    tmCifSrchKey.iMcfCode = ilMcfCode
    tmCifSrchKey.sName = slInvNo
    tmCifSrchKey.sCut = slCut
    ilRet = btrGetGreaterOrEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmCif.iMcfCode = ilMcfCode) And (Trim$(tmCif.sName) = slInvNo) And (Trim$(tmCif.sCut) = slCut)
        tmAdfSrchKey.iCode = tmCif.iAdfCode
        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            ilFound = False
            '10/12/09:  Allow for same advertisers being reassigned the same cart number
            'For ilLoop = 0 To lbcAdvt.ListCount - 1 Step 1
            '    'If (StrComp(Trim$(tmAdf.sName), lbcAdvt.List(ilLoop), 1) = 0) Then
            '    If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
            '        slAdvtName = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
            '    Else
            '        slAdvtName = Trim$(tmAdf.sName)
            '    End If
            '    If (InStr(1, Trim$(lbcAdvt.List(ilLoop)), Trim$(slAdvtName), 1) = 1) Then
            '        ilFound = True
            '        Exit For
            '    End If
            'Next ilLoop
            If Not ilFound Then
                If tmCif.sPurged = "H" Then
                    slStatus = " History "
                ElseIf tmCif.sPurged = "P" Then
                    slStatus = " Purged "
                Else
                    slStatus = " Active "
                End If
                gUnpackDate tmCif.iRotStartDate(0), tmCif.iRotStartDate(1), slStartDate
                gUnpackDate tmCif.iRotEndDate(0), tmCif.iRotEndDate(1), slEndDate
                If Len(slStartDate) > 0 Then
                    slDateRange = slStartDate & "-" & slEndDate
                Else
                    slDateRange = ""
                End If
                If tmCif.lcpfCode > 0 Then
                    tmCpfSrchKey.lCode = tmCif.lcpfCode
                    ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        slProduct = Trim$(tmCpf.sName)
                    Else
                        slProduct = ""
                    End If
                Else
                    slProduct = ""
                End If
                ''lbcAdvt.AddItem gAlignStringByPixel(Trim$(tmAdf.sName) & "|" & slStatus & "|" & slDateRange & "|" & slProduct, "|", imListField(), imListFieldChar())
                '10/12/09:  Add sort
                If Len(slEndDate) > 0 Then
                    tmFdCartNoInfo(UBound(tmFdCartNoInfo)).lEndDate = gDateValue(slEndDate)
                Else
                    tmFdCartNoInfo(UBound(tmFdCartNoInfo)).lEndDate = 999999
                End If
                If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                '    lbcAdvt.AddItem Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & "|" & slStatus & "|" & slDateRange & "|" & slProduct
                    tmFdCartNoInfo(UBound(tmFdCartNoInfo)).sListInfo = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID) & "|" & slStatus & "|" & slDateRange & "|" & slProduct
                Else
                '    lbcAdvt.AddItem Trim$(tmAdf.sName) & "|" & slStatus & "|" & slDateRange & "|" & slProduct
                    tmFdCartNoInfo(UBound(tmFdCartNoInfo)).sListInfo = Trim$(tmAdf.sName) & "|" & slStatus & "|" & slDateRange & "|" & slProduct
                End If
                ReDim Preserve tmFdCartNoInfo(0 To UBound(tmFdCartNoInfo) + 1) As FDCRTNOINFO
            End If
        End If
        ilRet = btrGetNext(hmCif, tmCif, imCifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If UBound(tmFdCartNoInfo) > 0 Then
        'Sort descending
        ArraySortTyp fnAV(tmFdCartNoInfo(), 0), UBound(tmFdCartNoInfo), 1, LenB(tmFdCartNoInfo(0)), 0, -2, 0
    End If
    For ilLoop = 0 To UBound(tmFdCartNoInfo) - 1 Step 1
        lbcAdvt.AddItem tmFdCartNoInfo(ilLoop).sListInfo
    Next ilLoop
    pbcLbcAdvt_Paint
End Sub
Private Sub edcFindInv_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
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
'    gShowBranner
    Me.KeyPreview = True
    FdCartNo.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = True
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
        mTerminate
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    Erase tmFdCartNoInfo
    
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    btrExtClear hmMcf   'Clear any previous extend operation
    ilRet = btrClose(hmMcf)
    btrDestroy hmMcf
    btrExtClear hmCpf   'Clear any previous extend operation
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    btrExtClear hmCif   'Clear any previous extend operation
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    
    Set FdCartNo = Nothing   'Remove data segment
    
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
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
    imTerminate = False
    imFirstActivate = True

    Screen.MousePointer = vbHourglass
    imLBCtrls = 1

    'mParseCmmdLine
    FdCartNo.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone FdCartNo
    'FdCartNo.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imFirstFocus = True
    ReDim tmFdCartNoInfo(0 To 0) As FDCRTNOINFO
    hmCif = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", FdCartNo
    On Error GoTo 0
    imCifRecLen = Len(tmCif)
    hmCpf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cpf.Btr)", FdCartNo
    On Error GoTo 0
    imCpfRecLen = Len(tmCpf)
    hmAdf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", FdCartNo
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    tmAdf.iCode = 0
    hmMcf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Mcf.Btr)", FdCartNo
    On Error GoTo 0
    ReDim tmMcf(0 To 0) As MCF
    imMcfRecLen = Len(tmMcf(0))
    ilRet = btrGetFirst(hmMcf, tmMcf(0), imMcfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tmMcf(0 To UBound(tmMcf) + 1) As MCF
        ilRet = btrGetNext(hmMcf, tmMcf(UBound(tmMcf)), imMcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If imTerminate Then
        Exit Sub
    End If
    imListField(1) = 15
    imListField(2) = 28 * igAlignCharWidth + 40 '27
    imListField(3) = 37 * igAlignCharWidth '35
    imListField(4) = 57 * igAlignCharWidth '52
    imListField(5) = 117 * igAlignCharWidth '110
'    gCenterModalForm FdCartNo
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
    Unload FdCartNo
    igManUnload = NO
End Sub

Private Sub lbcAdvt_Scroll()
    pbcLbcAdvt_Paint
End Sub

Private Sub pbcClickFocus_GotFocus()
    If imFirstFocus Then
        imFirstFocus = False
    End If
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub

Private Sub pbcLbcAdvt_Paint()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilAdvtEnd As Integer
    Dim ilField As Integer
    Dim llWidth As Long
    Dim slFields(0 To 3) As String
    Dim llFgColor As Long
    Dim ilListIndex As Integer
    
    ilAdvtEnd = lbcAdvt.TopIndex + lbcAdvt.Height \ fgListHtArial825
    If ilAdvtEnd > lbcAdvt.ListCount Then
        ilAdvtEnd = lbcAdvt.ListCount
    End If
    If lbcAdvt.ListCount <= lbcAdvt.Height \ fgListHtArial825 Then
        llWidth = lbcAdvt.Width - 30
    Else
        llWidth = lbcAdvt.Width - igScrollBarWidth - 30
    End If
    pbcLbcAdvt.Width = llWidth
    pbcLbcAdvt.Cls
    llFgColor = pbcLbcAdvt.ForeColor
    For ilLoop = lbcAdvt.TopIndex To ilAdvtEnd - 1 Step 1
        pbcLbcAdvt.ForeColor = llFgColor
        If lbcAdvt.MultiSelect = 0 Then
            If lbcAdvt.ListIndex = ilLoop Then
                gPaintArea pbcLbcAdvt, CSng(0), CSng((ilLoop - lbcAdvt.TopIndex) * fgListHtArial825), CSng(pbcLbcAdvt.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcAdvt.ForeColor = vbWhite
            End If
        Else
            If lbcAdvt.Selected(ilLoop) Then
                gPaintArea pbcLbcAdvt, CSng(0), CSng((ilLoop - lbcAdvt.TopIndex) * fgListHtArial825), CSng(pbcLbcAdvt.Width), CSng(fgListHtArial825) - 15, vbHighlight 'WHITE
                pbcLbcAdvt.ForeColor = vbWhite
            End If
        End If
        slStr = lbcAdvt.List(ilLoop)
        gParseItemFields slStr, "|", slFields()
        ilListIndex = 0
        For ilField = imLBCtrls To UBound(imListField) - 1 Step 1
            pbcLbcAdvt.CurrentX = imListField(ilField)
            pbcLbcAdvt.CurrentY = (ilLoop - lbcAdvt.TopIndex) * fgListHtArial825 + 15
            slStr = slFields(ilListIndex)
            gAdjShowLen pbcLbcAdvt, slStr, imListField(ilField + 1) - imListField(ilField)
            pbcLbcAdvt.Print slStr
            ilListIndex = ilListIndex + 1
        Next ilField
        pbcLbcAdvt.ForeColor = llFgColor
    Next ilLoop
End Sub

Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcAdvt_Paint()
    plcAdvt.CurrentX = 0
    plcAdvt.CurrentY = 0
    plcAdvt.Print " Advertiser                             Status    Rotation Dates        Product"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Copy Numbers"
End Sub
