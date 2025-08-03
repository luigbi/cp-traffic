VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ARInvNo 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4125
   ClientLeft      =   885
   ClientTop       =   2415
   ClientWidth     =   9315
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
   ScaleHeight     =   4125
   ScaleWidth      =   9315
   Begin ComctlLib.ListView lbcInvNoList 
      Height          =   2205
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1320
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   3889
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Contract #"
         Object.Width           =   1341
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Agency"
         Object.Width           =   3529
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Advertiser"
         Object.Width           =   3263
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Type"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Inv. Date"
         Object.Width           =   1164
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Net $"
         Object.Width           =   1429
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "From"
         Object.Width           =   653
      EndProperty
   End
   Begin VB.PictureBox plcAdvt 
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
      Left            =   1125
      ScaleHeight     =   255
      ScaleWidth      =   4410
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1890
      Width           =   4410
      Begin VB.PictureBox plcAdvtName 
         BackColor       =   &H0080FFFF&
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
         Left            =   1095
         ScaleHeight     =   195
         ScaleWidth      =   3195
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.PictureBox plcAgy 
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
      Height          =   270
      Left            =   1125
      ScaleHeight     =   270
      ScaleWidth      =   4410
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1470
      Width           =   4410
      Begin VB.PictureBox plcAgyName 
         BackColor       =   &H0080FFFF&
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
         Left            =   1095
         ScaleHeight     =   195
         ScaleWidth      =   3195
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   3255
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
      Left            =   450
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3600
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
      Left            =   705
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3600
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
      Left            =   1005
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3600
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
      ScaleWidth      =   2370
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -30
      Width           =   2370
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   4185
      TabIndex        =   5
      Top             =   3720
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
      Left            =   1500
      ScaleHeight     =   855
      ScaleWidth      =   6270
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   255
      Width           =   6330
      Begin VB.CommandButton cmcFind 
         Appearance      =   0  'Flat
         Caption         =   "&Find"
         Height          =   285
         Left            =   2655
         TabIndex        =   4
         Top             =   495
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
         Left            =   4590
         TabIndex        =   3
         Top             =   105
         Width           =   1560
      End
      Begin VB.Label lacFind 
         Appearance      =   0  'Flat
         Caption         =   "Find Agency/Advertiser Associated With Invoice #"
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
         Width           =   4395
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   3630
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "ARInvNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Arinvno.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ARInvNo.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim tmRvf As RVF        'Rvf record image
Dim hmRvf As Integer    'Receivable file handle
Dim imRvfRecLen As Integer        'RvF record length
Dim hmPhf As Integer    'Receivable file handle
Dim tmAdf As ADF        'Adf record image
Dim tmAdfSrchKey As INTKEY0    'Rvf key record image
Dim hmAdf As Integer    'Advertiser file handle
Dim imAdfRecLen As Integer        'ADF record length
Dim tmAgf As AGF        'Agf record image
Dim tmAgfSrchKey As INTKEY0    'Agf key record image
Dim hmAgf As Integer    'Agency file handle
Dim imAgfRecLen As Integer        'AGF record length
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer

'*******************************************************
'*                                                     *
'*      Procedure Name:mFind                           *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Find agency/advertiser         *
'*                      associated with invoice #      *
'*                                                     *
'*******************************************************
Private Function mFind(llInvNo As Long) As Integer
'
'   mFind
'   Where:
'
    Dim llNoRec As Long         'Number of records in Sof
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim tlLTypeBuff As POPLCODE   'Type field record
    Dim ilPass As Integer
    Dim hlFile As Integer
    Dim mItem As ListItem
    Dim slDate As String
    Dim slFrom As String
    Dim slNet As String

    lbcInvNoList.ListItems.Clear
    For ilPass = 0 To 1 Step 1
        If ilPass = 0 Then
            hlFile = hmRvf
            slFrom = "A/R"
        Else
            hlFile = hmPhf
            slFrom = "Hist"
        End If
        ilExtLen = Len(tmRvf)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlFile) 'Obtain number of records
        btrExtClear hlFile   'Clear any previous extend operation
        ilRet = btrGetFirst(hlFile, tmRvf, imRvfRecLen, 5, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            mFind = False
            Exit Function
        End If
        Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "RVF", "") 'Set extract limits (all records)
        tlLTypeBuff.lCode = llInvNo    'ilAgyCode
        ilOffSet = gFieldOffset("Rvf", "RvfInvNo")
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlLTypeBuff, 4)
        ilRet = btrExtAddField(hlFile, 0, imRvfRecLen)  'Extract First Name field
        If ilRet <> BTRV_ERR_NONE Then
            mFind = False
            Exit Function
        End If
        'ilRet = btrExtGetNextExt(hlFile)    'Extract record
        ilRet = btrExtGetNext(hlFile, tmRvf, imRvfRecLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlFile, tmRvf, imRvfRecLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    If tmRvf.iAdfCode > 0 Then
                        If tmRvf.iAdfCode <> tmAdf.iCode Then
                            tmAdfSrchKey.iCode = tmRvf.iAdfCode
                            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then
                                tmAdf.sName = "Missing" & str$(tmRvf.iAdfCode)
                            End If
                        End If
                    Else
                        tmAdf.sName = ""
                    End If
                    If tmRvf.iAgfCode > 0 Then
                        If tmRvf.iAgfCode <> tmAgf.iCode Then
                            tmAgfSrchKey.iCode = tmRvf.iAgfCode
                            ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then
                                tmAgf.sName = "Missing" & str$(tmRvf.iAgfCode)
                                tmAgf.sCityID = ""
                            End If
                        End If
                    Else
                        tmAgf.iCode = 0
                        tmAgf.sName = ""
                        tmAgf.sCityID = ""
                    End If
                    Set mItem = lbcInvNoList.ListItems.Add()
                    mItem.Text = Trim$(str$(tmRvf.lCntrNo))
                    If tmRvf.iAgfCode > 0 Then
                        mItem.SubItems(1) = Trim$(tmAgf.sName) & ", " & Trim$(tmAgf.sCityID)
                    Else
                        mItem.SubItems(1) = ""
                    End If
                    If (tmAdf.sBillAgyDir = "D") And (Trim$(tmAdf.sAddrID) <> "") Then
                        mItem.SubItems(2) = Trim$(tmAdf.sName) & ", " & Trim$(tmAdf.sAddrID)
                    Else
                        mItem.SubItems(2) = Trim$(tmAdf.sName)
                    End If
                    mItem.SubItems(3) = Trim$(tmRvf.sTranType)
                    gUnpackDate tmRvf.iInvDate(0), tmRvf.iInvDate(1), slDate
                    mItem.SubItems(4) = slDate
                    gPDNToStr tmRvf.sNet, 2, slNet
                    mItem.SubItems(5) = slNet
                    mItem.SubItems(6) = slFrom
                    ilRet = btrExtGetNext(hlFile, tmRvf, imRvfRecLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlFile, tmRvf, imRvfRecLen, llRecPos)
                    Loop
                Loop
            End If
        End If
    Next ilPass
    mFind = True
    Exit Function
End Function


Private Sub cmcDone_Click()
    Dim ilRet As Integer
    Dim ilRow As Integer
    Dim slName As String

    sgInvNoName = ""
    If lbcInvNoList.ListItems.Count <= 0 Then
        mTerminate
        Exit Sub
    End If
    ilRet = 0
    On Error GoTo cmcDoneErr
    ilRow = lbcInvNoList.SelectedItem.Index
    If ilRet <> 0 Then
        mTerminate
        Exit Sub
    End If
    On Error GoTo 0
    If ilRow >= 1 Then
        slName = lbcInvNoList.ListItems(ilRow).SubItems(1)
        If Trim$(slName) = "" Then
            slName = lbcInvNoList.ListItems(ilRow).SubItems(2)
        End If
        sgInvNoName = slName
    End If
    mTerminate
    Exit Sub
cmcDoneErr:
    ilRet = 1
    Resume Next
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcFind_Click()
    Dim ilRet As Integer
    Dim llInvNo As Long
    llInvNo = Val(edcFindInv.Text)
    ilRet = mFind(llInvNo)
    If lbcInvNoList.ListItems.Count <= 0 Then
        lbcInvNoList.ListItems.Add , , "None"
    End If
'    If ilRet Then
'        'plcAgyName.Caption = Trim$(tmAgf.sName) & ", " & Trim$(tmAgf.sCityID)
'        'plcAdvtName.Caption = Trim$(tmAdf.sName)
'        plcAgyName.Cls
'        plcAgyName.CurrentX = 0
'        plcAgyName.CurrentY = 0
'        plcAgyName.Print Trim$(tmAgf.sName) & ", " & Trim$(tmAgf.sCityID)
'        plcAdvtName.Cls
'        plcAdvtName.CurrentX = 0
'        plcAdvtName.CurrentY = 0
'        plcAdvtName.Print Trim$(tmAdf.sName)
'    Else
'        'plcAgyName.Caption = "Invoice number not found"
'        'plcAdvtName.Caption = ""
'        plcAgyName.Cls
'        plcAgyName.CurrentX = 0
'        plcAgyName.CurrentY = 0
'        plcAgyName.Print "Invoice number not found"
'        plcAdvtName.Cls
'        plcAdvtName.CurrentX = 0
'        plcAdvtName.CurrentY = 0
'        plcAdvtName.Print ""
'    End If
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
    ARInvNo.Refresh
    Me.KeyPreview = True
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
    
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    btrExtClear hmAgf   'Clear any previous extend operation
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    btrExtClear hmPhf   'Clear any previous extend operation
    ilRet = btrClose(hmPhf)
    btrDestroy hmPhf
    btrExtClear hmRvf   'Clear any previous extend operation
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf
    
    Set ARInvNo = Nothing   'Remove data segment
    
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
    Dim llRet As Long
    imFirstActivate = True
    imTerminate = False

    Screen.MousePointer = vbHourglass
    'mParseCmmdLine
    ARInvNo.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone ARInvNo
    'ARInvNo.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imFirstFocus = True
    hmRvf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rvf.Btr)", ARInvNo
    On Error GoTo 0
    imRvfRecLen = Len(tmRvf)
    hmPhf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Phf.Btr)", ARInvNo
    On Error GoTo 0
    hmAdf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", ARInvNo
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    hmAgf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Agf.Btr)", ARInvNo
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)
    tmAdf.iCode = 0
    tmAgf.iCode = 0
    If imTerminate Then
        Exit Sub
    End If
    llRet = SendMessageByNum(lbcInvNoList.hwnd, LV_SETEXTENDEDLISTVIEWSTYLE, 0, LV_FULLROWSSELECT)
'    gCenterModalForm ARInvNo
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
    'Unload Traffic
    Unload ARInvNo
    igManUnload = NO
End Sub

Private Sub lbcInvNoList_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lbcInvNoList.SortKey = ColumnHeader.Index - 1
    lbcInvNoList.Sorted = True
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
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcAdvt_Paint()
    plcAdvt.CurrentX = 0
    plcAdvt.CurrentY = 0
    plcAdvt.Print " Advertiser"
End Sub
Private Sub plcAgy_Paint()
    plcAgy.CurrentX = 0
    plcAgy.CurrentY = 0
    plcAgy.Print " Agency"
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Collection Invoice Numbers"
End Sub
