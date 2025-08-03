VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form BlockVw 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4125
   ClientLeft      =   885
   ClientTop       =   2415
   ClientWidth     =   11010
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
   ScaleWidth      =   11010
   Begin VB.CommandButton cmcRefresh 
      Appearance      =   0  'Flat
      Caption         =   "&Refresh"
      Height          =   285
      Left            =   5040
      TabIndex        =   1
      Top             =   3720
      Width           =   945
   End
   Begin VB.CommandButton cmcErase 
      Appearance      =   0  'Flat
      Caption         =   "&Erase"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6285
      TabIndex        =   8
      Top             =   3720
      Width           =   945
   End
   Begin ComctlLib.ListView lbcView 
      Height          =   2820
      Left            =   105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   705
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   4974
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Time"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Task"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "User Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Item"
         Object.Width           =   3722
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Code"
         Object.Width           =   0
      EndProperty
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   1770
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   3810
      TabIndex        =   0
      Top             =   3720
      Width           =   945
   End
   Begin VB.Label lbcScreen 
      Caption         =   "View Blocks"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   45
      Width           =   1965
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
Attribute VB_Name = "BlockVw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of BlockVw.frm on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmRlfSrchKey1                 imPopReqd                     imBypassSetting           *
'*  imShowHelpMsg                                                                         *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: BlockVw.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
'Contract line
Dim hmRlf As Integer        'Contract line file handle
Dim tmRlf As RLF            'CHF record image
Dim tmRlfSrchKey As LONGKEY0 'CHF key record image
Dim imRlfRecLen As Integer     'CHF record length
Dim tmRlfView() As RLFVIEW

'Inventory
Dim hmCif As Integer    'Inventory file handle
Dim tmCifSrchKey As LONGKEY0    'Cif key record image
Dim imCifRecLen As Integer        'Cif record length
Dim tmCif As CIF

Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer



Private Sub cmcDone_Click()
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcErase_Click()
    Dim ilRet As Integer
    Dim ilRow As Integer

    If lbcView.ListItems.Count <= 0 Then
        Exit Sub
    End If
    ilRet = 0
    On Error GoTo cmcEraseErr
    ilRow = lbcView.SelectedItem.Index
    If ilRet <> 0 Then
        Exit Sub
    End If
    On Error GoTo 0
    If ilRow >= 1 Then
        ilRet = MsgBox("Erase Selected row", vbYesNo + vbQuestion, "Remove Row")
        If ilRet = vbYes Then
            tmRlfSrchKey.lCode = lbcView.ListItems(ilRow).SubItems(6)
            ilRet = btrGetEqual(hmRlf, tmRlf, imRlfRecLen, tmRlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hmRlf)
                If ilRet = BTRV_ERR_NONE Then
                    lbcView.ListItems.Remove ilRow
                Else
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    MsgBox "Unable to Erase record because of Error = " & ilRet
                End If
                cmcErase.Enabled = False
            End If
        End If
    End If
    Exit Sub
cmcEraseErr:
    ilRet = 1
    Resume Next
End Sub

Private Sub cmcRefresh_Click()
    mPopulate
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
    BlockVw.Refresh
    Me.KeyPreview = True
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilReSet                                                                               *
'******************************************************************************************


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
    
    Erase tmRlfView
    
    btrExtClear hmRlf   'Clear any previous extend operation
    ilRet = btrClose(hmRlf)
    btrDestroy hmRlf
    
    btrExtClear hmCif   'Clear any previous extend operation
    ilRet = btrClose(hmCif)
    btrDestroy hmCif
    
    Set BlockVw = Nothing   'Remove data segment
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
    imFirstActivate = True
    imTerminate = False

    Screen.MousePointer = vbHourglass
    'mParseCmmdLine
    BlockVw.height = cmcDone.Top + 5 * cmcDone.height / 3
    gCenterStdAlone BlockVw
    'BlockVw.Show
    Screen.MousePointer = vbHourglass
'    mInitDDE
    'imcHelp.Picture = Traffic!imcHelp.Picture
    imFirstFocus = True
    hmRlf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRlf, "", sgDBPath & "Rlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Rlf.Btr)", BlockVw
    On Error GoTo 0
    imRlfRecLen = Len(tmRlf)  'Get and save CHF record length
    
    hmCif = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cif.Btr)", BlockVw
    On Error GoTo 0
    imCifRecLen = Len(tmCif)  'Get and save CHF record length
    
    ilRet = gObtainUrf()
    ilRet = gObtainVef()
    ilRet = gObtainMCF()

    mPopulate
    If imTerminate Then
        Exit Sub
    End If
'    gCenterModalForm BlockVw
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
    Unload BlockVw
    igManUnload = NO
End Sub


Private Sub lbcView_Click()
    Dim ilRet As Integer
    Dim ilRow As Integer

    If lbcView.ListItems.Count <= 0 Then
        Exit Sub
    End If
    ilRet = 0
    On Error GoTo lbcViewErr
    ilRow = lbcView.SelectedItem.Index
    If ilRet <> 0 Then
        Exit Sub
    End If
    On Error GoTo 0
    If ilRow >= 1 Then
        cmcErase.Enabled = True
    Else
        cmcErase.Enabled = False
    End If
    Exit Sub
lbcViewErr:
    ilRet = 1
    Resume Next
End Sub

Private Sub lbcView_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lbcView.SortKey = ColumnHeader.Index - 1
    lbcView.Sorted = True
End Sub

Private Sub pbcClickFocus_GotFocus()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCode                                                                                *
'******************************************************************************************

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

Private Sub mPopulate()

    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim slTime As String
    Dim llTime As Long
    Dim ilIndex As Integer
    Dim mItem As ListItem
    Dim llRecCode As Long
    Dim llDate As Long
    Dim ilVefCode As Integer
    Dim ilVef As Integer
    Dim llRet As Long
    Dim ilCol As Integer

    imRlfRecLen = Len(tmRlf)
    lbcView.ListItems.Clear
    ReDim tmRlfView(0 To 0) As RLFVIEW
    llRet = SendMessageByNum(lbcView.hWnd, LV_SETEXTENDEDLISTVIEWSTYLE, 0, LV_FULLROWSSELECT)

    ilRet = btrGetFirst(hmRlf, tmRlf, imRlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE)
        If tmRlf.sType <> "A" Then
            gUnpackDateForSort tmRlf.iEnteredDate(0), tmRlf.iEnteredDate(1), slDate
            gUnpackTimeLong tmRlf.iEnteredTime(0), tmRlf.iEnteredTime(1), False, llTime
            slTime = Trim$(str$(llTime))
            Do While Len(slTime) < 6
                slTime = "0" & slTime
            Loop
            tmRlfView(UBound(tmRlfView)).sKey = slDate & slTime
            tmRlfView(UBound(tmRlfView)).tRlf = tmRlf
            ReDim Preserve tmRlfView(0 To UBound(tmRlfView) + 1) As RLFVIEW
        End If
        ilRet = btrGetNext(hmRlf, tmRlf, imRlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop
    If UBound(tmRlfView) - 1 > 0 Then
        ArraySortTyp fnAV(tmRlfView(), 0), UBound(tmRlfView), 1, LenB(tmRlfView(0)), 0, LenB(tmRlfView(0).sKey), 0
    End If
    lbcView.ColumnHeaders.item(1).Width = lbcView.Width / 13    'Date
    lbcView.ColumnHeaders.item(2).Width = lbcView.Width / 11    'Time
    lbcView.ColumnHeaders.item(3).Width = lbcView.Width / 9    'Task
    'lbcView.ColumnHeaders.Item(4).Width = lbcView.Width / 10    'User
    lbcView.ColumnHeaders.item(5).Width = lbcView.Width / 4.3   'Contract or Vehicle
    lbcView.ColumnHeaders.item(6).Width = lbcView.Width / 13    'Date
    lbcView.ColumnHeaders.item(7).Width = 0
    lbcView.ColumnHeaders.item(4).Width = lbcView.Width - 2 * igScrollBarWidth - 2 * fgPadDeltaX - 7 * 120
    For ilCol = 1 To 6 Step 1
        If ilCol <> 4 Then
            If lbcView.ColumnHeaders.item(4).Width > lbcView.ColumnHeaders.item(ilCol).Width Then
                lbcView.ColumnHeaders.item(4).Width = lbcView.ColumnHeaders.item(4).Width - lbcView.ColumnHeaders.item(ilCol).Width
            Else
                Exit For
            End If
        End If
    Next ilCol

    For ilIndex = 0 To UBound(tmRlfView) - 1 Step 1
        tmRlf = tmRlfView(ilIndex).tRlf
        Set mItem = lbcView.ListItems.Add()
        gUnpackDate tmRlf.iEnteredDate(0), tmRlf.iEnteredDate(1), slDate
        mItem.Text = slDate
        gUnpackTime tmRlf.iEnteredTime(0), tmRlf.iEnteredTime(1), "A", "1", slTime
        mItem.SubItems(1) = slTime
        mItem.SubItems(3) = ""
        For ilLoop = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
            If tgPopUrf(ilLoop).iCode = tmRlf.iUrfCode Then
                mItem.SubItems(3) = Trim$(tgPopUrf(ilLoop).sRept)
                Exit For
            End If
        Next ilLoop
        If tmRlf.sType = "C" Then
            If tmRlf.sSubType = "C" Then
                mItem.SubItems(2) = "in Order"
            Else
                mItem.SubItems(2) = "Scheduling"
            End If
            mItem.SubItems(4) = Trim$(str$(tmRlf.lRecCode))
        ElseIf tmRlf.sType = "S" Then
            If tmRlf.sSubType = "C" Then
                mItem.SubItems(2) = "Scheduling"
            ElseIf tmRlf.sSubType = "P" Then
                mItem.SubItems(2) = "Posting"
            Else
                mItem.SubItems(2) = "Spots"
            End If
            llRecCode = tmRlf.lRecCode
            llDate = llRecCode Mod 65536
            ilVefCode = (llRecCode - llDate) \ 65536
            ilVef = gBinarySearchVef(ilVefCode)
            If ilVef <> -1 Then
                mItem.SubItems(4) = Trim$(tgMVef(ilVef).sName)
            Else
                mItem.SubItems(4) = ""
            End If
            If llDate < 1000 Then
                'Game Number
                mItem.SubItems(5) = "# " & Trim$(str$(llDate))
            Else
                mItem.SubItems(5) = Format$(llDate, "m/d/yy")
            End If
        ElseIf tmRlf.sType = "I" Then
            If tmRlf.sSubType = "U" Then
                mItem.SubItems(2) = "Invoicing"
                mItem.SubItems(4) = "Undo"
            Else
                mItem.SubItems(2) = "Invoicing"
                mItem.SubItems(4) = "Final"
                mItem.SubItems(5) = Format$(tmRlf.lRecCode, "m/d/yy")
            End If
        ElseIf tmRlf.sType = "M" Then
            mItem.SubItems(2) = "Copy Inventory"
            tmCifSrchKey.lCode = tmRlf.lRecCode
            ilRet = btrGetEqual(hmCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                For ilLoop = LBound(tgMCF) To UBound(tgMCF) - 1 Step 1
                    If tgMCF(ilLoop).iCode = tmCif.iMcfCode Then
                        mItem.SubItems(4) = Trim$(tgMCF(ilLoop).sName) & Trim$(tmCif.sName)
                        Exit For
                    End If
                Next ilLoop
            End If
        ElseIf tmRlf.sType = "R" Then
            If tmRlf.sSubType = "R" Then
                mItem.SubItems(2) = "Collections"
                mItem.SubItems(4) = "Reconcile"
            ElseIf tmRlf.sSubType = "Z" Then
                mItem.SubItems(2) = "Collections"
                mItem.SubItems(4) = "Zero-Purge"
            Else
                mItem.SubItems(2) = "Collections"
                mItem.SubItems(4) = ""
            End If
            mItem.SubItems(5) = ""
        ElseIf tmRlf.sType = "Y" And tmRlf.sSubType = "S" Then
            mItem.SubItems(2) = "Copy Save"
        End If
        mItem.SubItems(6) = tmRlf.lCode
    Next ilIndex
End Sub

