VERSION 5.00
Begin VB.Form LinksDef 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5880
   ClientLeft      =   1215
   ClientTop       =   1485
   ClientWidth     =   8460
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5880
   ScaleWidth      =   8460
   Begin VB.Timer tmcScroll 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   960
      Top             =   5415
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   285
      Left            =   4350
      TabIndex        =   10
      Top             =   5490
      Width           =   1140
   End
   Begin VB.CommandButton cmcReport 
      Appearance      =   0  'Flat
      Caption         =   "&Report"
      Height          =   285
      Left            =   5880
      TabIndex        =   11
      Top             =   5490
      Width           =   1140
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   5370
   End
   Begin VB.PictureBox pbcIconLink 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragIcon        =   "Linksdef.frx":0000
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
      Height          =   195
      Left            =   7515
      ScaleHeight     =   165
      ScaleWidth      =   150
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5355
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox pbcIconStd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   195
      Left            =   7350
      ScaleHeight     =   165
      ScaleWidth      =   150
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5595
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox pbcIconTrash 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragIcon        =   "Linksdef.frx":030A
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
      Height          =   195
      Left            =   6840
      ScaleHeight     =   165
      ScaleWidth      =   150
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5340
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox pbcIconMove 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragIcon        =   "Linksdef.frx":0614
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
      Height          =   195
      Left            =   7230
      ScaleHeight     =   165
      ScaleWidth      =   150
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5340
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CheckBox ckcShow 
      Caption         =   "Show Discrepancy Avails Only"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5385
      TabIndex        =   12
      Top             =   -15
      Width           =   2805
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   5490
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Top             =   5490
      Width           =   1050
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   3135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3135
   End
   Begin VB.PictureBox plcNetworks 
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
      Height          =   5070
      Left            =   220
      ScaleHeight     =   5010
      ScaleWidth      =   7980
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   320
      Width           =   8040
      Begin VB.ListBox lbcSelling 
         Appearance      =   0  'Flat
         Height          =   1710
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   260
         Width           =   1935
      End
      Begin VB.HScrollBar hbcAiring 
         Height          =   240
         LargeChange     =   3
         Left            =   270
         SmallChange     =   3
         TabIndex        =   7
         Top             =   4560
         Width           =   7485
      End
      Begin VB.HScrollBar hbcSelling 
         Height          =   240
         LargeChange     =   3
         Left            =   285
         SmallChange     =   3
         TabIndex        =   4
         Top             =   2160
         Width           =   7485
      End
      Begin VB.ListBox lbcAiring 
         Appearance      =   0  'Flat
         Height          =   1710
         Index           =   0
         ItemData        =   "Linksdef.frx":091E
         Left            =   240
         List            =   "Linksdef.frx":0920
         TabIndex        =   6
         Top             =   2640
         Width           =   1940
      End
      Begin VB.Label lacMess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Order of Selling vehicle within Airing vehicle indicates order of spots on Log"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   540
         TabIndex        =   17
         Top             =   4830
         Width           =   6450
      End
      Begin VB.Label lacAiring 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Air"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   2445
         Width           =   1940
      End
      Begin VB.Label lacSelling 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Sell"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   60
         Width           =   1940
      End
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   135
      Top             =   5415
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   7920
      Picture         =   "Linksdef.frx":0922
      Top             =   5370
      Width           =   480
   End
End
Attribute VB_Name = "LinksDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Linksdef.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'******************************************************
'
'            LINKSDEF MODULE VARIABLES
'
'   Created : 4/24/94       By : D. Hannifan
'   Modified :              By :
'
'******************************************************
Option Explicit
Option Compare Text
'LinksDef Module Flags
Dim imReinitFlag As Integer     'True=Reinitialize LinksDef ; False=First Initialization process
Dim imFirstActivate As Integer
Dim imTerminate As Integer      'True = terminating task, False= OK
Dim imUpdateFlag As Integer     'True = changes have been made ; False = no changes have been made
Dim imLegal As Integer          'True = move or swap is valid ; False = illegal call
'Drag Events Module Variables
Dim imDragSource As Integer     '0=Selling; 1= Airing
Dim imDragType As Integer       '0=Avail; 1= Link
Dim imDragListIndex As Integer  'Index list-item where drag event initiated
Dim imDragIndex As Integer      'Selling or Airing control index
Dim imSellClickSource As Integer     '0=Selling; 1= Airing
Dim imSellClickType As Integer       '0=Avail; 1= Link
Dim imSellClickListIndex As Integer  'Index list-item where drag event initiated
Dim imSellClickIndex As Integer      'Selling or Airing control index
Dim imAirClickSource As Integer     '0=Selling; 1= Airing
Dim imAirClickType As Integer       '0=Avail; 1= Link
Dim imAirClickListIndex As Integer  'Index list-item where drag event initiated
Dim imAirClickIndex As Integer      'Selling or Airing control index
Dim imSwapFlag As Integer       'False = link or drag mode ; True = swap mode
'Drag Scroll trigger variables
Dim smDSourceType As String * 1 'Vehicle type of drag source "A"=Airing  "S"=Selling
Dim imDType As Integer          '0=Link or Drag operation ; 1=Swap or Move operation ; -1 = Invalid Operation
Dim imDTop As Integer           'Top limit of scroll box trigger area
Dim imDLeft As Integer          'Left limit of scroll box trigger area
Dim smLeaveType As String * 1   'Target list box type A=Airing  ; S = selling
Dim imBoxWidth As Integer       'Width of scroll box trigger area
Dim imHeight As Integer         'Height of scroll box trigger area
Dim imDSourceIndex As Integer   'Index of drag source list box
Dim imDBottom As Integer        'Bottom limit of scroll box trigger area
Dim imDRight As Integer         'Right limit of scroll box trigger area
Dim imDExitDirect As Integer    '0 = left list box from top  ; 1= left list box from bottom
Dim imLeaveIndex As Integer     'Index of list box from last leave event
'Modular variables imported from Links
Dim imDateCode As Integer       'Date Code Active From Links module 0=M-F, 6=Sa, 7=Su
Dim smLinksDefStatus As String * 1 'Vlf:P=Pending, C=Current status from Links module
Dim smDateFilter As String      'Date filter from Links
Dim lmDateFilter As Long        'Date filter from links
Dim smEndDate As String         'End Date, If TFN this will be blank
Dim imNoSelling As Integer      'Number of selling networks selected
Dim imNoAiring As Integer       'Number of airing networks selected
Dim imDate0 As Integer          'Byte 0 of smDateFilter
Dim imDate1 As Integer          'Byte 1 of smDateFilter
'LCF Variables
Dim hmLcf As Integer            'Log calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim tmDLcf As LCF
Dim tmLcfSrchKey As LCFKEY0     'LCF Key 0 image
Dim imLcfRecLen As Integer         'LCF record length
'LEF Variables
Dim hmLef As Integer            'Log event file handle
Dim tmLef As LEF                'LEF record image
Dim tmLefSrchKey As LEFKEY0     'LEF Key 0 image
Dim imLefRecLen As Integer      'LEF record length
'LVF Variables
Dim hmLvf As Integer            'Log version file handle
Dim tmLvf As LVF                'LVF record image
Dim tmLvfSrchKey As LONGKEY0     'LVF Key 0 image
Dim imLvfRecLen As Integer      'LVF record length
'ANF Variables
Dim hmAnf As Integer            'Avail name file handle
Dim tmAnf As ANF                'ANF record image
Dim tmAnfSrchKey As INTKEY0     'ANF Key 0 image
Dim imAnfRecLen As Integer      'ANF record length
'Dim tmRec As LPOPREC
'VCF Variables
Dim hmVcf As Integer
Dim tmVcf As VCF
Dim tmVcfSrchKey0 As VCFKEY0
Dim imVcfRecLen As Integer
'VLF variables
Dim tmVlf As VLF
Dim tmVlfSrchKey0 As VLFKEY0
Dim tmVlfSrchKey1 As VLFKEY1
Dim tmVlfSrchKey2 As LONGKEY0
Dim tmVlfSrchKey3 As VLFKEY3
Dim tmVlfSrchKey4 As VLFKEY4
Dim tmVlfUpdate() As VLF        'VLF update record image
Dim hmVlfPop As Integer         'VLF file handle
Dim tmVlfPop() As VLF           'VLF record image (temp redimensionable file)
Dim imUpperBound As Integer     'Upper Bound for tmVLFPop
Dim imVlfPopRecLen As Integer   'VLF record length
Dim imSellPending() As Integer   '-VefCode=VLF Pending exist; +VefCode=No pending VLF
Dim imAirPending() As Integer   '-VefCode=VLF Pending exist; +VefCode=No pending VLF
'Temporary string arrays and dimension parameters (for reinit and show discreps procedures)
Dim smSellingLists()  As String 'Selling vehicles temp array of list box
Dim smAiringLists()  As String  'Selling vehicles temp array of list box
Dim imSellCount() As Integer    'Selling vehicles list item count for temp array
Dim imAirCount() As Integer     'Airing vehicles list item count for temp array
Dim imSellUpperB As Integer     'upperbound of imSellingLists array
Dim imAirUpperB As Integer      'upperBound of imAiringLists array
Dim imUpdateUpB As Integer      'tmVlfUpdate array size
'List Population arrays
Dim tmCLLC() As LLC             'Current LLC image
Dim tmPLLC() As LLC             'Pending LLC image
Dim imVefCode As Integer        'Vehicle code number
Dim imVpfIndex As Integer       'VPF index value
Dim imSSmallChange As Integer
Dim imASmallChange As Integer
Dim lmDelVlf() As Long
Dim imSellDelVef() As Integer
Dim imAirlDelVef() As Integer
Dim imChkSellTermVlf() As Integer
Dim imChkAirTermVlf() As Integer
Dim smScreenCaption As String
Dim imUpdateAllowed As Integer

Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor

Private bmFirstCallToVpfFind As Boolean





'****************************************************************
'
'           Procedure Name : ckcShow_Click
'
'       Date Created : 4/24/94             By: D. Hannifan
'       Date Modified :                    By:
'
'       Comments : Show discpencies or reinitialize linksdef
'
'*****************************************************************
'
Private Sub ckcShow_Click()
    'Code added because Value removed as parameter
    Dim Value As Integer
    Value = False
    If ckcShow.Value = vbChecked Then
        Value = True
    End If
    'End of coded added
    If Value <> False Then
        mShowDiscreps
    Else
        mReInitLinksDef
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
    End If
    mSetCommands
End Sub
Private Sub cmcCancel_Click()
    mTerminate     'Exit LinksDef
End Sub
'****************************************************************
'
'           Procedure Name : cmcDone_Click
'
'       Date Created : 4/24/94             By: D. Hannifan
'       Date Modified :                    By:
'
'       Comments : Exit LinksDef : If changes were made prompt
'                  for a VLF file update
'*****************************************************************
'
Private Sub cmcDone_Click()
    Dim ilRet As Integer

    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If imUpdateFlag <> False Then
        ilRet = MsgBox("Save Changes Before Leaving ?", 35, "LINKS DEFINITION")
        If ilRet = 6 Then  'Yes
            Screen.MousePointer = vbHourglass
            mUpdateVlf             'Update VLF
            mWriteVlf
            imReinitFlag = False
            Screen.MousePointer = vbDefault
            mTerminate
            Exit Sub
        ElseIf ilRet = 2 Then  'Cancel : Return to main screen
            Exit Sub
        ElseIf ilRet = 7 Then   'No  : Exit without VLF update
            imReinitFlag = False
            mTerminate
            Exit Sub
        End If
    Else
        imReinitFlag = False
        mTerminate
        Exit Sub
    End If
End Sub
Private Sub cmcReport_Click()
    Dim slStr As String
    'Unload IconTraf
    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
    '    Exit Sub
    'End If
    igRptCallType = PROGRAMMINGJOB
    igRptType = 0   'Selling to Airing
    ''Screen.MousePointer = vbHourGlass  'Wait
    'igChildDone = False
    'Links!edcLinkSrceDoneMsg.Text = ""
    'If (Not igStdAloneMode) And (igShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Links^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        Else
            slStr = "Links^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType)) & "\" & Trim$(str$(igRptType))
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Links^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    Else
    '        slStr = "Links^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType)) & "\" & Trim$(Str$(igRptType))
    '    End If
    'End If
    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
    'LinksDef.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    'slStr = sgDoneMsg
    'LinksDef.Enabled = True
    'Links!edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    sgCommandStr = slStr
    RptList.Show vbModal
    ''Screen.MousePointer = vbDefault    'Default
End Sub
'****************************************************************
'
'           Procedure Name : cmcUpdate_Click
'
'       Date Created : 4/24/94             By: D. Hannifan
'       Date Modified :                    By:
'
'       Comments : Update VLF file
'
'*****************************************************************
'
Private Sub cmcUpdate_Click()
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mUpdateVlf
    mWriteVlf
    imUpdateFlag = False
    Screen.MousePointer = vbDefault
    mSetCommands
End Sub
Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    imFirstActivate = False
'    gShowBranner
    If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    gShowBranner imUpdateAllowed
    Me.KeyPreview = True
    LinksDef.Refresh
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
    imTerminate = False
    imReinitFlag = False
    If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100) / Me.Height
        Me.Height = (lgPercentAdjH * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    End If
    mInit
    If imTerminate Then
        'cmcCancel_Click
        mTerminate
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    On Error Resume Next
    mCloseFiles
    ilRet = btrClose(hmAnf)
    btrDestroy hmAnf
    ilRet = btrClose(hmLvf)
    btrDestroy hmLvf
    ilRet = btrClose(hmVcf)
    btrDestroy hmVcf
    ilRet = btrClose(hmVlfPop)
    btrDestroy hmVlfPop
    'Unload lists from mem
    If imNoSelling > 0 Then
        For ilLoop = imNoSelling - 1 To 1 Step -1
            Unload lacSelling(ilLoop)
            Unload lbcSelling(ilLoop)
        Next ilLoop
    End If
    imNoSelling = 0
    If imNoAiring > 0 Then
        For ilLoop = imNoAiring - 1 To 1 Step -1
            Unload lacAiring(ilLoop)
            Unload lbcAiring(ilLoop)
        Next ilLoop
    End If
    imNoAiring = 0
    'Erase arrays from mem
    Erase imChkSellTermVlf
    Erase imChkAirTermVlf
    Erase tmVlfPop
    Erase smSellingLists
    Erase smAiringLists
    Erase imSellCount
    Erase imAirCount
    Erase tmCLLC
    Erase tmPLLC
    Erase lmDelVlf
    Erase imSellDelVef
    Erase imAirlDelVef
    
    Set LinksDef = Nothing
    
End Sub

'****************************************************************
'
'           Procedure Name : hbcAiring_Change (H-Scroll Bar)
'
'       Date Created : 4/24/94             By: D. Hannifan
'       Date Modified :                    By:
'
'       Comments : Shift visible list box indexes
'
'*****************************************************************
'
Private Sub hbcAiring_Change()
    Call mMoveLBox(2, CInt(hbcAiring.Value))
End Sub
'****************************************************************
'
'           Procedure Name : hbcSelling_Change (H-Scroll Bar)
'
'       Date Created : 4/24/94             By: D. Hannifan
'       Date Modified :                    By:
'
'       Comments : Shift visible list box indexes
'
'*****************************************************************
'
Private Sub hbcSelling_Change()
    Call mMoveLBox(1, CInt(hbcSelling.Value))
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
'****************************************************************
'
'           Procedure Name : imcTrash_Click
'
'       Date Created : ?            By: D. LeVine
'       Date Modified : 4/24/94     By: D. Hannifan
'
'       Comments : Delete a link
'
'*****************************************************************
'
Private Sub imcTrash_Click()

    imSellClickIndex = -1
    imAirClickIndex = -1
    If imDragType = 1 Then  'Link
        If imDragSource = 0 Then    'Selling
            mDeleteSelling
        Else    'Airing
            mDeleteAiring
        End If
    End If
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
End Sub
Private Sub imcTrash_DragDrop(Source As control, X As Single, Y As Single)

    If imDragSource = 0 Then    'Selling
        lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconStd.DragIcon
    Else    'Airing
        lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconStd.DragIcon
    End If
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    If (State = vbEnter) And (imDragType = 1) Then
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
        If imDragSource = 0 Then    'Selling
            lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        Else    'Airing
            lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconDwnArrow.DragIcon
        End If
    ElseIf (State = vbLeave) And (imDragType = 1) Then    'DragLeave
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
        If imDragSource = 0 Then    'Selling
            lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
        Else    'Airing
            lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
        End If
    End If
End Sub


Private Sub lacAiring_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lacAiring(Index).ToolTipText = lacAiring(Index).Caption
End Sub

Private Sub lacSelling_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lacSelling(Index).ToolTipText = lacSelling(Index).Caption
End Sub

'****************************************************************
'
'                   Procedure Name : lbcAiring_DragDrop
'
'       Date Created : ?            By: D. LeVine
'       Date Modified :4/24/94      By: D. Hannifan
'
'       Comments : Determine drag call type and process a
'                  swap, move or link event
'
'*****************************************************************
'
Private Sub lbcAiring_DragDrop(Index As Integer, Source As control, X As Single, Y As Single)
    Dim ilDragType As Integer           '1=Link , 0=Swap Or Move
    Dim ilDragListIndex As Integer      'Index of target
    Dim slName1 As String               'List item string vehicle 1
    Dim slName2 As String               'List item string vehicle 2
    Dim ilSaveListIndex As Integer      'Temporary list item index for swap/move
    Dim ilSaveIndex As Integer          'Temporary list index for swap/move
    Dim ilSaveDragListIndex As Integer  'Temporary source list index for swap/move
    Dim ilIndexShift As Integer         'Increment/Decrement value for list item
    Dim ilLoop As Integer               'DragListIndex counter
    Dim ilDragIndex As Integer          'Index of drag source
    Dim slStr As String                 'Parsing string
    Dim ilRet As Integer                'Return value from call
    Dim ilSameTime As Integer           'True=swap in same listbox with same sell time
    Dim ilSetIndex1 As Integer          'Listbox1 Index for mSetSelected call
    Dim ilSetIndex2 As Integer          'Listbox2 Index for mSetSelected call
    Dim ilSetList1 As Integer           'Listbox1 List Index for mSetSelected call
    Dim ilSetList2 As Integer           'Listbox2 ListIndex for mSetSelected call
    Dim ilDelimLoc As Integer           'Delimeter location for parsing
    Dim ilLen As Integer                'String length
    Dim slStr2 As String                'parsing string
    On Error GoTo lbcAiringErr


    imSellClickIndex = -1
    imAirClickIndex = -1
    ilDragIndex = imDragIndex 'Save imDragIndex at call
    imSwapFlag = False
    If imDragSource = 0 Then    'Selling
        lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconStd.DragIcon
    Else                        'Airing
        lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconStd.DragIcon
    End If
    ilDragType = mDragType(lbcAiring(Index), X, Y, ilDragListIndex)

    If imDragSource = 1 Then    'Airing
        If (imDragType = 1) And (ilDragType = 1) Then  'Swap links
            imSwapFlag = True
            imLegal = True
            If (Index <> imDragIndex) Then 'swap in different list box
                ilSetIndex1 = Index
                ilSetIndex2 = imDragIndex
                ilSetList1 = ilDragListIndex
                ilSetList2 = imDragListIndex
                ilSaveListIndex = ilDragListIndex
                ilSaveIndex = imDragIndex
                ilSaveDragListIndex = imDragListIndex - 1
                mMoveAiring Index, ilDragListIndex
                If Not (imLegal) Then
                    Exit Sub
                End If
                imDragIndex = Index
                imDragListIndex = ilSaveListIndex
                If imLegal Then
                    mMoveAiring ilSaveIndex, ilSaveDragListIndex
                    imUpdateFlag = True
                    mSetSelected ilSetIndex1, ilSetIndex2, ilSetList1, ilSetList2, "A", "A"
                Else
                    Exit Sub
                End If
                tmcScroll.Enabled = False
                imLegal = True
            End If
            If (ilDragIndex = imDragIndex) Then
                ilSameTime = mCheckTime(ilDragIndex, ilDragListIndex, imDragListIndex, "A")
                If (ilDragListIndex < imDragListIndex) And Not (ilSameTime) Then 'swap within the same list box std order
                    ilSaveListIndex = ilDragListIndex
                    ilSaveIndex = imDragIndex
                    ilSaveDragListIndex = imDragListIndex
                    mMoveAiring Index, ilDragListIndex
                    If Not (imLegal) Then
                        Exit Sub
                    End If
                    imDragIndex = Index
                    imDragListIndex = ilSaveListIndex
                    If imLegal Then
                        mMoveAiring ilSaveIndex, ilSaveDragListIndex
                        mSetSelected 1, 1, 1, 1, "C", "C"
                        imUpdateFlag = True
                    Else
                        Exit Sub
                    End If
                    imLegal = True
                    tmcScroll.Enabled = False
                End If
                If (ilDragListIndex > imDragListIndex) And Not (ilSameTime) Then 'swap within the same list box reverse order
                    ilSaveListIndex = ilDragListIndex
                    ilSaveIndex = imDragIndex
                    ilSaveDragListIndex = imDragListIndex
                    imDragIndex = Index
                    imDragListIndex = ilSaveListIndex
                    mMoveAiring ilSaveIndex, ilSaveDragListIndex
                    If imLegal Then
                        mMoveAiring Index, ilDragListIndex
                        mSetSelected 1, 1, 1, 1, "C", "C"
                        imUpdateFlag = True
                    Else
                        Exit Sub
                    End If
                    imLegal = True
                    tmcScroll.Enabled = False
                End If
                If (ilSameTime) Then
                    slStr2 = Trim$(lbcAiring(ilDragIndex).List(ilDragListIndex))
                    ilLen = Len(slStr2)
                    ilDelimLoc = InStr(1, slStr2, " ")
                    slStr2 = Trim$(right$(slStr2, ilLen - ilDelimLoc))
                    ilRet = gParseItem(slStr2, 1, "@", slStr2)
                    slStr2 = Trim$(slStr2)
                    imLegal = False
                    For ilRet = 0 To imNoSelling - 1 Step 1
                        If (slStr2 = Trim$(lacSelling(ilRet).Caption)) Then
                            imLegal = True
                            Exit For
                        End If
                    Next ilRet
                    If imLegal Then
                        slStr = lbcAiring(ilDragIndex).List(ilDragListIndex)
                        lbcAiring(ilDragIndex).List(ilDragListIndex) = lbcAiring(ilDragIndex).List(imDragListIndex)
                        lbcAiring(ilDragIndex).List(imDragListIndex) = slStr
                        mSetSelected 1, 1, 1, 1, "C", "C"
                        imUpdateFlag = True
                        tmcScroll.Enabled = False
                        ilSameTime = False
                    Else
                        imLegal = True
                        Exit Sub
                    End If
                End If
            End If
            imSwapFlag = False
            imLegal = True
        End If
        If (imDragType = 1) And (ilDragType = 0) Then   'Move link
            imSwapFlag = True
            mMoveAiring Index, ilDragListIndex
            If imLegal Then
                imUpdateFlag = True
            Else
                Exit Sub
            End If
            imLegal = True
        End If
    Else    'Selling
        If (imDragType = 0) And (ilDragType = 0) Then 'Add links
            'Construct list strings
            ilRet = gParseItem(lbcSelling(imDragIndex).List(imDragListIndex), 1, " ", slStr)
            slStr = Trim$(slStr)
            slName1 = CStr("  " & Trim$(slStr) & " " & Trim$(lacSelling(imDragIndex).Caption))
            ilRet = gParseItem(lbcSelling(imDragIndex).List(imDragListIndex), 2, "@", slStr)
            slName1 = CStr(slName1 & "@" & Trim$(slStr))
            ilRet = gParseItem(lbcAiring(Index).List(ilDragListIndex), 1, " ", slStr)
            slStr = Trim$(slStr)
            slName2 = CStr("  " & Trim$(slStr) & " " & Trim$(lacAiring(Index).Caption))
            ilRet = gParseItem(lbcAiring(Index).List(ilDragListIndex), 2, "@", slStr)
            slName2 = CStr(slName2 & "@" & Trim$(slStr))
'Check for collating order before adding to selling list box
                ilIndexShift = 1
                If (imDragSource = 0) Then
                Else
                    ilLoop = imDragListIndex + 1
                    Do While ilLoop < lbcSelling(imDragIndex).ListCount
                       If (Left$(lbcSelling(imDragIndex).List(imDragListIndex + ilIndexShift), 2) = "  ") Then
                           If (lbcSelling(imDragIndex).List(imDragListIndex + ilIndexShift) = slName2) Then 'Duplicate
                                Beep
                                mSetSelected 1, 1, 1, 1, "C", "C"
                                DoEvents
                                Exit Sub
                           End If
                            ilIndexShift = ilIndexShift + 1
                            ilLoop = ilLoop + 1
                        Else: Exit Do
                        End If
                    Loop
                    ' Add to selling list box
                    lbcSelling(imDragIndex).AddItem mFillTo100(slName2), imDragListIndex + ilIndexShift
                    mSetSelected imDragIndex, 1, imDragListIndex + ilIndexShift, 1, "S", "C"
                    imUpdateFlag = True
                    tmcScroll.Enabled = False
                End If

    ' Check for collate order before adding to airing list box
                ilIndexShift = 1
                If (imDragSource = 1) Then
                Else
                    ilLoop = ilDragListIndex + 1
                    Do While ilLoop < lbcAiring(Index).ListCount
                        If (Left$(lbcAiring(Index).List(ilDragListIndex + ilIndexShift), 2) = "  ") Then
                           If (lbcAiring(Index).List(ilDragListIndex + ilIndexShift) = slName1) Then 'Duplicate
                                Beep
                                mSetSelected 1, 1, 1, 1, "C", "C"
                                DoEvents
                                Exit Sub
                           End If
                            ilIndexShift = ilIndexShift + 1
                            ilLoop = ilLoop + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    'Add to airing list box
                    lbcAiring(Index).AddItem mFillTo100(slName1), ilDragListIndex + ilIndexShift
                    mSetSelected Index, 1, ilDragListIndex + ilIndexShift, 1, "A", "C"
                    imUpdateFlag = True
                    tmcScroll.Enabled = False
                End If
                ilIndexShift = 1
                If (imDragSource = 0) Then
                    ilLoop = imDragListIndex + 1
                    Do While ilLoop < lbcSelling(imDragIndex).ListCount
                        If (Left$(lbcSelling(imDragIndex).List(imDragListIndex + ilIndexShift), 2) = "  ") Then
                            If (lbcSelling(imDragIndex).List(imDragListIndex + ilIndexShift) = slName2) Then 'Duplicate
                                Beep
                                mSetSelected 1, 1, 1, 1, "C", "C"
                                DoEvents
                                Exit Sub
                            End If
                            ilIndexShift = ilIndexShift + 1
                            ilLoop = ilLoop + 1
                        Else: Exit Do
                        End If
                    Loop
                    lbcSelling(imDragIndex).AddItem mFillTo100(slName2), imDragListIndex + ilIndexShift
                    mSetSelected imDragIndex, 1, imDragListIndex + ilIndexShift, 1, "S", "C"
                    imUpdateFlag = True
                    tmcScroll.Enabled = False
                End If
                ilIndexShift = 1
                If (imDragSource = 1) Then
                    lbcAiring(Index).AddItem mFillTo100(slName1), ilDragListIndex + ilIndexShift
                    mSetSelected Index, 1, ilDragListIndex + ilIndexShift, 1, "A", "C"
                    imUpdateFlag = True
                    tmcScroll.Enabled = False
                End If
            End If
        End If
        tmcScroll.Enabled = False
        mSetCommands
        Exit Sub
lbcAiringErr:
    On Error GoTo 0
    imSwapFlag = False
    Exit Sub
End Sub
'****************************************************************
'
'                   Procedure Name : lbcAiring_DragOver
'
'       Date Created : ?            By: D. LeVine
'       Date Modified :4/24/94      By: D. Hannifan
'
'       Comments : Determine drag call type and drag icon to show
'                  Turn On/Off scroll timer during drag events
'
'*****************************************************************
'
Private Sub lbcAiring_DragOver(Index As Integer, Source As control, X As Single, Y As Single, State As Integer)
    Dim ilDragType As Integer    '1=Link 0=Swap or Move
    Dim ilListIndex As Integer   'Target index
    Dim ilSourceType As Integer  '0=Selling ; 1=Airing

    If imDType = 2 Then
        Exit Sub
    End If

    ilDragType = mDragType(lbcAiring(Index), X, Y, ilListIndex)
    If (State = vbEnter) Or (State = vbOver) Then
        If (ilListIndex < 0) Or (ilListIndex > lbcAiring(Index).ListCount - 1) Then
            Exit Sub
        End If
    End If
    If (Source.Top > 500) Then
        ilSourceType = 1 ' Airing
    Else
        ilSourceType = 0 ' Selling
    End If

    If (State = vbEnter) Then   'Turn off scroll
        If tmcScroll.Enabled Then
            tmcScroll.Enabled = False
        End If
    End If
    If (State = vbLeave) Then   'Set scroll variables
        imLeaveIndex = Index
        smLeaveType = "A"
        If Y < 500 Then
            imDExitDirect = 0
            imDTop = lbcAiring(Index).Top - imHeight
            imDBottom = lbcAiring(Index).Top
        Else
            imDExitDirect = 1
            imDTop = lbcAiring(Index).Top + lbcAiring(Index).Height
            imDBottom = lbcAiring(Index).Top + lbcAiring(Index).Height + imHeight
        End If
        imDLeft = lbcAiring(Index).Left
        imDRight = lbcAiring(Index).Left + imBoxWidth
        lbcAiring(Index).ListIndex = -1
    End If
    'Determine the type of drag event to process (drag, swap or link)
    If (State = vbEnter) Or (State = vbOver) Then
        If imDragSource = 1 Then    'Airing
            If (imDragType = 1) And (ilDragType = 1) Then  'Swap links
                If (imDragListIndex <> ilListIndex) Or (Index <> imDragIndex) Then
                    lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconSwap.DragIcon
                    If (State = vbOver) And Not (lbcAiring(Index).Selected(ilListIndex)) And (ilListIndex <= lbcAiring(Index).ListCount - 1) Then
                        lbcAiring(Index).Selected(ilListIndex) = True
                    End If
                Else
                    lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
                End If
            ElseIf (imDragType = 1) And (ilDragType = 0) Then   'Move link
                lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconMove.DragIcon
                If (State = vbOver) And Not (lbcAiring(Index).Selected(ilListIndex)) And (ilListIndex <= lbcAiring(Index).ListCount - 1) Then
                    lbcAiring(Index).Selected(ilListIndex) = True
                End If
            Else
                lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
            End If
        Else    'Selling
            If (imDragType = 0) And (ilDragType = 0) Then
                lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconLink.DragIcon
                If (State = vbOver) And Not (lbcAiring(Index).Selected(ilListIndex)) And (ilListIndex <= lbcAiring(Index).ListCount - 1) Then
                    lbcAiring(Index).Selected(ilListIndex) = True
                End If
            Else
                lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
            End If
        End If
    Else    'DragLeave
        If (imDragIndex <= imNoAiring - 1) Then
            lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
        End If
    End If
End Sub
'****************************************************************
'
'           Procedure Name : lbcAiring_MouseDown
'
'       Date Created : ?            By: D. LeVine
'       Date Modified : 4/24/94     By: D. Hannifan
'
'       Comments : Enable Drag control timer & initialize
'                  drag event variables
'*****************************************************************
'
Private Sub lbcAiring_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Index = index of source list box for drag event

    Dim ilRet As Integer
    Dim slName1 As String
    Dim slStr As String
    Dim ilFound As Integer
    Dim ilLen As Integer


    mSetSelected Index, 1, 1, 1, "A", "ALLOTHERS"
    imDSourceIndex = Index
    smDSourceType = "A"  'Airing
    ilFound = True
    If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        Exit Sub
    End If
    imDragSource = 1
    imDragIndex = Index
    imDragType = mDragType(lbcAiring(Index), X, Y, imDragListIndex)

    'Check if a valid link was selected
    If (Left$(lbcAiring(Index).List(imDragListIndex), 2) = "  ") Then
        imDType = 1 'swap or move operation
        ilLen = Len(Trim$(lbcAiring(Index).List(imDragListIndex)))
        ilRet = InStr(1, Trim$(lbcAiring(Index).List(imDragListIndex)), " ")
        slStr = right$(Trim$(lbcAiring(Index).List(imDragListIndex)), ilLen - ilRet)
        slStr = Trim$(slStr)
        ilRet = gParseItem(Trim$(slStr), 1, "@", slStr)
        slName1 = Trim$(slStr)
        For ilRet = 0 To imNoSelling - 1 Step 1
            If (slName1 = Trim$(lacSelling(ilRet).Caption)) Then
                ilFound = True
                Exit For
            Else
                ilFound = False
            End If
        Next ilRet
    Else
        imDType = 0 'link or drag operation
    End If
    If Not ilFound Then
        tmcDrag.Enabled = False
        imDType = 2  'invalid operation
        Exit Sub
    End If
    imAirClickSource = imDragSource
    imAirClickIndex = imDragIndex
    imAirClickType = imDragType
    imAirClickListIndex = imDragListIndex
    tmcDrag.Enabled = True
    mSetCommands
End Sub
'****************************************************************
'
'            Procedure Name : lbcAiring_MouseUp
'
'       Date Created : ?            By: D. LeVine
'       Date Modified : 4/24/94     By: D. Hannifan
'
'       Comments : Disable Drag & scroll control timers
'                  & reinitialize dragevent counters
'
'*****************************************************************
'
Private Sub lbcAiring_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If tmcDrag.Enabled Then
        tmcDrag.Enabled = False
        If imDragSource = 0 Then    'Selling
            lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconStd.DragIcon
        Else                        'Airing
            lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconStd.DragIcon
            If (Index = imAirClickIndex) And (imSellClickIndex >= 0) Then
                imDragSource = imSellClickSource
                imDragIndex = imSellClickIndex
                imDragType = imSellClickType
                imDragListIndex = imSellClickListIndex
                lbcAiring_DragDrop Index, lbcSelling(imSellClickIndex), X, Y
                imSellClickIndex = -1
                imAirClickIndex = -1
            ElseIf (Index = imAirClickIndex) And (imAirClickListIndex = imDragListIndex) And (imAirClickSource = imDragSource) And ((Shift And vbCtrlMask) = vbCtrlMask) Then
                lbcAiring(Index).ListIndex = -1
                imSellClickIndex = -1
                imAirClickIndex = -1
            End If
        End If
    End If

    tmcDrag.Enabled = False
    lbcAiring(Index).Enabled = True
    tmcScroll.Enabled = False
    mSetCommands
End Sub
'****************************************************************
'
'                   Procedure Name : lbcSelling_DragDrop
'
'       Date Created : ?            By: D. LeVine
'       Date Modified :4/24/94      By: D. Hannifan
'
'       Comments : Determine drag call type and process a
'                  swap, move or link event
'
'*****************************************************************
Private Sub lbcSelling_DragDrop(Index As Integer, Source As control, X As Single, Y As Single)
    Dim ilDragType As Integer           '1=Link , 0=Swap Or Move
    Dim ilDragListIndex As Integer      'Index of target
    Dim slName1 As String               'List item string vehicle 1
    Dim slName2 As String               'List item string vehicle 2
    Dim ilSaveListIndex As Integer      'Temporary list item index for swap/move
    Dim ilSaveIndex As Integer          'Temporary list index for swap/move
    Dim ilSaveDragListIndex As Integer  'Temporary source list index for swap/move
    Dim ilIndexShift As Integer         'Increment/Decrement value for list item
    Dim ilLoop As Integer               'DragListIndex counter
    Dim ilDragIndex As Integer          'Index of drag source
    Dim ilRet As Integer                'Return value from call
    Dim slStr As String                 'parse string
    Dim ilSameTime As Integer           'True=swap in same listbox with same sell time
    Dim ilSetIndex1 As Integer          'Listbox1 Index for mSetSelected call
    Dim ilSetIndex2 As Integer          'Listbox2 Index for mSetSelected call
    Dim ilSetList1 As Integer           'Listbox1 List Index for mSetSelected call
    Dim ilSetList2 As Integer           'Listbox2 ListIndex for mSetSelected call
    Dim ilDelimLoc As Integer           'Delimeter location for parsing
    Dim ilLen As Integer                'String length
    Dim slStr2 As String                'parsing string
    On Error GoTo lbcSellingErr

    imSellClickIndex = -1
    imAirClickIndex = -1
    ilDragIndex = imDragIndex 'Save imDragIndex at call
    imSwapFlag = False
    If imDragSource = 0 Then    'Selling
        lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconStd.DragIcon
    Else                        'Airing
        lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconStd.DragIcon
    End If

    ilDragType = mDragType(lbcSelling(Index), X, Y, ilDragListIndex)
    If imDragSource = 0 Then    'Selling
        If (imDragType = 1) And (ilDragType = 1) Then  'Swap links
            imSwapFlag = True
            imLegal = True
            If (Index <> imDragIndex) Then  'swap in different list box
                ilSetIndex1 = Index
                ilSetIndex2 = imDragIndex
                ilSetList1 = ilDragListIndex
                ilSetList2 = imDragListIndex
                ilSaveListIndex = ilDragListIndex
                ilSaveIndex = imDragIndex
                ilSaveDragListIndex = imDragListIndex - 1
                mMoveSelling Index, ilDragListIndex
                If Not (imLegal) Then
                    Exit Sub
                End If
                imDragIndex = Index
                imDragListIndex = ilSaveListIndex
                If imLegal Then
                    mMoveSelling ilSaveIndex, ilSaveDragListIndex
                    mSetSelected ilSetIndex1, ilSetIndex2, ilSetList1, ilSetList2, "S", "S"
                    imUpdateFlag = True
                Else
                    Exit Sub
                End If
                imLegal = True
                tmcScroll.Enabled = False
            End If
            If (ilDragIndex = imDragIndex) Then
                ilSameTime = mCheckTime(ilDragIndex, ilDragListIndex, imDragListIndex, "S")
                If (ilDragListIndex < imDragListIndex) And Not (ilSameTime) Then 'swap within the same list box std order
                    imSwapFlag = True
                    ilSaveListIndex = ilDragListIndex
                    ilSaveIndex = imDragIndex
                    ilSaveDragListIndex = imDragListIndex
                    mMoveSelling Index, ilDragListIndex
                    If Not (imLegal) Then
                        Exit Sub
                    End If
                    imDragIndex = Index
                    imDragListIndex = ilSaveListIndex
                    If imLegal Then
                        mMoveSelling ilSaveIndex, ilSaveDragListIndex
                        mSetSelected 1, 1, 1, 1, "C", "C"
                        imUpdateFlag = True
                    Else
                        Exit Sub
                    End If
                    imLegal = True
                    tmcScroll.Enabled = False
                End If
                If (ilDragListIndex > imDragListIndex) And Not (ilSameTime) Then 'swap within the same list box in reverse order
                    imSwapFlag = True
                    ilSaveListIndex = ilDragListIndex
                    ilSaveIndex = imDragIndex
                    ilSaveDragListIndex = imDragListIndex
                    imDragIndex = Index
                    imDragListIndex = ilSaveListIndex
                    mMoveSelling ilSaveIndex, ilSaveDragListIndex
                    If imLegal Then
                        mMoveSelling Index, ilDragListIndex
                        mSetSelected 1, 1, 1, 1, "C", "C"
                        imUpdateFlag = True
                    Else
                        Exit Sub
                    End If
                    imLegal = True
                    tmcScroll.Enabled = False
                End If
                If (ilSameTime) Then
                    slStr2 = Trim$(lbcSelling(ilDragIndex).List(ilDragListIndex))
                    ilLen = Len(slStr2)
                    ilDelimLoc = InStr(1, slStr2, " ")
                    slStr2 = Trim$(right$(slStr2, ilLen - ilDelimLoc))
                    ilRet = gParseItem(slStr2, 1, "@", slStr2)
                    slStr2 = Trim$(slStr2)
                    imLegal = False
                    For ilRet = 0 To imNoAiring - 1 Step 1
                        If (slStr2 = Trim$(lacAiring(ilRet).Caption)) Then
                            imLegal = True
                            Exit For
                        End If
                    Next ilRet
                    If imLegal Then
                        slStr = lbcSelling(ilDragIndex).List(ilDragListIndex)
                        lbcSelling(ilDragIndex).List(ilDragListIndex) = lbcSelling(ilDragIndex).List(imDragListIndex)
                        lbcSelling(ilDragIndex).List(imDragListIndex) = slStr
                        imUpdateFlag = True
                        tmcScroll.Enabled = False
                        ilSameTime = False
                        mSetSelected 1, 1, 1, 1, "C", "C"
                    Else
                        imLegal = True
                        Exit Sub
                    End If
                End If
            End If
            imSwapFlag = False
            imLegal = True
        End If
        If (imDragType = 1) And (ilDragType = 0) Then   'Move link
            imSwapFlag = True
            mMoveSelling Index, ilDragListIndex
            If imLegal Then
                imUpdateFlag = True
            Else
                Exit Sub
            End If
            imLegal = True
        End If
    Else    'Airing
        If (imDragType = 0) And (ilDragType = 0) Then 'Link
            'Construct link strings
            ilRet = gParseItem(lbcAiring(imDragIndex).List(imDragListIndex), 1, " ", slStr)
            slStr = Trim$(slStr)
            slName1 = CStr("  " & Trim$(slStr) & " " & Trim$(lacAiring(imDragIndex).Caption))
            ilRet = gParseItem(lbcAiring(imDragIndex).List(imDragListIndex), 2, "@", slStr)
            slName1 = CStr(slName1 & "@" & Trim$(slStr))
            ilRet = gParseItem(lbcSelling(Index).List(ilDragListIndex), 1, " ", slStr)
            slStr = Trim$(slStr)
            slName2 = CStr("  " & Trim$(slStr) & " " & Trim$(lacSelling(Index).Caption))
            ilRet = gParseItem(lbcSelling(Index).List(ilDragListIndex), 2, "@", slStr)
            slName2 = CStr(slName2 & "@" & Trim$(slStr))
' Check for collate order before adding to airing list box
            ilIndexShift = 1
            If (imDragSource = 1) Then
            Else
                ilLoop = imDragListIndex + 1
                Do While ilLoop < lbcAiring(imDragIndex).ListCount
                    If (Left$(lbcAiring(imDragIndex).List(imDragListIndex + ilIndexShift), 2) = "  ") Then
                        If (lbcAiring(imDragIndex).List(imDragListIndex + ilIndexShift) = slName2) Then 'Duplicate
                            Beep
                            mSetSelected 1, 1, 1, 1, "C", "C"
                            DoEvents
                            Exit Sub
                        End If
                        ilIndexShift = ilIndexShift + 1
                        ilLoop = ilLoop + 1
                    Else
                        Exit Do
                    End If
                Loop
                'Add to airing list box
                lbcAiring(imDragIndex).AddItem mFillTo100(slName2), imDragListIndex + ilIndexShift
                mSetSelected imDragIndex, 1, imDragListIndex + ilIndexShift, 1, "A", "C"
                imUpdateFlag = True
                tmcScroll.Enabled = False
            End If
' Check for collate order before adding to selling list box
            ilIndexShift = 1
            If (imDragSource = 0) Then
            Else
                ilLoop = ilDragListIndex + 1
                Do While ilLoop < lbcSelling(Index).ListCount
                    If (Left$(lbcSelling(Index).List(ilDragListIndex + ilIndexShift), 2) = "  ") Then
                        If (lbcSelling(Index).List(ilDragListIndex + ilIndexShift) = slName1) Then 'Duplicate
                            Beep
                            mSetSelected 1, 1, 1, 1, "C", "C"
                            DoEvents
                            Exit Sub
                        End If
                        ilIndexShift = ilIndexShift + 1
                        ilLoop = ilLoop + 1
                    Else: Exit Do
                    End If
                Loop
               ' Add to selling list box
                lbcSelling(Index).AddItem mFillTo100(slName1), ilDragListIndex + ilIndexShift
                mSetSelected Index, 1, ilDragListIndex + ilIndexShift, 1, "S", "C"
                imUpdateFlag = True
                tmcScroll.Enabled = False
            End If
            ilIndexShift = 1
            If (imDragSource = 1) Then
                ilLoop = imDragListIndex + 1
                Do While ilLoop < lbcAiring(imDragIndex).ListCount
                    If (Left$(lbcAiring(imDragIndex).List(imDragListIndex + ilIndexShift), 2) = "  ") Then
                        If (lbcAiring(imDragIndex).List(imDragListIndex + ilIndexShift) = slName2) Then 'Duplicate
                            Beep
                            mSetSelected 1, 1, 1, 1, "C", "C"
                            DoEvents
                            Exit Sub
                        End If
                        ilIndexShift = ilIndexShift + 1
                        ilLoop = ilLoop + 1
                    Else: Exit Do
                    End If
                Loop
                lbcAiring(imDragIndex).AddItem mFillTo100(slName2), imDragListIndex + ilIndexShift
                mSetSelected imDragIndex, 1, imDragListIndex + ilIndexShift, 1, "A", "C"
                imUpdateFlag = True
                tmcScroll.Enabled = False
            End If
            ilIndexShift = 1
            If (imDragSource = 0) Then
                lbcSelling(Index).AddItem mFillTo100(slName1), ilDragListIndex + ilIndexShift
                mSetSelected Index, 1, ilDragListIndex + ilIndexShift, 1, "S", "C"
                imUpdateFlag = True
                tmcScroll.Enabled = False
            End If
        End If
    End If
    tmcScroll.Enabled = False
    mSetCommands
    Exit Sub
lbcSellingErr:
    On Error GoTo 0
    imSwapFlag = False
    Exit Sub
End Sub
'****************************************************************
'
'                   Procedure Name : lbcSelling_DragOver
'
'       Date Created : ?            By: D. LeVine
'       Date Modified :4/24/94      By: D. Hannifan
'
'       Comments : Determine drag call type and drag icon to show
'                  Turn On/Off scroll timer during drag events
'
'*****************************************************************
'
Private Sub lbcSelling_DragOver(Index As Integer, Source As control, X As Single, Y As Single, State As Integer)
    Dim ilDragType As Integer    '1=Link 0=Swap or Move
    Dim ilListIndex As Integer   'Target index
    Dim ilSourceType As Integer  '0=Selling ; 1=Airing

    If imDType = 2 Then
        Exit Sub
    End If

    ilDragType = mDragType(lbcSelling(Index), X, Y, ilListIndex)
    If (State = vbEnter) Or (State = vbOver) Then
        If (ilListIndex < 0) Or (ilListIndex > lbcSelling(Index).ListCount - 1) Then
            Exit Sub
        End If
    End If
    If (Source.Top > 500) Then
        ilSourceType = 1 ' Airing
    Else
        ilSourceType = 0 ' Selling
    End If

    If (State = vbEnter) Then    'Turn off scroll
        If tmcScroll.Enabled Then
            tmcScroll.Enabled = False
        End If
    End If
    If (State = vbLeave) Then   'Set scroll variables
        imLeaveIndex = Index
        smLeaveType = "S"
        If Y < 500 Then
            imDExitDirect = 0
            imDTop = lbcSelling(Index).Top - imHeight
            imDBottom = lbcSelling(Index).Top
        Else
            imDExitDirect = 1
            imDTop = lbcSelling(Index).Top + lbcSelling(Index).Height
            imDBottom = lbcSelling(Index).Top + lbcSelling(Index).Height + imHeight
        End If
        imDLeft = lbcSelling(Index).Left
        imDRight = lbcSelling(Index).Left + imBoxWidth
        lbcSelling(Index).ListIndex = -1
    End If
    'Determine the type of drag process and set dragicon (swap , move or link)
    If (State = vbEnter) Or (State = vbOver) Then
        If imDragSource = 0 Then    'Selling
            If (imDragType = 1) And (ilDragType = 1) Then  'Swap links
                If (imDragListIndex <> ilListIndex) Or (Index <> imDragIndex) Then
                    lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconSwap.DragIcon
                    If (State = vbOver) And Not (lbcSelling(Index).Selected(ilListIndex)) And (ilListIndex <= lbcSelling(Index).ListCount - 1) Then
                        lbcSelling(Index).Selected(ilListIndex) = True
                    End If
                Else
                    lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
                End If
            ElseIf (imDragType = 1) And (ilDragType = 0) Then   'Move link
                lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconMove.DragIcon
                If (State = vbOver) And Not (lbcSelling(Index).Selected(ilListIndex)) And (ilListIndex <= lbcSelling(Index).ListCount - 1) Then
                    lbcSelling(Index).Selected(ilListIndex) = True
                End If
            Else
                lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
            End If
        Else    'Airing
            If (imDragType = 0) And (ilDragType = 0) Then
                lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconLink.DragIcon
                If (State = vbOver) And Not (lbcSelling(Index).Selected(ilListIndex)) And (ilListIndex <= lbcSelling(Index).ListCount - 1) Then
                    lbcSelling(Index).Selected(ilListIndex) = True
                End If
            Else
                lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
            End If
        End If
    Else    'DragLeave
        If (imDragIndex <= imNoSelling - 1) Then
            lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
        End If
    End If
End Sub
'****************************************************************
'
'           Procedure Name : lbcSelling_MouseDown
'
'       Date Created : ?            By: D. LeVine
'       Date Modified : 4/24/94     By: D. Hannifan
'
'       Comments : Enable Drag control timer & initialize
'                  drag event variables
'*****************************************************************
'
Private Sub lbcSelling_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Initialize dragevent variables
    Dim ilRet As Integer    'Return value from call
    Dim slName1 As String   'list string selected
    Dim slStr As String     'parse string
    Dim ilFound As Integer  'True = matching air list box exists
    Dim ilLen As Integer
    mSetSelected Index, 1, 1, 1, "S", "ALLOTHERS"
    imDSourceIndex = Index
    smDSourceType = "S"  'Selling
    ilFound = True
    If (igWinStatus(PROGRAMMINGJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        Exit Sub
    End If
    imDragSource = 0
    imDragIndex = Index
    imDragType = mDragType(lbcSelling(Index), X, Y, imDragListIndex)
    'Check to find if a valid link was selected
    If (Left$(lbcSelling(Index).List(imDragListIndex), 2) = "  ") Then
        imDType = 1 'swap or move operation
        ilLen = Len(Trim$(lbcSelling(Index).List(imDragListIndex)))
        ilRet = InStr(1, Trim$(lbcSelling(Index).List(imDragListIndex)), " ")
        slStr = right$(Trim$(lbcSelling(Index).List(imDragListIndex)), ilLen - ilRet)
        slStr = Trim$(slStr)
        ilRet = gParseItem(Trim$(slStr), 1, "@", slName1)
        slName1 = Trim$(slName1)
        For ilRet = 0 To imNoAiring - 1 Step 1
            If (slName1 = Trim$(lacAiring(ilRet).Caption)) Then
                ilFound = True
                Exit For
            Else
                ilFound = False
            End If
        Next ilRet
    Else
        imDType = 0 'drag or link operation
    End If
    If ilFound = False Then
        imDType = 2  'invalid operation
        tmcDrag.Enabled = False
        Exit Sub
    End If
    imSellClickSource = imDragSource
    imSellClickIndex = imDragIndex
    imSellClickType = imDragType
    imSellClickListIndex = imDragListIndex
    tmcDrag.Enabled = True
    mSetCommands
End Sub
'****************************************************************
'
'            Procedure Name : lbcSelling_MouseUp
'
'       Date Created : ?            By: D. LeVine
'       Date Modified : 4/24/94     By: D. Hannifan
'
'       Comments : Disable Drag & Scroll control timer
'                  & reinitialize dragevent counters
'
'*****************************************************************
'
Private Sub lbcSelling_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                                                                               *
'******************************************************************************************


    tmcScroll.Enabled = False
    If tmcDrag.Enabled Then
        tmcDrag.Enabled = False
        If imDragSource = 0 Then    'Selling
            lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconStd.DragIcon
            If (Index = imSellClickIndex) And (imAirClickIndex >= 0) Then
                imDragSource = imAirClickSource
                imDragIndex = imAirClickIndex
                imDragType = imAirClickType
                imDragListIndex = imAirClickListIndex
                lbcSelling_DragDrop Index, lbcAiring(imAirClickIndex), X, Y
                imSellClickIndex = -1
                imAirClickIndex = -1
            ElseIf (Index = imSellClickIndex) And (imSellClickListIndex = imDragListIndex) And (imSellClickSource = imDragSource) And ((Shift And vbCtrlMask) = vbCtrlMask) Then
                lbcSelling(Index).ListIndex = -1
                imSellClickIndex = -1
                imAirClickIndex = -1
            End If
        Else                        'Airing
            lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconStd.DragIcon
        End If
    End If
    tmcDrag.Enabled = False
    tmcScroll.Enabled = False
    mSetCommands
End Sub
'**********************************************************************
'
'       Procedure Name : mAnyPendForAiring
'
'       Created : 4/24/94       By: D. Hannifan
'       Modified :              By:
'
'       Comments : Determine if any pending exist for selling
'            Note:
'
'**********************************************************************
Private Function mAnyPendForAiring(ilVefCode As Integer) As Integer
    'Dim tlVlfSrchKey As VLFKEY1
    Dim ilRet As Integer
    Dim tlVlf As VLF
    Dim ilVlfRecLen As Integer
    ilVlfRecLen = Len(tlVlf)
    'tlVlfSrchKey.iAirCode = ilVefCode
    'tlVlfSrchKey.iAirDay = imDateCode
    'tlVlfSrchKey.iEffDate(0) = 0
    'tlVlfSrchKey.iEffDate(1) = 0
    'tlVlfSrchKey.iAirTime(0) = 0
    'tlVlfSrchKey.iAirTime(1) = 0
    'tlVlfSrchKey.iAirPosNo = 0
    'tlVlfSrchKey.iAirSeq = 0
    'ilRet = btrGetGreaterOrEqual(hmVlfPop, tlVlf, ilVlfRecLen, tlVlfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get current record
    'Do While (ilRet = BTRV_ERR_NONE) And (tlVlf.iAirCode = ilVefCode)
    '    If tlVlf.iSellDay = imDateCode Then
    '        If tlVlf.sStatus = "P" Then
    '            mAnyPendForAiring = True
    '            Exit Function
    '        End If
    '    End If
    '    ilRet = btrGetNext(hmVlfPop, tlVlf, ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    'Loop
    
    tmVlfSrchKey4.iAirCode = ilVefCode
    tmVlfSrchKey4.iAirDay = imDateCode
    tmVlfSrchKey4.sStatus = "P"
    ilRet = btrGetEqual(hmVlfPop, tlVlf, ilVlfRecLen, tmVlfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)  'Get current record
    If (ilRet = BTRV_ERR_NONE) Then
        mAnyPendForAiring = True
        Exit Function
    End If

    mAnyPendForAiring = False
End Function
'**********************************************************************
'
'       Procedure Name : mAnyPendForSelling
'
'       Created : 4/24/94       By: D. Hannifan
'       Modified :              By:
'
'       Comments : Determine if any pending exist for selling
'            Note:
'
'**********************************************************************
Private Function mAnyPendForSelling(ilVefCode As Integer) As Integer
    'Dim tlVlfSrchKey As VLFKEY0
    Dim ilRet As Integer
    Dim tlVlf As VLF
    Dim ilVlfRecLen As Integer
    ilVlfRecLen = Len(tlVlf)
    'tlVlfSrchKey.iSellCode = ilVefCode
    'tlVlfSrchKey.iSellDay = imDateCode
    'tlVlfSrchKey.iEffDate(0) = 0
    'tlVlfSrchKey.iEffDate(1) = 0
    'tlVlfSrchKey.iSellTime(0) = 0
    'tlVlfSrchKey.iSellTime(1) = 0
    'tlVlfSrchKey.iSellPosNo = 0
    'tlVlfSrchKey.iSellSeq = 0
    'ilRet = btrGetGreaterOrEqual(hmVlfPop, tlVlf, ilVlfRecLen, tlVlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
    'Do While (ilRet = BTRV_ERR_NONE) And (tlVlf.iSellCode = ilVefCode)
    '    If tlVlf.iSellDay = imDateCode Then
    '        If tlVlf.sStatus = "P" Then
    '            mAnyPendForSelling = True
    '            Exit Function
    '        End If
    '    End If
    '    ilRet = btrGetNext(hmVlfPop, tlVlf, ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    'Loop
    
    tmVlfSrchKey3.iSellCode = ilVefCode
    tmVlfSrchKey3.iSellDay = imDateCode
    tmVlfSrchKey3.sStatus = "P"
    ilRet = btrGetEqual(hmVlfPop, tlVlf, ilVlfRecLen, tmVlfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)  'Get current record
    If (ilRet = BTRV_ERR_NONE) Then
        mAnyPendForSelling = True
        Exit Function
    End If
    
    mAnyPendForSelling = False
End Function
Private Function mCheckTime(ilDragIndex As Integer, ilDragListIndex As Integer, ilDragListIndex2 As Integer, slType As String) As Integer
    'ilDragIndex (I)        : listbox index
    'ilDragListIndex (I)    : listbox item 1
    'ilDragListIndex2 (I)   : listbox item 2
    'slType (I)             : "S" = selling  ; "A" = airing
    '
    'mCheckTime (O)         : True= same time found for both list items
    Dim ilCount As Integer      'List item counter
    Dim slTime As String        'air time associated with ilDragListIndex
    Dim slTime2 As String       'air time associated with ilDragListIndex2
    Dim ilRet As Integer
    mCheckTime = False
    If slType = "A" Then 'airing
        For ilCount = ilDragListIndex To 0 Step -1
            If (Left$(lbcAiring(ilDragIndex).List(ilCount), 2) <> "  ") Then
                ilRet = gParseItem(LTrim$(lbcAiring(ilDragIndex).List(ilCount)), 1, " ", slTime)
                slTime = Trim$(slTime)
                Exit For
            End If
        Next ilCount
        For ilCount = ilDragListIndex2 To 0 Step -1
            If (Left$(lbcAiring(ilDragIndex).List(ilCount), 2) <> "  ") Then
                ilRet = gParseItem(LTrim$(lbcAiring(ilDragIndex).List(ilCount)), 1, " ", slTime2)
                slTime2 = Trim$(slTime2)
                Exit For
            End If
        Next ilCount
        If (slTime = slTime2) Then
            mCheckTime = True
            Exit Function
        Else
            mCheckTime = False
            Exit Function
        End If
    Else 'selling
        For ilCount = ilDragListIndex To 0 Step -1
            If (Left$(lbcSelling(ilDragIndex).List(ilCount), 2) <> "  ") Then
                ilRet = gParseItem(LTrim$(lbcSelling(ilDragIndex).List(ilCount)), 1, " ", slTime)
                slTime = Trim$(slTime)
                Exit For
            End If
        Next ilCount
        For ilCount = ilDragListIndex2 To 0 Step -1
            If (Left$(lbcSelling(ilDragIndex).List(ilCount), 2) <> "  ") Then
                ilRet = gParseItem(LTrim$(lbcSelling(ilDragIndex).List(ilCount)), 1, " ", slTime2)
                slTime2 = Trim$(slTime2)
                Exit For
            End If
        Next ilCount
        If (slTime = slTime2) Then
            mCheckTime = True
            Exit Function
        Else
            mCheckTime = False
            Exit Function
        End If
    End If
End Function
'**********************************************************************
'
'       Procedure Name : mCloseFiles
'
'       Created : 4/24/94       By: D. Hannifan
'       Modified :              By:
'
'       Comments : Procedure to close files resident in LinksDef
'            Note: hmVlfPop is still open...closed in mTerminate
'
'**********************************************************************
'
Private Sub mCloseFiles()
    Dim ilRet As Integer

    'Close LCF and LEF after initialization ; leave tmVlfPop open
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmLef)
    btrDestroy hmLef
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDeleteAiring                   *
'*                                                     *
'*             Created:10/16/93      By:D. LeVine      *
'*            Modified:4/24/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Remove a link                  *
'*                                                     *
'*******************************************************
Private Sub mDeleteAiring()
    Dim slName1 As String       'String from source
    Dim slName2 As String       'String to match in selling
    Dim slName As String        'Parsed string from source
    Dim slTime As String        'Time and maxunits string from source
    Dim ilFound As Integer      'Search flag
    Dim ilIndex As Integer      'Index of target list box
    Dim ilLoop As Integer       'Loop Counter for list item index
    Dim ilRet As Integer        'General return variable
    Dim slSrchTime As String    'Time string to search for in selling
    Dim ilLen As Integer        'String length
    Dim ilCount As Integer      'List Item index incrementor/decrementor
    Dim slTime2 As String       'Time string found in target list
    Dim ilCount2 As Integer     'Matching string found flag
    Dim slStr As String         'Parsing string

    On Error GoTo mDeleteAiringErr

    'Get vehicle name and time from source
    slName1 = lbcAiring(imDragIndex).List(imDragListIndex)
    slStr = LTrim$(slName1)
    ilLen = Len(slStr)
    ilRet = InStr(1, slStr, " ")
    slName = right$(slStr, ilLen - ilRet)
    slName = LTrim$(slName)
    ilRet = gParseItem(slName, 1, "@", slName)
    slName = Trim$(slName)
    ilRet = gParseItem(slStr, 1, " ", slTime)
    slSrchTime = Trim$(slTime) 'Time to search for in selling

    ilFound = False
    ilIndex = 0

    Do
        If lacSelling(ilIndex).Caption = slName Then
            ilFound = True
            Exit Do
        End If
        ilIndex = ilIndex + 1
    Loop Until ilFound  'If not found- an error will occur
    ilFound = False
    ilLoop = imDragListIndex - 1
    Do
        If Left$(lbcAiring(imDragIndex).List(ilLoop), 2) <> "  " Then
            ilFound = True
            Exit Do
        End If
        ilLoop = ilLoop - 1
    Loop Until ilFound
    ilRet = gParseItem(LTrim$(lbcAiring(imDragIndex).List(ilLoop)), 1, " ", slStr)
    slStr = Trim$(slStr)
    slName2 = CStr("  " & Trim$(slStr) & " " & Trim$(lacAiring(imDragIndex).Caption))
    ilRet = gParseItem(lbcAiring(imDragIndex).List(ilLoop), 2, "@", slStr)
    slName2 = mFillTo100(RTrim$(slName2 & "@" & slStr))
    gFindMatch slName2, 0, lbcSelling(ilIndex)
    If gLastFound(lbcSelling(ilIndex)) < 0 Then
        Exit Sub
    End If
    'Check for matching string : if entire string is not the same then find next
    ilCount2 = 0
    ilCount = 1
    Do
        If (Left$(lbcSelling(ilIndex).List(gLastFound(lbcSelling(ilIndex)) - ilCount), 2) <> "  ") Then
            ilRet = gParseItem(LTrim$(lbcSelling(ilIndex).List(gLastFound(lbcSelling(ilIndex)) - ilCount)), 1, " ", slTime2)
            slTime2 = Trim$(slTime2)
            Exit Do
        Else
            ilCount = ilCount + 1
        End If
    Loop
    Do While ilCount2 = 0
        If slTime2 <> slSrchTime Then
            gFndNext lbcSelling(ilIndex), slName2
            If gLastFound(lbcSelling(ilIndex)) < 0 Then
                Exit Sub
            End If
            ilCount = 1
lRepeatSrchSell:
            Do
                If (Left$(lbcSelling(ilIndex).List(gLastFound(lbcSelling(ilIndex)) - ilCount), 2) <> "  ") Then
                    ilRet = gParseItem(LTrim$(lbcSelling(ilIndex).List(gLastFound(lbcSelling(ilIndex)) - ilCount)), 1, " ", slTime2)
                    slTime2 = Trim$(slTime2)
                    Exit Do
                Else
                    ilCount = ilCount + 1
                End If
            Loop
            If (slTime2 <> slSrchTime) And (gLastFound(lbcSelling(ilIndex)) - ilCount > 0) Then
                ilCount = ilCount + 1
                GoTo lRepeatSrchSell
            End If
        Else
            ilCount2 = 1 'Match found
        End If
    Loop
    'Delete Link
    lbcAiring(imDragIndex).RemoveItem imDragListIndex
    lbcSelling(ilIndex).RemoveItem gLastFound(lbcSelling(ilIndex))
    imUpdateFlag = True
    mSetCommands
    Exit Sub
mDeleteAiringErr:
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDeleteSelling                  *
'*                                                     *
'*             Created:10/16/93      By:D. LeVine      *
'*            Modified:4/24/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Remove links                   *
'*                                                     *
'*******************************************************
Private Sub mDeleteSelling()
    Dim slName1 As String       'String from source
    Dim slName2 As String       'String to match in airing
    Dim slName As String        'Parsed string from source
    Dim slTime As String        'Time and maxunits string from source
    Dim ilFound As Integer      'Search flag
    Dim ilIndex As Integer      'Index of target list box
    Dim ilLoop As Integer       'Loop Counter for list item index
    Dim ilRet As Integer        'General return variable
    Dim slSrchTime As String    'Time string to search for in airing
    Dim ilLen As Integer        'String length
    Dim ilCount As Integer      'List Item index incrementor/decrementor
    Dim slTime2 As String       'Time string found in target list
    Dim ilCount2 As Integer     'Matching string found flag
    Dim slStr As String         'Parsing string

    On Error GoTo mDeleteSellingErr

    'Get vehicle name and time from source
    slName1 = lbcSelling(imDragIndex).List(imDragListIndex)
    slStr = LTrim$(slName1)
    ilLen = Len(slStr)
    ilRet = InStr(1, slStr, " ")
    slName = right$(slStr, ilLen - ilRet)
    slName = LTrim$(slName)
    ilRet = gParseItem(slName, 1, "@", slName)
    slName = Trim$(slName)
    ilRet = gParseItem(slStr, 1, " ", slTime)
    slSrchTime = Trim$(slTime)  'Time to search for in selling

    ilFound = False
    ilIndex = 0
    slName = Trim$(slName)
    Do
        If lacAiring(ilIndex).Caption = slName Then
            ilFound = True
            Exit Do
        End If
        ilIndex = ilIndex + 1
    Loop Until ilFound  'If not found- an error will occur
    ilFound = False
    ilLoop = imDragListIndex - 1
    Do
        If Left$(lbcSelling(imDragIndex).List(ilLoop), 2) <> "  " Then
            ilFound = True
            Exit Do
        End If
        ilLoop = ilLoop - 1
    Loop Until ilFound
    'Create string to match in selling list box
    ilRet = gParseItem(LTrim$(lbcSelling(imDragIndex).List(ilLoop)), 1, " ", slStr)
    slStr = Trim$(slStr)
    slName2 = CStr("  " & Trim$(slStr) & " " & Trim$(lacSelling(imDragIndex).Caption))
    ilRet = gParseItem(lbcSelling(imDragIndex).List(ilLoop), 2, "@", slStr)
    slName2 = mFillTo100(slName2 & "@" & slStr)
    gFindMatch slName2, 0, lbcAiring(ilIndex)
    If gLastFound(lbcAiring(ilIndex)) < 0 Then
        Exit Sub
    End If
    'Check for matching string : if entire string is not the same then find next
    ilCount2 = 0
    ilCount = 1
    Do
        If (Left$(lbcAiring(ilIndex).List(gLastFound(lbcAiring(ilIndex)) - ilCount), 2) <> "  ") Then
            ilRet = gParseItem(LTrim$(lbcAiring(ilIndex).List(gLastFound(lbcAiring(ilIndex)) - ilCount)), 1, " ", slTime2)
            slTime2 = Trim$(slTime2)
            Exit Do
        Else
            ilCount = ilCount + 1
        End If
    Loop
    Do While ilCount2 = 0
        If slTime2 <> slSrchTime Then
            gFndNext lbcAiring(ilIndex), slName2
            If gLastFound(lbcAiring(ilIndex)) < 0 Then
                Exit Sub
            End If
            ilCount = 1
lRepeatSrchAir:
            Do
                If (Left$(lbcAiring(ilIndex).List(gLastFound(lbcAiring(ilIndex)) - ilCount), 2) <> "  ") Then
                    ilRet = gParseItem(LTrim$(lbcAiring(ilIndex).List(gLastFound(lbcAiring(ilIndex)) - ilCount)), 1, " ", slTime2)
                    slTime2 = Trim$(slTime2)
                    Exit Do
                Else
                    ilCount = ilCount + 1
                End If
            Loop
            If (slTime2 <> slSrchTime) And (gLastFound(lbcAiring(ilIndex)) - ilCount > 0) Then
                ilCount = ilCount + 1
                GoTo lRepeatSrchAir
            End If
        Else
            ilCount2 = 1  'Match found
        End If
    Loop
    'Delete Link
    lbcSelling(imDragIndex).RemoveItem imDragListIndex
    lbcAiring(ilIndex).RemoveItem gLastFound(lbcAiring(ilIndex))
    imUpdateFlag = True
    mSetCommands
    Exit Sub
mDeleteSellingErr:
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDragType                       *
'*                                                     *
'*             Created:10/15/93      By:D. LeVine      *
'*            Modified:3/11/94       By:D. Hannifan    *
'*                                                     *
'*      Comments: Determine type of list box item      *
'*                                                     *
'*******************************************************
Private Function mDragType(lbcCtrl As control, flX As Single, flY As Single, ilListIndex As Integer) As Integer
'
'   ilRet = mDragType(lbcCtrl, flX, flY, ilListIndex As Integer)
'   Where:
'       lbcCtrl (I)- List box control where mouse pointer is at
'       flX (I)- mouse pointer X location within control
'       flY (I)- Mouse pointer Y location within control
'       ilListIndex (O)- list index
'       ilRet (O)- 0=>Avail; 1=>Link
'

    Dim slName As String    'First two characters of list string
    ilListIndex = (flY - 15) \ fgListHtArial825 + lbcCtrl.TopIndex
    slName = Left$(lbcCtrl.List(ilListIndex), 2)
    If slName <> "  " Then  'Avails
        mDragType = 0
    Else                    'Link
        mDragType = 1
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/22/93       By:D. LeVine      *
'*            Modified:4/24/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Initialize LinksDef module     *
'*                                                     *
'*******************************************************
Private Sub mInit()

    Dim ilLoop As Integer  'List box index counter
    Dim ilRet As Integer   'Btrieve returns
    Dim slStr As String
    Dim slStr2 As String
    Dim slStr1 As String
    Dim ilSpaceBetweenButtons As Integer

    Screen.MousePointer = vbHourglass
    bmFirstCallToVpfFind = True
    imFirstActivate = True
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    imSwapFlag = False      'Initialize swap event flag
    imLegal = True
    imHeight = fgListHtArial825
    imBoxWidth = 1940
    If Not imReinitFlag Then  'Not a reinitialization process
                                'Initialize variables and h-scroll bars
        imUpdateFlag = False    'Initialize VLF update flag
        imUpperBound = 1
        imNoSelling = 0
        imNoAiring = 0
        hbcSelling.Visible = False
        hbcAiring.Visible = False
        ReDim smSellingLists(0, 0) As String
        ReDim smAiringLists(0, 0) As String
        ReDim imSellCount(1) As Integer
        ReDim imAirCount(1) As Integer
        ilRet = gObtainVef()
        imSellClickIndex = -1
        imAirClickIndex = -1

        cmcDone.Top = LinksDef.Height - cmcDone.Height - 180
        cmcCancel.Top = cmcDone.Top
        cmcUpdate.Top = cmcDone.Top
        cmcReport.Top = cmcDone.Top

        plcNetworks.Width = fmAdjFactorW * plcNetworks.Width
        plcNetworks.Height = cmcDone.Top - plcNetworks.Top - 120 'fmAdjFactorH * plcNetworks.Height
        lbcSelling(0).Width = fmAdjFactorW * lbcSelling(0).Width
        lbcSelling(0).Height = plcNetworks.Height / 2 - lacSelling(0).Height - 2 * hbcSelling.Height - 60
        lacSelling(0).Width = lbcSelling(0).Width
        lbcAiring(0).Width = lbcSelling(0).Width
        lbcAiring(0).Height = lbcSelling(0).Height
        hbcSelling.Top = lbcSelling(0).Top + lbcSelling(0).Height + 90
        lacMess.Top = plcNetworks.Height - lacMess.Height - 60
        lacMess.Left = plcNetworks.Width / 2 - lacMess.Width / 2
        hbcAiring.Top = lacMess.Top - hbcAiring.Height - 60
        lbcAiring(0).Top = hbcAiring.Top - lbcAiring(0).Height - (hbcSelling.Top - lbcSelling(0).Top - lbcSelling(0).Height)
        lacAiring(0).Top = lbcAiring(0).Top - lacAiring(0).Height - (lbcSelling(0).Top - lacSelling(0).Top - lacSelling(0).Height)
        lacAiring(0).Width = lbcAiring(0).Width
        cmcDone.Top = LinksDef.Height - cmcDone.Height - 180
        cmcCancel.Top = cmcDone.Top
        cmcUpdate.Top = cmcDone.Top
        cmcReport.Top = cmcDone.Top
        ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
        Do While ilSpaceBetweenButtons Mod 15 <> 0
            ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
        Loop
        cmcDone.Left = (LinksDef.Width - cmcDone.Width - cmcCancel.Width - cmcUpdate.Width - cmcReport.Width - 3 * ilSpaceBetweenButtons) / 2
        cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
        cmcUpdate.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
        cmcReport.Left = cmcUpdate.Left + cmcUpdate.Width + ilSpaceBetweenButtons
        imcTrash.Top = plcNetworks.Top + plcNetworks.Height + 30
        imcTrash.Left = LinksDef.Width - imcTrash.Width - 120
        ckcShow.Left = plcNetworks.Left + plcNetworks.Width - ckcShow.Width
        imBoxWidth = lbcSelling(0).Width
    End If

    slStr = Links!lacStatus.Caption
    ilRet = gParseItem(slStr, 1, "\", slStr1)
    ilRet = gParseItem(slStr, 2, "\", slStr2)
    If slStr2 = "P" Then 'Pending status mode (Vlf pending)
        smLinksDefStatus = "P"
    Else
        smLinksDefStatus = "C"  'Current status mode
    End If
    If (slStr1 = "P") And (slStr2 = "C") Then   'No changes are required
        imUpdateFlag = True    'Initialize VLF update flag
    End If
    smDateFilter = Trim$(Links!edcStartDate.Text)   'Store Effective Date
    lmDateFilter = gDateValue(smDateFilter)
    gPackDate smDateFilter, imDate0, imDate1
    smEndDate = Trim$(Links!edcEndDate.Text)   'End Date
    If smEndDate = "TFN" Then
        smEndDate = ""
    End If
    If Links!rbcDay(0).Value Then     'M-F
        imDateCode = 0
        smScreenCaption = "Link Definitions : Monday-Friday"
    ElseIf Links!rbcDay(1).Value Then 'Sa
        imDateCode = 6
        smScreenCaption = "Link Definitions : Saturday"
    Else                              'Su
        imDateCode = 7
        smScreenCaption = "Link Definitions : Sunday"
    End If
    If Not imReinitFlag Then  ' Not a reinitialization process
        imLcfRecLen = Len(tmLcf)  'Get and save LCF record length
        hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()        'Create LCF object handle
        On Error GoTo 0
        ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo 0
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", LinksDef
        On Error GoTo mInitErr

        imLefRecLen = Len(tmLef)  'Get and save LEF record length
        hmLef = CBtrvTable(ONEHANDLE) 'CBtrvObj()        'Create LEF object handle
        ilRet = btrOpen(hmLef, "", sgDBPath & "Lef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo 0
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Lef.Btr)", LinksDef
        On Error GoTo mInitErr

        imLvfRecLen = Len(tmLvf)  'Get and save LVF record length
        hmLvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()        'Create LVF object handle
        ilRet = btrOpen(hmLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo 0
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Lvf.Btr)", LinksDef
        On Error GoTo mInitErr
        imAnfRecLen = Len(tmAnf)  'Get and save ANF record length
        hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()        'Create ANF object handle
        ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo 0
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Anf.Btr)", LinksDef
        On Error GoTo mInitErr
        imVcfRecLen = Len(tmVcf)  'Get and save LCF record length
        hmVcf = CBtrvTable(TWOHANDLES) 'CBtrvObj()        'Create LCF object handle
        On Error GoTo 0
        ilRet = btrOpen(hmVcf, "", sgDBPath & "Vcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo 0
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Vcf.Btr)", LinksDef
        On Error GoTo mInitErr
        ReDim tmVlfPop(0 To 200) As VLF      'ReInitialize tmVLFPop image
        imUpperBound = UBound(tmVlfPop)    'Set initial upperbound for tmVlfPop
        imVlfPopRecLen = Len(tmVlfPop(0))  'Get VLF record length
        hmVlfPop = CBtrvTable(TWOHANDLES) 'CBtrvObj()              'Create VLF object handle
        ilRet = btrOpen(hmVlfPop, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo 0
        gBtrvErrorMsg ilRet, "mInit (btrOpen: Vlf.Btr)", LinksDef
        On Error GoTo mInitErr

        ' Locate Selling vehicles selected within Links
        For ilLoop = 0 To Links!lbcSelling.ListCount - 1 Step 1
            If Links!lbcSelling.Selected(ilLoop) Then 'Selected vehicle found
                If imNoSelling <> 0 Then
                    Load lacSelling(imNoSelling)
                    Load lbcSelling(imNoSelling)
                    If imNoSelling > 3 Then
                        lacSelling(imNoSelling).Visible = False
                        lbcSelling(imNoSelling).Visible = False
                    Else
                        lacSelling(imNoSelling).Visible = True
                        lbcSelling(imNoSelling).Visible = True
                    End If
                End If
                lacSelling(imNoSelling).DragIcon = IconTraf!imcIconDrag.DragIcon
                lacSelling(imNoSelling).Caption = Links!lbcSelling.List(ilLoop)
                imNoSelling = imNoSelling + 1  'Inc counter for the number of selling vehicles
            End If
        Next ilLoop
        If imNoSelling - 4 > 0 Then
            hbcSelling.Max = imNoSelling - 4
        Else
            hbcSelling.Max = 0
        End If
        ' Locate Airing Vehicles selected within Links
        For ilLoop = 0 To Links!lbcAiring.ListCount - 1 Step 1
            If Links!lbcAiring.Selected(ilLoop) Then 'Selected vehicle found
                If imNoAiring <> 0 Then
                    Load lacAiring(imNoAiring)
                    Load lbcAiring(imNoAiring)
                    If imNoAiring > 3 Then
                        lacAiring(imNoAiring).Visible = False
                        lbcAiring(imNoAiring).Visible = False
                    Else
                        lacAiring(imNoAiring).Visible = True
                        lbcAiring(imNoAiring).Visible = True
                    End If
                End If
                lacAiring(imNoAiring).Caption = Links!lbcAiring.List(ilLoop)
                imNoAiring = imNoAiring + 1  'Inc counter for the number of airing vehicles
            End If
        Next ilLoop
        If imNoAiring - 4 > 0 Then
            hbcAiring.Max = imNoAiring - 4
        Else
            hbcAiring.Max = 0
        End If
        If (imNoSelling <= 0) Or (imNoAiring <= 0) Then
            If (imNoSelling <= 0) Then
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("No Selling Vehicle Defined", vbOKOnly + vbExclamation, "Links")
            Else
                Screen.MousePointer = vbDefault
                ilRet = MsgBox("No Airing Vehicle Defined", vbOKOnly + vbExclamation, "Links")
            End If
            imTerminate = True
            Exit Sub
        End If
        mSetLBoxTabs    'Set tabs for list Boxes (with 30 Char @ delimeter spacing)
        mMoveLBox 0, 0  'Position Selling & Airing list boxes
        mPopLnkDef      'Populate List Boxes from LCF VEF & VLF (first initialization only)
    End If
    If imReinitFlag Then 'Lists already populated ; Repopulate via list arrays (smAiringLists(),smSellingLists())
        mRePopLnkDef
    End If

    imReinitFlag = True               'Reset initialization flag
    'LinksDef.Height = cmcCancel.Top + 5 * cmcCancel.Height / 3
    gCenterStdAlone LinksDef         'Position form
'    Traffic!plcHelp.Caption = ""
    mSetCommands                      'Initialize Controls
    Screen.MousePointer = vbDefault
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    'imcHelp.Picture = IconTraf!imcHelp.Picture
    mCloseFiles                       'Close LCF & LEF files ; leave VLF open
    plcScreen_Paint
    If imTerminate Then
        'mTerminate
        Exit Sub
    End If
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    'mTerminate
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveAiring                     *
'*                                                     *
'*             Created:10/16/93      By:D. LeVine      *
'*            Modified:4/24/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Move links                     *
'*                                                     *
'*******************************************************
Private Sub mMoveAiring(Index As Integer, ilDragListIndex As Integer)
    'Index(I) : Index of target list
    'ilDragListIndex : Index of target list item

    Dim slName1 As String       'Drag source string
    Dim slName2 As String       'Target list string
    Dim slName As String        'Drag source vehicle name
    Dim slTime As String        'Drag source time
    Dim ilFound As Integer      'String match found flag
    Dim ilIndex As Integer      'Target list item index
    Dim ilLoop As Integer       'List item index counter
    Dim ilRet As Integer        'Function call return
    Dim ilLastFound As Integer  'Index of list string found during search
    Dim ilIndexShift As Integer 'Increment value to offset target list item index for insert
    Dim slSrchTime As String    'Time string to search for in selling
    Dim ilLen As Integer        'String length
    Dim ilCount As Integer      'List Item index incrementor/decrementor
    Dim slTime2 As String       'Time string found in target list
    Dim ilCount2 As Integer     'Matching string found flag
    Dim slStr As String         'Parsing string
    Dim ilSell1Index As Integer          'Index1 for sell vehicle
    Dim ilSell1ListIndex As Integer      'List Index for sell vehicle
    Dim slDummy As String                'Parsing string
    On Error GoTo mMoveAiringErr
    imUpdateFlag = True    'Reset VLF update flag
    slName1 = lbcAiring(imDragIndex).List(imDragListIndex)
    slDummy = lbcAiring(Index).List(ilDragListIndex)
    If (imSwapFlag) Then 'Test if swap or move is legal...no duplicate links
        imLegal = True
        If (Left$(slDummy, 2) <> "  ") Then 'move operation
            For ilCount = ilDragListIndex + 1 To lbcAiring(Index).ListCount - 1 Step 1
                If (Left$(lbcAiring(Index).List(ilCount), 2) <> "  ") Then
                    Exit For
                Else
                    If (lbcAiring(Index).List(ilCount) = slName1) Then
                        imLegal = False
                        Exit For
                    End If
                End If
            Next ilCount
            If (imLegal = False) Then
                Beep
                Exit Sub
            End If
        End If
        If (Left$(slDummy, 2) = "  ") Then  'swap operation
            ilRet = 0
            For ilCount = ilDragListIndex - 1 To 0 Step -1
                If (Left$(lbcAiring(Index).List(ilCount), 2) <> "  ") Then
                    ilRet = ilCount
                    Exit For
                End If
            Next ilCount
            For ilCount = ilRet + 1 To lbcAiring(Index).ListCount - 1 Step 1
                If (Left$(lbcAiring(Index).List(ilCount), 2) <> "  ") Then
                    Exit For
                Else
                    If (lbcAiring(Index).List(ilCount) = slName1) Then
                        imLegal = False
                        Exit For
                    End If
                End If
            Next ilCount
            If (imLegal = False) Then
                Beep
                Exit Sub
            End If
        End If
    End If

    slStr = LTrim$(slName1)
    ilLen = Len(slStr)
    ilRet = InStr(1, slStr, " ")
    slName = right$(slStr, ilLen - ilRet)
    slName = LTrim$(slName)
    ilRet = gParseItem(slName, 1, "@", slName)
    slName = Trim$(slName)                    'Source vehicle name
    ilRet = gParseItem(slStr, 1, " ", slTime)
    slSrchTime = Trim$(slTime) 'Time to search for in selling
    imUpdateFlag = True    'Reset VLF update flag
    ilFound = False
    ilIndex = 0
    slName = Trim$(slName)
    Do
        If lacSelling(ilIndex).Caption = slName Then
            ilFound = True
            Exit Do
        End If
        ilIndex = ilIndex + 1
    Loop Until ilFound  'If not found- an error will occur
    ilFound = False
    ilLoop = imDragListIndex - 1
    Do
        If Left$(lbcAiring(imDragIndex).List(ilLoop), 2) <> "  " Then
            ilFound = True
            Exit Do
        End If
        ilLoop = ilLoop - 1
    Loop Until ilFound
    'Create string to search for in selling
    ilRet = gParseItem(LTrim$(lbcAiring(imDragIndex).List(ilLoop)), 1, " ", slStr)
    slStr = Trim$(slStr)
    slName2 = CStr("  " & Trim$(slStr) & " " & Trim$(lacAiring(imDragIndex).Caption))
    ilRet = gParseItem(lbcAiring(imDragIndex).List(ilLoop), 2, "@", slStr)
    slName2 = mFillTo100(RTrim$(slName2 & "@" & slStr))


    gFindMatch slName2, 0, lbcSelling(ilIndex)
    If gLastFound(lbcSelling(ilIndex)) < 0 Then
        Exit Sub
    End If

    'Check for matching string : if entire string is not the same then find next
    ilCount2 = 0
    ilCount = 1
    Do
        If (Left$(lbcSelling(ilIndex).List(gLastFound(lbcSelling(ilIndex)) - ilCount), 2) <> "  ") Then
            ilRet = gParseItem(LTrim$(lbcSelling(ilIndex).List(gLastFound(lbcSelling(ilIndex)) - ilCount)), 1, " ", slTime2)
            slTime2 = Trim$(slTime2)
            Exit Do
        Else
            ilCount = ilCount + 1
        End If
    Loop
    Do While ilCount2 = 0
        If slTime2 <> slSrchTime Then
            gFndNext lbcSelling(ilIndex), slName2
            If gLastFound(lbcSelling(ilIndex)) < 0 Then
                Exit Sub
            End If
            ilCount = 1
lRepeatSrchS:
            Do
                If (Left$(lbcSelling(ilIndex).List(gLastFound(lbcSelling(ilIndex)) - ilCount), 2) <> "  ") Then
                    ilRet = gParseItem(LTrim$(lbcSelling(ilIndex).List(gLastFound(lbcSelling(ilIndex)) - ilCount)), 1, " ", slTime2)
                    slTime2 = Trim$(slTime2)
                    Exit Do
                Else
                    ilCount = ilCount + 1
                End If
            Loop
            If (slTime2 <> slSrchTime) And (gLastFound(lbcSelling(ilIndex)) - ilCount > 0) Then
                ilCount = ilCount + 1
                GoTo lRepeatSrchS
            End If
        Else
            ilCount2 = 1 'Match found
        End If
    Loop

    'Delete airing
    ilLastFound = gLastFound(lbcSelling(ilIndex))
    lbcAiring(imDragIndex).RemoveItem imDragListIndex
    imUpdateFlag = True
    If (ilDragListIndex > imDragListIndex) And (imDragIndex = Index) Then
        ilDragListIndex = ilDragListIndex - 1
    End If

    ilFound = False
    ilLoop = ilLastFound - 1
    Do
        If Left$(lbcSelling(ilIndex).List(ilLoop), 2) <> "  " Then
            ilFound = True
            Exit Do
        End If
        ilLoop = ilLoop - 1
    Loop Until ilFound

    If imSwapFlag Then
        ilSell1Index = ilIndex
        ilSell1ListIndex = gLastFound(lbcSelling(ilIndex))
    Else
        lbcSelling(ilIndex).RemoveItem gLastFound(lbcSelling(ilIndex))
    End If

    imUpdateFlag = True
    imDragListIndex = ilDragListIndex
    ilDragListIndex = ilLoop
    ilFound = False
    ilLoop = imDragListIndex
    Do
        If Left$(lbcAiring(Index).List(ilLoop), 2) <> "  " Then
            ilFound = True
            Exit Do
        End If
        ilLoop = ilLoop - 1
    Loop Until ilFound

    'Create sell and air strings to add
    ilRet = gParseItem(lbcAiring(Index).List(ilLoop), 1, " ", slStr)
    slName1 = Trim$(slStr)
    ilRet = gParseItem(lbcAiring(Index).List(ilLoop), 2, "@", slStr)
    slName1 = CStr("  " & slName1 & " " & Trim$(lacAiring(Index).Caption) & "@" & Trim$(slStr))
    ilRet = gParseItem(lbcSelling(ilIndex).List(ilDragListIndex), 1, " ", slStr)
    slName2 = Trim$(slStr)
    ilRet = gParseItem(lbcSelling(ilIndex).List(ilDragListIndex), 2, "@", slStr)
    slName2 = CStr("  " & slName2 & " " & Trim$(lacSelling(ilIndex).Caption) & "@" & Trim$(slStr))
'Check for collate order before dropping
    ilIndexShift = 0
    If imSwapFlag Then   ' swap operation within same list box
        lbcSelling(ilSell1Index).List(ilSell1ListIndex) = slName1
        mSetSelected ilSell1Index, 1, ilSell1ListIndex, 1, "S", "C"
        imUpdateFlag = True
    Else
        ilLoop = ilLastFound
        Do While ilLoop < lbcSelling(ilIndex).ListCount
            If ((Left$(lbcSelling(ilIndex).List(ilLastFound + ilIndexShift), 2)) = "  ") Then
                If (lbcSelling(ilIndex).List(ilLastFound + ilIndexShift) = slName1) Then 'Duplicate
                    Beep
                    DoEvents
                    Exit Sub
                End If
                ilIndexShift = ilIndexShift + 1
                ilLoop = ilLoop + 1
            Else
                Exit Do
            End If
        Loop
        'Drop in list
        lbcSelling(ilIndex).AddItem mFillTo100(slName1), ilLastFound + ilIndexShift
        mSetSelected ilIndex, 1, ilLastFound + ilIndexShift, 1, "S", "C"
        imUpdateFlag = True
    End If
'Check for collate order before dropping
    ilIndexShift = 1
    If imSwapFlag Then ' swap operation within same list box
        lbcAiring(Index).AddItem mFillTo100(slName2), imDragListIndex + ilIndexShift
        mSetSelected Index, 1, imDragListIndex + ilIndexShift, 1, "A", "C"
        imUpdateFlag = True
    Else
        ilLoop = imDragListIndex
        Do While ilLoop < lbcAiring(Index).ListCount
            If (Left$(lbcAiring(Index).List(imDragListIndex + ilIndexShift), 2) = "  ") Then
                If (lbcAiring(Index).List(imDragListIndex + ilIndexShift) = slName2) Then 'Duplicate
                    Beep
                    DoEvents
                    Exit Sub
                End If
                ilIndexShift = ilIndexShift + 1
                ilLoop = ilLoop + 1
            Else
                Exit Do
            End If
        Loop
        ' Drop in List
        lbcAiring(Index).AddItem mFillTo100(slName2), imDragListIndex + ilIndexShift
        mSetSelected Index, 1, imDragListIndex + ilIndexShift, 1, "A", "C"
        imUpdateFlag = True
    End If
    mSetCommands
    Exit Sub
mMoveAiringErr:
    On Error GoTo 0
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveEvents                     *
'*                                                     *
'*            Created:4/4/94         By:D. LeVine      *
'*            Modified:4/24/94        By:D. Hannifan    *
'*                                                     *
'*            Comments: populate list box with pending *
'*                      events taking precedence       *
'*******************************************************
Private Sub mMoveEvents(slType As String, tlCLLC() As LLC, tlPLLC() As LLC, lbcCtrl As control)
'
'   mMoveEvents tmCLLC(), tmPLLC() ,lbcCtrl
'   Where:
'       slType   (I)- Vehicle type "S"=selling ; "A"=airing
'       tlCLLC() (I)- Current event records to be processed
'       tlPLLC() (I)- Current event records to be processed
'       lbcCtrl  (I)- list box control to populate
'

    Dim ilCIndex As Integer             'index for current event
    Dim ilPIndex As Integer             'index for pending event
    Dim clGStartTime As Currency        'general start time reference
    Dim clGEndTime As Currency          'general end time reference
    Dim clCEvtTime As Currency          'Current event start time
    Dim clCEvtEndTime As Currency       'Current event end time
    Dim clPEvtTime As Currency          'Pending event start time
    Dim clPEvtEndTime As Currency       'Pending event start time
    Dim slTime As String                'Time string
    Dim ilCFindEvt As Integer           'Current event found flag
    Dim ilPFindEvt As Integer           'Pending event found flag
    Dim ilRet As Integer                'Return value from call
    Dim ilShowCurrent As Integer        '0=Show current event; 1=show pending only;
                                        '-1=Increment pending as current and pending
                                        'intersect (pending will be shown once current is beyond pending)
    Dim slString As String              'String to add to list box
    Dim ilRealTime0 As Integer          'Time Byte 0
    Dim ilRealTime1 As Integer          'Time byte 1
    Dim ilLogDate0 As Integer           'Effective date byte 0
    Dim ilLogDate1 As Integer           'Effective date byte 1
    Dim ilEndDate0 As Integer           'Effective date byte 0
    Dim ilEndDate1 As Integer           'Effective date byte 1
    Dim slAvailName As String
    Dim slXMid As String

    gPackDate smDateFilter, ilLogDate0, ilLogDate1
    If smEndDate = "" Then
        ilEndDate0 = 0
        ilEndDate1 = 0
    Else
        gPackDate smEndDate, ilEndDate0, ilEndDate1
    End If
    ilCIndex = LBound(tlCLLC)  'Save lower bound for current
    ilPIndex = LBound(tlPLLC)  'Save lower bound for pending
        clGStartTime = 0
        clGEndTime = 86399
        Do While (tlCLLC(ilCIndex).iDay = 0) Or ((tlPLLC(ilPIndex).iDay = 0))
            ilCFindEvt = False
            Do While tlCLLC(ilCIndex).iDay = 0
                If (tlCLLC(ilCIndex).sType = "L") Then
                    clCEvtTime = gTimeToCurrency(tlCLLC(ilCIndex).sStartTime, False)
                    gAddTimeLength tlCLLC(ilCIndex).sStartTime, tlCLLC(ilCIndex).sLength, "A", "1", slTime, slXMid
                    clCEvtEndTime = gTimeToCurrency(slTime, True) - 1
                    If (clCEvtEndTime < clGStartTime) Or (clCEvtTime > clGEndTime) Then
                        If (tlCLLC(ilCIndex).iDay = -1) Or (ilCIndex >= UBound(tlCLLC)) Then
                            Exit Do
                        End If
                        ilCIndex = ilCIndex + 1
                    Else
                        ilCFindEvt = True
                        Exit Do
                    End If
                Else
                    If (tlCLLC(ilCIndex).iDay = -1) Or (ilCIndex >= UBound(tlCLLC)) Then
                        Exit Do
                    End If
                    ilCIndex = ilCIndex + 1
                End If
            Loop
            ilPFindEvt = False
            Do While tlPLLC(ilPIndex).iDay = 0
                If (tlPLLC(ilPIndex).sType = "L") Then
                    clPEvtTime = gTimeToCurrency(tlPLLC(ilPIndex).sStartTime, False)
                    gAddTimeLength tlPLLC(ilPIndex).sStartTime, tlPLLC(ilPIndex).sLength, "A", "1", slTime, slXMid
                    clPEvtEndTime = gTimeToCurrency(slTime, True) - 1
                    If (clPEvtEndTime < clGStartTime) Or (clPEvtTime > clGEndTime) Then
                        If (tlPLLC(ilPIndex).iDay = -1) Or (ilPIndex >= UBound(tlPLLC)) Then
                            Exit Do
                        End If
                        ilPIndex = ilPIndex + 1
                    Else
                        ilPFindEvt = True
                        Exit Do
                    End If
                Else
                    If (tlPLLC(ilPIndex).iDay = -1) Or (ilPIndex >= UBound(tlPLLC)) Then
                        Exit Do
                    End If
                    ilPIndex = ilPIndex + 1
                End If
            Loop
            If ilCFindEvt And ilPFindEvt Then
                If clCEvtEndTime < clPEvtTime Then
                    ilCIndex = ilCIndex + 1
                    ilShowCurrent = 0
                ElseIf clPEvtEndTime < clCEvtTime Then
                    ilPIndex = ilPIndex + 1
                    ilShowCurrent = 1
                Else
                    ilCIndex = ilCIndex + 1
                    ilShowCurrent = -1
                End If
            ElseIf ilCFindEvt And Not ilPFindEvt Then
                ilCIndex = ilCIndex + 1
                ilShowCurrent = 0
            ElseIf Not ilCFindEvt And ilPFindEvt Then
                ilPIndex = ilPIndex + 1
                ilShowCurrent = 1
            End If
            If ilShowCurrent = 0 Then
                Do While tlCLLC(ilCIndex).iDay = 0
                    If (tlCLLC(ilCIndex).sType <> "L") Then
                        slString = tlCLLC(ilCIndex).sStartTime
                        gPackTime slString, ilRealTime0, ilRealTime1
                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                            slString = slString & " " & Trim$(str$(tlCLLC(ilCIndex).iUnits)) & "/" & tlCLLC(ilCIndex).sLength
                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                            slString = slString & " " & Trim$(str$(tlCLLC(ilCIndex).iUnits)) & "/" & tlCLLC(ilCIndex).sLength
                        Else
                            slString = slString & " " & Trim$(str$(tlCLLC(ilCIndex).iUnits))
                        End If
                        slAvailName = ""
                        If tlCLLC(ilCIndex).iAvailInfo <> tmAnf.iCode Then
                            tmAnfSrchKey.iCode = tlCLLC(ilCIndex).iAvailInfo
                            ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                            If (ilRet = BTRV_ERR_NONE) Then
                                slAvailName = Trim$(tmAnf.sName)
                            Else
                                tmAnf.sName = ""
                            End If
                        Else
                            slAvailName = Trim$(tmAnf.sName)
                        End If
                        slString = slString & " " & slAvailName & "@" & CStr(imUpperBound)
                        lbcCtrl.AddItem mFillTo100(slString)
                        If (slType = "S") Then  'Selling
                            '6/6/16: Replaced GoSub
                            'GoSub lAddSellItems 'Update tmVlfPop
                            mAddSellItems ilRealTime0, ilRealTime1, ilLogDate0, ilLogDate1, ilEndDate0, ilEndDate1
                        Else
                            '6/6/16: Replaced GoSub
                            'GoSub lAddAirItems
                            mAddAirItems ilRealTime0, ilRealTime1, ilLogDate0, ilLogDate1, ilEndDate0, ilEndDate1
                        End If
                        If (tlCLLC(ilCIndex).iDay = -1) Or (ilCIndex >= UBound(tlCLLC)) Then
                            Exit Do
                        End If
                        ilCIndex = ilCIndex + 1
                    Else
                        Exit Do
                    End If
                Loop
            ElseIf ilShowCurrent = 1 Then
                Do While tlPLLC(ilPIndex).iDay = 0
                    If (tlPLLC(ilPIndex).sType <> "L") Then
                        slString = tlPLLC(ilPIndex).sStartTime
                        gPackTime slString, ilRealTime0, ilRealTime1
                        If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                            slString = slString & " " & Trim$(str$(tlPLLC(ilPIndex).iUnits)) & "/" & tlPLLC(ilPIndex).sLength
                        ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                            slString = slString & " " & Trim$(str$(tlPLLC(ilPIndex).iUnits)) & "/" & tlPLLC(ilPIndex).sLength
                        Else
                            slString = slString & " " & Trim$(str$(tlPLLC(ilPIndex).iUnits))
                        End If
                        slAvailName = ""
                        If tlPLLC(ilPIndex).iAvailInfo <> tmAnf.iCode Then
                            tmAnfSrchKey.iCode = tlPLLC(ilPIndex).iAvailInfo
                            ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                            If (ilRet = BTRV_ERR_NONE) Then
                                slAvailName = Trim$(tmAnf.sName)
                            Else
                                tmAnf.sName = ""
                            End If
                        Else
                            slAvailName = Trim$(tmAnf.sName)
                        End If
                        slString = slString & " " & slAvailName & "@" & CStr(imUpperBound)
                        lbcCtrl.AddItem mFillTo100(slString)
                        If (slType = "S") Then  'Selling
                            '6/6/16: Replaced GoSub
                            'GoSub lAddSellItems 'Update tmVlfPop
                            mAddSellItems ilRealTime0, ilRealTime1, ilLogDate0, ilLogDate1, ilEndDate0, ilEndDate1
                        Else
                            '6/6/16: Replaced GoSub
                            'GoSub lAddAirItems
                            mAddAirItems ilRealTime0, ilRealTime1, ilLogDate0, ilLogDate1, ilEndDate0, ilEndDate1
                        End If
                        If (tlPLLC(ilPIndex).iDay = -1) Or (ilPIndex >= UBound(tlPLLC)) Then
                            Exit Do
                        End If
                        ilPIndex = ilPIndex + 1
                    Else
                        Exit Do
                    End If
                Loop
            End If
        Loop
Exit Sub
'lAddSellItems: 'Add temp Selling record to tmVlfPop from LCF & LEF
'    tmVlfPop(imUpperBound).iSellCode = imVefCode
'    tmVlfPop(imUpperBound).iSellDay = imDateCode                '0=M-F etc
'    tmVlfPop(imUpperBound).iSellTime(0) = ilRealTime0
'    tmVlfPop(imUpperBound).iSellTime(1) = ilRealTime1
'    tmVlfPop(imUpperBound).iSellPosNo = 0
'    tmVlfPop(imUpperBound).iSellSeq = 0
'    tmVlfPop(imUpperBound).sStatus = "P"
'    tmVlfPop(imUpperBound).iAirCode = 0
'    tmVlfPop(imUpperBound).iAirDay = imDateCode
'    tmVlfPop(imUpperBound).iAirTime(0) = 0
'    tmVlfPop(imUpperBound).iAirTime(1) = 0
'    tmVlfPop(imUpperBound).iAirPosNo = 0
'    tmVlfPop(imUpperBound).iAirSeq = 0
'    tmVlfPop(imUpperBound).iEffDate(0) = ilLogDate0
'    tmVlfPop(imUpperBound).iEffDate(1) = ilLogDate1
'    tmVlfPop(imUpperBound).iTermDate(0) = ilEndDate0
'    tmVlfPop(imUpperBound).iTermDate(1) = ilEndDate1
'    tmVlfPop(imUpperBound).sDelete = ""
'
'    imUpperBound = imUpperBound + 1  'inc upperbound of tmVlfPop
'    If (imUpperBound > UBound(tmVlfPop)) Then
'        ReDim Preserve tmVlfPop(0 To imUpperBound) As VLF 'Redim tmVLFPop
'    End If
'    Return
'lAddAirItems: 'Add temp Airing record to temp VLF from LCF & LEF
'    tmVlfPop(imUpperBound).iSellCode = 0
'    tmVlfPop(imUpperBound).iSellDay = imDateCode     '0=M-F , 6=Sa, 7=Su
'    tmVlfPop(imUpperBound).iSellTime(0) = 0
'    tmVlfPop(imUpperBound).iSellTime(1) = 0
'    tmVlfPop(imUpperBound).iSellPosNo = 0
'    tmVlfPop(imUpperBound).iSellSeq = 0
'    tmVlfPop(imUpperBound).sStatus = "P"
'    tmVlfPop(imUpperBound).iAirCode = imVefCode
'    tmVlfPop(imUpperBound).iAirDay = imDateCode
'    tmVlfPop(imUpperBound).iAirTime(0) = ilRealTime0
'    tmVlfPop(imUpperBound).iAirTime(1) = ilRealTime1
'    tmVlfPop(imUpperBound).iAirPosNo = 0
'    tmVlfPop(imUpperBound).iAirSeq = 0
'    tmVlfPop(imUpperBound).iEffDate(0) = ilLogDate0
'    tmVlfPop(imUpperBound).iEffDate(1) = ilLogDate1
'    tmVlfPop(imUpperBound).iTermDate(0) = ilEndDate0
'    tmVlfPop(imUpperBound).iTermDate(1) = ilEndDate1
'    tmVlfPop(imUpperBound).sDelete = ""
'
'    imUpperBound = imUpperBound + 1  'inc upperbound of tmVlfPop
'    If (imUpperBound > UBound(tmVlfPop)) Then
'        ReDim Preserve tmVlfPop(0 To imUpperBound) As VLF 'Redim tmVLFPop
'    End If
'    Return
End Sub
'************************************************************
'          Procedure Name : nMoveLBox
'
'    Created : 4/24/94      By : D. Hannifan
'    Modified :             By :
'
'    Comments: Position list boxes on LinksDef form
'              (& H-Scroll bars if required)
'
'************************************************************
Private Sub mMoveLBox(ilHSB As Integer, ilHSBPos As Integer)
'ilHSB = H-Scroll bar calling (0=none, 1=Selling, 2=Airing)
'ilHSBPos = H-Scroll bar button position
    Dim ilSellSpace As Integer   ' Blank space increment for Selling Vehicles
    Dim ilAirSpace As Integer    ' Blank space increment for Airing Vehicles
    Dim ilSellTop As Integer     ' Top of selling list box
    Dim ilAirTop As Integer      ' Top of Airing list box
    Dim imBoxWidth As Integer    ' Width of list boxes
    Dim ilSLabelTop As Integer   ' Sell vehicle list box label Top
    Dim ilALabelTop As Integer   ' Airing vehicle list box label Top
    Dim ilLBSellShift As Integer ' Starting listbox number (listbox(x+ilLBSellShift))
    Dim ilLBAirShift As Integer  ' Starting listbox number (listbox(x+ilLBAirShift))
    Dim ilLoop As Integer        ' List box index counter
On Error GoTo mMoveLBoxErr
'Make all listboxes invisible
    For ilLoop = 0 To imNoSelling - 1 Step 1
        If ilHSB = 1 Or ilHSB = 0 Then
            lbcSelling(ilLoop).Visible = False
            lacSelling(ilLoop).Visible = False
        End If
    Next ilLoop
    For ilLoop = 0 To imNoAiring - 1 Step 1
        If ilHSB = 2 Or ilHSB = 0 Then
            lbcAiring(ilLoop).Visible = False
            lacAiring(ilLoop).Visible = False
        End If
    Next ilLoop
' Calculate Spacing For Left Positions of List Boxes and set Top Values
' Panel = 8040 width and Box Width = 1940

        imBoxWidth = lbcSelling(0).Width    '1940
        ilSellSpace = mSpaceSize(plcNetworks.Width, "S", imBoxWidth)    '8040, "S", imBoxWidth)
        ilSellTop = lbcSelling(0).Top   '285
        ilAirSpace = mSpaceSize(plcNetworks.Width, "A", imBoxWidth) '8040, "A", imBoxWidth)
        ilAirTop = lbcAiring(0).Top '2685
        ilSLabelTop = lacSelling(0).Top '60
        ilALabelTop = lacAiring(0).Top  '2445
' Check for horizontal scroll bar calls
    ilLBSellShift = 0        ' init default
    ilLBAirShift = 0
    If ilHSB > 0 Then     'Call from h-scroll bar
       'If (ilHSB = 1) And (imSSmallChange > 0) Then 'hbcSelling.LargeChange > 0 Then
       If (ilHSB = 1) Then  'hbcSelling.LargeChange > 0 Then
            ilLBSellShift = CInt(hbcSelling.Value) 'CInt(hbcSelling.Value / (imSSmallChange - 1)) '(hbcSelling.SmallChange - 1))
       End If
       'If (ilHSB = 2) And (imASmallChange) Then 'hbcAiring.LargeChange > 0 Then
       If (ilHSB = 2) Then  'hbcAiring.LargeChange > 0 Then
            ilLBAirShift = CInt(hbcAiring.Value) 'CInt(hbcAiring.Value / (imASmallChange - 1)) '(hbcAiring.SmallChange - 1))
       End If
    End If
' Position List Boxes, List Box Labels & make visible
' Selling Vehicles
If ilHSB = 2 Then
    GoTo doaironly
End If
    If imNoSelling = 1 Then
        lbcSelling(0 + ilLBSellShift).Move ilSellSpace, ilSellTop
        lacSelling(0 + ilLBSellShift).Move lbcSelling(0 + ilLBSellShift).Left, ilSLabelTop
        lbcSelling(0 + ilLBSellShift).Visible = True
        lacSelling(0 + ilLBSellShift).Visible = True
    ElseIf imNoSelling = 2 Then
        lbcSelling(0 + ilLBSellShift).Move ilSellSpace, ilSellTop
        lacSelling(0 + ilLBSellShift).Move lbcSelling(0 + ilLBSellShift).Left, ilSLabelTop
        lbcSelling(0 + ilLBSellShift).Visible = True
        lacSelling(0 + ilLBSellShift).Visible = True
        lbcSelling(1 + ilLBSellShift).Move ilSellSpace + lbcSelling(0 + ilLBSellShift).Left + imBoxWidth, ilSellTop
        lacSelling(1 + ilLBSellShift).Move lbcSelling(1 + ilLBSellShift).Left, ilSLabelTop
        lbcSelling(1 + ilLBSellShift).Visible = True
        lacSelling(1 + ilLBSellShift).Visible = True
    ElseIf imNoSelling = 3 Then
        lbcSelling(0 + ilLBSellShift).Move ilSellSpace, ilSellTop
        lacSelling(0 + ilLBSellShift).Move ilSellSpace, ilSLabelTop
        lbcSelling(0 + ilLBSellShift).Visible = True
        lacSelling(0 + ilLBSellShift).Visible = True
        lbcSelling(1 + ilLBSellShift).Move ilSellSpace + lbcSelling(0 + ilLBSellShift).Left + imBoxWidth, ilSellTop
        lacSelling(1 + ilLBSellShift).Move lbcSelling(1 + ilLBSellShift).Left, ilSLabelTop
        lbcSelling(1 + ilLBSellShift).Visible = True
        lacSelling(1 + ilLBSellShift).Visible = True
        lbcSelling(2 + ilLBSellShift).Move ilSellSpace + lbcSelling(1 + ilLBSellShift).Left + imBoxWidth, ilSellTop
        lacSelling(2 + ilLBSellShift).Move lbcSelling(2 + ilLBSellShift).Left, ilSLabelTop
        lbcSelling(2 + ilLBSellShift).Visible = True
        lacSelling(2 + ilLBSellShift).Visible = True
   Else
        lbcSelling(0 + ilLBSellShift).Move ilSellSpace + CInt(fgBevelX), ilSellTop
        lacSelling(0 + ilLBSellShift).Move ilSellSpace + CInt(fgBevelX), ilSLabelTop
        lbcSelling(0 + ilLBSellShift).Visible = True
        lacSelling(0 + ilLBSellShift).Visible = True
        lbcSelling(1 + ilLBSellShift).Move ilSellSpace + lbcSelling(0 + ilLBSellShift).Left + imBoxWidth, ilSellTop
        lacSelling(1 + ilLBSellShift).Move lbcSelling(1 + ilLBSellShift).Left, ilSLabelTop
        lbcSelling(1 + ilLBSellShift).Visible = True
        lacSelling(1 + ilLBSellShift).Visible = True
        lbcSelling(2 + ilLBSellShift).Move ilSellSpace + lbcSelling(1 + ilLBSellShift).Left + imBoxWidth, ilSellTop
        lacSelling(2 + ilLBSellShift).Move lbcSelling(2 + ilLBSellShift).Left, ilSLabelTop
        lbcSelling(2 + ilLBSellShift).Visible = True
        lacSelling(2 + ilLBSellShift).Visible = True
        lbcSelling(3 + ilLBSellShift).Move ilSellSpace + lbcSelling(2 + ilLBSellShift).Left + imBoxWidth, ilSellTop
        lacSelling(3 + ilLBSellShift).Move lbcSelling(3 + ilLBSellShift).Left, ilSLabelTop
        lbcSelling(3 + ilLBSellShift).Visible = True
        lacSelling(3 + ilLBSellShift).Visible = True
    End If
' If more than 4 selling vehicles are selected show h-scroll bar
    If imNoSelling > 4# Then
        imSSmallChange = 3 '32767 / (imNoSelling - 4)
        'hbcSelling.LargeChange = 3 * (imSSmallChange)   '32767 / (imNoSelling - 4)
        'hbcSelling.SmallChange = 3 * imSSmallChange '32767 / (imNoSelling - 4)
        hbcSelling.Move lbcSelling(0 + ilLBSellShift).Left, hbcSelling.Top, (4 * imBoxWidth) + (3 * ilSellSpace)
        hbcSelling.Visible = True
    End If
doaironly:
If ilHSB = 1 Then
    GoTo skipairing
End If
' Airing Vehicles
    If imNoAiring = 1 Then
        lbcAiring(0 + ilLBAirShift).Move ilAirSpace, ilAirTop
        lacAiring(0 + ilLBAirShift).Move lbcAiring(0 + ilLBAirShift).Left, ilALabelTop
        lbcAiring(0 + ilLBAirShift).Visible = True
        lacAiring(0 + ilLBAirShift).Visible = True
    ElseIf imNoAiring = 2 Then
        lbcAiring(0 + ilLBAirShift).Move ilAirSpace, ilAirTop
        lacAiring(0 + ilLBAirShift).Move lbcAiring(0 + ilLBAirShift).Left, ilALabelTop
        lbcAiring(0 + ilLBAirShift).Visible = True
        lacAiring(0 + ilLBAirShift).Visible = True
        lbcAiring(1 + ilLBAirShift).Move ilAirSpace + lbcAiring(0 + ilLBAirShift).Left + imBoxWidth, ilAirTop
        lacAiring(1 + ilLBAirShift).Move lbcAiring(1 + ilLBAirShift).Left, ilALabelTop
        lbcAiring(1 + ilLBAirShift).Visible = True
        lacAiring(1 + ilLBAirShift).Visible = True
    ElseIf imNoAiring = 3 Then
        lbcAiring(0 + ilLBAirShift).Move ilAirSpace, ilAirTop
        lacAiring(0 + ilLBAirShift).Move ilAirSpace, ilALabelTop
        lbcAiring(0 + ilLBAirShift).Visible = True
        lacAiring(0 + ilLBAirShift).Visible = True
        lbcAiring(1 + ilLBAirShift).Move ilAirSpace + lbcAiring(0 + ilLBAirShift).Left + imBoxWidth, ilAirTop
        lacAiring(1 + ilLBAirShift).Move lbcAiring(1 + ilLBAirShift).Left, ilALabelTop
        lbcAiring(1 + ilLBAirShift).Visible = True
        lacAiring(1 + ilLBAirShift).Visible = True
        lbcAiring(2 + ilLBAirShift).Move ilAirSpace + lbcAiring(1 + ilLBAirShift).Left + imBoxWidth, ilAirTop
        lacAiring(2 + ilLBAirShift).Move lbcAiring(2 + ilLBAirShift).Left, ilALabelTop
        lbcAiring(2 + ilLBAirShift).Visible = True
        lacAiring(2 + ilLBAirShift).Visible = True
   Else
        lbcAiring(0 + ilLBAirShift).Move ilAirSpace + CInt(fgBevelX), ilAirTop
        lacAiring(0 + ilLBAirShift).Move ilAirSpace + CInt(fgBevelX), ilALabelTop
        lbcAiring(0 + ilLBAirShift).Visible = True
        lacAiring(0 + ilLBAirShift).Visible = True
        lbcAiring(1 + ilLBAirShift).Move ilAirSpace + lbcAiring(0 + ilLBAirShift).Left + imBoxWidth, ilAirTop
        lacAiring(1 + ilLBAirShift).Move lbcAiring(1 + ilLBAirShift).Left, ilALabelTop
        lbcAiring(1 + ilLBAirShift).Visible = True
        lacAiring(1 + ilLBAirShift).Visible = True
        lbcAiring(2 + ilLBAirShift).Move ilAirSpace + lbcAiring(1 + ilLBAirShift).Left + imBoxWidth, ilAirTop
        lacAiring(2 + ilLBAirShift).Move lbcAiring(2 + ilLBAirShift).Left, ilALabelTop
        lbcAiring(2 + ilLBAirShift).Visible = True
        lacAiring(2 + ilLBAirShift).Visible = True
        lbcAiring(3 + ilLBAirShift).Move ilAirSpace + lbcAiring(2 + ilLBAirShift).Left + imBoxWidth, ilAirTop
        lacAiring(3 + ilLBAirShift).Move lbcAiring(3 + ilLBAirShift).Left, ilALabelTop
        lbcAiring(3 + ilLBAirShift).Visible = True
        lacAiring(3 + ilLBAirShift).Visible = True
    End If
    If imNoAiring > 4# Then
        imASmallChange = 3 '32767 / (imNoAiring - 4)
        'hbcAiring.LargeChange = 3 * imASmallChange  '32767 / (imNoAiring - 4)
        'hbcAiring.SmallChange = 3 * imASmallChange  '32767 / (imNoAiring - 4)
        hbcAiring.Move lbcAiring(0 + ilLBAirShift).Left, hbcAiring.Top, (4 * imBoxWidth) + (3 * ilAirSpace)
        hbcAiring.Visible = True
    End If

    Exit Sub
skipairing:
' Reset scroll calls
    ilHSB = 0
    ilHSBPos = 0
    Exit Sub
mMoveLBoxErr:
    On Error GoTo 0
    ilHSB = 0
    ilHSBPos = 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveSelling                    *
'*                                                     *
'*             Created:10/16/93      By:D. LeVine      *
'*            Modified:4/24/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: Move links                     *
'*                                                     *
'*******************************************************
Private Sub mMoveSelling(Index As Integer, ilDragListIndex As Integer)
    'Index(I) : Index of target list
    'ilDragListIndex : Index of target list item

    Dim slName1 As String       'Drag source string
    Dim slName2 As String       'Target list string
    Dim slName As String        'Drag source vehicle name
    Dim slTime As String        'Drag source time
    Dim ilFound As Integer      'String match found flag
    Dim ilIndex As Integer      'Target list item index
    Dim ilLoop As Integer       'List item index counter
    Dim ilRet As Integer        'Function call return
    Dim ilLastFound As Integer  'Index of list string found during search
    Dim ilIndexShift As Integer 'Increment value to offset target list item index for insert
    Dim slSrchTime As String    'Time string to search for in airing
    Dim ilLen As Integer        'String length
    Dim ilCount As Integer      'List Item index incrementor/decrementor
    Dim slTime2 As String       'Time string found in target list
    Dim ilCount2 As Integer     'Matching string found flag
    Dim slStr As String         'Parse string
    Dim ilAir1Index As Integer          'Air vehicle index
    Dim ilAir1ListIndex As Integer      'Air vehicle list index
    Dim slDummy As String               'parsing string

    On Error GoTo mMoveSellingErr
    slName1 = lbcSelling(imDragIndex).List(imDragListIndex)
    slDummy = lbcSelling(Index).List(ilDragListIndex)
    If (imSwapFlag) Then 'Test if swap or move is legal...no duplicate links
        imLegal = True
        If (Left$(slDummy, 2) <> "  ") Then 'move operation
            For ilCount = ilDragListIndex + 1 To lbcSelling(Index).ListCount - 1 Step 1
                If (Left$(lbcSelling(Index).List(ilCount), 2) <> "  ") Then
                    Exit For
                Else
                    If (lbcSelling(Index).List(ilCount) = slName1) Then
                        imLegal = False
                        Exit For
                    End If
                End If
            Next ilCount
            If (imLegal = False) Then
                Beep
                Exit Sub
            End If
        End If
        If (Left$(slDummy, 2) = "  ") Then  'swap operation
            ilRet = 0
            For ilCount = ilDragListIndex - 1 To 0 Step -1
                If (Left$(lbcSelling(Index).List(ilCount), 2) <> "  ") Then
                    ilRet = ilCount
                    Exit For
                End If
            Next ilCount
            For ilCount = ilRet + 1 To lbcSelling(Index).ListCount - 1 Step 1
                If (Left$(lbcSelling(Index).List(ilCount), 2) <> "  ") Then
                    Exit For
                Else
                    If (lbcSelling(Index).List(ilCount) = slName1) Then
                        imLegal = False
                        Exit For
                    End If
                End If
            Next ilCount
            If (imLegal = False) Then
                Beep
                Exit Sub
            End If
        End If
    End If
    slStr = LTrim$(slName1)
    ilLen = Len(slStr)
    ilRet = InStr(1, slStr, " ")
    slName = right$(slStr, ilLen - ilRet)
    slName = LTrim$(slName)
    ilRet = gParseItem(slName, 1, "@", slName)
    slName = Trim$(slName)
    ilRet = gParseItem(slStr, 1, " ", slTime)
    slSrchTime = Trim$(slTime)  'Time to search for in selling

    imUpdateFlag = True    'Reset VLF update flag
    ilFound = False
    ilIndex = 0
    slName = Trim$(slName)
    Do
        If lacAiring(ilIndex).Caption = slName Then
            ilFound = True
            Exit Do
        End If
        ilIndex = ilIndex + 1
    Loop Until ilFound  'If not found- an error will occur
    ilFound = False
    ilLoop = imDragListIndex - 1
    Do
        If Left$(lbcSelling(imDragIndex).List(ilLoop), 2) <> "  " Then
            ilFound = True
            Exit Do
        End If
        ilLoop = ilLoop - 1
    Loop Until ilFound
    ilRet = gParseItem(LTrim$(lbcSelling(imDragIndex).List(ilLoop)), 1, " ", slStr)
    slStr = Trim$(slStr)
    slName2 = CStr("  " & Trim$(slStr) & " " & Trim$(lacSelling(imDragIndex).Caption))
    ilRet = gParseItem(lbcSelling(imDragIndex).List(ilLoop), 2, "@", slStr)
    slName2 = mFillTo100(slName2 & "@" & slStr)

    gFindMatch slName2, 0, lbcAiring(ilIndex)
    If gLastFound(lbcAiring(ilIndex)) < 0 Then
        Exit Sub
    End If
    'Check for matching string : if entire string is not the same then find next
    ilCount2 = 0
    ilCount = 1
    Do
        If (Left$(lbcAiring(ilIndex).List(gLastFound(lbcAiring(ilIndex)) - ilCount), 2) <> "  ") Then
            ilRet = gParseItem(LTrim$(lbcAiring(ilIndex).List(gLastFound(lbcAiring(ilIndex)) - ilCount)), 1, " ", slTime2)
            slTime2 = Trim$(slTime2)

            Exit Do
        Else
            ilCount = ilCount + 1
        End If
    Loop
    Do While ilCount2 = 0
        If slTime2 <> slSrchTime Then
            gFndNext lbcAiring(ilIndex), slName2
            If gLastFound(lbcAiring(ilIndex)) < 0 Then
                Exit Sub
            End If
            ilCount = 1
lRepeatSrchA:
            Do
                If (Left$(lbcAiring(ilIndex).List(gLastFound(lbcAiring(ilIndex)) - ilCount), 2) <> "  ") Then
                    ilRet = gParseItem(LTrim$(lbcAiring(ilIndex).List(gLastFound(lbcAiring(ilIndex)) - ilCount)), 1, " ", slTime2)
                    slTime2 = Trim$(slTime2)
                    Exit Do
                Else
                    ilCount = ilCount + 1
                End If
            Loop
            If (slTime2 <> slSrchTime) And (gLastFound(lbcAiring(ilIndex)) - ilCount > 0) Then
                ilCount = ilCount + 1
                GoTo lRepeatSrchA
            End If
        Else
            ilCount2 = 1  'Match found
        End If
    Loop
    'Delete selling
    ilLastFound = gLastFound(lbcAiring(ilIndex))
    lbcSelling(imDragIndex).RemoveItem imDragListIndex
    imUpdateFlag = True
    If (ilDragListIndex > imDragListIndex) And (imDragIndex = Index) Then
        ilDragListIndex = ilDragListIndex - 1
    End If
    ilFound = False
    ilLoop = ilLastFound - 1
    Do
        If Left$(lbcAiring(ilIndex).List(ilLoop), 2) <> "  " Then
            ilFound = True
            Exit Do
        End If
        ilLoop = ilLoop - 1
    Loop Until ilFound

    If imSwapFlag Then
        ilAir1Index = ilIndex
        ilAir1ListIndex = gLastFound(lbcAiring(ilIndex))
    Else
        lbcAiring(ilIndex).RemoveItem gLastFound(lbcAiring(ilIndex))
    End If

    imUpdateFlag = True
    imDragListIndex = ilDragListIndex
    ilDragListIndex = ilLoop
    ilFound = False
    ilLoop = imDragListIndex
    Do
        If Left$(lbcSelling(Index).List(ilLoop), 2) <> "  " Then
            ilFound = True
            Exit Do
        End If
        ilLoop = ilLoop - 1
    Loop Until ilFound
     'Create strings for list box add
     ilRet = gParseItem(lbcSelling(Index).List(ilLoop), 1, " ", slStr)
     slName1 = Trim$(slStr)
     ilRet = gParseItem(lbcSelling(Index).List(ilLoop), 2, "@", slStr)
     slName1 = CStr("  " & slName1 & " " & Trim$(lacSelling(Index).Caption) & "@" & Trim$(slStr))
     ilRet = gParseItem(lbcAiring(ilIndex).List(ilDragListIndex), 1, " ", slStr)
     slName2 = Trim$(slStr)
     ilRet = gParseItem(lbcAiring(ilIndex).List(ilDragListIndex), 2, "@", slStr)
     slName2 = CStr("  " & slName2 & " " & Trim$(lacAiring(ilIndex).Caption) & "@" & Trim$(slStr))
' Check for collating order before adding to the airing list box
    ilIndexShift = 0
    If imSwapFlag Then  'Swap within same list box
        lbcAiring(ilAir1Index).List(ilAir1ListIndex) = slName1
        mSetSelected ilAir1Index, 1, ilAir1ListIndex, 1, "A", "C"
        imUpdateFlag = True
    Else
        ilLoop = ilLastFound
        Do While ilLoop < lbcAiring(ilIndex).ListCount
            If ((Left$(lbcAiring(ilIndex).List(ilLastFound + ilIndexShift), 2)) = "  ") Then
                If (lbcAiring(ilIndex).List(ilLastFound + ilIndexShift) = slName1) Then 'Duplicate
                    Beep
                    DoEvents
                    Exit Sub
                End If
                ilIndexShift = ilIndexShift + 1
                ilLoop = ilLoop + 1
            Else
                Exit Do
            End If
        Loop
        ' Add to the airing list box
        lbcAiring(ilIndex).AddItem mFillTo100(slName1), ilLastFound + ilIndexShift
        mSetSelected ilIndex, 1, ilLastFound + ilIndexShift, 1, "A", "C"
        imUpdateFlag = True
    End If
' Check for collating order before adding to the selling list box
    ilIndexShift = 1
    If imSwapFlag Then 'Swap within same list box
        lbcSelling(Index).AddItem mFillTo100(slName2), imDragListIndex + ilIndexShift
        mSetSelected Index, 1, imDragListIndex + ilIndexShift, 1, "S", "C"
        imUpdateFlag = True
    Else
        ilLoop = imDragListIndex '+ 1
        Do While ilLoop < lbcSelling(Index).ListCount
            If (Left$(lbcSelling(Index).List(imDragListIndex + ilIndexShift), 2) = "  ") Then
                If (lbcSelling(Index).List(imDragListIndex + ilIndexShift) = slName2) Then 'Duplicate
                    Beep
                    DoEvents
                    Exit Sub
                End If
                ilIndexShift = ilIndexShift + 1
                ilLoop = ilLoop + 1
            Else
                Exit Do
            End If
        Loop
        'Add to the selling list box
        lbcSelling(Index).AddItem mFillTo100(slName2), imDragListIndex + ilIndexShift
        mSetSelected Index, 1, imDragListIndex + ilIndexShift, 1, "S", "C"
        imUpdateFlag = True
    End If
    mSetCommands
    Exit Sub
mMoveSellingErr:
    On Error GoTo 0
    Exit Sub
End Sub
'**********************************************************************
'
'           Procedure Name : mPopLnkDef
'
'       Created : 4/4/94        By: D. Hannifan
'       Modified :              By:
'
' Comments : Populate list boxes in LinksDef using LEF,LCF & VLF Files
'            Store VLF Values in tmVlfPop(1). Record index is added
'            as a tabset in the list box item with a @ delimeter.
'            (@1 =  tmVlfPop Record 1)
'
'***********************************************************************
Private Sub mPopLnkDef()
    Dim ilLoop As Integer        'List item counter
    Dim ilRet  As Integer        'Return value from call
    Dim slNameCode As String     'Vehicle name
    Dim slCode As String         'Vehicle code
    Dim ilIndex As Integer       'List Box Index
    On Error GoTo mPopLnkDefErr

    imUpperBound = 1

    'Process Selected selling
    ilIndex = 0
    ReDim imSellPending(0 To imNoSelling - 1) As Integer   'Neg VefCode=VLF Pending exist; Pos VefCode=VLF no pending (use current)
    For ilLoop = 0 To Links!lbcSelling.ListCount - 1 Step 1
        If Links!lbcSelling.Selected(ilLoop) Then
            slNameCode = tgUserVehicle(ilLoop).sKey 'Links!lbcVehName.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imVefCode = Val(slCode)
            If bmFirstCallToVpfFind Then
                imVpfIndex = gVpfFind(LinksDef, imVefCode)
                bmFirstCallToVpfFind = False
            Else
                imVpfIndex = gVpfFindIndex(imVefCode)
            End If
            mReadLcfLefLnf "S", "C", tmCLLC()  'Get current events
            mReadLcfLefLnf "S", "P", tmPLLC()  'Get pending events
            mMoveEvents "S", tmCLLC(), tmPLLC(), lbcSelling(ilIndex) 'Pop list box
            'ilRet = gCodeChrRefExist(LinksDef, "Vlf.Btr", imVefCode, "vlfSellCode", "P", "vlfStatus")
            ilRet = mAnyPendForSelling(imVefCode)
            If ilRet Then
                imSellPending(ilIndex) = -imVefCode
            Else
                imSellPending(ilIndex) = imVefCode
            End If
            ilIndex = ilIndex + 1
        End If
    Next ilLoop
    'Process Selected airing
    ReDim imAirPending(0 To imNoAiring - 1) As Integer     'True=Pending; False=Current
    ilIndex = 0
    For ilLoop = 0 To Links!lbcAiring.ListCount - 1 Step 1
        If Links!lbcAiring.Selected(ilLoop) Then
            slNameCode = tgVehicle(ilLoop).sKey  'Links!lbcVehMName.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imVefCode = Val(slCode)
            If bmFirstCallToVpfFind Then
                imVpfIndex = gVpfFind(LinksDef, imVefCode)
                bmFirstCallToVpfFind = False
            Else
                imVpfIndex = gVpfFindIndex(imVefCode)
            End If
            mReadLcfLefLnf "A", "C", tmCLLC()  'Get current events
            mReadLcfLefLnf "A", "P", tmPLLC()  'Get pending events
            mMoveEvents "A", tmCLLC(), tmPLLC(), lbcAiring(ilIndex)  'Pop list box
            'ilRet = gCodeChrRefExist(LinksDef, "Vlf.Btr", imVefCode, "vlfAirCode", "P", "vlfStatus")
            ilRet = mAnyPendForAiring(imVefCode)
            If ilRet Then
                imAirPending(ilIndex) = -imVefCode
            Else
                imAirPending(ilIndex) = imVefCode
            End If
            ilIndex = ilIndex + 1
        End If
    Next ilLoop
    imUpperBound = imUpperBound - 1 'Reset imUpperBound before exit
    If (imUpperBound < UBound(tmVlfPop)) Then
        ReDim Preserve tmVlfPop(0 To imUpperBound) As VLF
    End If
    mScanVlf  'Intersperce VLF records into list boxes
    Exit Sub
mPopLnkDefErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadLcfLcfLnf                  *
'*                                                     *
'*             Created:10/17/93      By:D. LeVine      *
'*            Modified: 4/24/94      By:D. Hannifan    *
'*                                                     *
'*            Comments: Read in all events for a date  *
'*                                                     *
'*******************************************************
Private Sub mReadLcfLefLnf(slSAType As String, sLCP As String, tlLLC() As LLC)
'
'   slSAType(I)- Vehicle type "S"=selling ; "A"=airing
'   slCP  (I) : status key "C" = current "P"=pending
'   tlLLC (I) : current or pending LCC events array
'
'   smDateFilter contains the effective date
'


    Dim ilUpper As Integer          'Upperbound of tlLLC array
    Dim ilSeqNo As Integer          'Sequence number
    Dim ilType As Integer            'Vehicle type
    Dim ilRet As Integer            'Return from call
    Dim ilIndex As Integer          'List index
    Dim slStartTime As String       'Effective start time
    Dim slStr As String             'Parse string
    Dim ilDate0 As Integer          'Byte 0 start date
    Dim ilDate1 As Integer          'Byte 1 start date
    ReDim tlLLC(0 To 0) As LLC      'LLC image
    Dim ilFound As Integer          'True=valid avail found
    Dim ilDay As Integer
    Dim slDate As String
    Dim ilDel As Integer
    Dim ilDeleted As Integer
    Dim ilTestDel As Integer
    Dim slXMid As String
    On Error GoTo mReadLcfLefErr

    ilUpper = UBound(tlLLC)
    tlLLC(ilUpper).iDay = -1
    ilType = 0
    ilSeqNo = 1
    gPackDate smDateFilter, ilDate0, ilDate1
    ilDay = gWeekDayStr(smDateFilter)
    If slSAType = "A" Then    'Determine effective date
        ilFound = False
        tmLcfSrchKey.iType = ilType
        tmLcfSrchKey.sStatus = sLCP
        tmLcfSrchKey.iVefCode = imVefCode
        tmLcfSrchKey.iLogDate(0) = ilDate0
        tmLcfSrchKey.iLogDate(1) = ilDate1
        tmLcfSrchKey.iSeqNo = ilSeqNo
        ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
        Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = imVefCode) And (tmLcf.iType = ilType) And (tmLcf.sStatus = sLCP)
            gUnpackDate tmLcf.iLogDate(0), tmLcf.iLogDate(1), slDate
            If ilDay = gWeekDayStr(slDate) Then
                ilDate0 = tmLcf.iLogDate(0)
                ilDate1 = tmLcf.iLogDate(1)
                ilFound = True
                Exit Do
            End If
            ilRet = btrGetNext(hmLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If Not ilFound Then
            'Use TFN
            tmLcfSrchKey.iType = ilType
            tmLcfSrchKey.sStatus = sLCP
            tmLcfSrchKey.iVefCode = imVefCode
            tmLcfSrchKey.iLogDate(0) = ilDay + 1
            tmLcfSrchKey.iLogDate(1) = 0
            tmLcfSrchKey.iSeqNo = ilSeqNo
            ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
            If (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = imVefCode) And (tmLcf.iType = ilType) And (tmLcf.sStatus = sLCP) Then
                If (tmLcf.iLogDate(0) <= 7) And (tmLcf.iLogDate(1) = 0) Then
                    If ilDay + 1 = tmLcf.iLogDate(0) Then
                        ilDate0 = tmLcf.iLogDate(0)
                        ilDate1 = tmLcf.iLogDate(1)
                        ilFound = True
                    End If
                End If
            End If
        End If
        If Not ilFound Then
            Exit Sub
        End If
    Else
        ilFound = False
        tmLcfSrchKey.iType = ilType
        tmLcfSrchKey.sStatus = sLCP
        tmLcfSrchKey.iVefCode = imVefCode
        tmLcfSrchKey.iLogDate(0) = ilDate0
        tmLcfSrchKey.iLogDate(1) = ilDate1
        tmLcfSrchKey.iSeqNo = ilSeqNo
        'using greater or equal works best for pending- this allow any input date
        'then finds the next valid day after date
        ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
        Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = imVefCode) And (tmLcf.iType = ilType) And (tmLcf.sStatus = sLCP)
            gUnpackDate tmLcf.iLogDate(0), tmLcf.iLogDate(1), slDate
            If ilDay = gWeekDayStr(slDate) Then
                ilDate0 = tmLcf.iLogDate(0)
                ilDate1 = tmLcf.iLogDate(1)
                ilFound = True
                Exit Do
            End If
            ilRet = btrGetNext(hmLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If Not ilFound Then
            'Use TFN
            tmLcfSrchKey.iType = ilType
            tmLcfSrchKey.sStatus = sLCP
            tmLcfSrchKey.iVefCode = imVefCode
            tmLcfSrchKey.iLogDate(0) = ilDay + 1
            tmLcfSrchKey.iLogDate(1) = 0
            tmLcfSrchKey.iSeqNo = ilSeqNo
            ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
            If (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = imVefCode) And (tmLcf.iType = ilType) And (tmLcf.sStatus = sLCP) Then
                If (tmLcf.iLogDate(0) <= 7) And (tmLcf.iLogDate(1) = 0) Then
                    If ilDay + 1 = tmLcf.iLogDate(0) Then
                        ilDate0 = tmLcf.iLogDate(0)
                        ilDate1 = tmLcf.iLogDate(1)
                        ilFound = True
                    End If
                End If
            End If
        End If
        If Not ilFound Then
            Exit Sub
        End If
    End If
    ilTestDel = True
    tmLcfSrchKey.iType = ilType
    tmLcfSrchKey.sStatus = "D"
    tmLcfSrchKey.iVefCode = imVefCode
    tmLcfSrchKey.iLogDate(0) = ilDate0
    tmLcfSrchKey.iLogDate(1) = ilDate1
    tmLcfSrchKey.iSeqNo = 1
    ilRet = btrGetEqual(hmLcf, tmDLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
    If ilRet <> BTRV_ERR_NONE Then
        ilTestDel = False
    End If
    Do
        tmLcfSrchKey.iType = ilType
        tmLcfSrchKey.sStatus = sLCP
        tmLcfSrchKey.iVefCode = imVefCode
        tmLcfSrchKey.iLogDate(0) = ilDate0
        tmLcfSrchKey.iLogDate(1) = ilDate1
        tmLcfSrchKey.iSeqNo = ilSeqNo
        ilRet = btrGetEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        If ilRet = BTRV_ERR_NONE Then
            ilSeqNo = ilSeqNo + 1
            For ilIndex = LBound(tmLcf.lLvfCode) To UBound(tmLcf.lLvfCode) Step 1
                If tmLcf.lLvfCode(ilIndex) <> 0 Then
                    ilDeleted = False
                    If ilTestDel Then
                        For ilDel = LBound(tmDLcf.lLvfCode) To UBound(tmDLcf.lLvfCode) Step 1
                            If tmDLcf.lLvfCode(ilDel) <> 0 Then
                                If tmLcf.lLvfCode(ilIndex) = tmDLcf.lLvfCode(ilDel) Then
                                    If (tmLcf.iTime(0, ilIndex) = tmDLcf.iTime(0, ilDel)) And (tmLcf.iTime(1, ilIndex) = tmDLcf.iTime(1, ilDel)) Then
                                        ilDeleted = True
                                    End If
                                End If
                            End If
                        Next ilDel
                    End If
                    If Not ilDeleted Then
                        tlLLC(ilUpper).iDay = 0
                        tlLLC(ilUpper).sType = "L"
                        gUnpackTime tmLcf.iTime(0, ilIndex), tmLcf.iTime(1, ilIndex), "A", "1", tlLLC(ilUpper).sStartTime
                        slStartTime = tlLLC(ilUpper).sStartTime
                        'Read in Lnf to obtain name and length
                        tmLvfSrchKey.lCode = tmLcf.lLvfCode(ilIndex)
                        ilRet = btrGetEqual(hmLvf, tmLvf, imLvfRecLen, tmLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'Get current record
                        If ilRet = BTRV_ERR_NONE Then
                            gUnpackLength tmLvf.iLen(0), tmLvf.iLen(1), "3", False, tlLLC(ilUpper).sLength
                            ilUpper = ilUpper + 1
                            ReDim Preserve tlLLC(0 To ilUpper) As LLC
                        End If
                        tlLLC(ilUpper).iDay = -1
                        'Read in all the event record (Lef)
                        tmLefSrchKey.lLvfCode = tmLcf.lLvfCode(ilIndex)
                        tmLefSrchKey.iStartTime(0) = 0
                        tmLefSrchKey.iStartTime(1) = 0
                        tmLefSrchKey.iSeqNo = 0
                        ilRet = btrGetGreaterOrEqual(hmLef, tmLef, imLefRecLen, tmLefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmLef.lLvfCode = tmLcf.lLvfCode(ilIndex))
                            tlLLC(ilUpper).iDay = 0
                            gUnpackLength tmLef.iStartTime(0), tmLef.iStartTime(1), "3", False, slStr
                            gAddTimeLength slStartTime, slStr, "A", "1", tlLLC(ilUpper).sStartTime, slXMid
                            ilFound = False
                            Select Case tmLef.iEtfCode
                                Case 1  'Program
                                Case 2, 6 To 9 'Contract Avail
                                    ilFound = True
                                    tlLLC(ilUpper).sType = Trim$(str$(tmLef.iEtfCode))
                                    tlLLC(ilUpper).iUnits = tmLef.iMaxUnits
                                    If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                                        gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                    ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                                        gUnpackLength tmLef.iLen(0), tmLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                                    Else
                                        tlLLC(ilUpper).sLength = "0S"
                                    End If
                                    tlLLC(ilUpper).iAvailInfo = tmLef.ianfCode
                                Case 3
                                Case 4
                                Case 5
                                'Case 6  'Cmml Promo
                                'Case 7  'Feed avail
                                'Case 8  'PSA (Avail)
                                'Case 9  'Promo
                                Case 10  'Page eject, Line space 1, 2 or 3
                                Case 11
                                Case 12
                                Case 13
                                Case Else   'Other
                            End Select
                            If ilFound Then
                                ilUpper = ilUpper + 1
                                ReDim Preserve tlLLC(0 To ilUpper) As LLC
                            End If
                            tlLLC(ilUpper).iDay = -1
                            ilRet = btrGetNext(hmLef, tmLef, imLefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                Else
                    ilSeqNo = -1
                    Exit For
                End If
            Next ilIndex
        Else
            ilSeqNo = -1
        End If
    Loop While ilSeqNo > 0
Exit Sub
mReadLcfLefErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'**********************************************************************
'
'       Procedure Name : mReInitLinksDef
'
'       Created : 3/19/94       By: D. Hannifan
'       Modified :              By:
'
'       Comments : Procedure to reinitialize LinksDef list boxes after
'                  a Show Discrepencies event
'
'
'**********************************************************************
Private Sub mReInitLinksDef()

    imReinitFlag = True
    mSetCommands
    Call mInit
End Sub
'**********************************************************************
'
'       Procedure Name : mRePopLnkDef
'
'       Created : 4/24/94       By: D. Hannifan
'       Modified :              By:
'
'       Comments : Procedure to populate LinksDef list boxes after
'                  a Show Discrepencies event
'
'
'**********************************************************************
'
Private Sub mRePopLnkDef()

    Dim ilCount As Integer              'General index counters
    Dim ilCount2 As Integer
    Dim ilCount3 As Integer
    Dim ilCount4 As Integer
    Dim slParseString1 As String * 2    'First two characters of list string
    Dim slParseString2 As String * 2    'First two characters of list string
    Dim slString1 As String             'First list string
    Dim slString2 As String             'Second list string
    Dim slString As String              'First loop Test string
    Dim slSaveString As String          'string to save for overwrite
    Dim ilInsertCount As Integer        'Number of new links inserted
    Dim ilInsertOffset As Integer       'Number of places to move new link insert from matching record
    On Error GoTo mRePopLnkDefErr

    For ilCount = 0 To imNoSelling - 1 Step 1 'Get discrepency Selling list box records
        ilInsertCount = 0  'initialize link insert counter
        For ilCount2 = 0 To lbcSelling(ilCount).ListCount - 1 Step 1
            slParseString1 = Left$(lbcSelling(ilCount).List(ilCount2), 2)
            slParseString2 = Left$(lbcSelling(ilCount).List(ilCount2 + 1), 2)
            If (slParseString1 <> "  ") And (slParseString2 <> "  ") Then
            Else
                If ((slParseString1 <> "  ") And (slParseString2 = "  ")) Or ((slParseString1 = "  ") And (slParseString2 = "  ")) Then 'Insert a new link
                    'New link found at ilCount2 +1
                    ilInsertCount = ilInsertCount + 1
                    slString = lbcSelling(ilCount).List(ilCount2)
                    If (Left$(slString, 2) <> "  ") Then
                        slString1 = lbcSelling(ilCount).List(ilCount2)
                        ilInsertOffset = 1
                    Else
                        ilInsertOffset = ilInsertOffset + 1
                    End If
                    slString2 = lbcSelling(ilCount).List(ilCount2 + 1)
                    If (imSellCount(ilCount) + ilInsertCount >= imSellUpperB) Then ' Redimension Selling array
                        imSellUpperB = imSellCount(ilCount) + ilInsertCount
                        ReDim Preserve smSellingLists(imNoSelling - 1, imSellUpperB) As String
                    End If
                    For ilCount3 = 0 To imSellCount(ilCount) - 1 Step 1
                        'Scan smSell for valid link insert location
                        If (slString1 = smSellingLists(ilCount, ilCount3)) Then
                            'matching record found at ilCount3
                            For ilCount4 = (ilCount3 + ilInsertOffset) To (imSellCount(ilCount) - 1 + ilInsertCount) Step 1
                                slSaveString = smSellingLists(ilCount, ilCount4) 'save record in link insert target
                                smSellingLists(ilCount, ilCount4) = slString2    'add new link
                                slString2 = slSaveString 'prepare to write saved record in next location
                            Next ilCount4
                            Exit For
                        End If
                    Next ilCount3
                    imSellCount(ilCount) = imSellCount(ilCount) + ilInsertCount
                End If
            End If
        Next ilCount2
        If (imSellCount(ilCount) >= imSellUpperB) Then  ' Redimension Selling array
            imSellUpperB = imSellCount(ilCount)
            ReDim Preserve smSellingLists(imNoSelling - 1, imSellUpperB) As String
        End If
    Next ilCount
    'Airing scan
    For ilCount = 0 To imNoAiring - 1 Step 1 'Get Airing list box discrepency records
        ilInsertCount = 0
        For ilCount2 = 0 To lbcAiring(ilCount).ListCount - 1 Step 1
            slParseString1 = Left$(lbcAiring(ilCount).List(ilCount2), 2)
            slParseString2 = Left$(lbcAiring(ilCount).List(ilCount2 + 1), 2)
            If (slParseString1 <> "  ") And (slParseString2 <> "  ") Then
            Else
                If ((slParseString1 <> "  ") And (slParseString2 = "  ")) Or ((slParseString1 = "  ") And (slParseString2 = "  ")) Then 'Insert a new link
                    'New link found at ilCount2 +1
                    ilInsertCount = ilInsertCount + 1
                    slString = lbcAiring(ilCount).List(ilCount2)
                    If (Left$(slString, 2) <> "  ") Then
                        slString1 = lbcAiring(ilCount).List(ilCount2)
                        ilInsertOffset = 1
                    Else
                        ilInsertOffset = ilInsertOffset + 1
                    End If
                    slString2 = lbcAiring(ilCount).List(ilCount2 + 1)
                    If (imAirCount(ilCount) + ilInsertCount >= imAirUpperB) Then ' Redimension Airing array
                        imAirUpperB = imAirCount(ilCount) + ilInsertCount
                        ReDim Preserve smAiringLists(imNoAiring - 1, imAirUpperB) As String
                    End If
                    For ilCount3 = 0 To imAirCount(ilCount) - 1 Step 1
                        'Scan smAir for valid new link insert location
                        If (slString1 = smAiringLists(ilCount, ilCount3)) Then
                            'matching record location found at ilCount3
                            For ilCount4 = (ilCount3 + ilInsertOffset) To (imAirCount(ilCount) - 1 + ilInsertCount) Step 1
                                slSaveString = smAiringLists(ilCount, ilCount4)
                                smAiringLists(ilCount, ilCount4) = slString2
                                slString2 = slSaveString ' smAiringLists(ilCount, ilCount4 + 1)
                            Next ilCount4
                            Exit For
                        End If
                    Next ilCount3
                    imAirCount(ilCount) = imAirCount(ilCount) + ilInsertCount
                End If
            End If
        Next ilCount2
        If (imAirCount(ilCount) >= imAirUpperB) Then  ' Redimension Airing array
            imAirUpperB = imAirCount(ilCount)
            ReDim Preserve smAiringLists(imNoAiring - 1, imAirUpperB) As String
        End If
    Next ilCount
    For ilCount = 0 To imNoSelling - 1 Step 1 ' Clear selling list boxes
        lbcSelling(ilCount).ListIndex = -1
        lbcSelling(ilCount).Clear
    Next ilCount
    For ilCount = 0 To imNoAiring - 1 Step 1 ' Clear airing list boxes
        lbcAiring(ilCount).ListIndex = -1
        lbcAiring(ilCount).Clear
    Next ilCount

    For ilCount = 0 To imNoSelling - 1 Step 1  'Repopulate selling lists
        For ilCount2 = 0 To imSellCount(ilCount) - 1 Step 1
            slString = smSellingLists(ilCount, ilCount2)
            If Len(slString) > 3 Then
                lbcSelling(ilCount).AddItem mFillTo100(smSellingLists(ilCount, ilCount2))
            End If
        Next ilCount2
    Next ilCount
    For ilCount = 0 To imNoAiring - 1 Step 1   'Repopulate airing lists
        For ilCount2 = 0 To imAirCount(ilCount) - 1 Step 1
            slString = smAiringLists(ilCount, ilCount2)
            If Len(slString) > 3 Then
                lbcAiring(ilCount).AddItem mFillTo100(smAiringLists(ilCount, ilCount2))
            End If
        Next ilCount2
    Next ilCount
    imSellUpperB = 0        'Clean up arrays and dimensioning variables
    imAirUpperB = 0
    Erase smSellingLists
    Erase smAiringLists
    Erase imSellCount
    Erase imAirCount
    Exit Sub
mRePopLnkDefErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*********************************************************
'      Procedure Name : mScanVlf
'
'      Created : 4/24/94         By: D. Hannifan
'      Modified :               By:
'
'      Comments : Scan VLF and add valid links to lists
'
'**********************************************************
Private Sub mScanVlf()
    Dim slSellName As String        'Selling vehicle name
    Dim slAirName As String         'Airing vehicle name
    Dim ilRet As Integer            'Return from btrieve call
    Dim ilSellIndex As Integer      'Selling List box index
    Dim ilSellListIndex As Integer  'Selling List item index
    Dim ilAirIndex As Integer       'Airing List box index
    Dim ilAirListIndex As Integer   'Airing List item index
    Dim slSellTime As String        'Selling time string
    Dim slAirTime As String         'Airing time string
    Dim ilSellRecNo As Integer      'tmVlfPop record number for selling
    Dim ilAirRecNo As Integer       'tmVlfPop record number for airing
    Dim llNoVlfRecs As Long         'number of VLF records
    Dim llRecCount As Long          'Vlf record counter
    Dim ilSellFound As Integer      '1=matching selling vehicle found ; 0=none found
    Dim ilAirFound As Integer       '1=matching airing vehicle found ; 0=none found
    Dim slType As String            'S=selling ; A=Airing
    Dim slTime As String            'time value from selling list string
    Dim slTime2 As String           'time value from airing list string
    Dim slAirRecNo As String        'tmVlfPop(record number) for airing
    Dim slSellRecNo As String       'tmVlfPop(record number) for selling
    Dim slSellString As String      'Link string to be added to selling
    Dim slAirString As String       'Link string to be added to airing
    Dim ilCount As Integer          'List item index counter
    Dim ilSellAdd As Integer        'True=link added to selling ;False=no link added
    Dim ilAirAdd As Integer         'True=link added to airing  ;False=no link added
    Dim ilAirTimeFound As Integer   'True = air time found in list
    Dim ilSellTimeFound As Integer  'True = sell time found in list
    Dim slStr As String             'Parsing string
    Dim ilVlfOk As Integer          'Indicator if Vlf record is Ok
    Dim ilSeqNo As Integer
    Dim slSeqNo As String
    Dim slEffDate As String
    Dim slEndDate As String
    Dim slLinksDefStatus As String
    Dim slStr1 As String
    Dim slStr2 As String
    Dim llEndDate As Long
    Dim tlVlf As VLF
    Dim ilCurVlfVefCode As Integer
    Dim blBrandToBottomLoop As Boolean
    On Error GoTo mScanVlfErr

    ilSellFound = 0
    ilAirFound = 0
    llNoVlfRecs = btrRecords(hmVlfPop) 'Save the number of VLF records
    If llNoVlfRecs = 0 Then
        Exit Sub
    End If
    If (smEndDate <> "") And (smEndDate <> "TNF") Then
        llEndDate = gDateValue(smEndDate)
    Else
        llEndDate = 0
    End If
'   ilRet = gCodeChrRefExist (MainForm, slFileName, ilMatchCode, slCodeFieldName, slMatchChr, sChrFieldName)

    '1/21/13: Replace loopiung thru all vlf record with key read instead
    'For llRecCount = 1 To llNoVlfRecs Step 1
    ilCurVlfVefCode = -1
    'For llRecCount = 0 To imNoSelling - 1 Step 1
    llRecCount = 0
    Do
        blBrandToBottomLoop = False
        'If llRecCount = 1 Then    'Get First VLF record
        '    ilRet = btrGetFirst(hmVlfPop, tlVlf, imVlfPopRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'Get first record
        '    If ilRet = BTRV_ERR_END_OF_FILE Then
        '        GoTo mScanVlfErr
        '    End If
        'End If
        'If (llRecCount > 1) Then   'Get Next VLF record
        '    ilRet = btrGetNext(hmVlfPop, tlVlf, imVlfPopRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        '    If ilRet = BTRV_ERR_END_OF_FILE Then
        '        GoTo mScanVlfErr
        '    End If
        'End If
        If Abs(imSellPending(llRecCount)) <> ilCurVlfVefCode Then
            ilCurVlfVefCode = Abs(imSellPending(llRecCount))
            tmVlfSrchKey3.iSellCode = ilCurVlfVefCode
            tmVlfSrchKey3.iSellDay = imDateCode
            tmVlfSrchKey3.sStatus = ""
            ilRet = btrGetGreaterOrEqual(hmVlfPop, tlVlf, imVlfPopRecLen, tmVlfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)  'Get current record
        Else
            ilRet = btrGetNext(hmVlfPop, tlVlf, imVlfPopRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        End If
        If (ilRet <> BTRV_ERR_NONE) Or (tlVlf.iSellCode <> ilCurVlfVefCode) Or (tlVlf.iSellDay <> imDateCode) Then
            llRecCount = llRecCount + 1
            'GoTo lGetNextVlfRec
            blBrandToBottomLoop = True
        End If
        If Not blBrandToBottomLoop Then
            ilSellAdd = False
            ilAirAdd = False
            ilVlfOk = False
            slLinksDefStatus = ""
            For ilSellIndex = 0 To imNoSelling - 1 Step 1
                If Abs(imSellPending(ilSellIndex)) = tlVlf.iSellCode Then
                    If imSellPending(ilSellIndex) < 0 Then
                        slLinksDefStatus = "P"
                    Else
                        slLinksDefStatus = "C"
                    End If
                    Exit For
                End If
            Next ilSellIndex
            For ilAirIndex = 0 To imNoAiring - 1 Step 1
                If Abs(imAirPending(ilAirIndex)) = tlVlf.iAirCode Then
                    If imAirPending(ilAirIndex) < 0 Then
                        slLinksDefStatus = "P"
                    Else
                        slLinksDefStatus = "C"
                    End If
                    Exit For
                End If
            Next ilAirIndex
            If slLinksDefStatus = "C" Then
                'Select only records which are TFN and have Effective date prior to specified date
                If llEndDate = 0 Then
                    If (tlVlf.iSellDay = imDateCode) And (tlVlf.iTermDate(0) = 0) And (tlVlf.iTermDate(1) = 0) And (tlVlf.sStatus = slLinksDefStatus) Then
                        gUnpackDate tlVlf.iEffDate(0), tlVlf.iEffDate(1), slEffDate
                        If gDateValue(slEffDate) <= lmDateFilter Then
                            ilVlfOk = True
                        End If
                    End If
                Else
                    gUnpackDate tlVlf.iEffDate(0), tlVlf.iEffDate(1), slEffDate
                    gUnpackDate tlVlf.iTermDate(0), tlVlf.iTermDate(1), slEndDate
                    If slEndDate = "" Then
                        slEndDate = "12/31/2060"
                    End If
                    If (tlVlf.iSellDay = imDateCode) And (gDateValue(slEffDate) <= lmDateFilter) And (gDateValue(slEndDate) >= llEndDate) And (tlVlf.sStatus = slLinksDefStatus) Then
                        ilVlfOk = True
                    End If
                End If
            Else
                If (tlVlf.iSellDay = imDateCode) And (tlVlf.iEffDate(0) = imDate0) And (tlVlf.iEffDate(1) = imDate1) And (tlVlf.sStatus = slLinksDefStatus) Then
                    ilVlfOk = True
                End If
            End If
            If ilVlfOk Then
                gObtainVehicleName tlVlf.iSellCode, slSellName, slType
                gObtainVehicleName tlVlf.iAirCode, slAirName, slType
                slSellName = Trim$(slSellName)  'Save Selling Name
                slAirName = Trim$(slAirName)    'Save Airing Name
                gUnpackTime tlVlf.iSellTime(0), tlVlf.iSellTime(1), "A", "1", slSellTime
                gUnpackTime tlVlf.iAirTime(0), tlVlf.iAirTime(1), "A", "1", slAirTime
                slSellTime = Trim$(slSellTime)  'Save Selling Time
                slAirTime = Trim$(slAirTime)    'Save Airing Time
                ilSellFound = False
                ilAirFound = False
                'Changed 12/2/8/98 by Dick- to aviod endless loop when times Not found
                ilSellTimeFound = False 'True
                ilAirTimeFound = False  'True
                For ilSellIndex = 0 To imNoSelling - 1 Step 1
                    If (slSellName = lacSelling(ilSellIndex).Caption) Then
                        ilSellFound = True  'save selling lbc index (ilSellIndex)
                        For ilCount = 0 To lbcSelling(ilSellIndex).ListCount - 1 Step 1
                            If (Left$(lbcSelling(ilSellIndex).List(ilCount), 2) <> "  ") Then
                                ilRet = gParseItem(Trim$(lbcSelling(ilSellIndex).List(ilCount)), 1, " ", slStr)
                                slStr = Trim$(slStr)
                                If (slStr = slSellTime) Then
                                    ilSellTimeFound = True
                                    Exit For
                                Else
                                    ilSellTimeFound = False
                                End If
                            End If
                        Next ilCount
                        Exit For
                    End If
                Next ilSellIndex
                For ilAirIndex = 0 To imNoAiring - 1 Step 1
                    If (slAirName = lacAiring(ilAirIndex).Caption) Then
                        ilAirFound = True   'save airing lbc index (ilAirIndex)
                        For ilCount = 0 To lbcAiring(ilAirIndex).ListCount - 1 Step 1
                            If (Left$(lbcAiring(ilAirIndex).List(ilCount), 2) <> "  ") Then
                                ilRet = gParseItem(Trim$(lbcAiring(ilAirIndex).List(ilCount)), 1, " ", slStr)
                                slStr = Trim$(slStr)
                                If (slStr = slAirTime) Then
                                    ilAirTimeFound = True
                                    Exit For
                                Else
                                    ilAirTimeFound = False
                                End If
                            End If
                        Next ilCount
                        Exit For
                    End If
                Next ilAirIndex
                If Not ilSellTimeFound Or Not ilAirTimeFound Then
                    'GoTo lGetNextVlfRec
                    blBrandToBottomLoop = True
                End If
                If Not ilSellTimeFound And Not ilAirTimeFound Then
                    'GoTo lGetNextVlfRec
                    blBrandToBottomLoop = True
                End If
                If (Not blBrandToBottomLoop) And (ilAirFound) And (ilSellFound) And (ilSellTimeFound) And (ilAirTimeFound) Then   'Both vehicles selected
                    For ilSellListIndex = 0 To lbcSelling(ilSellIndex).ListCount - 1 Step 1
                        If (Left$(lbcSelling(ilSellIndex).List(ilSellListIndex), 2) <> "  ") Then
                            ilRet = gParseItem(lbcSelling(ilSellIndex).List(ilSellListIndex), 1, " ", slTime)
                            slTime = Trim$(slTime)
                            If (slTime = slSellTime) Then    'Matching time found in selling
                                ilRet = gParseItem(lbcSelling(ilSellIndex).List(ilSellListIndex), 2, "@", slSellRecNo)
                                ilSellRecNo = Val(slSellRecNo) 'Save tmvlfpop record # for selling
                                tmVlfPop(ilSellRecNo).iAirSeq = tlVlf.iAirSeq 'Set tmvlfPop air sequence number
                                tmVlfPop(ilSellRecNo).iSellSeq = tlVlf.iSellSeq  'set tmvlfpop air seq number
                                'get air rec # ...set  air seqno to sell seq #
                                For ilAirListIndex = 0 To lbcAiring(ilAirIndex).ListCount - 1 Step 1
                                    If (Left$(lbcAiring(ilAirIndex).List(ilAirListIndex), 2) <> "  ") Then
                                        ilRet = gParseItem(lbcAiring(ilAirIndex).List(ilAirListIndex), 1, " ", slTime2)
                                        slTime2 = Trim$(slTime2)
                                        If (slTime2 = slAirTime) Then 'matching time found in airing
                                            ilRet = gParseItem(lbcAiring(ilAirIndex).List(ilAirListIndex), 2, "@", slAirRecNo)
                                            ilAirRecNo = Val(slAirRecNo) 'save airing link record number
                                            tmVlfPop(ilAirRecNo).iSellSeq = tlVlf.iSellSeq  'set tmvlfpop air seq number
                                            'Added 4/10/01 to obtain records in correct order
                                            tmVlfPop(ilAirRecNo).iAirSeq = tlVlf.iAirSeq  'set tmvlfpop air seq number
                                            Exit For
                                        End If
                                    End If
                                Next ilAirListIndex
                                'make strings...check for collating...and add
                                slSellString = CStr("  " & slTime2 & " " & slAirName & "@" & ilAirRecNo & "@" & tlVlf.iSellSeq) 'link string to add to selling
                                slAirString = CStr("  " & slTime & " " & slSellName & "@" & ilSellRecNo & "@" & tlVlf.iAirSeq) 'link string to add to airing
                                If (ilSellListIndex + 1 > lbcSelling(ilSellIndex).ListCount - 1) Then
                                    lbcSelling(ilSellIndex).AddItem mFillTo100(slSellString), ilSellListIndex + 1
                                    ilSellListIndex = ilSellListIndex + 1
                                    ilSellAdd = True
                                End If
                                If Not ilSellAdd Then
                                    For ilCount = ilSellListIndex + 1 To lbcSelling(ilSellIndex).ListCount - 1 Step 1
                                        If (Left$(lbcSelling(ilSellIndex).List(ilCount), 2) <> "  ") Then
                                            lbcSelling(ilSellIndex).AddItem mFillTo100(slSellString), ilCount
                                            ilSellAdd = True
                                            Exit For
                                        Else
                                            'Changed index from ilSellListIndex to ilCount 4/10/01 to get sequence numbers correct
                                            'ilRet = gParseItem(LTrim$(lbcSelling(ilSellIndex).List(ilSellListIndex)), 2, "@", slAirRecNo)
                                            ilRet = gParseItem(LTrim$(lbcSelling(ilSellIndex).List(ilCount)), 3, "@", slSeqNo)
                                            ilSeqNo = Val(slSeqNo)
                                            If (ilSeqNo > tlVlf.iSellSeq) Then 'insert here
                                                lbcSelling(ilSellIndex).AddItem mFillTo100(slSellString), ilCount
                                                ilSellAdd = True
                                                Exit For
                                            'Changed 'and' to 'or' 4/10/01 to add if at end
                                            'ElseIf (tmVlfPop(ilAirRecNo).iSellSeq < tlVLF.iSellSeq) Or (ilCount = lbcSelling(ilSellIndex).ListCount - 1) Then
                                            ElseIf (ilCount = lbcSelling(ilSellIndex).ListCount - 1) Then
                                                lbcSelling(ilSellIndex).AddItem mFillTo100(slSellString), ilCount + 1
                                                ilSellAdd = True
                                                Exit For
                                            End If
                                        End If
                                    Next ilCount
                                End If
                                If (ilAirListIndex + 1 > lbcAiring(ilAirIndex).ListCount - 1) Then
                                    lbcAiring(ilAirIndex).AddItem mFillTo100(slAirString), ilAirListIndex + 1
                                    ilAirListIndex = ilAirListIndex + 1
                                    ilAirAdd = True
                                End If
                                If Not ilAirAdd Then
                                    For ilCount = ilAirListIndex + 1 To lbcAiring(ilAirIndex).ListCount - 1 Step 1
                                        If (Left$(lbcAiring(ilAirIndex).List(ilCount), 2) <> "  ") Then
                                            lbcAiring(ilAirIndex).AddItem mFillTo100(slAirString), ilCount
                                            ilAirAdd = True
                                            Exit For
                                        Else
                                            'Changed index from ilAirListIndex to ilCount 4/10/01 to get sequence numbers correct
                                            'ilRet = gParseItem(lbcAiring(ilAirIndex).List(ilAirListIndex), 2, "@", slSellRecNo)
                                            ilRet = gParseItem(lbcAiring(ilAirIndex).List(ilCount), 3, "@", slSeqNo)
                                            ilSeqNo = Val(slSeqNo)
                                            If (ilSeqNo > tlVlf.iAirSeq) Then 'insert here
                                                lbcAiring(ilAirIndex).AddItem mFillTo100(slAirString), ilCount
                                                ilAirAdd = True
                                                Exit For
                                            'Changed 'and' to 'or' 4/10/01 to add if at end
                                            'ElseIf (tmVlfPop(ilSellRecNo).iAirSeq < tlVLF.iAirSeq) Or (ilCount = lbcAiring(ilAirIndex).ListCount - 1) Then
                                            ElseIf (ilCount = lbcAiring(ilAirIndex).ListCount - 1) Then
                                                lbcAiring(ilAirIndex).AddItem mFillTo100(slAirString), ilCount + 1
                                                ilAirAdd = True
                                                Exit For
                                            End If
                                        End If
                                    Next ilCount
                                End If
                            End If
                            If ilAirAdd And ilSellAdd Then
                                'GoTo lGetNextVlfRec
                                blBrandToBottomLoop = True
                            End If
                        End If
                        If blBrandToBottomLoop Then
                            Exit For
                        End If
                    Next ilSellListIndex
                ElseIf (Not blBrandToBottomLoop) And (ilAirFound) And (Not ilSellFound) Then ' sell vehicle not selected
                    For ilAirListIndex = 0 To lbcAiring(ilAirIndex).ListCount - 1 Step 1
                        If (Left$(lbcAiring(ilAirIndex).List(ilAirListIndex), 2) <> "  ") Then
                            ilRet = gParseItem(lbcAiring(ilAirIndex).List(ilAirListIndex), 1, " ", slTime2)
                            slTime2 = Trim$(slTime2)
                            If (slTime2 = slAirTime) Then 'matching time found in airing
                                ilRet = gParseItem(lbcAiring(ilAirIndex).List(ilAirListIndex), 2, "@", slAirRecNo)
                                ilAirRecNo = Val(slAirRecNo) 'save airing link record number
                                tmVlfPop(ilAirRecNo).iSellSeq = tlVlf.iSellSeq  'set tmvlfpop air seq number
                                tmVlfPop(ilAirRecNo).iAirSeq = tlVlf.iAirSeq  'set tmvlfpop air seq number
                                Exit For
                            End If
                        End If
                    Next ilAirListIndex
                    imUpperBound = imUpperBound + 1
                    '6/6/16: Replaced GoSub
                    'GoSub lAddRecord
                    mAddRecord tlVlf
                    ilSellRecNo = imUpperBound
                    slAirString = CStr("  " & slSellTime & " " & slSellName & "@" & imUpperBound & "@" & tlVlf.iAirSeq) '& ilSellRecNo) 'link string to add to airing
                    For ilCount = ilAirListIndex + 1 To lbcAiring(ilAirIndex).ListCount - 1 Step 1
                        If (Left$(lbcAiring(ilAirIndex).List(ilCount), 2) <> "  ") Then
                            lbcAiring(ilAirIndex).AddItem mFillTo100(slAirString), ilCount
                            ilAirAdd = True
                            Exit For
                        Else
                            'If (tmVlfPop(ilSellRecNo).iAirSeq > tlVLF.iAirSeq) Then 'insert here
                            '    lbcAiring(ilAirIndex).AddItem slAirString, ilCount
                            '    ilAirAdd = True
                            '    Exit For
                            'End If
                            ilRet = gParseItem(lbcAiring(ilAirIndex).List(ilCount), 3, "@", slSeqNo)
                            ilSeqNo = Val(slSeqNo)
                            If (ilSeqNo > tlVlf.iAirSeq) Then 'insert here
                                lbcAiring(ilAirIndex).AddItem mFillTo100(slAirString), ilCount
                                ilAirAdd = True
                                Exit For
                            ElseIf (ilCount = lbcAiring(ilAirIndex).ListCount - 1) Then
                                lbcAiring(ilAirIndex).AddItem mFillTo100(slAirString), ilCount + 1
                                ilAirAdd = True
                                Exit For
                            End If
                        End If
                    Next ilCount
                    If (ilAirAdd) And Not (ilSellAdd) Then
                        'GoTo lGetNextVlfRec
                        blBrandToBottomLoop = True
                    End If
                ElseIf (Not blBrandToBottomLoop) And (Not ilAirFound) And (ilSellFound) Then ' Air vehicle not selected
                    For ilSellListIndex = 0 To lbcSelling(ilSellIndex).ListCount - 1 Step 1
                        If (Left$(lbcSelling(ilSellIndex).List(ilSellListIndex), 2) <> "  ") Then
                            ilRet = gParseItem(lbcSelling(ilSellIndex).List(ilSellListIndex), 1, " ", slTime)
                            slTime = Trim$(slTime)
                            If (slTime = slSellTime) Then    'Matching time found in selling
                                ilRet = gParseItem(lbcSelling(ilSellIndex).List(ilSellListIndex), 2, "@", slSellRecNo)
                                ilSellRecNo = Val(slSellRecNo) 'Save tmvlfpop record # for selling
                                imUpperBound = imUpperBound + 1
                                '6/6/16: Replaced GoSub
                                'GoSub lAddRecord
                                mAddRecord tlVlf
                                ilAirRecNo = imUpperBound
                                slSellString = CStr("  " & slAirTime & " " & slAirName & "@" & imUpperBound & "@" & tlVlf.iSellSeq) 'link string to add to selling
                                For ilCount = ilSellListIndex + 1 To lbcSelling(ilSellIndex).ListCount - 1 Step 1
                                    If (Left$(lbcSelling(ilSellIndex).List(ilCount), 2) <> "  ") Then
                                        lbcSelling(ilSellIndex).AddItem mFillTo100(slSellString), ilCount
                                        ilSellAdd = True
                                        Exit For
                                    Else
                                        'If (tmVlfPop(ilAirRecNo).iSellSeq > tlVLF.iSellSeq) Then 'insert here
                                        '    lbcSelling(ilSellIndex).AddItem slSellString, ilCount
                                        '    ilSellAdd = True
                                        '    Exit For
                                        'End If
                                        ilRet = gParseItem(LTrim$(lbcSelling(ilSellIndex).List(ilCount)), 3, "@", slSeqNo)
                                        ilSeqNo = Val(slSeqNo)
                                        If (ilSeqNo > tlVlf.iSellSeq) Then 'insert here
                                            lbcSelling(ilSellIndex).AddItem mFillTo100(slSellString), ilCount
                                            ilSellAdd = True
                                            Exit For
                                        'Changed 'and' to 'or' 4/10/01 to add if at end
                                        'ElseIf (tmVlfPop(ilAirRecNo).iSellSeq < tlVLF.iSellSeq) Or (ilCount = lbcSelling(ilSellIndex).ListCount - 1) Then
                                        ElseIf (ilCount = lbcSelling(ilSellIndex).ListCount - 1) Then
                                            lbcSelling(ilSellIndex).AddItem mFillTo100(slSellString), ilCount + 1
                                            ilSellAdd = True
                                            Exit For
                                        End If
                                    End If
                                Next ilCount
                            End If
                            If Not (ilAirAdd) And (ilSellAdd) Then
                                'GoTo lGetNextVlfRec
                                blBrandToBottomLoop = True
                            End If
                        End If
                        If blBrandToBottomLoop Then
                            Exit For
                        End If
                    Next ilSellListIndex
                End If
            End If
        End If
'lGetNextVlfRec:
    ''Next llRecCount
    Loop While llRecCount <= (imNoSelling - 1)
    'Remove sequence number
    For ilSellIndex = 0 To imNoSelling - 1 Step 1
        For ilSellListIndex = 0 To lbcSelling(ilSellIndex).ListCount - 1 Step 1
            ilRet = gParseItemNoTrim(lbcSelling(ilSellIndex).List(ilSellListIndex), 1, "@", slStr1)
            ilRet = gParseItemNoTrim(lbcSelling(ilSellIndex).List(ilSellListIndex), 2, "@", slStr2)
            lbcSelling(ilSellIndex).List(ilSellListIndex) = slStr1 & "@" & slStr2
        Next ilSellListIndex
    Next ilSellIndex
    For ilAirIndex = 0 To imNoAiring - 1 Step 1
        For ilAirListIndex = 0 To lbcAiring(ilAirIndex).ListCount - 1 Step 1
            ilRet = gParseItemNoTrim(lbcAiring(ilAirIndex).List(ilAirListIndex), 1, "@", slStr1)
            ilRet = gParseItemNoTrim(lbcAiring(ilAirIndex).List(ilAirListIndex), 2, "@", slStr2)
            lbcAiring(ilAirIndex).List(ilAirListIndex) = slStr1 & "@" & slStr2
        Next ilAirListIndex
    Next ilAirIndex
    Exit Sub
'lAddRecord:
'    If UBound(tmVlfPop) < imUpperBound Then
'        ReDim Preserve tmVlfPop(0 To imUpperBound)
'    End If
'    tmVlfPop(imUpperBound).iSellCode = tlVlf.iSellCode
'    tmVlfPop(imUpperBound).iSellDay = tlVlf.iSellDay                '0=M-F etc
'    tmVlfPop(imUpperBound).iSellTime(0) = tlVlf.iSellTime(0)
'    tmVlfPop(imUpperBound).iSellTime(1) = tlVlf.iSellTime(1)
'    tmVlfPop(imUpperBound).iSellPosNo = 0
'    tmVlfPop(imUpperBound).iSellSeq = tlVlf.iSellSeq
'    tmVlfPop(imUpperBound).sStatus = tlVlf.sStatus
'    tmVlfPop(imUpperBound).iAirCode = tlVlf.iAirCode
'    tmVlfPop(imUpperBound).iAirDay = tlVlf.iAirDay
'    tmVlfPop(imUpperBound).iAirTime(0) = tlVlf.iAirTime(0)
'    tmVlfPop(imUpperBound).iAirTime(1) = tlVlf.iAirTime(1)
'    tmVlfPop(imUpperBound).iAirPosNo = 0
'    tmVlfPop(imUpperBound).iAirSeq = tlVlf.iAirSeq
'    tmVlfPop(imUpperBound).iEffDate(0) = tlVlf.iEffDate(0)
'    tmVlfPop(imUpperBound).iEffDate(1) = tlVlf.iEffDate(1)
'    tmVlfPop(imUpperBound).iTermDate(0) = tlVlf.iTermDate(0)
'    tmVlfPop(imUpperBound).iTermDate(1) = tlVlf.iTermDate(1)
'    tmVlfPop(imUpperBound).sDelete = ""
'    Return
mScanVlfErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'**********************************************************************
'
'       Procedure Name : mSetCommands
'
'       Created : 4/24/94       By: D. Hannifan
'       Modified :              By:
'
'       Comments : Procedure to set control properties in LinksDef
'
'
'**********************************************************************
'
Private Sub mSetCommands()

    tmcScroll.Enabled = False
    If imTerminate Then   'Critical Error has occurred
        mTerminate
        Exit Sub
    End If
    If (ckcShow.Value = vbChecked) Then     'Currently showing discrepencies
        cmcDone.Enabled = False
        cmcUpdate.Enabled = False
    Else
        cmcDone.Enabled = True
        If (imUpdateFlag = False) Or (Not imUpdateAllowed) Then  'No changes made
            cmcUpdate.Enabled = False
        Else
            cmcUpdate.Enabled = True   'Changes were made since last call
        End If
    End If
End Sub
'**********************************************************************
'
'       Procedure Name : mSetBoxTabs
'
'       Created : 4/24/94       By: D. Hannifan
'       Modified :              By:
'
'       Comments : Procedure to set "SetTabs" options for
'                  list box arrays in LinksDef
'
'**********************************************************************
'
Private Sub mSetLBoxTabs()
   ' Dim ilCount As Integer 'List box index counter
   '     For ilCount = 1 To imNoSelling - 1 Step 1
   '         lbcSelling(ilCount).TabCharacter = lbcSelling(0).TabCharacter
   '         lbcSelling(ilCount).TabPos(0) = lbcSelling(0).TabPos(0)
   '         lbcSelling(ilCount).TabScale = lbcSelling(0).TabScale
   '         lbcSelling(ilCount).TabType(0) = lbcSelling(0).TabType(0)
   '     Next ilCount
   '     For ilCount = 1 To imNoAiring - 1 Step 1
   '         lbcAiring(ilCount).TabCharacter = lbcAiring(0).TabCharacter
   '         lbcAiring(ilCount).TabPos(0) = lbcAiring(0).TabPos(0)
   '         lbcAiring(ilCount).TabScale = lbcAiring(0).TabScale
   '         lbcAiring(ilCount).TabType(0) = lbcAiring(0).TabType(0)
   '     Next ilCount
End Sub
'****************************************************************
'
'          Procedure Name : mSetSelected
'
'
'       Created : 4/9/94        By: D. Hannifan
'       Modified :              By:
'
'       Comments : Routine to select/deselect list box items
'                  based upon last operation performed
'
'****************************************************************
Private Sub mSetSelected(ilIndex As Integer, ilIndex2 As Integer, ilListIndex As Integer, ilListIndex2 As Integer, slType As String, slType2 As String)
    ' ilIndex (I)           : List box index
    ' ilIndex2 (I)          : List box 2 index
    ' ilListIndex (I)       : List item index
    ' ilListIndex2 (I)      : Listbox 2 item index
    ' slType (I)            : "A"=airing  "S"=selling   "C" = clear all
    ' slType2 (I)           : "A"=airing  "S"=selling   "C" = work on only one type  "ALLOTHERS" = both sell and air except the one calling
    Dim ilCount As Integer      'Index Counter
    Dim ilValue As Integer

    ilValue = False
    slType = Trim$(slType)
    slType2 = Trim$(slType2)
    If (slType = "A") And (slType2 = "A") Then
        For ilCount = 0 To imNoAiring - 1 Step 1
            If (ilIndex = ilCount) Then
                lbcAiring(ilCount).Selected(ilListIndex) = True
            End If
            If (ilIndex2 = ilCount) Then
                lbcAiring(ilCount).Selected(ilListIndex2) = True
            End If
            If (ilCount <> ilIndex) And (ilCount <> ilIndex2) Then
                ''lbcAiring(ilCount).Selected(-1) = False
                'llRg = CLng(lbcAiring(ilCount).ListCount - 1) * &H10000 Or 0
                'llRet = SendMessageByNum(lbcAiring(ilCount).hwnd, LB_SELITEMRANGE, ilValue, llRg)
                lbcAiring(ilCount).ListIndex = -1
            End If
        Next ilCount
        For ilCount = 0 To imNoSelling - 1 Step 1
            'lbcSelling(ilCount).Selected(-1) = False
            'llRg = CLng(lbcSelling(ilCount).ListCount - 1) * &H10000 Or 0
            'llRet = SendMessageByNum(lbcSelling(ilCount).hwnd, LB_SELITEMRANGE, ilValue, llRg)
            lbcSelling(ilCount).ListIndex = -1
        Next ilCount
    End If
    If (slType = "S") And (slType2 = "C") Then
        For ilCount = 0 To imNoSelling - 1 Step 1
            If (ilIndex = ilCount) Then
                lbcSelling(ilCount).Selected(ilListIndex) = True
            End If
            If (ilCount <> ilIndex) Then
                ''lbcSelling(ilCount).Selected(-1) = False
                'llRg = CLng(lbcSelling(ilCount).ListCount - 1) * &H10000 Or 0
                'llRet = SendMessageByNum(lbcSelling(ilCount).hwnd, LB_SELITEMRANGE, ilValue, llRg)
                lbcSelling(ilCount).ListIndex = -1
            End If
        Next ilCount
    End If
    If (slType = "A") And (slType2 = "C") Then
        For ilCount = 0 To imNoAiring - 1 Step 1
            If (ilIndex = ilCount) Then
                lbcAiring(ilCount).Selected(ilListIndex) = True
            End If
            If (ilCount <> ilIndex) Then
                ''lbcAiring(ilCount).Selected(-1) = False
                'llRg = CLng(lbcAiring(ilCount).ListCount - 1) * &H10000 Or 0
                'llRet = SendMessageByNum(lbcAiring(ilCount).hwnd, LB_SELITEMRANGE, ilValue, llRg)
                lbcAiring(ilCount).ListIndex = -1
            End If
        Next ilCount
    End If
    If (slType = "S") And (slType2 = "S") Then
        For ilCount = 0 To imNoSelling - 1 Step 1
            If (ilIndex = ilCount) Then
                lbcSelling(ilCount).Selected(ilListIndex) = True
            End If
            If (ilIndex2 = ilCount) Then
                lbcSelling(ilCount).Selected(ilListIndex2) = True
            End If
            If (ilCount <> ilIndex) And (ilCount <> ilIndex2) Then
                ''lbcSelling(ilCount).Selected(-1) = False
                'llRg = CLng(lbcSelling(ilCount).ListCount - 1) * &H10000 Or 0
                'llRet = SendMessageByNum(lbcSelling(ilCount).hwnd, LB_SELITEMRANGE, ilValue, llRg)
                lbcSelling(ilCount).ListIndex = -1
            End If
        Next ilCount
        For ilCount = 0 To imNoAiring - 1 Step 1
            ''lbcAiring(ilCount).Selected(-1) = False
            'llRg = CLng(lbcAiring(ilCount).ListCount - 1) * &H10000 Or 0
            'llRet = SendMessageByNum(lbcAiring(ilCount).hwnd, LB_SELITEMRANGE, ilValue, llRg)
            lbcAiring(ilCount).ListIndex = -1
        Next ilCount
    End If
    If (slType = "C") And (slType2 = "C") Then
        For ilCount = 0 To imNoAiring - 1 Step 1
            ''lbcAiring(ilCount).Selected(-1) = False
            'llRg = CLng(lbcAiring(ilCount).ListCount - 1) * &H10000 Or 0
            'llRet = SendMessageByNum(lbcAiring(ilCount).hwnd, LB_SELITEMRANGE, ilValue, llRg)
            lbcAiring(ilCount).ListIndex = -1
        Next ilCount
        For ilCount = 0 To imNoSelling - 1 Step 1
            ''lbcSelling(ilCount).Selected(-1) = False
            'llRg = CLng(lbcSelling(ilCount).ListCount - 1) * &H10000 Or 0
            'llRet = SendMessageByNum(lbcSelling(ilCount).hwnd, LB_SELITEMRANGE, ilValue, llRg)
            lbcSelling(ilCount).ListIndex = -1
        Next ilCount
    End If
    If (slType = "A") And (slType2 = "ALLOTHERS") Then
        For ilCount = 0 To imNoAiring - 1 Step 1
            If (ilCount <> ilIndex) Then
                ''lbcAiring(ilCount).Selected(-1) = False
                'llRg = CLng(lbcAiring(ilCount).ListCount - 1) * &H10000 Or 0
                'llRet = SendMessageByNum(lbcAiring(ilCount).hwnd, LB_SELITEMRANGE, ilValue, llRg)
                lbcAiring(ilCount).ListIndex = -1
            End If
        Next ilCount
        For ilCount = 0 To imNoSelling - 1 Step 1
            ''lbcSelling(ilCount).Selected(-1) = False
            'llRg = CLng(lbcSelling(ilCount).ListCount - 1) * &H10000 Or 0
            'llRet = SendMessageByNum(lbcSelling(ilCount).hwnd, LB_SELITEMRANGE, ilValue, llRg)
            lbcSelling(ilCount).ListIndex = -1
        Next ilCount
    End If
    If (slType = "S") And (slType2 = "ALLOTHERS") Then
        For ilCount = 0 To imNoAiring - 1 Step 1
            ''lbcAiring(ilCount).Selected(-1) = False
            'llRg = CLng(lbcAiring(ilCount).ListCount - 1) * &H10000 Or 0
            'llRet = SendMessageByNum(lbcAiring(ilCount).hwnd, LB_SELITEMRANGE, ilValue, llRg)
            lbcAiring(ilCount).ListIndex = -1
        Next ilCount
        For ilCount = 0 To imNoSelling - 1 Step 1
            If (ilCount <> ilIndex) Then
                ''lbcSelling(ilCount).Selected(-1) = False
                'llRg = CLng(lbcSelling(ilCount).ListCount - 1) * &H10000 Or 0
                'llRet = SendMessageByNum(lbcSelling(ilCount).hwnd, LB_SELITEMRANGE, ilValue, llRg)
                lbcSelling(ilCount).ListIndex = -1
            End If
        Next ilCount
    End If

End Sub
'**********************************************************************
'
'       Procedure Name : mShowDiscreps
'
'       Created : 3/19/94       By: D. Hannifan
'       Modified :              By:
'
'       Comments : Procedure to clear, initialize and populate
'                  LinksDef list boxes for a Show Discrepencies event
'
'**********************************************************************
'
Private Sub mShowDiscreps()
    Dim ilCount As Integer                          'List Box index counter
    Dim ilCount2 As Integer                         'List item index counter
    ReDim imSellCount(imNoSelling - 1) As Integer   'Number of list items in selling array
    ReDim imAirCount(imNoAiring - 1) As Integer     'Number of list items in airing array
    Dim slString1 As String * 2                     'First two characters of list string
    Dim slString2 As String * 2                     'First two characters of list string
    On Error GoTo mShowDisCrepsErr
    imSellUpperB = 0
    imAirUpperB = 0

    For ilCount = 0 To imNoSelling - 1 Step 1   'Find Max # of list items in selling list boxes
        imSellCount(ilCount) = lbcSelling(ilCount).ListCount
        If (imSellUpperB < lbcSelling(ilCount).ListCount) Then
            imSellUpperB = lbcSelling(ilCount).ListCount
        End If
    Next ilCount
    For ilCount = 0 To imNoAiring - 1 Step 1  'Find Max # of list items in airing list boxes
        imAirCount(ilCount) = lbcAiring(ilCount).ListCount
        If (imAirUpperB < lbcAiring(ilCount).ListCount) Then
            imAirUpperB = lbcAiring(ilCount).ListCount
        End If
    Next ilCount
    ReDim smSellingLists(imNoSelling - 1, imSellUpperB) 'Dim array to hold selling lists temporarily
    ReDim smAiringLists(imNoAiring - 1, imAirUpperB)    'Dim array to hold airing lists temporarily
    For ilCount = 0 To imNoSelling - 1 Step 1 'Load selling list into selling array
        For ilCount2 = 0 To lbcSelling(ilCount).ListCount - 1 Step 1
            smSellingLists(ilCount, ilCount2) = lbcSelling(ilCount).List(ilCount2)
        Next ilCount2
    Next ilCount
    For ilCount = 0 To imNoAiring - 1 Step 1  'Load airing list into airing array
        For ilCount2 = 0 To lbcAiring(ilCount).ListCount - 1 Step 1
            smAiringLists(ilCount, ilCount2) = lbcAiring(ilCount).List(ilCount2)
        Next ilCount2
    Next ilCount
    For ilCount = 0 To imNoSelling - 1 Step 1 ' Clear selling list boxes
        lbcSelling(ilCount).Clear
        lbcSelling(ilCount).ListIndex = -1
    Next ilCount
    For ilCount = 0 To imNoAiring - 1 Step 1 ' Clear airing list boxes
        lbcAiring(ilCount).Clear
        lbcAiring(ilCount).ListIndex = -1
    Next ilCount
    For ilCount = 0 To imNoSelling - 1 Step 1 'Find list items not linked and add to discrepency list
        For ilCount2 = 0 To imSellCount(ilCount) - 2 Step 1
            slString1 = Left$(smSellingLists(ilCount, ilCount2), 2)
            slString2 = Left$(smSellingLists(ilCount, ilCount2 + 1), 2)
            If (slString1 <> "  ") And (slString2 <> "  ") Then  'No link exists for slString1
                lbcSelling(ilCount).AddItem mFillTo100(smSellingLists(ilCount, ilCount2))
            End If
            If (slString2 <> "  ") And (imSellCount(ilCount) - 1 = ilCount2 + 1) Then 'No link exists for slString1
                lbcSelling(ilCount).AddItem mFillTo100(smSellingLists(ilCount, ilCount2 + 1))
            End If
        Next ilCount2
    Next ilCount
    For ilCount = 0 To imNoAiring - 1 Step 1 'Find list items not linked and add to discrepency list
        For ilCount2 = 0 To imAirCount(ilCount) - 2 Step 1
            slString1 = Left$(smAiringLists(ilCount, ilCount2), 2)
            slString2 = Left$(smAiringLists(ilCount, ilCount2 + 1), 2)
            If (slString1 <> "  ") And (slString2 <> "  ") Then  'No link exists for slString1
                lbcAiring(ilCount).AddItem mFillTo100(smAiringLists(ilCount, ilCount2))
            End If
            If (slString2 <> "  ") And (imAirCount(ilCount) - 1 = ilCount2 + 1) Then 'No link exists for slString1
                lbcAiring(ilCount).AddItem mFillTo100(smAiringLists(ilCount, ilCount2 + 1))
            End If
        Next ilCount2
    Next ilCount
    Exit Sub
mShowDisCrepsErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'****************************************************************
'
'                   Procedure Name : mSpaceSize
'
'       Date Created : 4/24/94      By: D. Hannifan
'       Date Modified :             By:
'
'       Comments : Determine spacing and placement for list boxes
'                  shown in LinksDef screen
'*****************************************************************
'
Private Function mSpaceSize(ilPanWidth As Integer, slVehType As String, ilLBoxWidth As Integer) As Integer
Dim ilTotalSpace As Integer
Dim ilTotalLBWidth As Integer
Dim ilNoBoxes As Integer
' Where :
        '       ilPanWidth(I) = width of background panel
        '       slVehType(I) = Airing or Selling vehicle call ;"S"=selling, "A"=Airing
        '       ilLBoxWidth(I) = width of list boxes (all listbox widths are the same)
'               ilNoBoxes(I) = Number of list boxes to display (4 max)
'
'               mSpaceSize(O) = space increment between list boxes & panel edges
    If slVehType = "S" Then
        ilNoBoxes = imNoSelling
    Else
        ilNoBoxes = imNoAiring
    End If
    If ilNoBoxes > 4 Then
        ilNoBoxes = 4
    End If
    ilTotalLBWidth = (ilNoBoxes * ilLBoxWidth) + (2 * CInt(fgBevelX))
    ilTotalSpace = ilPanWidth - ilTotalLBWidth
    mSpaceSize = CSng(ilTotalSpace / (ilNoBoxes + 1))
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:4/24/94       By:D. Hannifan    *
'*                                                     *
'*            Comments: terminate LinksDef form        *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    'Close Files
    imReinitFlag = False

    'Unload form
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload LinksDef
    igManUnload = NO
End Sub
'**********************************************************************
'
'       Procedure Name : mUpdateVlf
'
'       Created : 4/24/94       By: D. Hannifan
'       Modified :              By:
'
'       Comments : Procedure to update links in the temporary VLF file
'                  TmVlfUpdate
'
'**********************************************************************
'
Private Sub mUpdateVlf()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilFound                       ilLoop                        ilTest                    *
'*  llDate                                                                                *
'******************************************************************************************

    Dim slCurSellVefName As String      'Current list item Selling Vehicle Name
    Dim ilCurSellRecNo As Integer       'Current Selling vehicle tmVlfPop record index
    Dim slCurAirVefName As String       'Current list item airing Vehicle Name
    Dim ilCurAirRecNo As Integer        'Current airing vehicle tmVlfPop record index
    Dim ilCurIndex As Integer           'Current lbcSelling index
    Dim ilCurListIndex As Integer       'Current list index
    Dim slCurListString As String       'String in current list index selected
    Dim slParsedString As String        'String following @ delimeter
    Dim slCurListType As String * 1     'List string type "S"=Selling time or "A"=Airing Link
    ReDim tmVlfUpdate(0) As VLF         'Clear array for update
    Dim ilSellSeqNo As Integer          'Sequence number for selling vehicle
    Dim ilAirSeqNo As Integer           'Sequence number for airing vehicle
    Dim ilRet As Integer                'Return from call
    Dim ilCount As Integer              'list index counter
    Dim ilFoundAir As Integer           'True=airing list box item found
    Dim ilCount2 As Integer             'List item index counter
    Dim slTimeStr1 As String            'Link string to find
    Dim slTimeStr2 As String            'Link string to match
    Dim ilCount3 As Integer             'list item index counter
    Dim slCurSellTime As String         'selling time string
    Dim slSrchString As String          'string to search for in list
    Dim slStr As String                 'parse string
    Dim ilAirSeqSet As Integer          'True = airing link sequence number was set
    Dim slLinksDefStatus As String
    Dim ilSellIndex As Integer

    If ckcShow.Value = vbChecked Then     'Discrepencies displayed
        ckcShow.Value = vbUnchecked
        DoEvents
    End If

    imUpdateUpB = 0
    ilSellSeqNo = 0
    ilAirSeqNo = 1
    For ilCurIndex = 0 To imNoSelling - 1 Step 1  'Check selling list box
        ilSellSeqNo = 0
        ilAirSeqNo = 0
        For ilCurListIndex = 0 To lbcSelling(ilCurIndex).ListCount - 1 Step 1
            slCurListString = lbcSelling(ilCurIndex).List(ilCurListIndex)
            If (Left$(slCurListString, 2) = "  ") Then 'Airing link string found
                slCurListType = "A"
                '6/6/16: Replaced GoSub
                'GoSub lParseAirLink   'Get record number for air link
                mParseAirLink slCurListString, slParsedString, ilCurAirRecNo, slCurAirVefName, slCurListType
                If imUpdateUpB = 0 Then
                    ilAirSeqNo = 1
                End If
                ilFoundAir = True
                ilSellSeqNo = ilSellSeqNo + 1
                '6/6/16: Replaced GoSub
                'GoSub lCheckSeqNo
                mCheckSeqNo ilFoundAir, ilAirSeqSet, slCurAirVefName, slCurListString, ilAirSeqNo, slSrchString, ilCurAirRecNo
                If ilAirSeqSet Then
                    '6/6/16: Replaced GoSub
                    'GoSub lCreateLink 'Selling time string found
                    mCreateLink ilCurSellRecNo, ilSellSeqNo, slLinksDefStatus, slCurListType, ilCurAirRecNo, ilAirSeqNo
                End If
            Else
                ilRet = gParseItem(slCurListString, 1, " ", slCurSellTime)
                slCurSellTime = Trim$(slCurSellTime)
                slCurListType = "S"
                ilSellSeqNo = 0
                If imUpdateUpB = 0 Then
                    ilAirSeqNo = 1
                End If
                '6/6/16: Replace GoSub
                'GoSub lParseSellTime   'Get record number for sell vehicle
                mParseSellTime slCurListString, slParsedString, ilCurSellRecNo, slCurSellVefName, slCurListType
                slSrchString = CStr(slCurSellTime & " " & Trim$(slCurSellVefName))
            End If
        Next ilCurListIndex
    Next ilCurIndex
    imUpdateUpB = imUpdateUpB - 1
    Exit Sub
'lParseAirLink:  'Get air vehicle record number & name
'    slCurListString = LTrim$(slCurListString)
'    ilRet = gParseItem(slCurListString, 2, "@", slParsedString)
'    ilCurAirRecNo = Val(Trim$(slParsedString))
'    gObtainVehicleName tmVlfPop(ilCurAirRecNo).iAirCode, slCurAirVefName, slCurListType
'
'    Return
'lParseSellTime: 'Get air vehicle record number & name
'    slCurListString = LTrim$(slCurListString)
'    ilRet = gParseItem(slCurListString, 2, "@", slParsedString)
'    ilCurSellRecNo = Val(Trim$(slParsedString))
'    gObtainVehicleName tmVlfPop(ilCurSellRecNo).iSellCode, slCurSellVefName, slCurListType
'
'    Return
'lCreateLink:
'
'    tmVlfUpdate(imUpdateUpB).iSellCode = tmVlfPop(ilCurSellRecNo).iSellCode
'    tmVlfUpdate(imUpdateUpB).iSellDay = tmVlfPop(ilCurSellRecNo).iSellDay
'    tmVlfUpdate(imUpdateUpB).iSellTime(0) = tmVlfPop(ilCurSellRecNo).iSellTime(0)
'    tmVlfUpdate(imUpdateUpB).iSellTime(1) = tmVlfPop(ilCurSellRecNo).iSellTime(1)
'    tmVlfUpdate(imUpdateUpB).iSellPosNo = 0
'    tmVlfUpdate(imUpdateUpB).iSellSeq = ilSellSeqNo
'    slLinksDefStatus = ""
'    For ilSellIndex = 0 To imNoSelling - 1 Step 1
'        If Abs(imSellPending(ilSellIndex)) = tmVlfUpdate(imUpdateUpB).iSellCode Then
'            If imSellPending(ilSellIndex) < 0 Then
'                slLinksDefStatus = "P"
'            Else
'                slLinksDefStatus = "C"
'            End If
'        End If
'    Next ilSellIndex
'    If slLinksDefStatus = "C" Then
'        tmVlfUpdate(imUpdateUpB).sStatus = "P"
'    ElseIf slCurListType = "S" Then
'        tmVlfUpdate(imUpdateUpB).sStatus = tmVlfPop(ilCurSellRecNo).sStatus
'    ElseIf slCurListType = "A" Then
'        tmVlfUpdate(imUpdateUpB).sStatus = tmVlfPop(ilCurAirRecNo).sStatus
'    End If
'    tmVlfUpdate(imUpdateUpB).iAirCode = tmVlfPop(ilCurAirRecNo).iAirCode
'    tmVlfUpdate(imUpdateUpB).iAirDay = tmVlfPop(ilCurAirRecNo).iAirDay
'    tmVlfUpdate(imUpdateUpB).iAirTime(0) = tmVlfPop(ilCurAirRecNo).iAirTime(0)
'    tmVlfUpdate(imUpdateUpB).iAirTime(1) = tmVlfPop(ilCurAirRecNo).iAirTime(1)
'    tmVlfUpdate(imUpdateUpB).iAirPosNo = 0
'    tmVlfUpdate(imUpdateUpB).iAirSeq = ilAirSeqNo
'    tmVlfUpdate(imUpdateUpB).iEffDate(0) = tmVlfPop(ilCurSellRecNo).iEffDate(0)
'    tmVlfUpdate(imUpdateUpB).iEffDate(1) = tmVlfPop(ilCurSellRecNo).iEffDate(1)
'    tmVlfUpdate(imUpdateUpB).iTermDate(0) = tmVlfPop(ilCurSellRecNo).iTermDate(0)
'    tmVlfUpdate(imUpdateUpB).iTermDate(1) = tmVlfPop(ilCurSellRecNo).iTermDate(1)
'    tmVlfUpdate(imUpdateUpB).sDelete = tmVlfPop(ilCurSellRecNo).sDelete
'    imUpdateUpB = imUpdateUpB + 1
'    ReDim Preserve tmVlfUpdate(imUpdateUpB) As VLF
'    Return
'lCheckSeqNo:
'
'    'If imUpdateUpB = 0 Then
'    '    ilAirSeqSet = True
'    '    Return
'    'End If
'    ilFoundAir = False
'    ilAirSeqSet = False
'    For ilCount = 0 To imNoAiring - 1 Step 1
'        If (Trim$(lacAiring(ilCount).Caption) = Trim$(slCurAirVefName)) Then
'            ilFoundAir = True
'            Exit For
'        End If
'    Next ilCount
'    If ilFoundAir Then
'        ilRet = gParseItem(LTrim$(slCurListString), 1, " ", slTimeStr1)
'        slTimeStr1 = Trim$(slTimeStr1)
'        For ilCount2 = 0 To lbcAiring(ilCount).ListCount - 1 Step 1
'            If (Left$(lbcAiring(ilCount).List(ilCount2), 2) <> "  ") Then
'                ilRet = gParseItem(lbcAiring(ilCount).List(ilCount2), 1, " ", slTimeStr2)
'                slTimeStr2 = Trim$(slTimeStr2)
'                If (slTimeStr1 = slTimeStr2) Then
'                    ilAirSeqNo = 1
'                    For ilCount3 = ilCount2 + 1 To lbcAiring(ilCount).ListCount - 1 Step 1
'                        If (Left$(lbcAiring(ilCount).List(ilCount3), 2) = "  ") Then
'                            ilRet = gParseItem(Trim$(lbcAiring(ilCount).List(ilCount3)), 1, "@", slStr)
'                            slStr = Trim$(slStr)
'                            If (slStr = slSrchString) Then
'                                ilAirSeqSet = True
'                                Exit For
'                            Else
'                                ilAirSeqNo = ilAirSeqNo + 1
'                            End If
'                        Else
'                            ilAirSeqSet = True
'                            Exit For
'                        End If
'                    Next ilCount3
'                End If
'                If ilAirSeqSet Then
'                    Exit For
'                End If
'            End If
'        Next ilCount2
'    Else
'        ilAirSeqNo = tmVlfPop(ilCurAirRecNo).iAirSeq
'        ilAirSeqSet = True
'    End If
'    Return
End Sub
'**********************************************************************
'
'       Procedure Name : mWriteVlf
'
'       Created : 4/24/94       By: D. Hannifan
'       Modified :              By:
'
'       Comments : Procedure to update VLF file
'
'
'**********************************************************************
'
Private Sub mWriteVlf()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slDate                                                                                *
'******************************************************************************************

    Dim ilCount As Integer          'Index for tmVlfPop array
    Dim ilRet As Integer            'Return from btrieve call
    Dim llNoVlfRec As Long          'Number of VLF records
    Dim llRecCount As Long          'Number of VLF records
    Dim slVefName As String         'Vehicle name from List box
    Dim slType As String * 1        'Vehicle type from tmVLFUpdate array
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim llDateFilter As Long
    Dim llEffDate As Long
    Dim llTermDate As Long

    On Error GoTo mWriteVlfErr
    llDateFilter = gDateValue(smDateFilter)
    llNoVlfRec = btrRecords(hmVlfPop) 'Get number of VLF records

    ilRet = btrBeginTrans(hmVlfPop, 1000)
    If ilRet <> BTRV_ERR_NONE Then
    End If
    If (llNoVlfRec = 0) Then ' No records exist...add all records
        For ilCount = 0 To imUpdateUpB Step 1
            If (tmVlfUpdate(ilCount).iSellCode > 0) And (tmVlfUpdate(ilCount).iAirCode > 0) Then
                tmVlfUpdate(ilCount).lCode = 0
                ilRet = btrInsert(hmVlfPop, tmVlfUpdate(ilCount), imVlfPopRecLen, INDEXKEY0)
                If ilRet <> BTRV_ERR_NONE Then
                    '6/6/16: Replaced GoSub
                    'GoSub mAbortWrite
                    mAbortWrite
                    Exit Sub
                End If
            End If
        Next ilCount
        imUpdateFlag = False
        ilRet = btrEndTrans(hmVlfPop)
        Exit Sub

    Else  'Records Exist ; Delete VLF records matching initial population criteria
        ReDim imSellDelVef(0 To 0) As Integer
        For ilCount = 0 To imNoSelling - 1 Step 1
            slVefName = Trim$(lacSelling(ilCount).Caption)
            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If Trim$(tgMVef(ilLoop).sName) = slVefName Then
                    imSellDelVef(UBound(imSellDelVef)) = tgMVef(ilLoop).iCode
                    ReDim Preserve imSellDelVef(0 To UBound(imSellDelVef) + 1) As Integer
                End If
            Next ilLoop
        Next ilCount
        ReDim imAirDelVef(0 To 0) As Integer
        For ilCount = 0 To imNoAiring - 1 Step 1
            slVefName = Trim$(lacAiring(ilCount).Caption)
            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If Trim$(tgMVef(ilLoop).sName) = slVefName Then
                    imAirDelVef(UBound(imAirDelVef)) = tgMVef(ilLoop).iCode
                    ReDim Preserve imAirDelVef(0 To UBound(imAirDelVef) + 1) As Integer
                End If
            Next ilLoop
        Next ilCount
        ReDim lmDelVlf(0 To 0) As Long
        ReDim imChkSellTermVlf(0 To 0) As Integer
        ReDim imChkAirTermVlf(0 To 0) As Integer
        For llRecCount = 1 To llNoVlfRec Step 1
            If llRecCount = 1 Then
                ilRet = btrGetFirst(hmVlfPop, tmVlfPop(0), imVlfPopRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'Get first record
                If ilRet <> BTRV_ERR_NONE Then
                    '6/6/16: Replaced GoSub
                    'GoSub mAbortWrite
                    mAbortWrite
                    Exit Sub
                End If
            End If
            If llRecCount > llNoVlfRec Then
                Exit For
            End If
            If (llRecCount > 1) Then
                ilRet = btrGetNext(hmVlfPop, tmVlfPop(0), imVlfPopRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    '6/6/16: Replaced GoSub
                    'GoSub mAbortWrite
                    mAbortWrite
                    Exit Sub
                End If
            End If
            'gObtainVehicleName tmVlfPop(0).iSellCode, slName, slType
            'slName = Trim$(slName)
            'If (slName = slVefName) Then
            ilTest = False
            For ilLoop = 0 To UBound(imSellDelVef) - 1 Step 1
                If tmVlfPop(0).iSellCode = imSellDelVef(ilLoop) Then
                    ilTest = True
                    Exit For
                End If
            Next ilLoop
            If ilTest Then
                gUnpackDateLong tmVlfPop(0).iEffDate(0), tmVlfPop(0).iEffDate(1), llEffDate
                If (tmVlfPop(0).sStatus = "P") And (llEffDate = llDateFilter) Then
                    If (tmVlfPop(0).iSellDay = imDateCode) Then   'Matching record found
                        lmDelVlf(UBound(lmDelVlf)) = tmVlfPop(0).lCode
                        ReDim Preserve lmDelVlf(0 To UBound(lmDelVlf) + 1) As Long
                    End If
                End If
                gUnpackDateLong tmVlfPop(0).iTermDate(0), tmVlfPop(0).iTermDate(1), llTermDate
                If llTermDate = 0 Then
                    llTermDate = 999999
                End If
                If (tmVlfPop(0).sStatus <> "P") And (llDateFilter >= llEffDate) And (llDateFilter <= llTermDate) And (tmVlfPop(0).iSellDay = imDateCode) Then
                    ilTest = True
                    For ilCount = 0 To imUpdateUpB Step 1   'Set index for tmVlfPop
                        If (tmVlfUpdate(ilCount).iSellCode = tmVlfPop(0).iSellCode) Then
                            ilTest = False
                            Exit For
                        End If
                    Next ilCount
                    If ilTest Then
                        For ilLoop = 0 To UBound(imChkSellTermVlf) - 1 Step 1
                            If tmVlfPop(0).iSellCode = imChkSellTermVlf(ilLoop) Then
                                ilTest = False
                                Exit For
                            End If
                        Next ilLoop
                        If ilTest Then
                            imChkSellTermVlf(UBound(imChkSellTermVlf)) = tmVlfPop(0).iSellCode
                            ReDim Preserve imChkSellTermVlf(0 To UBound(imChkSellTermVlf) + 1) As Integer
                        End If
                    End If
                End If
            End If
            ilTest = False
            For ilLoop = 0 To UBound(imAirDelVef) - 1 Step 1
                If tmVlfPop(0).iAirCode = imAirDelVef(ilLoop) Then
                    ilTest = True
                    Exit For
                End If
            Next ilLoop
            If ilTest Then
                gUnpackDateLong tmVlfPop(0).iEffDate(0), tmVlfPop(0).iEffDate(1), llEffDate
                gUnpackDateLong tmVlfPop(0).iTermDate(0), tmVlfPop(0).iTermDate(1), llTermDate
                If llTermDate = 0 Then
                    llTermDate = 999999
                End If
                If (tmVlfPop(0).sStatus <> "P") And (llDateFilter >= llEffDate) And (llDateFilter <= llTermDate) And (tmVlfPop(0).iSellDay = imDateCode) Then
                    ilTest = True
                    For ilCount = 0 To imUpdateUpB Step 1   'Set index for tmVlfPop
                        If (tmVlfUpdate(ilCount).iAirCode = tmVlfPop(0).iAirCode) Then
                            ilTest = False
                            Exit For
                        End If
                    Next ilCount
                    If ilTest Then
                        For ilLoop = 0 To UBound(imChkAirTermVlf) - 1 Step 1
                            If tmVlfPop(0).iAirCode = imChkAirTermVlf(ilLoop) Then
                                ilTest = False
                                Exit For
                            End If
                        Next ilLoop
                        If ilTest Then
                            imChkAirTermVlf(UBound(imChkAirTermVlf)) = tmVlfPop(0).iAirCode
                            ReDim Preserve imChkAirTermVlf(0 To UBound(imChkAirTermVlf) + 1) As Integer
                        End If
                    End If
                End If
            End If
        Next llRecCount
        'Delete Previous Pending
        For ilLoop = 0 To UBound(lmDelVlf) - 1 Step 1
            Do
                'tmVlfPop(0).lCode = lmDelVlf(ilLoop)
                'tmRec = tmVlfPop(0)
                'ilRet = gGetByKeyForUpdate("VLF", hmVlfPop, tmRec)
                'tmVlfPop(0) = tmRec
                tmVlfSrchKey2.lCode = lmDelVlf(ilLoop)
                ilRet = btrGetEqual(hmVlfPop, tmVlfPop(0), imVlfPopRecLen, tmVlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    '6/6/16: Replaced GoSub
                    'GoSub mAbortWrite
                    mAbortWrite
                    Exit Sub
                End If
                ilRet = btrDelete(hmVlfPop)
                'If ilRet = BTRV_ERR_CONFLICT Then
                '    ilCRet = btrGetDirect(hmVlfPop, tmVlfPop(0), imVlfPopRecLen, llVlfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                '    If ilCRet <> BTRV_ERR_NONE Then
                '        GoSub mAbortWrite
                '        Exit Sub
                '    End If
                'End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                '6/6/16: Replaced GoSub
                'GoSub mAbortWrite
                mAbortWrite
                Exit Sub
            End If
        Next ilLoop
        'Terminate or delete unlinked vehicles
        For ilLoop = 0 To UBound(imChkSellTermVlf) - 1 Step 1
            tmVlfSrchKey0.iSellCode = imChkSellTermVlf(ilLoop)
            tmVlfSrchKey0.iSellDay = imDateCode
            gPackDateLong llDateFilter, tmVlfSrchKey0.iEffDate(0), tmVlfSrchKey0.iEffDate(1)
            tmVlfSrchKey0.iSellTime(0) = 0
            tmVlfSrchKey0.iSellTime(1) = 6144  '24*256
            tmVlfSrchKey0.iSellPosNo = 32000
            ilRet = btrGetLessOrEqual(hmVlfPop, tmVlf, imVlfPopRecLen, tmVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (ilRet = BTRV_ERR_NONE) Then
                gUnpackDateLong tmVlf.iEffDate(0), tmVlf.iEffDate(1), llEffDate
                gUnpackDateLong tmVlf.iTermDate(0), tmVlf.iTermDate(1), llTermDate
                If llTermDate = 0 Then
                    llTermDate = 999999
                End If
            End If
            If (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = imChkSellTermVlf(ilLoop)) And (llDateFilter >= llEffDate) And (llDateFilter <= llTermDate) And (tmVlf.iSellDay = imDateCode) Then
                tmVlfSrchKey0.iSellCode = imChkSellTermVlf(ilLoop)
                tmVlfSrchKey0.iSellDay = imDateCode
                tmVlfSrchKey0.iEffDate(0) = tmVlf.iEffDate(0)
                tmVlfSrchKey0.iEffDate(1) = tmVlf.iEffDate(1)
                tmVcfSrchKey0.iEffDate(0) = tmVlf.iEffDate(0)
                tmVcfSrchKey0.iEffDate(1) = tmVlf.iEffDate(1)
                tmVlfSrchKey0.iSellTime(0) = 0
                tmVlfSrchKey0.iSellTime(1) = 0
                tmVlfSrchKey0.iSellPosNo = 0
                ilRet = btrGetGreaterOrEqual(hmVlfPop, tmVlf, imVlfPopRecLen, tmVlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    gUnpackDateLong tmVlf.iEffDate(0), tmVlf.iEffDate(1), llEffDate
                    gUnpackDateLong tmVlf.iTermDate(0), tmVlf.iTermDate(1), llTermDate
                    If llTermDate = 0 Then
                        llTermDate = 999999
                    End If
                End If
                Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iSellCode = imChkSellTermVlf(ilLoop)) And (llDateFilter >= llEffDate) And (llDateFilter <= llTermDate) And (tmVlf.iSellDay = imDateCode)
                    gPackDateLong llDateFilter - 1, tmVlf.iTermDate(0), tmVlf.iTermDate(1)
                    ilRet = btrUpdate(hmVlfPop, tmVlf, imVlfPopRecLen)
                    ilRet = btrGetNext(hmVlfPop, tmVlf, imVlfPopRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    gUnpackDateLong tmVlf.iEffDate(0), tmVlf.iEffDate(1), llEffDate
                    gUnpackDateLong tmVlf.iTermDate(0), tmVlf.iTermDate(1), llTermDate
                    If llTermDate = 0 Then
                        llTermDate = 999999
                    End If
                Loop
                tmVcfSrchKey0.iSellDay = imChkSellTermVlf(ilLoop)
                tmVcfSrchKey0.iSellTime(0) = 0
                tmVcfSrchKey0.iSellTime(1) = 0
                tmVcfSrchKey0.iSellPosNo = 0
                ilRet = btrGetGreaterOrEqual(hmVcf, tmVcf, imVcfRecLen, tmVcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
                If ilRet = BTRV_ERR_NONE Then
                    gUnpackDateLong tmVcf.iEffDate(0), tmVcf.iEffDate(1), llEffDate
                    gUnpackDateLong tmVcf.iTermDate(0), tmVcf.iTermDate(1), llTermDate
                    If llTermDate = 0 Then
                        llTermDate = 999999
                    End If
                End If
                Do While (ilRet = BTRV_ERR_NONE) And (tmVcf.iSellCode = imChkSellTermVlf(ilLoop)) And (llDateFilter >= llEffDate) And (llDateFilter <= llTermDate) And (tmVcf.iSellDay = imDateCode)
                    gPackDateLong llDateFilter - 1, tmVcf.iTermDate(0), tmVcf.iTermDate(1)
                    ilRet = btrUpdate(hmVcf, tmVcf, imVcfRecLen)
                    ilRet = btrGetNext(hmVcf, tmVcf, imVcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    gUnpackDateLong tmVcf.iEffDate(0), tmVcf.iEffDate(1), llEffDate
                    gUnpackDateLong tmVcf.iTermDate(0), tmVcf.iTermDate(1), llTermDate
                    If llTermDate = 0 Then
                        llTermDate = 999999
                    End If
                Loop
            End If
        Next ilLoop
        For ilLoop = 0 To UBound(imChkAirTermVlf) - 1 Step 1
            tmVlfSrchKey1.iAirCode = imChkAirTermVlf(ilLoop)
            tmVlfSrchKey1.iAirDay = imDateCode
            gPackDateLong llDateFilter, tmVlfSrchKey1.iEffDate(0), tmVlfSrchKey1.iEffDate(1)
            tmVlfSrchKey1.iAirTime(0) = 0
            tmVlfSrchKey1.iAirTime(1) = 6144  '24*256
            tmVlfSrchKey1.iAirPosNo = 32000
            ilRet = btrGetLessOrEqual(hmVlfPop, tmVlf, imVlfPopRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (ilRet = BTRV_ERR_NONE) Then
                gUnpackDateLong tmVlf.iEffDate(0), tmVlf.iEffDate(1), llEffDate
                gUnpackDateLong tmVlf.iTermDate(0), tmVlf.iTermDate(1), llTermDate
                If llTermDate = 0 Then
                    llTermDate = 999999
                End If
            End If
            If (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = imChkAirTermVlf(ilLoop)) And (llDateFilter >= llEffDate) And (llDateFilter <= llTermDate) And (tmVlf.iAirDay = imDateCode) Then
                tmVlfSrchKey1.iAirCode = imChkAirTermVlf(ilLoop)
                tmVlfSrchKey1.iAirDay = imDateCode
                tmVlfSrchKey1.iEffDate(0) = tmVlf.iEffDate(0)
                tmVlfSrchKey1.iEffDate(1) = tmVlf.iEffDate(1)
                tmVlfSrchKey1.iAirTime(0) = 0
                tmVlfSrchKey1.iAirTime(1) = 0
                tmVlfSrchKey1.iAirPosNo = 0
                ilRet = btrGetGreaterOrEqual(hmVlfPop, tmVlf, imVlfPopRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    gUnpackDateLong tmVlf.iEffDate(0), tmVlf.iEffDate(1), llEffDate
                    gUnpackDateLong tmVlf.iTermDate(0), tmVlf.iTermDate(1), llTermDate
                    If llTermDate = 0 Then
                        llTermDate = 999999
                    End If
                End If
                Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = imChkAirTermVlf(ilLoop)) And (llDateFilter >= llEffDate) And (llDateFilter <= llTermDate) And (tmVlf.iAirDay = imDateCode)
                    gPackDateLong llDateFilter - 1, tmVlf.iTermDate(0), tmVlf.iTermDate(1)
                    ilRet = btrUpdate(hmVlfPop, tmVlf, imVlfPopRecLen)
                    ilRet = btrGetNext(hmVlfPop, tmVlf, imVlfPopRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    gUnpackDateLong tmVlf.iEffDate(0), tmVlf.iEffDate(1), llEffDate
                    gUnpackDateLong tmVlf.iTermDate(0), tmVlf.iTermDate(1), llTermDate
                    If llTermDate = 0 Then
                        llTermDate = 999999
                    End If
                Loop
            End If
        Next ilLoop
        ' Add link records
        For ilCount = 0 To imUpdateUpB Step 1   'Set index for tmVlfPop
            If (tmVlfUpdate(ilCount).iSellCode > 0) And (tmVlfUpdate(ilCount).iAirCode > 0) Then
                On Error GoTo 0
                tmVlfUpdate(ilCount).lCode = 0
                ilRet = btrInsert(hmVlfPop, tmVlfUpdate(ilCount), imVlfPopRecLen, INDEXKEY0)
                'Check if duplicate record write is being requested
                If ilRet = 5 Then
                    Screen.MousePointer = vbDefault
                    gObtainVehicleName tmVlfUpdate(ilCount).iSellCode, slVefName, slType
                    MsgBox CStr("You have made a duplicate link within " & slVefName & " Vehicle ... Please correct the conflict before updating records"), 0, "Links Definition"
                    imUpdateFlag = False
                    ilRet = btrAbortTrans(hmVlfPop)
                    Exit Sub
                End If
                If ilRet <> BTRV_ERR_NONE Then
                    '6/6/16: Replaced GoSub
                    'GoSub mAbortWrite
                    mAbortWrite
                    Exit Sub
                End If
            End If
        Next ilCount
        ilRet = btrEndTrans(hmVlfPop)
        imUpdateFlag = False
        Exit Sub
    End If
    Erase tmVlfUpdate
    Exit Sub
mWriteVlfErr:
    On Error GoTo 0
    imUpdateFlag = False
    Erase tmVlfUpdate
    imTerminate = True
    Exit Sub
'mAbortWrite:
'    ilRet = btrAbortTrans(hmVlfPop)
'    imUpdateFlag = False
'    Erase tmVlfUpdate
'    Screen.MousePointer = vbDefault
'    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Links")
'    imTerminate = True
'    Return
End Sub
Private Sub plcNetworks_DragDrop(Source As control, X As Single, Y As Single)
    tmcScroll.Enabled = False
End Sub
Private Sub plcNetworks_DragOver(Source As control, X As Single, Y As Single, State As Integer)
    If (imDType = 2) Then
        tmcScroll.Enabled = False
        Exit Sub
    End If
    'Turn on scroll timer if in trigger box
    If (State = vbOver) Then
        If (smDSourceType = smLeaveType) And (imDType = 1) Then
            If (X >= imDLeft) And (X <= imDRight) Then
                If (Y <= imDBottom) And (Y >= imDTop) Then
                    If Not tmcScroll.Enabled Then
                        tmcScroll.Enabled = True
                        Exit Sub
                    End If
                ElseIf tmcScroll.Enabled Then
                    tmcScroll.Enabled = False
                    Exit Sub
                End If
            ElseIf tmcScroll.Enabled Then
                tmcScroll.Enabled = False
                Exit Sub
            End If
        ElseIf (smDSourceType <> smLeaveType) And (imDType = 0) Then
            If (X >= imDLeft) And (X <= imDRight) Then
                If (Y <= imDBottom) And (Y >= imDTop) Then
                    If Not tmcScroll.Enabled Then
                        tmcScroll.Enabled = True
                        Exit Sub
                    End If
                ElseIf tmcScroll.Enabled Then
                    tmcScroll.Enabled = False
                    Exit Sub
                End If
            ElseIf tmcScroll.Enabled Then
                tmcScroll.Enabled = False
                Exit Sub
            End If
        End If
    End If

End Sub
Private Sub plcNetworks_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmcScroll.Enabled = False
End Sub
'**********************************************************************
'
'       Control Name : tmcDrag (Timer)
'
'       Created : 4/24/94       By: D. Hannifan
'       Modified :              By:
'
'       Comments : Drag event timer
'
'
'**********************************************************************
'
Private Sub tmcDrag_Timer()
    tmcDrag.Enabled = False

    imSellClickIndex = -1
    imAirClickIndex = -1
    If imDragSource = 0 Then  'Selling
        If (imDType = 2) Then
            lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
        Else
            lbcSelling(imDragIndex).DragIcon = IconTraf!imcIconLink.DragIcon
        End If
        lbcSelling(imDragIndex).Drag vbBeginDrag
    Else   'Airing
        If (imDType = 2) Then
            lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconDrag.DragIcon
        Else
            lbcAiring(imDragIndex).DragIcon = IconTraf!imcIconLink.DragIcon
        End If
        lbcAiring(imDragIndex).Drag vbBeginDrag
    End If
End Sub
'**********************************************************************
'
'       Control Name : tmcScroll (Timer)
'
'       Created : 4/4/94       By: D. Hannifan
'       Modified :             By:
'
'       Comments : Procedure to scroll list boxes during a
'                  drag/swap/link event
'
'**********************************************************************
'
Private Sub tmcScroll_Timer()
    Dim ilMaxListItems As Integer   'Number of list items in target list box
    Dim ilIncrement As Integer      '1=Increase lbc.TopIndex  -1=Decrease lbc.TopIndex
    If (imDExitDirect = 0) Then
        ilIncrement = -1
    Else
        ilIncrement = 1
    End If
    If (smLeaveType = "S") Then    'Scroll a selling vehicle
        If (imLeaveIndex <= imNoSelling - 1) Then
            ilMaxListItems = lbcSelling(imLeaveIndex).ListCount - 1
        Else
            Exit Sub
        End If
    Else                            'Scroll an airing vehicle
        If (imLeaveIndex <= imNoAiring - 1) Then
            ilMaxListItems = lbcAiring(imLeaveIndex).ListCount - 1
        Else
            Exit Sub
        End If
    End If
    'Do a scroll Event
    If (smLeaveType = "S") And (ilIncrement = 1) Then
        If (lbcSelling(imLeaveIndex).TopIndex < ilMaxListItems - 7) Then
            lbcSelling(imLeaveIndex).TopIndex = lbcSelling(imLeaveIndex).TopIndex + ilIncrement
        End If
    End If
    If (smLeaveType = "S") And (ilIncrement = -1) Then
        If (lbcSelling(imLeaveIndex).TopIndex > 0) Then
            lbcSelling(imLeaveIndex).TopIndex = lbcSelling(imLeaveIndex).TopIndex + ilIncrement
        End If
    End If
    If (smLeaveType = "A") And (ilIncrement = 1) Then
        If (lbcAiring(imLeaveIndex).TopIndex < ilMaxListItems - 7) Then
            lbcAiring(imLeaveIndex).TopIndex = lbcAiring(imLeaveIndex).TopIndex + ilIncrement
        End If
    End If
    If (smLeaveType = "A") And (ilIncrement = -1) Then
        If (lbcAiring(imLeaveIndex).TopIndex > 0) Then
            lbcAiring(imLeaveIndex).TopIndex = lbcAiring(imLeaveIndex).TopIndex + ilIncrement
        End If
    End If
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print smScreenCaption
End Sub

Public Function mFillTo100(slInStr As String) As String
    Dim ilPos As Integer
    Dim slStr As String

    ilPos = InStr(1, slInStr, "@", vbTextCompare)
    If ilPos > 0 Then
        slStr = Left$(slInStr, ilPos - 1)
        While Len(slStr) < 105
            slStr = slStr & " "
        Wend
        mFillTo100 = slStr & Mid$(slInStr, ilPos)
    Else
        mFillTo100 = slInStr
    End If
End Function

Private Sub mAddSellItems(ilRealTime0 As Integer, ilRealTime1 As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, ilEndDate0 As Integer, ilEndDate1 As Integer)
    tmVlfPop(imUpperBound).iSellCode = imVefCode
    tmVlfPop(imUpperBound).iSellDay = imDateCode                '0=M-F etc
    tmVlfPop(imUpperBound).iSellTime(0) = ilRealTime0
    tmVlfPop(imUpperBound).iSellTime(1) = ilRealTime1
    tmVlfPop(imUpperBound).iSellPosNo = 0
    tmVlfPop(imUpperBound).iSellSeq = 0
    tmVlfPop(imUpperBound).sStatus = "P"
    tmVlfPop(imUpperBound).iAirCode = 0
    tmVlfPop(imUpperBound).iAirDay = imDateCode
    tmVlfPop(imUpperBound).iAirTime(0) = 0
    tmVlfPop(imUpperBound).iAirTime(1) = 0
    tmVlfPop(imUpperBound).iAirPosNo = 0
    tmVlfPop(imUpperBound).iAirSeq = 0
    tmVlfPop(imUpperBound).iEffDate(0) = ilLogDate0
    tmVlfPop(imUpperBound).iEffDate(1) = ilLogDate1
    tmVlfPop(imUpperBound).iTermDate(0) = ilEndDate0
    tmVlfPop(imUpperBound).iTermDate(1) = ilEndDate1
    tmVlfPop(imUpperBound).sDelete = ""

    imUpperBound = imUpperBound + 1  'inc upperbound of tmVlfPop
    If (imUpperBound > UBound(tmVlfPop)) Then
        ReDim Preserve tmVlfPop(0 To imUpperBound) As VLF 'Redim tmVLFPop
    End If
End Sub

Private Sub mAddAirItems(ilRealTime0 As Integer, ilRealTime1 As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, ilEndDate0 As Integer, ilEndDate1 As Integer)
    tmVlfPop(imUpperBound).iSellCode = 0
    tmVlfPop(imUpperBound).iSellDay = imDateCode     '0=M-F , 6=Sa, 7=Su
    tmVlfPop(imUpperBound).iSellTime(0) = 0
    tmVlfPop(imUpperBound).iSellTime(1) = 0
    tmVlfPop(imUpperBound).iSellPosNo = 0
    tmVlfPop(imUpperBound).iSellSeq = 0
    tmVlfPop(imUpperBound).sStatus = "P"
    tmVlfPop(imUpperBound).iAirCode = imVefCode
    tmVlfPop(imUpperBound).iAirDay = imDateCode
    tmVlfPop(imUpperBound).iAirTime(0) = ilRealTime0
    tmVlfPop(imUpperBound).iAirTime(1) = ilRealTime1
    tmVlfPop(imUpperBound).iAirPosNo = 0
    tmVlfPop(imUpperBound).iAirSeq = 0
    tmVlfPop(imUpperBound).iEffDate(0) = ilLogDate0
    tmVlfPop(imUpperBound).iEffDate(1) = ilLogDate1
    tmVlfPop(imUpperBound).iTermDate(0) = ilEndDate0
    tmVlfPop(imUpperBound).iTermDate(1) = ilEndDate1
    tmVlfPop(imUpperBound).sDelete = ""

    imUpperBound = imUpperBound + 1  'inc upperbound of tmVlfPop
    If (imUpperBound > UBound(tmVlfPop)) Then
        ReDim Preserve tmVlfPop(0 To imUpperBound) As VLF 'Redim tmVLFPop
    End If
End Sub

Private Sub mAddRecord(tlVlf As VLF)
    If UBound(tmVlfPop) < imUpperBound Then
        ReDim Preserve tmVlfPop(0 To imUpperBound)
    End If
    tmVlfPop(imUpperBound).iSellCode = tlVlf.iSellCode
    tmVlfPop(imUpperBound).iSellDay = tlVlf.iSellDay                '0=M-F etc
    tmVlfPop(imUpperBound).iSellTime(0) = tlVlf.iSellTime(0)
    tmVlfPop(imUpperBound).iSellTime(1) = tlVlf.iSellTime(1)
    tmVlfPop(imUpperBound).iSellPosNo = 0
    tmVlfPop(imUpperBound).iSellSeq = tlVlf.iSellSeq
    tmVlfPop(imUpperBound).sStatus = tlVlf.sStatus
    tmVlfPop(imUpperBound).iAirCode = tlVlf.iAirCode
    tmVlfPop(imUpperBound).iAirDay = tlVlf.iAirDay
    tmVlfPop(imUpperBound).iAirTime(0) = tlVlf.iAirTime(0)
    tmVlfPop(imUpperBound).iAirTime(1) = tlVlf.iAirTime(1)
    tmVlfPop(imUpperBound).iAirPosNo = 0
    tmVlfPop(imUpperBound).iAirSeq = tlVlf.iAirSeq
    tmVlfPop(imUpperBound).iEffDate(0) = tlVlf.iEffDate(0)
    tmVlfPop(imUpperBound).iEffDate(1) = tlVlf.iEffDate(1)
    tmVlfPop(imUpperBound).iTermDate(0) = tlVlf.iTermDate(0)
    tmVlfPop(imUpperBound).iTermDate(1) = tlVlf.iTermDate(1)
    tmVlfPop(imUpperBound).sDelete = ""
End Sub



Private Sub mAbortWrite()
    Dim ilRet As Integer
    ilRet = btrAbortTrans(hmVlfPop)
    imUpdateFlag = False
    Erase tmVlfUpdate
    Screen.MousePointer = vbDefault
    ilRet = MsgBox("Save Not Completed, Try Later", vbOKOnly + vbExclamation, "Links")
    imTerminate = True
End Sub

Private Sub mParseAirLink(slCurListString As String, slParsedString As String, ilCurAirRecNo As Integer, slCurAirVefName As String, slCurListType As String)
    Dim ilRet As Integer
    slCurListString = LTrim$(slCurListString)
    ilRet = gParseItem(slCurListString, 2, "@", slParsedString)
    ilCurAirRecNo = Val(Trim$(slParsedString))
    gObtainVehicleName tmVlfPop(ilCurAirRecNo).iAirCode, slCurAirVefName, slCurListType
End Sub

Private Sub mParseSellTime(slCurListString As String, slParsedString As String, ilCurSellRecNo As Integer, slCurSellVefName As String, slCurListType As String)
    Dim ilRet As Integer
    slCurListString = LTrim$(slCurListString)
    ilRet = gParseItem(slCurListString, 2, "@", slParsedString)
    ilCurSellRecNo = Val(Trim$(slParsedString))
    gObtainVehicleName tmVlfPop(ilCurSellRecNo).iSellCode, slCurSellVefName, slCurListType

End Sub

Private Sub mCreateLink(ilCurSellRecNo As Integer, ilSellSeqNo As Integer, slLinksDefStatus As String, slCurListType As String, ilCurAirRecNo As Integer, ilAirSeqNo As Integer)
    Dim ilSellIndex As Integer
    
    tmVlfUpdate(imUpdateUpB).iSellCode = tmVlfPop(ilCurSellRecNo).iSellCode
    tmVlfUpdate(imUpdateUpB).iSellDay = tmVlfPop(ilCurSellRecNo).iSellDay
    tmVlfUpdate(imUpdateUpB).iSellTime(0) = tmVlfPop(ilCurSellRecNo).iSellTime(0)
    tmVlfUpdate(imUpdateUpB).iSellTime(1) = tmVlfPop(ilCurSellRecNo).iSellTime(1)
    tmVlfUpdate(imUpdateUpB).iSellPosNo = 0
    tmVlfUpdate(imUpdateUpB).iSellSeq = ilSellSeqNo
    slLinksDefStatus = ""
    For ilSellIndex = 0 To imNoSelling - 1 Step 1
        If Abs(imSellPending(ilSellIndex)) = tmVlfUpdate(imUpdateUpB).iSellCode Then
            If imSellPending(ilSellIndex) < 0 Then
                slLinksDefStatus = "P"
            Else
                slLinksDefStatus = "C"
            End If
        End If
    Next ilSellIndex
    If slLinksDefStatus = "C" Then
        tmVlfUpdate(imUpdateUpB).sStatus = "P"
    ElseIf slCurListType = "S" Then
        tmVlfUpdate(imUpdateUpB).sStatus = tmVlfPop(ilCurSellRecNo).sStatus
    ElseIf slCurListType = "A" Then
        tmVlfUpdate(imUpdateUpB).sStatus = tmVlfPop(ilCurAirRecNo).sStatus
    End If
    tmVlfUpdate(imUpdateUpB).iAirCode = tmVlfPop(ilCurAirRecNo).iAirCode
    tmVlfUpdate(imUpdateUpB).iAirDay = tmVlfPop(ilCurAirRecNo).iAirDay
    tmVlfUpdate(imUpdateUpB).iAirTime(0) = tmVlfPop(ilCurAirRecNo).iAirTime(0)
    tmVlfUpdate(imUpdateUpB).iAirTime(1) = tmVlfPop(ilCurAirRecNo).iAirTime(1)
    tmVlfUpdate(imUpdateUpB).iAirPosNo = 0
    tmVlfUpdate(imUpdateUpB).iAirSeq = ilAirSeqNo
    tmVlfUpdate(imUpdateUpB).iEffDate(0) = tmVlfPop(ilCurSellRecNo).iEffDate(0)
    tmVlfUpdate(imUpdateUpB).iEffDate(1) = tmVlfPop(ilCurSellRecNo).iEffDate(1)
    tmVlfUpdate(imUpdateUpB).iTermDate(0) = tmVlfPop(ilCurSellRecNo).iTermDate(0)
    tmVlfUpdate(imUpdateUpB).iTermDate(1) = tmVlfPop(ilCurSellRecNo).iTermDate(1)
    tmVlfUpdate(imUpdateUpB).sDelete = tmVlfPop(ilCurSellRecNo).sDelete
    imUpdateUpB = imUpdateUpB + 1
    ReDim Preserve tmVlfUpdate(imUpdateUpB) As VLF
End Sub

Private Sub mCheckSeqNo(ilFoundAir As Integer, ilAirSeqSet As Integer, slCurAirVefName As String, slCurListString As String, ilAirSeqNo As Integer, slSrchString As String, ilCurAirRecNo As Integer)
    Dim ilCount As Integer
    Dim slTimeStr1 As String
    Dim slTimeStr2 As String
    Dim ilRet As Integer
    Dim ilCount2 As Integer
    Dim ilCount3 As Integer
    Dim slStr As String
    
    ilFoundAir = False
    ilAirSeqSet = False
    For ilCount = 0 To imNoAiring - 1 Step 1
        If (Trim$(lacAiring(ilCount).Caption) = Trim$(slCurAirVefName)) Then
            ilFoundAir = True
            Exit For
        End If
    Next ilCount
    If ilFoundAir Then
        ilRet = gParseItem(LTrim$(slCurListString), 1, " ", slTimeStr1)
        slTimeStr1 = Trim$(slTimeStr1)
        For ilCount2 = 0 To lbcAiring(ilCount).ListCount - 1 Step 1
            If (Left$(lbcAiring(ilCount).List(ilCount2), 2) <> "  ") Then
                ilRet = gParseItem(lbcAiring(ilCount).List(ilCount2), 1, " ", slTimeStr2)
                slTimeStr2 = Trim$(slTimeStr2)
                If (slTimeStr1 = slTimeStr2) Then
                    ilAirSeqNo = 1
                    For ilCount3 = ilCount2 + 1 To lbcAiring(ilCount).ListCount - 1 Step 1
                        If (Left$(lbcAiring(ilCount).List(ilCount3), 2) = "  ") Then
                            ilRet = gParseItem(Trim$(lbcAiring(ilCount).List(ilCount3)), 1, "@", slStr)
                            slStr = Trim$(slStr)
                            If (slStr = slSrchString) Then
                                ilAirSeqSet = True
                                Exit For
                            Else
                                ilAirSeqNo = ilAirSeqNo + 1
                            End If
                        Else
                            ilAirSeqSet = True
                            Exit For
                        End If
                    Next ilCount3
                End If
                If ilAirSeqSet Then
                    Exit For
                End If
            End If
        Next ilCount2
    Else
        ilAirSeqNo = tmVlfPop(ilCurAirRecNo).iAirSeq
        ilAirSeqSet = True
    End If
End Sub
