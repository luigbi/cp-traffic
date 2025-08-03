VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ViewList 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3540
   ClientLeft      =   2310
   ClientTop       =   1980
   ClientWidth     =   9240
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   9240
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
      TabIndex        =   1
      Top             =   1770
      Width           =   75
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Height          =   285
      Left            =   4200
      TabIndex        =   0
      Top             =   3195
      Width           =   945
   End
   Begin ComctlLib.ListView lbcView 
      Height          =   2820
      Left            =   195
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   270
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   4974
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Highlight"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Vehicle Name"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Dates"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Code"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lacScreen 
      Caption         =   "Dates Not Marked Completed"
      Height          =   195
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   5820
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   75
      Top             =   3165
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "ViewList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ViewList.Frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  imPopReqd                                                                             *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ViewList.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice number input screen code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim imFirstActivate As Integer
Dim imTerminate As Integer  'True = terminating task, False= OK
Private Sub cmcDone_Click()
    mTerminate
End Sub
Private Sub cmcDone_GotFocus()
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
'    gShowBranner
    Me.KeyPreview = True
    ViewList.Refresh
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

Private Sub Form_Load()
    mInit
    If imTerminate Then
        'cmcCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ViewList = Nothing   'Remove data segment
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
    imTerminate = False
    imFirstActivate = True


    Screen.MousePointer = vbHourglass
    ViewList.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    gCenterStdAlone ViewList
    'ViewList.Show
    Screen.MousePointer = vbHourglass
    mSetColumnWidth
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Exit Sub

    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mPopulateErr                                                                          *
'******************************************************************************************

'
'   mPopulate
'   Where:
'
    Dim mItem As ListItem
    Dim llLoop As Long
    Dim slVehName As String
    Dim slStr As String

    'llRet = SendMessageByNum(lbcView.hwnd, LV_SETEXTENDEDLISTVIEWSTYLE, 0, LV_FULLROWSSELECT)
    slVehName = ""
    For llLoop = LBound(tgNotMarkComplete) To UBound(tgNotMarkComplete) - 1 Step 1
        Set mItem = lbcView.ListItems.Add()
        If StrComp(Trim$(tgNotMarkComplete(llLoop).sVehName), slVehName, vbTextCompare) <> 0 Then
            mItem.Text = ""
            mItem.SubItems(1) = Trim$(tgNotMarkComplete(llLoop).sVehName)
            slVehName = Trim$(tgNotMarkComplete(llLoop).sVehName)
        End If
        If tgNotMarkComplete(llLoop).iGameNo > 0 Then
            slStr = "Event #: " & Trim$(str$(tgNotMarkComplete(llLoop).iGameNo)) & " @ "
        Else
            slStr = ""
        End If
        slStr = slStr & tgNotMarkComplete(llLoop).sDate
        mItem.SubItems(2) = slStr
        mItem.SubItems(3) = llLoop
    Next llLoop
    Exit Sub
mPopulateErr: 'VBC NR
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
    Unload ViewList
    igManUnload = NO
End Sub

Private Sub mSetColumnWidth()
    Dim ilNoColumns As Integer
    Dim ilCol As Integer
    Dim llWidth As Long


    ilNoColumns = 4
    lbcView.ColumnHeaders.Item(1).Width = 0
    lbcView.ColumnHeaders.Item(2).Width = (2 * lbcView.Width) / 3  'Vehicle
    lbcView.ColumnHeaders.Item(4).Width = 0
    For ilCol = 1 To 4 Step 1
        If ilCol <> 3 Then
            llWidth = llWidth + lbcView.ColumnHeaders.Item(ilCol).Width
        End If
    Next ilCol
    '150 was used to get scroll with correct
    lbcView.ColumnHeaders.Item(3).Width = lbcView.Width - llWidth - GRIDSCROLLWIDTH - ilNoColumns * 120 - 150
End Sub

Private Sub lbcView_Click()
    If lbcView.SelectedItem.Index >= 1 Then
        If lbcView.ListItems.Item(lbcView.SelectedItem.Index).Selected Then
            lbcView.ListItems.Item(lbcView.SelectedItem.Index).Selected = False
        End If
    End If
End Sub
