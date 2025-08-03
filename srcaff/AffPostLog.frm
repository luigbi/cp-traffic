VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPostLog 
   Caption         =   "Post Log"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   Icon            =   "AffPostLog.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   9180
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   75
      ScaleHeight     =   15
      ScaleWidth      =   0
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6195
      Width           =   0
   End
   Begin VB.ListBox lbcStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffPostLog.frx":08CA
      Left            =   7185
      List            =   "AffPostLog.frx":08CC
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2655
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCart 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      Index           =   3
      ItemData        =   "AffPostLog.frx":08CE
      Left            =   5280
      List            =   "AffPostLog.frx":08D0
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3525
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCart 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      Index           =   2
      ItemData        =   "AffPostLog.frx":08D2
      Left            =   4545
      List            =   "AffPostLog.frx":08D4
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3210
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCart 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      Index           =   1
      ItemData        =   "AffPostLog.frx":08D6
      Left            =   4110
      List            =   "AffPostLog.frx":08D8
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3105
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCart 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      Index           =   0
      ItemData        =   "AffPostLog.frx":08DA
      Left            =   3660
      List            =   "AffPostLog.frx":08DC
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2865
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.ListBox lbcCntr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffPostLog.frx":08DE
      Left            =   5970
      List            =   "AffPostLog.frx":08E0
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1890
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   105
      Picture         =   "AffPostLog.frx":08E2
      ScaleHeight     =   165
      ScaleWidth      =   90
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1515
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox pbcPostFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   75
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   15
      Top             =   840
      Width           =   60
   End
   Begin VB.PictureBox pbcPostTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000F&
      Height          =   90
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   27
      Top             =   5175
      Width           =   60
   End
   Begin VB.PictureBox pbcPostSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   105
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   17
      Top             =   1275
      Width           =   60
   End
   Begin VB.ListBox lbcAdvt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   810
      ItemData        =   "AffPostLog.frx":0BEC
      Left            =   3735
      List            =   "AffPostLog.frx":0BEE
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox txtDropdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   4080
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmcDropDown 
      Appearance      =   0  'Flat
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   5.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5025
      Picture         =   "AffPostLog.frx":0BF0
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2490
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.OptionButton rbcVeh 
      Caption         =   "All Veh."
      Height          =   240
      Index           =   1
      Left            =   6855
      TabIndex        =   35
      Top             =   870
      Width           =   1005
   End
   Begin VB.OptionButton rbcVeh 
      Caption         =   "Active Veh."
      Height          =   240
      Index           =   0
      Left            =   5595
      TabIndex        =   34
      Top             =   870
      Value           =   -1  'True
      Width           =   1290
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6930
      TabIndex        =   31
      Top             =   5835
      Width           =   1335
   End
   Begin VB.Frame fraShow 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   225
      Left            =   135
      TabIndex        =   8
      Top             =   930
      Width           =   4770
      Begin VB.OptionButton optShow 
         Caption         =   "PST"
         Height          =   195
         Index           =   3
         Left            =   3765
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   660
      End
      Begin VB.OptionButton optShow 
         Caption         =   "CST"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   720
      End
      Begin VB.OptionButton optShow 
         Caption         =   "EST"
         Height          =   195
         Index           =   0
         Left            =   1575
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton optShow 
         Caption         =   "MST"
         Height          =   195
         Index           =   2
         Left            =   3000
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Show Date/Time by"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.ComboBox cboSort 
         Height          =   315
         ItemData        =   "AffPostLog.frx":0CEA
         Left            =   135
         List            =   "AffPostLog.frx":0CEC
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   255
         Width           =   3240
      End
      Begin VB.TextBox txtWeek 
         Height          =   285
         Left            =   7395
         TabIndex        =   7
         Top             =   255
         Width           =   1110
      End
      Begin VB.CheckBox chkZone 
         Caption         =   "PST"
         Height          =   225
         Index           =   3
         Left            =   5865
         TabIndex        =   5
         Top             =   300
         Value           =   1  'Checked
         Width           =   780
      End
      Begin VB.CheckBox chkZone 
         Caption         =   "MST"
         Height          =   225
         Index           =   2
         Left            =   5085
         TabIndex        =   4
         Top             =   300
         Value           =   1  'Checked
         Width           =   720
      End
      Begin VB.CheckBox chkZone 
         Caption         =   "CST"
         Height          =   225
         Index           =   1
         Left            =   4320
         TabIndex        =   3
         Top             =   300
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chkZone 
         Caption         =   "EST"
         Height          =   225
         Index           =   0
         Left            =   3540
         TabIndex        =   2
         Top             =   300
         Value           =   1  'Checked
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   195
         Left            =   6810
         TabIndex        =   6
         Top             =   300
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6945
      Top             =   6150
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6255
      FormDesignWidth =   9180
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   6180
      TabIndex        =   29
      Top             =   5370
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPost 
      Height          =   3915
      Left            =   210
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1230
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6906
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   16
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7665
      TabIndex        =   30
      Top             =   5370
      Width           =   1335
   End
   Begin VB.Label lacGame 
      Alignment       =   2  'Center
      Height          =   195
      Left            =   630
      TabIndex        =   37
      Top             =   690
      Width           =   7740
   End
   Begin VB.Label Label5 
      Caption         =   "To 'Select': Click on Date or Time in Yellow area"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   810
      TabIndex        =   36
      Top             =   5400
      Width           =   5295
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   240
      Picture         =   "AffPostLog.frx":0CEE
      Top             =   5385
      Width           =   480
   End
   Begin VB.Image imcPrt 
      Height          =   480
      Left            =   8565
      Picture         =   "AffPostLog.frx":15B8
      Top             =   735
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "To Add Spot: 'Select' to insert after, then Click on Icon"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   810
      TabIndex        =   33
      Top             =   5835
      Width           =   5295
   End
   Begin VB.Label Label3 
      Caption         =   "To Swap Spots: 'Select' Spot, then 'Select' Spot to swap with"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   810
      TabIndex        =   32
      Top             =   5625
      Width           =   5250
   End
End
Attribute VB_Name = "frmPostLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmPostLog - Post Log (Allow spots to be swapped and replaced)
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imVefCode As Integer
Private imIntegralSet As Integer
Private imInChg As Integer
Private imBSMode As Integer
Private smFWkDate As String    'First week start start
Private smLWkDate As String  'Last week end date
Private imHeaderClick As Integer
Private imNoZones As Integer
Private imFromGetLst As Integer
Private imFillCntrAdfCode As Integer
Private imFillCartAdfCode As Integer
Private imMouseDown As Integer
Private imInFillCntr As Integer
Private imInFillCart As Integer
Private imFirstTime As Integer
Private imFieldChgd As Integer
Private imIgnoreChg As Integer
Private imInRowNo As Integer
Private tmPostInfo() As POSTINFO
Private tmFromPostInfo As POSTINFO
Private smFromCols(0 To 15) As String
Private smToCols(0 To 15) As String
Private tmToPostInfo As POSTINFO
Private tmCntrInfo() As CNTRINFO
Private tmCopyInfo() As COPYINFO

Private tmStatusTypes(0 To 14) As STATUSTYPES

Private lmSwapStartRow As Long
Private imSwapClickCount As Integer

Private bmCreateAbfRecord As Boolean
Private lmAbfDate() As Long
Private imAbfVefCode As Integer

Private rst_Lst As ADODB.Recordset
Private rst_chf As ADODB.Recordset
Private rst_Cif As ADODB.Recordset
Private rst_Gsf As ADODB.Recordset

'Grid Controls
Private imShowGridBox As Integer
Private imFromArrow As Integer
Private lmTopRow As Long            'Top row when cell clicked or - 1
Private lmEnableRow As Long         'Current or last row focus was on
Private lmEnableCol As Long         'Current or last column focus was on

Private lmMaxCol As Long
Const DATEINDEX = 0
Const TIMEINDEX = 1
Const ADVTINDEX = 3
Const CNTRNOINDEX = 4
Const LENINDEX = 5
Const CARTESTINDEX = 6
Const CARTCSTINDEX = 7
Const CARTMSTINDEX = 8
Const CARTPSTINDEX = 9
Const STATUSINDEX = 2
Const LSTCODEESTINDEX = 10
Const LSTCODECSTINDEX = 11
Const LSTCODEMSTINDEX = 12
Const LSTCODEPSTINDEX = 13
Const POSTINDEX = 14
Const ORIGTYPEINDEX = 15


Private Function mPostColAllowed(llCol As Long) As Integer
    Dim ilIndex As Integer
    Dim slStr As String
    Dim llRow As Long
    Dim ilLoop As Integer
    
    slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
    llRow = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
'    If llRow >= 0 Then
'        ilIndex = lbcStatus.ItemData(llRow)
'        If (tmStatusTypes(ilIndex).iPledged = 0) Then   'Live
'            mPostColAllowed = True
'        ElseIf (tmStatusTypes(ilIndex).iPledged = 1) Then   'Delayed
'            mPostColAllowed = True
'        ElseIf (tmStatusTypes(ilIndex).iPledged = 2) Then   'Not Aired
'            mPostColAllowed = True
'        ElseIf (tmStatusTypes(ilIndex).iPledged = 3) Then   'No Pledged Times
'            mPostColAllowed = True
'        End If
'    Else
'        slStr = grdPost.TextMatrix(grdPost.Row, ADVTINDEX)
'        llRow = SendMessageByString(lbcAdvt.hwnd, LB_FINDSTRING, -1, slStr)
'        If llRow >= 0 Then
'            slStr = grdPost.TextMatrix(grdPost.Row, CNTRNOINDEX)
'            llRow = SendMessageByString(lbcCntr.hwnd, LB_FINDSTRING, -1, slStr)
'            If llRow >= 0 Then
'                mPostColAllowed = True
'            Else
'                If llCol <= CNTRNOINDEX Then
'                    mPostColAllowed = True
'                Else
'                    mPostColAllowed = False
'                End If
'            End If
'        Else
'            If llCol = ADVTINDEX Then
'                mPostColAllowed = True
'            Else
'                mPostColAllowed = False
'            End If
'        End If
'    End If
    If llRow >= 0 Then
        slStr = grdPost.TextMatrix(grdPost.Row, ADVTINDEX)
        llRow = SendMessageByString(lbcAdvt.hwnd, LB_FINDSTRING, -1, slStr)
        If llRow >= 0 Then
            slStr = grdPost.TextMatrix(grdPost.Row, CNTRNOINDEX)
            'llRow = SendMessageByString(lbcCntr.hwnd, LB_FINDSTRING, -1, slStr)
            'If llRow >= 0 Then
            If Trim$(slStr) <> "" Then
                mPostColAllowed = True
            Else
                If llCol <= CNTRNOINDEX Then
                    mPostColAllowed = True
                Else
                    mPostColAllowed = False
                End If
            End If
        Else
            If llCol = ADVTINDEX Then
                mPostColAllowed = True
            Else
                mPostColAllowed = False
            End If
        End If
    Else
        If llCol = STATUSINDEX Then
            mPostColAllowed = True
        Else
            mPostColAllowed = False
        End If
    End If
End Function

Private Sub mPostEnableBox()
    Dim ilIndex As Integer
    Dim slStr As String
    Dim ilType As Integer
    
    If (grdPost.Row >= grdPost.FixedRows) And (grdPost.Row < grdPost.Rows) And (grdPost.Col >= DATEINDEX) And (grdPost.Col < grdPost.Cols - 1) Then
        lmEnableRow = grdPost.Row
        lmEnableCol = grdPost.Col
        imShowGridBox = True
        pbcArrow.Move grdPost.Left - pbcArrow.Width, grdPost.Top + grdPost.RowPos(grdPost.Row) + (grdPost.RowHeight(grdPost.Row) - pbcArrow.Height) / 2
        pbcArrow.Visible = True
        Select Case grdPost.Col
            Case ADVTINDEX  'Advertiser
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - cmcDropDown.Width - 30, grdPost.RowHeight(grdPost.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcAdvt.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcAdvt, 4
                slStr = grdPost.Text
                ilIndex = SendMessageByString(lbcAdvt.hwnd, LB_FINDSTRING, -1, slStr)
                If ilIndex >= 0 Then
                    lbcAdvt.ListIndex = ilIndex
                    txtDropdown.Text = lbcAdvt.List(lbcAdvt.ListIndex)
                Else
                    lbcAdvt.ListIndex = -1
                    txtDropdown.Text = ""
                End If
                If txtDropdown.Height > grdPost.RowHeight(grdPost.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdPost.RowHeight(grdPost.Row) - 15
                End If
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcAdvt.Visible = True
                txtDropdown.SetFocus
            Case CNTRNOINDEX  'Contract
                mFillCntr grdPost.Row
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - cmcDropDown.Width - 30, grdPost.RowHeight(grdPost.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcCntr.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + cmcDropDown.Width
                gSetListBoxHeight lbcCntr, 4
                slStr = grdPost.Text
                ilIndex = SendMessageByString(lbcCntr.hwnd, LB_FINDSTRING, -1, slStr)
                If ilIndex >= 0 Then
                    lbcCntr.ListIndex = ilIndex
                    txtDropdown.Text = lbcCntr.List(lbcCntr.ListIndex)
                Else
                    lbcCntr.ListIndex = -1
                    txtDropdown.Text = ""
                End If
                If txtDropdown.Height > grdPost.RowHeight(grdPost.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdPost.RowHeight(grdPost.Row) - 15
                End If
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcCntr.Visible = True
                txtDropdown.SetFocus
            Case LENINDEX  'Length
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - 30, grdPost.RowHeight(grdPost.Row) - 15
                txtDropdown.Text = grdPost.Text
                If txtDropdown.Height > grdPost.RowHeight(grdPost.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdPost.RowHeight(grdPost.Row) - 15
                End If
                txtDropdown.Visible = True
                txtDropdown.SetFocus
            Case CARTESTINDEX To CARTPSTINDEX '5, 6, 7, 8
                mFillCart grdPost.Row
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - cmcDropDown.Width - 30, grdPost.RowHeight(grdPost.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcCart(grdPost.Col - CARTESTINDEX).Move txtDropdown.Left - 3 * txtDropdown.Width, txtDropdown.Top + txtDropdown.Height, 4 * txtDropdown.Width + cmcDropDown.Width ' + grdPost.ColWidth(STATUSINDEX)
                gSetListBoxHeight lbcCart(grdPost.Col - CARTESTINDEX), 4
                slStr = grdPost.Text
                ilIndex = SendMessageByString(lbcCart(grdPost.Col - CARTESTINDEX).hwnd, LB_FINDSTRING, -1, slStr)
                If ilIndex >= 0 Then
                    lbcCart(grdPost.Col - CARTESTINDEX).ListIndex = ilIndex
                    txtDropdown.Text = lbcCart(grdPost.Col - CARTESTINDEX).List(lbcCart(grdPost.Col - CARTESTINDEX).ListIndex)
                Else
                    lbcCart(grdPost.Col - CARTESTINDEX).ListIndex = -1
                    txtDropdown.Text = ""
                End If
                If txtDropdown.Height > grdPost.RowHeight(grdPost.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdPost.RowHeight(grdPost.Row) - 15
                End If
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcCart(grdPost.Col - CARTESTINDEX).Visible = True
                txtDropdown.SetFocus
            Case STATUSINDEX
                'txtDropdown.Move grdPost.Left + imColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - 30, grdPost.RowHeight(grdPost.Row) - 15
                txtDropdown.Move grdPost.Left + grdPost.ColPos(STATUSINDEX) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(STATUSINDEX) - cmcDropDown.Width - 30, grdPost.RowHeight(grdPost.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcStatus.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + (7 * txtDropdown.Width) \ 2
                gSetListBoxHeight lbcStatus, 4
                slStr = grdPost.TextMatrix(grdPost.Row, STATUSINDEX)
                ilIndex = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
                If ilIndex >= 0 Then
                    lbcStatus.ListIndex = ilIndex
                Else
                    lbcStatus.ListIndex = 0
                End If
                txtDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
                If txtDropdown.Height > grdPost.RowHeight(grdPost.Row) - 15 Then
                    txtDropdown.FontName = "Arial"
                    txtDropdown.Height = grdPost.RowHeight(grdPost.Row) - 15
                End If
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcStatus.Visible = True
                txtDropdown.SetFocus
        End Select
    End If
End Sub


Private Sub mPostSetShow()
    Dim slStr As String
    Dim ilIndex As Integer
    Dim ilType As Integer
    
    If (lmEnableRow >= grdPost.FixedRows) And (lmEnableRow < grdPost.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case ADVTINDEX
                slStr = txtDropdown.Text
                If grdPost.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    imFieldChgd = True
                    grdPost.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    'ilIndex = lmEnableRow - grdPost.FixedRows
                    'If (tmPostInfo(ilIndex).iType = 3) Or (tmPostInfo(ilIndex).iType = 1) Then
                    ilIndex = Val(grdPost.TextMatrix(lmEnableRow, POSTINDEX))
                    If (tmPostInfo(ilIndex).iType = 3) Then
                        mAdjPositionNo
                    End If
                    grdPost.TextMatrix(lmEnableRow, CNTRNOINDEX) = ""
                    grdPost.TextMatrix(lmEnableRow, CARTESTINDEX) = ""
                    grdPost.TextMatrix(lmEnableRow, CARTCSTINDEX) = ""
                    grdPost.TextMatrix(lmEnableRow, CARTMSTINDEX) = ""
                    grdPost.TextMatrix(lmEnableRow, CARTPSTINDEX) = ""
                End If
            Case CNTRNOINDEX
                slStr = txtDropdown.Text
                If grdPost.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    imFieldChgd = True
                    grdPost.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                End If
            Case LENINDEX
                slStr = txtDropdown.Text
                If grdPost.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    imFieldChgd = True
                    grdPost.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                    'Update avail if new spot
                    ilIndex = Val(grdPost.TextMatrix(lmEnableRow, POSTINDEX))
                    If (tmPostInfo(ilIndex).iType = 2) Or (tmPostInfo(ilIndex).iType = 3) Then
                        mAdjUnit_Sec
                    End If
                End If
            Case CARTESTINDEX To CARTPSTINDEX
                slStr = txtDropdown.Text
                If grdPost.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    imFieldChgd = True
                    grdPost.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                End If
            Case STATUSINDEX
                slStr = txtDropdown.Text
                If grdPost.TextMatrix(lmEnableRow, lmEnableCol) <> slStr Then
                    imFieldChgd = True
                    grdPost.TextMatrix(lmEnableRow, lmEnableCol) = slStr
                End If
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imShowGridBox = False
    pbcArrow.Visible = False
    txtDropdown.Visible = False
    cmcDropDown.Visible = False
    lbcAdvt.Visible = False
    lbcCntr.Visible = False
    lbcCart(0).Visible = False
    lbcCart(1).Visible = False
    lbcCart(2).Visible = False
    lbcCart(3).Visible = False
    lbcStatus.Visible = False
    If imFieldChgd Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
End Sub








Private Sub mFillAdvt()
    Dim iNoWeeks As Integer
    Dim dLWeek As Date
    Dim dFWeek As Date
    Dim iFound As Integer
    Dim iLoop As Integer
    On Error GoTo ErrHand
    
    lbcAdvt.Clear
    For iLoop = 0 To UBound(tgAdvtInfo) - 1 Step 1
        lbcAdvt.AddItem Trim$(tgAdvtInfo(iLoop).sAdvtName)
        lbcAdvt.ItemData(lbcAdvt.NewIndex) = tgAdvtInfo(iLoop).iCode
    Next iLoop
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmPostLog-mFillAdvt"
End Sub

Private Sub mGridPaint(iNew As Integer)
    Dim iLoop As Integer
    Dim sDate As String
    Dim sTime As String
    Dim sStatus As String
    Dim slCntrProd As String
    Dim iTRow As Integer
    Dim llTRow As Long
    Dim llRow As Long
    
    llTRow = grdPost.TopRow
    grdPost.Redraw = False
    If iNew Then
        mClearGrid
    End If
    llRow = grdPost.FixedRows
    For iLoop = 0 To UBound(tmPostInfo) - 1 Step 1
        If optShow(1).Value Then
            sDate = tmPostInfo(iLoop).sDateZone(1)
            sTime = tmPostInfo(iLoop).sTimeZone(1)
        ElseIf optShow(2).Value Then
            sDate = tmPostInfo(iLoop).sDateZone(2)
            sTime = tmPostInfo(iLoop).sTimeZone(2)
        ElseIf optShow(3).Value Then
            sDate = tmPostInfo(iLoop).sDateZone(3)
            sTime = tmPostInfo(iLoop).sTimeZone(3)
        Else
            sDate = tmPostInfo(iLoop).sDateZone(0)
            sTime = tmPostInfo(iLoop).sTimeZone(0)
        End If
        If tmPostInfo(iLoop).iType = 0 Then
            If tmPostInfo(iLoop).iStatus < ASTEXTENDED_MG Then
                sStatus = Trim$(tmStatusTypes(tmPostInfo(iLoop).iStatus).sName)
            ElseIf tmPostInfo(iLoop).iStatus = ASTEXTENDED_MG Then
                sStatus = "MG"
            ElseIf tmPostInfo(iLoop).iStatus = ASTEXTENDED_BONUS Then
                sStatus = "Bonus"
            Else
                sStatus = ""
            End If
        Else
            sStatus = "Avail"
        End If
        If llRow + 1 > grdPost.Rows Then
            grdPost.AddItem ""
        End If
        grdPost.Row = llRow
        grdPost.Col = DATEINDEX
        grdPost.CellBackColor = LIGHTYELLOW
        grdPost.Col = TIMEINDEX
        grdPost.CellBackColor = LIGHTYELLOW
        If iNew Then
            If tmPostInfo(iLoop).iType = 0 Then
                If tmPostInfo(iLoop).lCntrNo > 0 Then
                    'grdPost.Text = sDate & vbTab & sTime & vbTab & Trim$(tmPostInfo(iLoop).sAdfName) & vbTab & tmPostInfo(iLoop).lCntrNo & " " & Trim$(tmPostInfo(iLoop).sProd) & vbTab & tmPostInfo(iLoop).iLen & vbTab & Trim$(tmPostInfo(iLoop).sCartZone(0)) & vbTab & Trim$(tmPostInfo(iLoop).sCartZone(1)) & vbTab & Trim$(tmPostInfo(iLoop).sCartZone(2)) & vbTab & Trim$(tmPostInfo(iLoop).sCartZone(3)) & vbTab & sStatus & vbTab & tmPostInfo(iLoop).lLstCodeZone(0) & vbTab & tmPostInfo(iLoop).lLstCodeZone(1) & vbTab & tmPostInfo(iLoop).lLstCodeZone(2) & vbTab & tmPostInfo(iLoop).lLstCodeZone(3) & vbTab & iLoop
                    grdPost.TextMatrix(llRow, DATEINDEX) = sDate
                    grdPost.TextMatrix(llRow, TIMEINDEX) = sTime
                    grdPost.TextMatrix(llRow, ADVTINDEX) = Trim$(tmPostInfo(iLoop).sAdfName)
                    grdPost.TextMatrix(llRow, CNTRNOINDEX) = tmPostInfo(iLoop).lCntrNo & " " & Trim$(tmPostInfo(iLoop).sProd)
                    grdPost.TextMatrix(llRow, LENINDEX) = tmPostInfo(iLoop).iLen
                    grdPost.TextMatrix(llRow, CARTESTINDEX) = Trim$(tmPostInfo(iLoop).sCartZone(0))
                    grdPost.TextMatrix(llRow, CARTCSTINDEX) = Trim$(tmPostInfo(iLoop).sCartZone(1))
                    grdPost.TextMatrix(llRow, CARTMSTINDEX) = Trim$(tmPostInfo(iLoop).sCartZone(2))
                    grdPost.TextMatrix(llRow, CARTPSTINDEX) = Trim$(tmPostInfo(iLoop).sCartZone(3))
                    grdPost.TextMatrix(llRow, STATUSINDEX) = sStatus
                    grdPost.TextMatrix(llRow, LSTCODEESTINDEX) = tmPostInfo(iLoop).lLstCodeZone(0)
                    grdPost.TextMatrix(llRow, LSTCODECSTINDEX) = tmPostInfo(iLoop).lLstCodeZone(1)
                    grdPost.TextMatrix(llRow, LSTCODEMSTINDEX) = tmPostInfo(iLoop).lLstCodeZone(2)
                    grdPost.TextMatrix(llRow, LSTCODEPSTINDEX) = tmPostInfo(iLoop).lLstCodeZone(3)
                    grdPost.TextMatrix(llRow, POSTINDEX) = iLoop
                    grdPost.TextMatrix(llRow, ORIGTYPEINDEX) = tmPostInfo(iLoop).iType
                Else
                    grdPost.TextMatrix(llRow, DATEINDEX) = sDate
                    grdPost.TextMatrix(llRow, TIMEINDEX) = sTime
                    grdPost.TextMatrix(llRow, ADVTINDEX) = Trim$(tmPostInfo(iLoop).sAdfName)
                    grdPost.TextMatrix(llRow, CNTRNOINDEX) = Trim$(tmPostInfo(iLoop).sProd)
                    grdPost.TextMatrix(llRow, LENINDEX) = tmPostInfo(iLoop).iLen
                    grdPost.TextMatrix(llRow, CARTESTINDEX) = Trim$(tmPostInfo(iLoop).sCartZone(0))
                    grdPost.TextMatrix(llRow, CARTCSTINDEX) = Trim$(tmPostInfo(iLoop).sCartZone(1))
                    grdPost.TextMatrix(llRow, CARTMSTINDEX) = Trim$(tmPostInfo(iLoop).sCartZone(2))
                    grdPost.TextMatrix(llRow, CARTPSTINDEX) = Trim$(tmPostInfo(iLoop).sCartZone(3))
                    grdPost.TextMatrix(llRow, STATUSINDEX) = sStatus
                    grdPost.TextMatrix(llRow, LSTCODEESTINDEX) = tmPostInfo(iLoop).lLstCodeZone(0)
                    grdPost.TextMatrix(llRow, LSTCODECSTINDEX) = tmPostInfo(iLoop).lLstCodeZone(1)
                    grdPost.TextMatrix(llRow, LSTCODEMSTINDEX) = tmPostInfo(iLoop).lLstCodeZone(2)
                    grdPost.TextMatrix(llRow, LSTCODEPSTINDEX) = tmPostInfo(iLoop).lLstCodeZone(3)
                    grdPost.TextMatrix(llRow, POSTINDEX) = iLoop
                    grdPost.TextMatrix(llRow, ORIGTYPEINDEX) = tmPostInfo(iLoop).iType
                End If
            Else
                grdPost.TextMatrix(llRow, DATEINDEX) = sDate
                grdPost.TextMatrix(llRow, TIMEINDEX) = sTime
                grdPost.TextMatrix(llRow, ADVTINDEX) = ""
                grdPost.TextMatrix(llRow, CNTRNOINDEX) = ""
                grdPost.TextMatrix(llRow, LENINDEX) = tmPostInfo(iLoop).iUnits & "/" & tmPostInfo(iLoop).iLen
                grdPost.TextMatrix(llRow, CARTESTINDEX) = ""
                grdPost.TextMatrix(llRow, CARTCSTINDEX) = ""
                grdPost.TextMatrix(llRow, CARTMSTINDEX) = ""
                grdPost.TextMatrix(llRow, CARTPSTINDEX) = ""
                grdPost.TextMatrix(llRow, STATUSINDEX) = sStatus
                grdPost.TextMatrix(llRow, LSTCODEESTINDEX) = tmPostInfo(iLoop).lLstCodeZone(0)
                grdPost.TextMatrix(llRow, LSTCODECSTINDEX) = tmPostInfo(iLoop).lLstCodeZone(1)
                grdPost.TextMatrix(llRow, LSTCODEMSTINDEX) = tmPostInfo(iLoop).lLstCodeZone(2)
                grdPost.TextMatrix(llRow, LSTCODEPSTINDEX) = tmPostInfo(iLoop).lLstCodeZone(3)
                grdPost.TextMatrix(llRow, POSTINDEX) = iLoop
                grdPost.TextMatrix(llRow, ORIGTYPEINDEX) = tmPostInfo(iLoop).iType
            End If
        Else
            grdPost.TextMatrix(llRow, DATEINDEX) = sDate
            grdPost.TextMatrix(llRow, TIMEINDEX) = sTime
        End If
        llRow = llRow + 1
    Next iLoop
    'Don't add extra row
'    If llRow >= grdPost.Rows Then
'        grdPost.AddItem ""
'    End If
    If Not iNew Then
        grdPost.TopRow = llTRow
    End If
    grdPost.Redraw = True
End Sub

Private Sub mClick()
    Dim iCol As Integer
    Dim iRow As Integer
    Dim lCode As Long
    Dim iColNum As Integer
    Dim iStatus As Integer
    Dim sStatus As String
    Dim iLoop As Integer
    
'    On Error GoTo ErrHand
'
'
'    DoEvents
'    If igPreOrPost Then
'        If (grdPostLog.Col <= 1) Or (grdPostLog.Col >= 10) Or imHeaderClick Then       'Don't change color of vehicle or #missing columns
'            cmdDone.SetFocus
'            imHeaderClick = False
'            Exit Sub
'        End If
'        If sgUstWin(3) <> "I" Then
'            cmdDone.SetFocus
'            imHeaderClick = False
'            Exit Sub
'        End If
'    Else
'        If (grdPostLog.Col <= 1) Or (grdPostLog.Col >= 10) Or imHeaderClick Then       'Don't change color of vehicle or #missing columns
'            cmdDone.SetFocus
'            imHeaderClick = False
'            Exit Sub
'        End If
'        If sgUstWin(5) <> "I" Then
'            cmdDone.SetFocus
'            imHeaderClick = False
'            Exit Sub
'        End If
'    End If
'    Screen.MousePointer = vbHourglass
'    imHeaderClick = False
'    iCol = grdPostLog.Col
'    Select Case iCol
'        Case 2  'Advertiser
'            'grdPostLog.DroppedDown = True
'        Case 3  'Contract
'            Screen.MousePointer = vbHourglass
'            mFillCntr grdPostLog.Row
'            Screen.MousePointer = vbDefault
'
'            'grdPostLog.DroppedDown = False
'            'grdPostLog.DroppedDown = True
'        Case 4  'Length
'        Case 5  'Cart-EST
'            Screen.MousePointer = vbHourglass
'            mFillCart grdPostLog.Row
'            Screen.MousePointer = vbDefault
'        Case 6  'Cart-CST
'            Screen.MousePointer = vbHourglass
'            mFillCart grdPostLog.Row
'            Screen.MousePointer = vbDefault
'        Case 7  'Cart-MST
'            Screen.MousePointer = vbHourglass
'            mFillCart grdPostLog.Row
'            Screen.MousePointer = vbDefault
'        Case 8  'Cart-PST
'            Screen.MousePointer = vbHourglass
'            mFillCart grdPostLog.Row
'            Screen.MousePointer = vbDefault
'        Case 9  'Status
'            'sStatus = grdPostLog.Columns(iCol).Text
'            'If StrComp(sStatus, "Missed", 1) = 0 Then
'            '    iStatus = 0
'            'Else
'            '    iStatus = 1
'            'End If
'            'cnn.BeginTrans
'            'For iLoop = 1 To 4 Step 1
'            '    iColNum = grdPostLog.Col + iLoop        'get cpttCode number from the column next to the date
'            '    If Trim$(grdPostLog.Columns(iColNum).Text) <> "" Then
'            '        lCode = grdPostLog.Columns(iColNum).Text
'            '        If lCode > 0 Then
'            '            SQLQuery = "UPDATE lst SET "
'            '            SQLQuery = SQLQuery + "lstStatus = " & iStatus & ""
'            '            SQLQuery = SQLQuery + " WHERE lstCode = " & lCode & ""
'            '            cnn.Execute SQLQuery, rdExecDirect
'            '        End If
'            '    End If
'            'Next iLoop
'            'If iStatus = 0 Then
'            '    grdPostLog.Columns(iCol).Text = "Aired"
'            'Else
'            '    grdPostLog.Columns(iCol).Text = "Missed"
'            'End If
'            'cnn.CommitTrans
'            If grdPostLog.Columns(9).ListIndex < 0 Then
'                grdPostLog.Columns(9).ListIndex = 0
'                grdPostLog_Change
'            End If
'    End Select
'    If (grdPostLog.Col = 2) Or (grdPostLog.Col = 3) Or (grdPostLog.Col = 5) Or (grdPostLog.Col = 6) Or (grdPostLog.Col = 7) Or (grdPostLog.Col = 8) Or (grdPostLog.Col = 9) Then
'        grdPostLog.DroppedDown = True
'    End If
    Screen.MousePointer = vbDefault
    'grdPostLog.Col = 0    'set to zero to aviod change when header clicked
    'cmdDone.SetFocus
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmPostLog-mClick"
End Sub


Private Sub cboSort_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long
    Dim iZone As Integer
    
    If imInChg Then
        Exit Sub
    End If
    imInChg = True
    Screen.MousePointer = vbHourglass
    sName = LTrim$(cboSort.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    lRow = SendMessageByString(cboSort.hwnd, CB_FINDSTRING, -1, sName)
    If lRow >= 0 Then
        'cboSort.Bookmark = lRow
        'cboSort.Text = cboSort.Columns(0).Text
        cboSort.ListIndex = lRow
        cboSort.SelStart = iLen
        cboSort.SelLength = Len(cboSort.Text)
        imVefCode = cboSort.ItemData(cboSort.ListIndex)
        chkZone(0).Enabled = False
        chkZone(1).Enabled = False
        chkZone(2).Enabled = False
        chkZone(3).Enabled = False
        optShow(0).Enabled = False
        optShow(1).Enabled = False
        optShow(2).Enabled = False
        optShow(3).Enabled = False
        For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            If tgVehicleInfo(iLoop).iCode = imVefCode Then
                For iZone = LBound(tgVehicleInfo(iLoop).sZone) To UBound(tgVehicleInfo(iLoop).sZone) Step 1
                    Select Case Left$(tgVehicleInfo(iLoop).sZone(iZone), 1)
                        Case "E"
                            If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
                                chkZone(0).Enabled = True
                            End If
                            'If (tgVehicleInfo(iLoop).sFed(iZone) <> "*") And (tgVehicleInfo(iLoop).sFed(iZone) <> " ") Then
                            If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
                                optShow(0).Enabled = True
                            End If
                        Case "C"
                            If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
                                chkZone(1).Enabled = True
                            End If
                            'If (tgVehicleInfo(iLoop).sFed(iZone) <> "*") And (tgVehicleInfo(iLoop).sFed(iZone) <> " ") Then
                            If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
                                optShow(1).Enabled = True
                            End If
                        Case "M"
                            If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
                                chkZone(2).Enabled = True
                            End If
                            'If (tgVehicleInfo(iLoop).sFed(iZone) <> "*") And (tgVehicleInfo(iLoop).sFed(iZone) <> " ") Then
                            If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
                                optShow(2).Enabled = True
                            End If
                        Case "P"
                            If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
                                chkZone(3).Enabled = True
                            End If
                            'If (tgVehicleInfo(iLoop).sFed(iZone) <> "*") And (tgVehicleInfo(iLoop).sFed(iZone) <> " ") Then
                            If tgVehicleInfo(iLoop).sFed(iZone) = "*" Then
                                optShow(3).Enabled = True
                            End If
                    End Select
                Next iZone
                Exit For
            End If
        Next iLoop
    End If
    Screen.MousePointer = vbDefault
    imInChg = False

End Sub

Private Sub cboSort_Click()
    
    cboSort_Change
End Sub

Private Sub cboSort_GotFocus()
    mPostSetShow
    mResetSwapColor
End Sub

Private Sub cboSort_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cboSort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cboSort.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub chkZone_GotFocus(Index As Integer)
    mPostSetShow
    mResetSwapColor
End Sub

Private Sub cmcDropDown_Click()
    
    Select Case grdPost.Col
        Case ADVTINDEX
            lbcAdvt.Visible = Not lbcAdvt.Visible
        Case CNTRNOINDEX
            lbcCntr.Visible = Not lbcCntr.Visible
        Case CARTESTINDEX To CARTPSTINDEX
            lbcCart(grdPost.Col - CARTESTINDEX).Visible = Not lbcCart(grdPost.Col - CARTESTINDEX).Visible
        Case STATUSINDEX
            lbcStatus.Visible = Not lbcStatus.Visible
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload frmPostLog
End Sub

Private Sub cmdCancel_GotFocus()
    mPostSetShow
    mResetSwapColor
End Sub

Private Sub cmdDone_Click()
    Dim iRet As Integer
    
    If imFieldChgd = True Then
        'If gMsgBox("Save all changes?", vbYesNo) = vbYes Then
            Screen.MousePointer = vbHourglass
            iRet = mPutLst(-1)
            If iRet = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        'End If
    End If
    Screen.MousePointer = vbDefault
    Unload frmPostLog
End Sub


Private Sub cmdDone_GotFocus()
    mPostSetShow
    mResetSwapColor
End Sub

Private Sub cmdSave_Click()
    Dim iRet As Integer
    
    Screen.MousePointer = vbHourglass
    If imFieldChgd Then
        mAddAbfRecords
        iRet = mPutLst(-1)
        If iRet = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        'Rebuild to get saved values and remove deleted avails
        Screen.MousePointer = vbHourglass
        mGetLst True
    End If
    imFieldChgd = False
    cmdSave.Enabled = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSave_GotFocus()
    mPostSetShow
    mResetSwapColor
End Sub

Private Sub Form_Activate()
    Dim ilCol As Integer
    Dim llRow As Long
    
    If imFirstTime Then
        bgLogVisible = True
        mSetColumnWidths True, True, True, True
        For ilCol = 0 To grdPost.Cols - 1 Step 1
            grdPost.ColAlignment(ilCol) = flexAlignLeftCenter
        Next ilCol
        gGrid_AlignAllColsLeft grdPost
        grdPost.TextMatrix(0, DATEINDEX) = "Date"
        grdPost.TextMatrix(0, TIMEINDEX) = "Time"
        grdPost.TextMatrix(0, ADVTINDEX) = "Advertiser"
        grdPost.TextMatrix(0, CNTRNOINDEX) = "Contract # Product"
        grdPost.TextMatrix(0, LENINDEX) = "Len"
        If sgSpfUseCartNo = "N" Then
            grdPost.TextMatrix(0, CARTESTINDEX) = "ISCI-EST"
            grdPost.TextMatrix(0, CARTCSTINDEX) = "ISCI-CST"
            grdPost.TextMatrix(0, CARTMSTINDEX) = "ISCI-MST"
            grdPost.TextMatrix(0, CARTPSTINDEX) = "ISCI-PST"
        Else
            grdPost.TextMatrix(0, CARTESTINDEX) = "Cart-EST"
            grdPost.TextMatrix(0, CARTCSTINDEX) = "Cart-CST"
            grdPost.TextMatrix(0, CARTMSTINDEX) = "Cart-MST"
            grdPost.TextMatrix(0, CARTPSTINDEX) = "Cart-PST"
        End If
        grdPost.TextMatrix(0, STATUSINDEX) = "Status"
        gGrid_IntegralHeight grdPost
        mClearGrid
        imFirstTime = False
        imFieldChgd = False
    End If
End Sub

Private Sub Form_Click()
    'Can't place into pbcClickFocus because of the swap color
    mPostSetShow
    mResetSwapColor
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.25
    Me.Top = (Screen.Height - Me.Height) / 1.8
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmPostLog
    gCenterForm frmPostLog
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    Dim ilLoop As Integer
    

    imFirstTime = True
    imVefCode = -1
    imIntegralSet = False
    imBSMode = False
    imInChg = False
    smFWkDate = ""
    smLWkDate = ""
    imHeaderClick = False
    imFillCntrAdfCode = -1
    imMouseDown = False
    imInFillCntr = False
    imFromGetLst = False
    imFieldChgd = False
    imIgnoreChg = False
    imInRowNo = -1
    imShowGridBox = False
    lmTopRow = -1
    imFromArrow = False
    lmEnableRow = -1
    lmEnableCol = -1
    lmSwapStartRow = -1
    imSwapClickCount = -1
    bmCreateAbfRecord = False
    imAbfVefCode = -1
    ReDim lmAbfDate(0 To 0) As Long
    lmAbfDate(0) = 0
    
    imcPrt.Picture = frmDirectory!imcPrinter.Picture
    ReDim tmPostInfo(0 To 0) As POSTINFO
    
    'D.S. 10/09/02 Clicking radio buttons before a veh. was selected caused a subscript error
    optShow(0).Enabled = False
    optShow(1).Enabled = False
    optShow(2).Enabled = False
    optShow(3).Enabled = False
    
    
    mPopVehBox
    If igPreOrPost = 0 Then
        frmPostLog.Caption = "Network Log - " & sgClientName
        For ilLoop = 0 To UBound(tgStatusTypes) Step 1
            tmStatusTypes(ilLoop) = tgStatusTypes(ilLoop)
            If InStr(1, tmStatusTypes(ilLoop).sName, "1-", vbTextCompare) = 1 Then
                tmStatusTypes(ilLoop).sName = "1-Air Live"
            End If
            If InStr(1, tmStatusTypes(ilLoop).sName, "2-", vbTextCompare) = 1 Then
                tmStatusTypes(ilLoop).sName = "2-Air Delay B'cast"  '"2-Air In Daypart"
            End If
            If InStr(1, tmStatusTypes(ilLoop).sName, "7-", vbTextCompare) = 1 Then
                tmStatusTypes(ilLoop).sName = "7-Air Outside Pledge"
            End If
            If InStr(1, tmStatusTypes(ilLoop).sName, "8-", vbTextCompare) = 1 Then
                tmStatusTypes(ilLoop).sName = "8-Air Not Pledged"
            End If
        Next ilLoop
    Else
        frmPostLog.Caption = "Network Log - " & sgClientName
        For ilLoop = 0 To UBound(tgStatusTypes) Step 1
            tmStatusTypes(ilLoop) = tgStatusTypes(ilLoop)
        Next ilLoop
    End If
    mClearGrid
    
    For iLoop = 0 To UBound(tmStatusTypes) Step 1
        If (tmStatusTypes(iLoop).iStatus < ASTEXTENDED_MG) Then
            '3/11/11: Remove 7-Air Outside Pledge and 8-Air not pledged
            If (tmStatusTypes(iLoop).iStatus <> 6) And (tmStatusTypes(iLoop).iStatus <> 7) Then
                lbcStatus.AddItem Trim$(tmStatusTypes(iLoop).sName)
                lbcStatus.ItemData(lbcStatus.NewIndex) = iLoop
            End If
        End If
    Next iLoop
    mFillAdvt
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    bgLogVisible = False
    mAddAbfRecords
    Erase tmPostInfo
    Erase tmCntrInfo
    Erase tmCopyInfo
    Erase lmAbfDate
    rst_Lst.Close
    rst_chf.Close
    rst_Cif.Close
    rst_Gsf.Close
    Set frmPostLog = Nothing
End Sub

Private Sub grdPost_Click()
    Dim llRow As Long
    Dim llCol As Long
    
    If igPreOrPost = 0 Then
        If sgUstWin(3) <> "I" Then
            On Error Resume Next
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    Else
        If sgUstWin(5) <> "I" Then
            On Error Resume Next
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    End If
    If UBound(tmPostInfo) <= LBound(tmPostInfo) Then
        On Error Resume Next
        pbcClickFocus.SetFocus
        Exit Sub
    End If
'    If grdPost.Col < 2 Then
'        If lmSwapStartRow = -1 Then
'
'            If imSwapClickCount <> 2 Then
'                If (Trim$(grdPost.TextMatrix(grdPost.Row, 2)) <> "") Then
'                    grdPost.Col = 1
'                    grdPost.CellForeColor = vbGreen
'                    lmSwapStartRow = grdPost.Row
'                    imSwapClickCount = 0
'                End If
'            Else
'                imSwapClickCount = -1
'            End If
'        Else
'            If Trim$(grdPost.TextMatrix(grdPost.Row, 2)) <> "" Then
'                If lmSwapStartRow <> grdPost.Row Then
'                    mSwap
'                Else
'                    'Unable to deselect because this code is re-entered after the first click without the user making a second click
'                    If imSwapClickCount <> 0 Then
'                        mResetSwapColor
'                        imSwapClickCount = 2
'                    Else
'                        imSwapClickCount = 1
'                    End If
'                End If
'            End If
'        End If
'        pbcClickFocus.SetFocus
'        Exit Sub
'    End If
'    mResetSwapColor
'    If Trim$(grdPost.TextMatrix(grdPost.Row, 2)) <> "" Then
'        If Not mPostColAllowed(grdPost.Col) Then
'            pbcClickFocus.SetFocus
'            Exit Sub
'        End If
'    Else
'        If grdPost.Col > 2 Then
'            pbcClickFocus.SetFocus
'            Exit Sub
'        End If
'    End If
'    If grdPost.Col > 9 Then
'        Exit Sub
'    End If
'    lmTopRow = grdPost.TopRow
'    llRow = grdPost.Row
'    If grdPost.TextMatrix(llRow, 2) = "" Then
'        grdPost.Redraw = False
'        Do
'            llRow = llRow - 1
'        Loop While grdPost.TextMatrix(llRow, 2) = ""
'        grdPost.Row = llRow + 1
'        'grdPost.Col = 0
'        grdPost.Redraw = True
'    End If
'    mPostEnableBox
End Sub

Private Sub grdPost_EnterCell()
    If UBound(tmPostInfo) <= LBound(tmPostInfo) Then
        On Error Resume Next
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    mPostSetShow
    If igPreOrPost = 0 Then
        If sgUstWin(3) <> "I" Then
            On Error Resume Next
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    Else
        If sgUstWin(5) <> "I" Then
            On Error Resume Next
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub grdPost_GotFocus()
    If UBound(tmPostInfo) <= LBound(tmPostInfo) Then
        On Error Resume Next
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdPost.Col >= grdPost.Cols - 1 Then
        Exit Sub
    End If
    'grdPost_Click
End Sub

Private Sub grdPost_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This was added to aviod seeing rows/columns blanked out temporary if user dragged mouse to different cell
    lmTopRow = grdPost.TopRow
    grdPost.Redraw = False
End Sub

Private Sub grdPost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    
    'grdPost.ToolTipText = ""
    If (Y > grdPost.RowHeight(0)) And (Y < grdPost.RowHeight(0)) Then
        grdPost.ToolTipText = ""
        Exit Sub
    End If
    ilFound = gGrid_GetRowCol(grdPost, X, Y, llRow, llCol)
    If ilFound Then
        If (llCol = CARTESTINDEX) Or (llCol = CARTCSTINDEX) Or (llCol = CARTMSTINDEX) Or (llCol = CARTPSTINDEX) Then
            grdPost.ToolTipText = Trim$(grdPost.TextMatrix(llRow, llCol))
            Exit Sub
        End If
    End If
    grdPost.ToolTipText = ""
End Sub

Private Sub grdPost_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim ilType As Integer
    
    If igPreOrPost = 0 Then
        If sgUstWin(3) <> "I" Then
            grdPost.Redraw = True
            On Error Resume Next
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    Else
        If sgUstWin(5) <> "I" Then
            grdPost.Redraw = True
            On Error Resume Next
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    End If
    If UBound(tmPostInfo) <= LBound(tmPostInfo) Then
        grdPost.Redraw = True
        On Error Resume Next
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    'Determine row and col mouse up onto
    ilFound = gGrid_DetermineRowCol(grdPost, X, Y)
    If Not ilFound Then
        grdPost.Redraw = True
        On Error Resume Next
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    If grdPost.Row - 1 >= UBound(tmPostInfo) Then
        grdPost.Redraw = True
        On Error Resume Next
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    ilType = grdPost.TextMatrix(grdPost.Row, ORIGTYPEINDEX)
    If grdPost.Col <= TIMEINDEX Then
        If ilType = 1 Then
            mResetSwapColor
            grdPost.Col = TIMEINDEX
            grdPost.CellForeColor = DARKGREEN   'vbGreen
            lmSwapStartRow = grdPost.Row
            imSwapClickCount = -1
        Else
            If ((lmSwapStartRow >= 0) And (imSwapClickCount = -1)) Then
                mResetSwapColor
            End If
            If (lmSwapStartRow = -1) Then
                If imSwapClickCount <> 2 Then
                    grdPost.Col = TIMEINDEX
                    grdPost.CellForeColor = DARKGREEN   'vbGreen
                    lmSwapStartRow = grdPost.Row
                    imSwapClickCount = 0
                Else
                    imSwapClickCount = -1
                End If
            Else
                'If Trim$(grdPost.TextMatrix(grdPost.Row, ADVTINDEX)) <> "" Then
                If lmSwapStartRow <> grdPost.Row Then
                    mSwap
                Else
                    'Unable to deselect because this code is re-entered after the first click without the user making a second click
                    If imSwapClickCount <> 0 Then
                        mResetSwapColor
                        imSwapClickCount = 2
                    Else
                        imSwapClickCount = 1
                    End If
                End If
            End If
        End If
        grdPost.Redraw = True
        On Error Resume Next
        pbcClickFocus.SetFocus
        Exit Sub
    End If
    mResetSwapColor
    'If Trim$(grdPost.TextMatrix(grdPost.Row, ADVTINDEX)) <> "" Then
    If (ilType = 0) Or (ilType = 2) Then
        If Not mPostColAllowed(grdPost.Col) Then
            grdPost.Redraw = True
            On Error Resume Next
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    Else
        'If grdPost.Col > ADVTINDEX Then
            grdPost.Redraw = True
            On Error Resume Next
            pbcClickFocus.SetFocus
            Exit Sub
        'End If
    End If
    If grdPost.Col > lmMaxCol Then
        Exit Sub
    End If
    lmTopRow = grdPost.TopRow
    grdPost.Redraw = True
    mPostEnableBox
End Sub

Private Sub grdPost_Scroll()
'    If (lmTopRow <> -1) And (lmTopRow <> grdPost.TopRow) Then
'        grdPost.TopRow = lmTopRow
'        lmTopRow = -1
'    End If
    If grdPost.Redraw = False Then
        grdPost.Redraw = True
        grdPost.TopRow = lmTopRow
        grdPost.Refresh
        grdPost.Redraw = False
    End If
    If (imShowGridBox) And (grdPost.Row >= grdPost.FixedRows) And (grdPost.Col >= DATEINDEX) And (grdPost.Col < grdPost.Cols - 1) Then
        If grdPost.RowIsVisible(grdPost.Row) Then
            pbcArrow.Move grdPost.Left - pbcArrow.Width, grdPost.Top + grdPost.RowPos(grdPost.Row) + (grdPost.RowHeight(grdPost.Row) - pbcArrow.Height) / 2
            pbcArrow.Visible = True
           If grdPost.Col = ADVTINDEX Then  'Advertiser
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - cmcDropDown.Width - 30, grdPost.RowHeight(grdPost.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcAdvt.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + cmcDropDown.Width
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcAdvt.Visible = True
                txtDropdown.SetFocus
           ElseIf grdPost.Col = CNTRNOINDEX Then  'Contract
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - cmcDropDown.Width - 30, grdPost.RowHeight(grdPost.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcCntr.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + cmcDropDown.Width
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcCntr.Visible = True
                txtDropdown.SetFocus
           ElseIf (grdPost.Col = LENINDEX) Then
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - 30, grdPost.RowHeight(grdPost.Row) - 15
                txtDropdown.Visible = True
                txtDropdown.SetFocus
            ElseIf (grdPost.Col = CARTESTINDEX) Or (grdPost.Col = CARTCSTINDEX) Or (grdPost.Col = CARTMSTINDEX) Or (grdPost.Col = CARTPSTINDEX) Then
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - cmcDropDown.Width - 30, grdPost.RowHeight(grdPost.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcCart(grdPost.Col - CARTESTINDEX).Move txtDropdown.Left - 3 * txtDropdown.Width, txtDropdown.Top + txtDropdown.Height, 4 * txtDropdown.Width + cmcDropDown.Width ' + grdPost.ColWidth(STATUSINDEX)
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcCart(grdPost.Col - CARTESTINDEX).Visible = True
                txtDropdown.SetFocus
            ElseIf grdPost.Col = STATUSINDEX Then
                txtDropdown.Move grdPost.Left + grdPost.ColPos(grdPost.Col) + 30, grdPost.Top + grdPost.RowPos(grdPost.Row) + 15, grdPost.ColWidth(grdPost.Col) - cmcDropDown.Width - 30, grdPost.RowHeight(grdPost.Row) - 15
                cmcDropDown.Move txtDropdown.Left + txtDropdown.Width, txtDropdown.Top, cmcDropDown.Width, txtDropdown.Height
                lbcStatus.Move txtDropdown.Left, txtDropdown.Top + txtDropdown.Height, txtDropdown.Width + (7 * txtDropdown.Width) \ 2
                txtDropdown.Visible = True
                cmcDropDown.Visible = True
                lbcStatus.Visible = True
                txtDropdown.SetFocus
            End If
        Else
            pbcPostFocus.SetFocus
            txtDropdown.Visible = False
            cmcDropDown.Visible = False
            lbcAdvt.Visible = False
            lbcCntr.Visible = False
            lbcCart(0).Visible = False
            lbcCart(1).Visible = False
            lbcCart(2).Visible = False
            lbcCart(3).Visible = False
            lbcStatus.Visible = False
            pbcArrow.Visible = False
        End If
    Else
        pbcPostFocus.SetFocus
        pbcArrow.Visible = False
        imFromArrow = False
    End If
End Sub














Private Sub imcInsert_Click()
    Dim ilRet As Integer
    mPostSetShow
    If lmSwapStartRow >= 0 Then
        ilRet = mInsert()
    End If
End Sub

Private Sub imcPrt_Click()
    Dim iLoop As Integer
    Dim sVehicle As String
    Dim sDate As String
    Dim sTime As String
    Dim sStatus As String
    Dim sCart As String
    Dim ilIdx As Integer
    Dim llRow As Long
    
    
    Screen.MousePointer = vbHourglass
    mPostSetShow
    mResetSwapColor
    Printer.Print ""
    Printer.Print Tab(65); Format$(Now)
    Printer.Print ""
    sVehicle = ""
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        If tgVehicleInfo(iLoop).iCode = imVefCode Then
            sVehicle = Trim$(tgVehicleInfo(iLoop).sVehicle)
            Exit For
        End If
    Next iLoop
    Printer.Print sVehicle & " " & smLWkDate & "-" & smFWkDate
    Printer.Print ""
    'Printer.Print "  Call Letters"; Tab(15); "Vehicle"; Tab(37); "Contact"; Tab(69); "Dates"
    
    'grdStation.MoveFirst
    'For iLoop = 0 To grdStation.Rows - 1 Step 1
    '    sContactPhone = Trim$(grdStation.Columns(2).Text)
    '    If Len(sContactPhone) > 16 Then
    '        sContactPhone = Left$(sContactPhone, 30 - Len(Trim$(grdStation.Columns(3)))) & " " & Trim$(grdStation.Columns(3).Text)
    '    Else
    '        sContactPhone = sContactPhone & " " & Trim$(grdStation.Columns(3).Text)
    '    End If
    '    Printer.Print "  " & Trim$(grdStation.Columns(0).Text); Tab(15); Trim$(grdStation.Columns(1).Text); Tab(37); sContactPhone; Tab(69); Trim$(grdStation.Columns(4).Text) & " " & Trim$(grdStation.Columns(5).Text) & " " & Trim$(grdStation.Columns(6).Text)
    '    grdStation.MoveNext
    'Next iLoop
    For iLoop = 0 To UBound(tmPostInfo) - 1 Step 1
        llRow = iLoop + grdPost.FixedRows
        If optShow(1).Value Then
            sDate = tmPostInfo(iLoop).sDateZone(1)
            sTime = tmPostInfo(iLoop).sTimeZone(1)
        ElseIf optShow(2).Value Then
            sDate = tmPostInfo(iLoop).sDateZone(2)
            sTime = tmPostInfo(iLoop).sTimeZone(2)
        ElseIf optShow(3).Value Then
            sDate = tmPostInfo(iLoop).sDateZone(3)
            sTime = tmPostInfo(iLoop).sTimeZone(3)
        Else
            sDate = tmPostInfo(iLoop).sDateZone(0)
            sTime = tmPostInfo(iLoop).sTimeZone(0)
        End If
        'If tmPostInfo(iLoop).iStatus = 1 Then
        '    sStatus = "Missed"
        'Else
        '    sStatus = "Aired"
        'End If
        sDate = Format$(sDate, sgShowDateForm)
        If (tmPostInfo(iLoop).iType = 0) Or (tmPostInfo(iLoop).iType = 2) Then
            If tmPostInfo(iLoop).iStatus < ASTEXTENDED_MG Then
                sStatus = grdPost.TextMatrix(llRow, STATUSINDEX)  '(tmStatusTypes(tmPostInfo(iLoop).iStatus).sName)
            ElseIf tmPostInfo(iLoop).iStatus = ASTEXTENDED_MG Then
                sStatus = "MG"
            ElseIf tmPostInfo(iLoop).iStatus = ASTEXTENDED_BONUS Then
                sStatus = "Bonus"
            Else
                sStatus = ""
            End If
            
            'sCart = Trim$(Trim$(tmPostInfo(iLoop).sCartZone(0)) & " " & Trim$(tmPostInfo(iLoop).sCartZone(1)) & " " & Trim$(tmPostInfo(iLoop).sCartZone(2)) & " " & Trim$(tmPostInfo(iLoop).sCartZone(3)))
            'D.S. 10/08/02  Was printing all time zone carts - now prints what's shown
            sCart = ""
            For ilIdx = CARTESTINDEX To CARTPSTINDEX Step 1
                If Trim$(grdPost.TextMatrix(llRow, ilIdx)) <> "" Then
                    sCart = sCart & Trim$(grdPost.TextMatrix(llRow, ilIdx)) & " "   'Trim$(tmPostInfo(iLoop).sCartZone(ilIdx - 5)) & " "
                End If
            Next ilIdx
                
            If tmPostInfo(iLoop).lCntrNo > 0 Then
                'Printer.Print "  " & sDate & " "; Tab(15); sTime & " "; Tab(25); Trim$(tmPostInfo(iLoop).sAdfName) & " "; Tab(50); tmPostInfo(iLoop).lCntrNo & " "; Tab(60); Trim$(tmPostInfo(iLoop).sProd) & " "; Tab(90); tmPostInfo(iLoop).iLen & " "; Tab(94); sCart & " "; Tab(117); sStatus
                Printer.Print "  " & sDate & " "; Tab(15); sTime & " "; Tab(27); Trim$(grdPost.TextMatrix(llRow, ADVTINDEX)) & " "; Tab(50); Trim$(grdPost.TextMatrix(llRow, CNTRNOINDEX)) & " "; Tab(90); Trim$(grdPost.TextMatrix(llRow, LENINDEX)) & " "; Tab(94); sCart & " "; Tab(117); sStatus
            Else
                Printer.Print "  " & sDate & " "; Tab(15); sTime & " "; Tab(27); Trim$(grdPost.TextMatrix(llRow, ADVTINDEX)) & " "; Tab(60); Trim$(grdPost.TextMatrix(llRow, CNTRNOINDEX)) & " "; Tab(90); Trim$(grdPost.TextMatrix(llRow, LENINDEX)) & " "; Tab(94); sCart & " "; Tab(117); sStatus
            End If
        ElseIf tmPostInfo(iLoop).iType = 0 Then
            'Printer.Print sDate & " "; Tab(13); sTime & " "; Tab(23); tmPostInfo(iLoop).iUnits & "/" & tmPostInfo(iLoop).iLen
            Printer.Print sDate & " "; Tab(15); sTime & " "; Tab(25); Trim$(grdPost.TextMatrix(llRow, LENINDEX))
        End If
    Next iLoop
    Printer.EndDoc
    Screen.MousePointer = vbDefault
    pbcClickFocus.SetFocus
End Sub

Private Sub lbcAdvt_Click()
    txtDropdown.Text = lbcAdvt.List(lbcAdvt.ListIndex)
    If (txtDropdown.Visible) And (txtDropdown.Enabled) Then
        txtDropdown.SetFocus
        lbcAdvt.Visible = False
    End If
End Sub


Private Sub lbcCart_Click(Index As Integer)
    txtDropdown.Text = lbcCart(Index).List(lbcCart(Index).ListIndex)
    If (txtDropdown.Visible) And (txtDropdown.Enabled) Then
        txtDropdown.SetFocus
        lbcCart(Index).Visible = False
    End If
End Sub

Private Sub lbcCntr_Click()
    txtDropdown.Text = lbcCntr.List(lbcCntr.ListIndex)
    If (txtDropdown.Visible) And (txtDropdown.Enabled) Then
        txtDropdown.SetFocus
        lbcCntr.Visible = False
    End If
End Sub

Private Sub lbcStatus_Click()
    txtDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
    If (txtDropdown.Visible) And (txtDropdown.Enabled) Then
        txtDropdown.SetFocus
        lbcStatus.Visible = False
    End If
End Sub

Private Sub optShow_Click(Index As Integer)
    If imFromGetLst Then
        mGridPaint True
        imFromGetLst = False
    Else
        mGridPaint False
    End If
End Sub

Private Sub optShow_GotFocus(Index As Integer)
    mPostSetShow
    mResetSwapColor
End Sub

Private Sub pbcInsert_KeyUp(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub pbcPostSTab_GotFocus()
    Dim ilType As Integer
    
    If GetFocus() <> pbcPostSTab.hwnd Then
        Exit Sub
    End If
    If igPreOrPost = 0 Then
        If sgUstWin(3) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    Else
        If sgUstWin(5) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    End If
    If imFromArrow Then
        imFromArrow = False
        mPostEnableBox
        Exit Sub
    End If
    If txtDropdown.Visible Then
        mPostSetShow
        mResetSwapColor
        'If grdPost.Col = ADVTINDEX Then
        If grdPost.Col = STATUSINDEX Then
            Do
                If grdPost.Row > grdPost.FixedRows Then
                    lmTopRow = -1
                    grdPost.Row = grdPost.Row - 1
                    If Not grdPost.RowIsVisible(grdPost.Row) Then
                        grdPost.TopRow = grdPost.TopRow - 1
                    End If
                    ilType = grdPost.TextMatrix(grdPost.Row, ORIGTYPEINDEX)
                    If (ilType <> 1) Then
                        'grdPost.Col = STATUSINDEX
                        grdPost.Col = lmMaxCol
                        mPostEnableBox
                        Exit Sub
                    End If
                Else
                    pbcClickFocus.SetFocus
                    Exit Sub
                End If
            Loop
        Else
            grdPost.Col = grdPost.Col - 1
            Do While grdPost.ColWidth(grdPost.Col) = 0
                grdPost.Col = grdPost.Col - 1
            Loop
            mPostEnableBox
        End If
    Else
        lmTopRow = -1
        grdPost.TopRow = grdPost.FixedRows
        grdPost.Row = grdPost.FixedRows - 1
        Do
            If grdPost.Row + 1 < grdPost.Rows Then
                lmTopRow = -1
                grdPost.Row = grdPost.Row + 1
                If Not grdPost.RowIsVisible(grdPost.Row) Then
                    grdPost.TopRow = grdPost.TopRow + 1
                End If
                ilType = grdPost.TextMatrix(grdPost.Row, ORIGTYPEINDEX)
                If (ilType <> 1) Then
                    'grdPost.Col = ADVTINDEX
                    grdPost.Col = STATUSINDEX
                    mPostEnableBox
                    Exit Sub
                End If
            Else
                pbcClickFocus.SetFocus
                Exit Sub
            End If
        Loop
    End If
End Sub

Private Sub pbcPostTab_GotFocus()
    Dim ilType As Integer
    Dim llRow As Long
    
    If GetFocus() <> pbcPostTab.hwnd Then
        Exit Sub
    End If
    If igPreOrPost = 0 Then
        If sgUstWin(3) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    Else
        If sgUstWin(5) <> "I" Then
            pbcClickFocus.SetFocus
            Exit Sub
        End If
    End If
    If txtDropdown.Visible Then
        mPostSetShow
        mResetSwapColor
        'If grdPost.Col = STATUSINDEX Then
        If grdPost.Col = lmMaxCol Then
            llRow = grdPost.Rows
            Do
                llRow = llRow - 1
            Loop While grdPost.TextMatrix(llRow, DATEINDEX) = ""
            llRow = llRow + 1
            Do
'                If grdPost.Row + 1 < grdPost.Rows Then
                If (grdPost.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdPost.Row = grdPost.Row + 1
                    If Not grdPost.RowIsVisible(grdPost.Row) Then
                        grdPost.TopRow = grdPost.TopRow + 1
                    End If
                    'If Trim$(grdPost.TextMatrix(grdPost.Row, ADVTINDEX)) <> "" Then
                    ilType = grdPost.TextMatrix(grdPost.Row, ORIGTYPEINDEX)
                    If (ilType <> 1) Then
                        'grdPost.Col = ADVTINDEX
                        grdPost.Col = STATUSINDEX
                        mPostEnableBox
                        Exit Sub
                    'Else
                    '    imFromArrow = True
                    '    pbcArrow.Move grdPost.Left - pbcArrow.Width, grdPost.Top + grdPost.RowPos(grdPost.Row) + (grdPost.RowHeight(grdPost.Row) - pbcArrow.Height) / 2
                    '    pbcArrow.Visible = True
                    '    pbcArrow.SetFocus
                    End If
                Else
                    pbcClickFocus.SetFocus
                    Exit Sub
                End If
            Loop
        Else
            grdPost.Col = grdPost.Col + 1
            Do While grdPost.ColWidth(grdPost.Col) = 0
                grdPost.Col = grdPost.Col + 1
            Loop
            mPostEnableBox
        End If
    Else
        lmTopRow = -1
        grdPost.TopRow = grdPost.FixedRows
        grdPost.Row = grdPost.FixedRows - 1
        llRow = grdPost.Rows
        Do
            llRow = llRow - 1
        Loop While grdPost.TextMatrix(llRow, DATEINDEX) = ""
        llRow = llRow + 1
        Do
'            If grdPost.Row + 1 < grdPost.Rows Then
            If grdPost.Row + 1 < llRow Then
                lmTopRow = -1
                grdPost.Row = grdPost.Row + 1
                If Not grdPost.RowIsVisible(grdPost.Row) Then
                    grdPost.TopRow = grdPost.TopRow + 1
                End If
                ilType = grdPost.TextMatrix(grdPost.Row, ORIGTYPEINDEX)
                If (ilType <> 1) Then
                    'grdPost.Col = ADVTINDEX
                    grdPost.Col = STATUSINDEX
                    mPostEnableBox
                    Exit Sub
                End If
            Else
                pbcClickFocus.SetFocus
                Exit Sub
            End If
        Loop
    End If
End Sub

Private Sub rbcVeh_Click(Index As Integer)

    mPopVehBox
    
End Sub

Private Sub rbcVeh_GotFocus(Index As Integer)
    mPostSetShow
    mResetSwapColor
End Sub

Private Sub tmcDelay_Timer()
End Sub

Private Sub tmcSwap_Timer()
End Sub

Private Sub txtDropdown_Change()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLen As Integer
    
    slStr = txtDropdown.Text
    ilLen = Len(slStr)
    If imBSMode Then
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        imBSMode = False
    End If
    Select Case grdPost.Col
        Case ADVTINDEX
            llRow = SendMessageByString(lbcAdvt.hwnd, LB_FINDSTRING, -1, slStr)
            If llRow >= 0 Then
                lbcAdvt.ListIndex = llRow
                txtDropdown.Text = lbcAdvt.List(lbcAdvt.ListIndex)
                txtDropdown.SelStart = ilLen
                txtDropdown.SelLength = Len(txtDropdown.Text)
            End If
        Case CNTRNOINDEX
            llRow = SendMessageByString(lbcCntr.hwnd, LB_FINDSTRING, -1, slStr)
            If llRow >= 0 Then
                lbcCntr.ListIndex = llRow
                txtDropdown.Text = lbcCntr.List(lbcCntr.ListIndex)
                txtDropdown.SelStart = ilLen
                txtDropdown.SelLength = Len(txtDropdown.Text)
            End If
        Case CARTESTINDEX To CARTPSTINDEX
            llRow = SendMessageByString(lbcCart(grdPost.Col - CARTESTINDEX).hwnd, LB_FINDSTRING, -1, slStr)
            If llRow >= 0 Then
                lbcCart(grdPost.Col - CARTESTINDEX).ListIndex = llRow
                txtDropdown.Text = lbcCart(grdPost.Col - CARTESTINDEX).List(lbcCart(grdPost.Col - CARTESTINDEX).ListIndex)
                txtDropdown.SelStart = ilLen
                txtDropdown.SelLength = Len(txtDropdown.Text)
            End If
        Case STATUSINDEX
            llRow = SendMessageByString(lbcStatus.hwnd, LB_FINDSTRING, -1, slStr)
            If llRow >= 0 Then
                lbcStatus.ListIndex = llRow
                txtDropdown.Text = lbcStatus.List(lbcStatus.ListIndex)
                txtDropdown.SelStart = ilLen
                txtDropdown.SelLength = Len(txtDropdown.Text)
            End If
    End Select

End Sub

Private Sub txtDropdown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtDropdown_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub txtDropdown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If txtDropdown.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub txtDropdown_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdPost.Col
            Case ADVTINDEX
                gProcessArrowKey Shift, KeyCode, lbcAdvt, True
            Case CNTRNOINDEX
                gProcessArrowKey Shift, KeyCode, lbcCntr, True
            Case CARTESTINDEX To CARTPSTINDEX
                gProcessArrowKey Shift, KeyCode, lbcCart(grdPost.Col - CARTESTINDEX), True
            Case STATUSINDEX
                gProcessArrowKey Shift, KeyCode, lbcStatus, True
        End Select
    End If
End Sub

Private Sub txtWeek_Change()
    mAddAbfRecords
    mClearGrid
End Sub

Private Sub txtWeek_GotFocus()
    mPostSetShow
    mResetSwapColor
    gCtrlGotFocus ActiveControl
End Sub

Private Sub mGetLst(ilFromSave As Integer)
    Dim sStatus As String
    Dim sCartOrISCI As String
    Dim sTime As String
    Dim iWkNo As Integer
    Dim iBreakNo As Integer
    Dim iPositionNo As Integer
    Dim iSeqNo As Integer
    Dim iFound As Integer
    Dim iUpper As Integer
    Dim iLoop As Integer
    Dim iTest As Integer
    Dim iInsert As Integer
    Dim iZone As Integer
    Dim sAdvtName As String
    Dim sVehicle As String
    Dim llVeh As Integer
    Dim ilLang As Integer
    Dim ilTeam As Integer
    Dim slStr As String
    Dim slEDate As String
    ReDim iZoneFd(0 To 3) As Integer
    
    ReDim tmPostInfo(0 To 0) As POSTINFO
    imFieldChgd = False
    imBSMode = False
    imInChg = False
    imInRowNo = -1
    imHeaderClick = False
    imMouseDown = False
    imIgnoreChg = False
    iZoneFd(0) = False
    iZoneFd(1) = False
    iZoneFd(2) = False
    iZoneFd(3) = False
    If imVefCode <= 0 Then
        Beep
        gMsgBox "Please enter a vehicle.", vbCritical
        cboSort.SetFocus
    End If
    If gIsDate(txtWeek.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        txtWeek.SetFocus
    Else
        smFWkDate = Format(txtWeek.Text, sgShowDateForm)
    End If
    On Error GoTo ErrHand
    smLWkDate = gObtainNextSunday(smFWkDate)
    smFWkDate = Format$(smFWkDate, sgShowDateForm)
    smLWkDate = Format$(smLWkDate, sgShowDateForm)
    '11/23/11: Include next day
    slEDate = smLWkDate
    llVeh = gBinarySearchVef(CLng(imVefCode))
    If llVeh <> -1 Then
        If tgVehicleInfo(llVeh).sVehType = "G" Then
            slEDate = DateAdd("d", 1, smLWkDate)
        End If
    End If
    SQLQuery = "SELECT lstType, lstSdfCode, lstLogDate, lstLogTime, lstProd, lstCntrNo, lstCifCode, lstISCI, lstZone, lstCart, lstLen, lstUnits, lstStatus, lstWkNo, lstBreakNo, lstPositionNo, lstSeqNo, lstAnfCode, lstCode, lstAdfCode, lstAgfCode, lstGsfCode FROM lst "
    'SQLQuery = SQLQuery + " WHERE (adf.adfCode = lst.lstAdfCode"
    SQLQuery = SQLQuery + " WHERE (lstLogVefCode = " & imVefCode
    'If chkZone(0).Value Then
    '    SQLQuery = SQLQuery + " AND lst.lstZone = '" & tgCPPosting(0).sZone & "'"
    'End If
    SQLQuery = SQLQuery + " AND lstBkoutLstCode = 0"
    '3/9/16: Fix the filter
    'SQLQuery = SQLQuery + " AND lstStatus < 20" 'Bypass MG/Bonus
    SQLQuery = SQLQuery + " AND Mod(lstStatus, 100) < " & ASTEXTENDED_MG 'Bypass MG/Bonus
    'SQLQuery = SQLQuery + " AND (lstLogDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "')" & ")"
    SQLQuery = SQLQuery + " AND (lstLogDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(slEDate, sgSQLDateForm) & "')" & ")"
    SQLQuery = SQLQuery + " ORDER BY lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
        
    Set rst_Lst = gSQLSelectCall(SQLQuery)
    If Not rst_Lst.EOF Then
        If Not ilFromSave Then
            lgSelGameGsfCode = -1
            lacGame.Caption = ""
            llVeh = gBinarySearchVef(CLng(imVefCode))
            If llVeh <> -1 Then
                If tgVehicleInfo(llVeh).sVehType = "G" Then
                    igGameVefCode = imVefCode
                    sgGameStartDate = smFWkDate
                    sgGameEndDate = smLWkDate
                    lgGameAttCode = -1
                    frmGetGame.Show vbModal
                    If lgSelGameGsfCode <= 0 Then
                        mClearGrid
                        Screen.MousePointer = vbDefault
                        cmdDone.SetFocus
                        Exit Sub
                    End If
                    SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfCode = " & lgSelGameGsfCode & ")"
                    Set rst_Gsf = gSQLSelectCall(SQLQuery)
                    If Not rst_Gsf.EOF Then
                        slStr = ""
                        'Feed Source
                        If ((Asc(sgSpfSportInfo) And USINGFEED) = USINGFEED) Then
                            If rst_Gsf!gsfFeedSource = "V" Then
                                slStr = "Feed: Visting"
                            Else
                                slStr = "Feed: Home"
                            End If
                        End If
                        'Language
                        If ((Asc(sgSpfSportInfo) And USINGLANG) = USINGLANG) Then
                            For ilLang = LBound(tgLangInfo) To UBound(tgLangInfo) - 1 Step 1
                                If tgLangInfo(ilLang).iCode = rst_Gsf!gsfLangMnfCode Then
                                    If slStr = "" Then
                                        slStr = Trim$(tgLangInfo(ilLang).sName)
                                    Else
                                        slStr = slStr & " " & Trim$(tgLangInfo(ilLang).sName)
                                    End If
                                    Exit For
                                End If
                            Next ilLang
                        End If
                        'Visiting Team
                        For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
                            If tgTeamInfo(ilTeam).iCode = rst_Gsf!gsfVisitMnfCode Then
                                If slStr = "" Then
                                    slStr = Trim$(tgTeamInfo(ilTeam).sName)
                                Else
                                    slStr = slStr & " " & Trim$(tgTeamInfo(ilTeam).sName)
                                End If
                                Exit For
                            End If
                        Next ilTeam
                        'Home Team
                        For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
                            If tgTeamInfo(ilTeam).iCode = rst_Gsf!gsfHomeMnfCode Then
                                If slStr = "" Then
                                    slStr = Trim$(tgTeamInfo(ilTeam).sName)
                                Else
                                    slStr = slStr & " @ " & Trim$(tgTeamInfo(ilTeam).sName)
                                End If
                                Exit For
                            End If
                        Next ilTeam
                        'Air Date
                        slStr = slStr & " on " & Format$(rst_Gsf!gsfAirDate, sgShowDateForm)
                        'Start Time
                        slStr = slStr & " at " & Format$(rst_Gsf!gsfAirTime, sgShowTimeWSecForm)
                        lacGame.Caption = slStr
                    End If
                End If
            End If
        End If
        While Not rst_Lst.EOF
            iZone = False
            If (lgSelGameGsfCode <= 0) Or ((lgSelGameGsfCode = rst_Lst!lstGsfCode) And (lgSelGameGsfCode > 0)) Then
                Select Case UCase$(Left$(rst_Lst!lstZone, 1))
                    Case "E"
                        If chkZone(0).Value = 1 Then
                            iZone = True
                        End If
                    Case "C"
                        If chkZone(1).Value = 1 Then
                            iZone = True
                        End If
                    Case "M"
                        If chkZone(2).Value = 1 Then
                            iZone = True
                        End If
                    Case "P"
                        If chkZone(3).Value = 1 Then
                            iZone = True
                        End If
                    Case Else
                        iZone = True
                End Select
            End If
            If iZone Then
                If rst_Lst!lstType = 0 Then 'Spot
                    iFound = False
                    sAdvtName = ""
                    For iLoop = 0 To UBound(tgAdvtInfo) - 1 Step 1
                        If rst_Lst!lstAdfCode = tgAdvtInfo(iLoop).iCode Then
                            sAdvtName = Trim$(tgAdvtInfo(iLoop).sAdvtName)
                            Exit For
                        End If
                    Next iLoop
                    For iLoop = 0 To UBound(tmPostInfo) - 1 Step 1
                        If Second(rst_Lst!lstLogTime) <> 0 Then
                            sTime = Format$(rst_Lst!lstLogTime, sgShowTimeWSecForm)
                        Else
                            sTime = Format$(rst_Lst!lstLogTime, sgShowTimeWOSecForm)
                        End If
                        If tmPostInfo(iLoop).iType = 0 Then
                            iFound = False
                            If StrComp(Trim$(tmPostInfo(iLoop).sAdfName), Trim$(sAdvtName), 1) = 0 Then
                                If tmPostInfo(iLoop).lCntrNo = rst_Lst!lstCntrNo Then
                                    'Need to test date/time in each zone
                                    Select Case UCase$(Left$(rst_Lst!lstZone, 1))
                                        Case "E"
                                            If Trim$(tmPostInfo(iLoop).sDateZone(0)) = "" Then
                                                If Trim$(tmPostInfo(iLoop).sDateZone(1)) <> "" Then
                                                    If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(1))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                                        If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(1), False) Then
                                                            If (rst_Lst!lstSdfCode = 0) Or (tmPostInfo(iLoop).lSdfCode = rst_Lst!lstSdfCode) Then
                                                                iFound = True
                                                                iUpper = iLoop
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If Trim$(tmPostInfo(iLoop).sDateZone(2)) <> "" Then
                                                        If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(2))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                                            If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(2), False) Then
                                                                If (rst_Lst!lstSdfCode = 0) Or (tmPostInfo(iLoop).lSdfCode = rst_Lst!lstSdfCode) Then
                                                                    iFound = True
                                                                    iUpper = iLoop
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        If Trim$(tmPostInfo(iLoop).sDateZone(3)) <> "" Then
                                                            If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(3))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                                                If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(3), False) Then
                                                                    If (rst_Lst!lstSdfCode = 0) Or (tmPostInfo(iLoop).lSdfCode = rst_Lst!lstSdfCode) Then
                                                                        iFound = True
                                                                        iUpper = iLoop
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Case "C"
                                            If Trim$(tmPostInfo(iLoop).sDateZone(1)) = "" Then
                                                If Trim$(tmPostInfo(iLoop).sDateZone(0)) <> "" Then
                                                    If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(0))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                                        If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(0), False) Then
                                                            If (rst_Lst!lstSdfCode = 0) Or (tmPostInfo(iLoop).lSdfCode = rst_Lst!lstSdfCode) Then
                                                                iFound = True
                                                                iUpper = iLoop
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If Trim$(tmPostInfo(iLoop).sDateZone(2)) <> "" Then
                                                        If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(2))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                                            If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(2), False) Then
                                                                If (rst_Lst!lstSdfCode = 0) Or (tmPostInfo(iLoop).lSdfCode = rst_Lst!lstSdfCode) Then
                                                                    iFound = True
                                                                    iUpper = iLoop
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        If Trim$(tmPostInfo(iLoop).sDateZone(3)) <> "" Then
                                                            If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(3))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                                                If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(3), False) Then
                                                                    If (rst_Lst!lstSdfCode = 0) Or (tmPostInfo(iLoop).lSdfCode = rst_Lst!lstSdfCode) Then
                                                                        iFound = True
                                                                        iUpper = iLoop
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Case "M"
                                            If Trim$(tmPostInfo(iLoop).sDateZone(2)) = "" Then
                                                If Trim$(tmPostInfo(iLoop).sDateZone(0)) <> "" Then
                                                    If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(0))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                                        If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(0), False) Then
                                                            If (rst_Lst!lstSdfCode = 0) Or (tmPostInfo(iLoop).lSdfCode = rst_Lst!lstSdfCode) Then
                                                                iFound = True
                                                                iUpper = iLoop
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If Trim$(tmPostInfo(iLoop).sDateZone(1)) <> "" Then
                                                        If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(1))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                                            If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(1), False) Then
                                                                If (rst_Lst!lstSdfCode = 0) Or (tmPostInfo(iLoop).lSdfCode = rst_Lst!lstSdfCode) Then
                                                                    iFound = True
                                                                    iUpper = iLoop
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        If Trim$(tmPostInfo(iLoop).sDateZone(3)) <> "" Then
                                                            If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(3))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                                                If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(3), False) Then
                                                                    If (rst_Lst!lstSdfCode = 0) Or (tmPostInfo(iLoop).lSdfCode = rst_Lst!lstSdfCode) Then
                                                                        iFound = True
                                                                        iUpper = iLoop
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Case "P"
                                            If Trim$(tmPostInfo(iLoop).sDateZone(3)) = "" Then
                                                If Trim$(tmPostInfo(iLoop).sDateZone(0)) <> "" Then
                                                    If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(0))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                                        If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(0), False) Then
                                                            If (rst_Lst!lstSdfCode = 0) Or (tmPostInfo(iLoop).lSdfCode = rst_Lst!lstSdfCode) Then
                                                                iFound = True
                                                                iUpper = iLoop
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If Trim$(tmPostInfo(iLoop).sDateZone(1)) <> "" Then
                                                        If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(1))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                                            If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(1), False) Then
                                                                If (rst_Lst!lstSdfCode = 0) Or (tmPostInfo(iLoop).lSdfCode = rst_Lst!lstSdfCode) Then
                                                                    iFound = True
                                                                    iUpper = iLoop
                                                                    Exit For
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        If Trim$(tmPostInfo(iLoop).sDateZone(2)) <> "" Then
                                                            If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(2))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                                                If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(2), False) Then
                                                                    If (rst_Lst!lstSdfCode = 0) Or (tmPostInfo(iLoop).lSdfCode = rst_Lst!lstSdfCode) Then
                                                                        iFound = True
                                                                        iUpper = iLoop
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Case Else
                                    End Select
                                End If
                            End If
                        End If
                    Next iLoop
                    If sgSpfUseCartNo <> "N" Then
                        If (IsNull(rst_Lst!lstCart) Or Left$(rst_Lst!lstCart, 1) = Chr$(0)) And IsNull(rst_Lst!lstISCI) Then
                            sCartOrISCI = ""
                        Else
                            If IsNull(rst_Lst!lstCart) Or Left$(rst_Lst!lstCart, 1) = Chr$(0) Then
                                sCartOrISCI = Trim$(rst_Lst!lstISCI)
                            Else
                                sCartOrISCI = Trim$(rst_Lst!lstCart) & " " & Trim$(rst_Lst!lstISCI)
                            End If
                        End If
                    Else
                        If IsNull(rst_Lst!lstISCI) Then
                            sCartOrISCI = ""
                        Else
                            sCartOrISCI = Trim$(rst_Lst!lstISCI)
                        End If
                    End If
                    If Second(rst_Lst!lstLogTime) <> 0 Then
                        sTime = Format$(rst_Lst!lstLogTime, sgShowTimeWSecForm)
                    Else
                        sTime = Format$(rst_Lst!lstLogTime, sgShowTimeWOSecForm)
                    End If
                    iWkNo = rst_Lst!lstWkNo
                    iBreakNo = rst_Lst!lstBreakNo
                    iPositionNo = rst_Lst!lstPositionNo
                    iSeqNo = rst_Lst!lstSeqNo
                    If Not iFound Then
                        iUpper = UBound(tmPostInfo)
                        tmPostInfo(iUpper).iType = rst_Lst!lstType
                        tmPostInfo(iUpper).lSdfCode = rst_Lst!lstSdfCode
                        If IsNull(rst_Lst!lstProd) Then
                            tmPostInfo(iUpper).sProd = ""
                        Else
                            tmPostInfo(iUpper).sProd = rst_Lst!lstProd
                        End If
                        tmPostInfo(iUpper).lCntrNo = rst_Lst!lstCntrNo
                        tmPostInfo(iUpper).iLen = rst_Lst!lstLen
                        tmPostInfo(iUpper).iUnits = 0
                        If rst_Lst!lstStatus < 0 Then
                            tmPostInfo(iUpper).iStatus = 0
                        Else
                            tmPostInfo(iUpper).iStatus = rst_Lst!lstStatus
                        End If
                        tmPostInfo(iUpper).sAdfName = sAdvtName
                        tmPostInfo(iUpper).iAdfCode = rst_Lst!lstAdfCode
                        tmPostInfo(iUpper).iAgfCode = rst_Lst!lstAgfCode
                        tmPostInfo(iUpper).iAnfCode = rst_Lst!lstAnfCode
                        tmPostInfo(iUpper).iChgd = False
                        Select Case UCase$(Left$(rst_Lst!lstZone, 1))
                            Case "E"
                                tmPostInfo(iUpper).sDateZone(0) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(0) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(0) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(0) = rst_Lst!lstCifCode
                                tmPostInfo(iUpper).sCartZone(0) = sCartOrISCI
                                tmPostInfo(iUpper).iWkNoZone(0) = iWkNo
                                tmPostInfo(iUpper).iBreakNoZone(0) = iBreakNo
                                tmPostInfo(iUpper).iPositionNoZone(0) = iPositionNo
                                tmPostInfo(iUpper).iSeqNoZone(0) = iSeqNo
                                tmPostInfo(iUpper).sDateZone(1) = ""
                                tmPostInfo(iUpper).sTimeZone(1) = ""
                                tmPostInfo(iUpper).lLstCodeZone(1) = 0
                                tmPostInfo(iUpper).lCifZone(1) = 0
                                tmPostInfo(iUpper).sCartZone(1) = ""
                                tmPostInfo(iUpper).iWkNoZone(1) = 0
                                tmPostInfo(iUpper).iBreakNoZone(1) = 0
                                tmPostInfo(iUpper).iPositionNoZone(1) = 0
                                tmPostInfo(iUpper).iSeqNoZone(1) = 0
                                tmPostInfo(iUpper).sDateZone(2) = ""
                                tmPostInfo(iUpper).sTimeZone(2) = ""
                                tmPostInfo(iUpper).lLstCodeZone(2) = 0
                                tmPostInfo(iUpper).lCifZone(2) = 0
                                tmPostInfo(iUpper).sCartZone(2) = ""
                                tmPostInfo(iUpper).iWkNoZone(2) = 0
                                tmPostInfo(iUpper).iBreakNoZone(2) = 0
                                tmPostInfo(iUpper).iPositionNoZone(2) = 0
                                tmPostInfo(iUpper).iSeqNoZone(2) = 0
                                tmPostInfo(iUpper).sDateZone(3) = ""
                                tmPostInfo(iUpper).sTimeZone(3) = ""
                                tmPostInfo(iUpper).lLstCodeZone(3) = 0
                                tmPostInfo(iUpper).lCifZone(3) = 0
                                tmPostInfo(iUpper).sCartZone(3) = ""
                                tmPostInfo(iUpper).iWkNoZone(3) = 0
                                tmPostInfo(iUpper).iBreakNoZone(3) = 0
                                tmPostInfo(iUpper).iPositionNoZone(3) = 0
                                tmPostInfo(iUpper).iSeqNoZone(3) = 0
                                iZoneFd(0) = True
                            Case "C"
                                tmPostInfo(iUpper).sDateZone(1) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(1) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(1) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(1) = rst_Lst!lstCifCode
                                tmPostInfo(iUpper).sCartZone(1) = sCartOrISCI
                                tmPostInfo(iUpper).iWkNoZone(1) = iWkNo
                                tmPostInfo(iUpper).iBreakNoZone(1) = iBreakNo
                                tmPostInfo(iUpper).iPositionNoZone(1) = iPositionNo
                                tmPostInfo(iUpper).iSeqNoZone(1) = iSeqNo
                                tmPostInfo(iUpper).sDateZone(0) = ""
                                tmPostInfo(iUpper).sTimeZone(0) = ""
                                tmPostInfo(iUpper).lLstCodeZone(0) = 0
                                tmPostInfo(iUpper).lCifZone(0) = 0
                                tmPostInfo(iUpper).sCartZone(0) = ""
                                tmPostInfo(iUpper).iWkNoZone(0) = 0
                                tmPostInfo(iUpper).iBreakNoZone(0) = 0
                                tmPostInfo(iUpper).iPositionNoZone(0) = 0
                                tmPostInfo(iUpper).iSeqNoZone(0) = 0
                                tmPostInfo(iUpper).sDateZone(2) = ""
                                tmPostInfo(iUpper).sTimeZone(2) = ""
                                tmPostInfo(iUpper).lLstCodeZone(2) = 0
                                tmPostInfo(iUpper).lCifZone(2) = 0
                                tmPostInfo(iUpper).sCartZone(2) = ""
                                tmPostInfo(iUpper).iWkNoZone(2) = 0
                                tmPostInfo(iUpper).iBreakNoZone(2) = 0
                                tmPostInfo(iUpper).iPositionNoZone(2) = 0
                                tmPostInfo(iUpper).iSeqNoZone(2) = 0
                                tmPostInfo(iUpper).sDateZone(3) = ""
                                tmPostInfo(iUpper).sTimeZone(3) = ""
                                tmPostInfo(iUpper).lLstCodeZone(3) = 0
                                tmPostInfo(iUpper).lCifZone(3) = 0
                                tmPostInfo(iUpper).sCartZone(3) = ""
                                tmPostInfo(iUpper).iWkNoZone(3) = 0
                                tmPostInfo(iUpper).iBreakNoZone(3) = 0
                                tmPostInfo(iUpper).iPositionNoZone(3) = 0
                                tmPostInfo(iUpper).iSeqNoZone(3) = 0
                                iZoneFd(1) = True
                            Case "M"
                                tmPostInfo(iUpper).sDateZone(2) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(2) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(2) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(2) = rst_Lst!lstCifCode
                                tmPostInfo(iUpper).sCartZone(2) = sCartOrISCI
                                tmPostInfo(iUpper).iWkNoZone(2) = iWkNo
                                tmPostInfo(iUpper).iBreakNoZone(2) = iBreakNo
                                tmPostInfo(iUpper).iPositionNoZone(2) = iPositionNo
                                tmPostInfo(iUpper).iSeqNoZone(2) = iSeqNo
                                tmPostInfo(iUpper).sDateZone(0) = ""
                                tmPostInfo(iUpper).sTimeZone(0) = ""
                                tmPostInfo(iUpper).lLstCodeZone(0) = 0
                                tmPostInfo(iUpper).lCifZone(0) = 0
                                tmPostInfo(iUpper).sCartZone(0) = ""
                                tmPostInfo(iUpper).iWkNoZone(0) = 0
                                tmPostInfo(iUpper).iBreakNoZone(0) = 0
                                tmPostInfo(iUpper).iPositionNoZone(0) = 0
                                tmPostInfo(iUpper).iSeqNoZone(0) = 0
                                tmPostInfo(iUpper).sDateZone(1) = ""
                                tmPostInfo(iUpper).sTimeZone(1) = ""
                                tmPostInfo(iUpper).lLstCodeZone(1) = 0
                                tmPostInfo(iUpper).lCifZone(1) = 0
                                tmPostInfo(iUpper).sCartZone(1) = ""
                                tmPostInfo(iUpper).iWkNoZone(1) = 0
                                tmPostInfo(iUpper).iBreakNoZone(1) = 0
                                tmPostInfo(iUpper).iPositionNoZone(1) = 0
                                tmPostInfo(iUpper).iSeqNoZone(1) = 0
                                tmPostInfo(iUpper).sDateZone(3) = ""
                                tmPostInfo(iUpper).sTimeZone(3) = ""
                                tmPostInfo(iUpper).lLstCodeZone(3) = 0
                                tmPostInfo(iUpper).lCifZone(3) = 0
                                tmPostInfo(iUpper).sCartZone(3) = ""
                                tmPostInfo(iUpper).iWkNoZone(3) = 0
                                tmPostInfo(iUpper).iBreakNoZone(3) = 0
                                tmPostInfo(iUpper).iPositionNoZone(3) = 0
                                tmPostInfo(iUpper).iSeqNoZone(3) = 0
                                iZoneFd(2) = True
                            Case "P"
                                tmPostInfo(iUpper).sDateZone(3) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(3) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(3) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(3) = rst_Lst!lstCifCode
                                tmPostInfo(iUpper).sCartZone(3) = sCartOrISCI
                                tmPostInfo(iUpper).iWkNoZone(3) = iWkNo
                                tmPostInfo(iUpper).iBreakNoZone(3) = iBreakNo
                                tmPostInfo(iUpper).iPositionNoZone(3) = iPositionNo
                                tmPostInfo(iUpper).iSeqNoZone(3) = iSeqNo
                                tmPostInfo(iUpper).sDateZone(0) = ""
                                tmPostInfo(iUpper).sTimeZone(0) = ""
                                tmPostInfo(iUpper).lLstCodeZone(0) = 0
                                tmPostInfo(iUpper).lCifZone(0) = 0
                                tmPostInfo(iUpper).sCartZone(0) = ""
                                tmPostInfo(iUpper).iWkNoZone(0) = 0
                                tmPostInfo(iUpper).iBreakNoZone(0) = 0
                                tmPostInfo(iUpper).iPositionNoZone(0) = 0
                                tmPostInfo(iUpper).iSeqNoZone(0) = 0
                                tmPostInfo(iUpper).sDateZone(1) = ""
                                tmPostInfo(iUpper).sTimeZone(1) = ""
                                tmPostInfo(iUpper).lLstCodeZone(1) = 0
                                tmPostInfo(iUpper).lCifZone(1) = 0
                                tmPostInfo(iUpper).sCartZone(1) = ""
                                tmPostInfo(iUpper).iWkNoZone(1) = 0
                                tmPostInfo(iUpper).iBreakNoZone(1) = 0
                                tmPostInfo(iUpper).iPositionNoZone(1) = 0
                                tmPostInfo(iUpper).iSeqNoZone(1) = 0
                                tmPostInfo(iUpper).sDateZone(2) = ""
                                tmPostInfo(iUpper).sTimeZone(2) = ""
                                tmPostInfo(iUpper).lLstCodeZone(2) = 0
                                tmPostInfo(iUpper).lCifZone(2) = 0
                                tmPostInfo(iUpper).sCartZone(2) = ""
                                tmPostInfo(iUpper).iWkNoZone(2) = 0
                                tmPostInfo(iUpper).iBreakNoZone(2) = 0
                                tmPostInfo(iUpper).iPositionNoZone(2) = 0
                                tmPostInfo(iUpper).iSeqNoZone(2) = 0
                                iZoneFd(3) = True
                           Case Else
                                tmPostInfo(iUpper).sDateZone(0) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(0) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(0) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(0) = rst_Lst!lstCifCode
                                tmPostInfo(iUpper).sCartZone(0) = sCartOrISCI
                                tmPostInfo(iUpper).iWkNoZone(0) = iWkNo
                                tmPostInfo(iUpper).iBreakNoZone(0) = iBreakNo
                                tmPostInfo(iUpper).iPositionNoZone(0) = iPositionNo
                                tmPostInfo(iUpper).iSeqNoZone(0) = iSeqNo
                                tmPostInfo(iUpper).sDateZone(1) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(1) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(1) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(1) = rst_Lst!lstCifCode
                                tmPostInfo(iUpper).sCartZone(1) = sCartOrISCI
                                tmPostInfo(iUpper).iWkNoZone(1) = iWkNo
                                tmPostInfo(iUpper).iBreakNoZone(1) = iBreakNo
                                tmPostInfo(iUpper).iPositionNoZone(1) = iPositionNo
                                tmPostInfo(iUpper).iSeqNoZone(1) = iSeqNo
                                tmPostInfo(iUpper).sDateZone(2) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(2) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(2) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(2) = rst_Lst!lstCifCode
                                tmPostInfo(iUpper).sCartZone(2) = sCartOrISCI
                                tmPostInfo(iUpper).iWkNoZone(2) = iWkNo
                                tmPostInfo(iUpper).iBreakNoZone(2) = iBreakNo
                                tmPostInfo(iUpper).iPositionNoZone(2) = iPositionNo
                                tmPostInfo(iUpper).iSeqNoZone(2) = iSeqNo
                                tmPostInfo(iUpper).sDateZone(3) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(3) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(3) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(3) = rst_Lst!lstCifCode
                                tmPostInfo(iUpper).sCartZone(3) = sCartOrISCI
                                tmPostInfo(iUpper).iWkNoZone(3) = iWkNo
                                tmPostInfo(iUpper).iBreakNoZone(3) = iBreakNo
                                tmPostInfo(iUpper).iPositionNoZone(3) = iPositionNo
                                tmPostInfo(iUpper).iSeqNoZone(3) = iSeqNo
                        End Select
                        ReDim Preserve tmPostInfo(0 To iUpper + 1) As POSTINFO
                    Else
                        Select Case UCase$(Left$(rst_Lst!lstZone, 1))
                            Case "E"
                                tmPostInfo(iUpper).sDateZone(0) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(0) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(0) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(0) = rst_Lst!lstCifCode
                                tmPostInfo(iUpper).sCartZone(0) = sCartOrISCI
                                iZoneFd(0) = True
                            Case "C"
                                tmPostInfo(iUpper).sDateZone(1) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(1) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(1) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(1) = rst_Lst!lstCifCode
                                tmPostInfo(iUpper).sCartZone(1) = sCartOrISCI
                                iZoneFd(1) = True
                            Case "M"
                                tmPostInfo(iUpper).sDateZone(2) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(2) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(2) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(2) = rst_Lst!lstCifCode
                                tmPostInfo(iUpper).sCartZone(2) = sCartOrISCI
                                iZoneFd(2) = True
                            Case "P"
                                tmPostInfo(iUpper).sDateZone(3) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(3) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(3) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(3) = rst_Lst!lstCifCode
                                tmPostInfo(iUpper).sCartZone(3) = sCartOrISCI
                                iZoneFd(3) = True
                            Case Else
                                If tmPostInfo(iUpper).lLstCodeZone(0) <= 0 Then
                                    tmPostInfo(iUpper).sDateZone(0) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                    tmPostInfo(iUpper).sTimeZone(0) = sTime
                                    tmPostInfo(iUpper).lLstCodeZone(0) = rst_Lst!lstCode
                                    tmPostInfo(iUpper).lCifZone(0) = rst_Lst!lstCifCode
                                    tmPostInfo(iUpper).sCartZone(0) = sCartOrISCI
                                End If
                                If tmPostInfo(iUpper).lLstCodeZone(1) <= 0 Then
                                    tmPostInfo(iUpper).sDateZone(1) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                    tmPostInfo(iUpper).sTimeZone(1) = sTime
                                    tmPostInfo(iUpper).lLstCodeZone(1) = rst_Lst!lstCode
                                    tmPostInfo(iUpper).lCifZone(1) = rst_Lst!lstCifCode
                                    tmPostInfo(iUpper).sCartZone(1) = sCartOrISCI
                                End If
                                If tmPostInfo(iUpper).lLstCodeZone(2) <= 0 Then
                                    tmPostInfo(iUpper).sDateZone(2) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                    tmPostInfo(iUpper).sTimeZone(2) = sTime
                                    tmPostInfo(iUpper).lLstCodeZone(2) = rst_Lst!lstCode
                                    tmPostInfo(iUpper).lCifZone(2) = rst_Lst!lstCifCode
                                    tmPostInfo(iUpper).sCartZone(2) = sCartOrISCI
                                End If
                                If tmPostInfo(iUpper).lLstCodeZone(3) <= 0 Then
                                    tmPostInfo(iUpper).sDateZone(3) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                    tmPostInfo(iUpper).sTimeZone(3) = sTime
                                    tmPostInfo(iUpper).lLstCodeZone(3) = rst_Lst!lstCode
                                    tmPostInfo(iUpper).lCifZone(3) = rst_Lst!lstCifCode
                                    tmPostInfo(iUpper).sCartZone(3) = sCartOrISCI
                                End If
                        End Select
                    End If
                Else
                    iFound = False
                    If Second(rst_Lst!lstLogTime) <> 0 Then
                        sTime = Format$(rst_Lst!lstLogTime, sgShowTimeWSecForm)
                    Else
                        sTime = Format$(rst_Lst!lstLogTime, sgShowTimeWOSecForm)
                    End If
                    For iLoop = 0 To UBound(tmPostInfo) - 1 Step 1
                        If tmPostInfo(iLoop).iType = 1 Then  'Avail
                            Select Case UCase$(Left$(rst_Lst!lstZone, 1))
                                Case "E"
                                    If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(0))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                        If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(0), False) Then
                                            iFound = True
                                            iUpper = iLoop
                                        End If
                                    End If
                                Case "C"
                                    If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(1))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                        If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(1), False) Then
                                            iFound = True
                                            iUpper = iLoop
                                        End If
                                    End If
                                Case "M"
                                    If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(2))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                        If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(2), False) Then
                                            iFound = True
                                            iUpper = iLoop
                                        End If
                                    End If
                                Case "P"
                                    If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(3))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                        If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(3), False) Then
                                            iFound = True
                                            iUpper = iLoop
                                        End If
                                    End If
                                Case Else
                                    If DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(0))) = DateValue(gAdjYear(rst_Lst!lstLogDate)) Then
                                        If gTimeToLong(sTime, False) = gTimeToLong(tmPostInfo(iLoop).sTimeZone(0), False) Then
                                            iFound = True
                                            iUpper = iLoop
                                        End If
                                    End If
                            End Select
                        End If
                    Next iLoop
                    If Not iFound Then
                        iUpper = UBound(tmPostInfo)
                        tmPostInfo(iUpper).iType = rst_Lst!lstType
                        tmPostInfo(iUpper).lSdfCode = 0
                        tmPostInfo(iUpper).sProd = ""
                        tmPostInfo(iUpper).lCntrNo = 0
                        tmPostInfo(iUpper).iLen = rst_Lst!lstLen
                        tmPostInfo(iUpper).iUnits = rst_Lst!lstUnits
                        If rst_Lst!lstStatus < 0 Then
                            tmPostInfo(iUpper).iStatus = 0
                        Else
                            tmPostInfo(iUpper).iStatus = rst_Lst!lstStatus
                        End If
                        tmPostInfo(iUpper).sAdfName = ""
                        tmPostInfo(iUpper).iAdfCode = 0
                        tmPostInfo(iUpper).iAnfCode = rst_Lst!lstAnfCode
                        Select Case UCase$(Left$(rst_Lst!lstZone, 1))
                            Case "E"
                                tmPostInfo(iUpper).sDateZone(0) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(0) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(0) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(0) = 0
                                tmPostInfo(iUpper).sCartZone(0) = ""
                                tmPostInfo(iUpper).sDateZone(2) = ""
                                tmPostInfo(iUpper).sTimeZone(2) = ""
                                tmPostInfo(iUpper).lLstCodeZone(2) = 0
                                tmPostInfo(iUpper).lCifZone(2) = 0
                                tmPostInfo(iUpper).sCartZone(2) = ""
                                tmPostInfo(iUpper).sDateZone(1) = ""
                                tmPostInfo(iUpper).sTimeZone(1) = ""
                                tmPostInfo(iUpper).lLstCodeZone(1) = 0
                                tmPostInfo(iUpper).lCifZone(1) = 0
                                tmPostInfo(iUpper).sCartZone(1) = ""
                                tmPostInfo(iUpper).sDateZone(3) = ""
                                tmPostInfo(iUpper).sTimeZone(3) = ""
                                tmPostInfo(iUpper).lLstCodeZone(3) = 0
                                tmPostInfo(iUpper).lCifZone(3) = 0
                                tmPostInfo(iUpper).sCartZone(3) = ""
                            Case "C"
                                tmPostInfo(iUpper).sDateZone(1) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(1) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(1) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(1) = 0
                                tmPostInfo(iUpper).sCartZone(1) = ""
                                tmPostInfo(iUpper).sDateZone(0) = ""
                                tmPostInfo(iUpper).sTimeZone(0) = ""
                                tmPostInfo(iUpper).lLstCodeZone(0) = 0
                                tmPostInfo(iUpper).lCifZone(0) = 0
                                tmPostInfo(iUpper).sCartZone(0) = ""
                                tmPostInfo(iUpper).sDateZone(2) = ""
                                tmPostInfo(iUpper).sTimeZone(2) = ""
                                tmPostInfo(iUpper).lLstCodeZone(2) = 0
                                tmPostInfo(iUpper).lCifZone(2) = 0
                                tmPostInfo(iUpper).sCartZone(2) = ""
                                tmPostInfo(iUpper).sDateZone(3) = ""
                                tmPostInfo(iUpper).sTimeZone(3) = ""
                                tmPostInfo(iUpper).lLstCodeZone(3) = 0
                                tmPostInfo(iUpper).lCifZone(3) = 0
                                tmPostInfo(iUpper).sCartZone(3) = ""
                             Case "M"
                                tmPostInfo(iUpper).sDateZone(2) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(2) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(2) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(2) = 0
                                tmPostInfo(iUpper).sCartZone(2) = ""
                                tmPostInfo(iUpper).sDateZone(0) = ""
                                tmPostInfo(iUpper).sTimeZone(0) = ""
                                tmPostInfo(iUpper).lLstCodeZone(0) = 0
                                tmPostInfo(iUpper).lCifZone(0) = 0
                                tmPostInfo(iUpper).sCartZone(0) = ""
                                tmPostInfo(iUpper).sDateZone(1) = ""
                                tmPostInfo(iUpper).sTimeZone(1) = ""
                                tmPostInfo(iUpper).lLstCodeZone(1) = 0
                                tmPostInfo(iUpper).lCifZone(1) = 0
                                tmPostInfo(iUpper).sCartZone(1) = ""
                                tmPostInfo(iUpper).sDateZone(3) = ""
                                tmPostInfo(iUpper).sTimeZone(3) = ""
                                tmPostInfo(iUpper).lLstCodeZone(3) = 0
                                tmPostInfo(iUpper).lCifZone(3) = 0
                                tmPostInfo(iUpper).sCartZone(3) = ""
                            Case "P"
                                tmPostInfo(iUpper).sDateZone(3) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(3) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(3) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(3) = 0
                                tmPostInfo(iUpper).sCartZone(3) = ""
                                tmPostInfo(iUpper).sDateZone(0) = ""
                                tmPostInfo(iUpper).sTimeZone(0) = ""
                                tmPostInfo(iUpper).lLstCodeZone(0) = 0
                                tmPostInfo(iUpper).lCifZone(0) = 0
                                tmPostInfo(iUpper).sCartZone(0) = ""
                                tmPostInfo(iUpper).sDateZone(1) = ""
                                tmPostInfo(iUpper).sTimeZone(1) = ""
                                tmPostInfo(iUpper).lLstCodeZone(1) = 0
                                tmPostInfo(iUpper).lCifZone(1) = 0
                                tmPostInfo(iUpper).sCartZone(1) = ""
                                tmPostInfo(iUpper).sDateZone(2) = ""
                                tmPostInfo(iUpper).sTimeZone(2) = ""
                                tmPostInfo(iUpper).lLstCodeZone(2) = 0
                                tmPostInfo(iUpper).lCifZone(2) = 0
                                tmPostInfo(iUpper).sCartZone(2) = ""
                            Case Else
                                tmPostInfo(iUpper).sDateZone(0) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(0) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(0) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(0) = 0
                                tmPostInfo(iUpper).sCartZone(0) = ""
                                tmPostInfo(iUpper).sDateZone(1) = ""
                                tmPostInfo(iUpper).sTimeZone(1) = ""
                                tmPostInfo(iUpper).lLstCodeZone(1) = 0
                                tmPostInfo(iUpper).lCifZone(1) = 0
                                tmPostInfo(iUpper).sCartZone(1) = ""
                                tmPostInfo(iUpper).sDateZone(2) = ""
                                tmPostInfo(iUpper).sTimeZone(2) = ""
                                tmPostInfo(iUpper).lLstCodeZone(2) = 0
                                tmPostInfo(iUpper).lCifZone(2) = 0
                                tmPostInfo(iUpper).sCartZone(2) = ""
                                tmPostInfo(iUpper).sDateZone(3) = ""
                                tmPostInfo(iUpper).sTimeZone(3) = ""
                                tmPostInfo(iUpper).lLstCodeZone(3) = 0
                                tmPostInfo(iUpper).lCifZone(3) = 0
                                tmPostInfo(iUpper).sCartZone(3) = ""
                        End Select
                        ReDim Preserve tmPostInfo(0 To iUpper + 1) As POSTINFO
                    Else
                        Select Case UCase$(Left$(rst_Lst!lstZone, 1))
                            Case "E"
                                tmPostInfo(iUpper).sDateZone(0) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(0) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(0) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(0) = 0
                                tmPostInfo(iUpper).sCartZone(0) = ""
                            Case "C"
                                tmPostInfo(iUpper).sDateZone(1) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(1) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(1) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(1) = 0
                                tmPostInfo(iUpper).sCartZone(1) = ""
                            Case "M"
                                tmPostInfo(iUpper).sDateZone(2) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(2) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(2) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(2) = 0
                                tmPostInfo(iUpper).sCartZone(2) = ""
                            Case "P"
                                tmPostInfo(iUpper).sDateZone(3) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                tmPostInfo(iUpper).sTimeZone(3) = sTime
                                tmPostInfo(iUpper).lLstCodeZone(3) = rst_Lst!lstCode
                                tmPostInfo(iUpper).lCifZone(3) = 0
                                tmPostInfo(iUpper).sCartZone(3) = ""
                            Case Else
                                If tmPostInfo(iUpper).lLstCodeZone(0) <= 0 Then
                                    tmPostInfo(iUpper).sDateZone(0) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                    tmPostInfo(iUpper).sTimeZone(0) = sTime
                                    tmPostInfo(iUpper).lLstCodeZone(0) = rst_Lst!lstCode
                                    tmPostInfo(iUpper).lCifZone(0) = 0
                                    tmPostInfo(iUpper).sCartZone(0) = ""
                                End If
                                If tmPostInfo(iUpper).lLstCodeZone(1) <= 0 Then
                                    tmPostInfo(iUpper).sDateZone(1) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                    tmPostInfo(iUpper).sTimeZone(1) = sTime
                                    tmPostInfo(iUpper).lLstCodeZone(1) = rst_Lst!lstCode
                                    tmPostInfo(iUpper).lCifZone(1) = 0
                                    tmPostInfo(iUpper).sCartZone(1) = ""
                                End If
                                If tmPostInfo(iUpper).lLstCodeZone(2) <= 0 Then
                                    tmPostInfo(iUpper).sDateZone(2) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                    tmPostInfo(iUpper).sTimeZone(2) = sTime
                                    tmPostInfo(iUpper).lLstCodeZone(2) = rst_Lst!lstCode
                                    tmPostInfo(iUpper).lCifZone(2) = 0
                                    tmPostInfo(iUpper).sCartZone(2) = ""
                                End If
                                If tmPostInfo(iUpper).lLstCodeZone(3) <= 0 Then
                                    tmPostInfo(iUpper).sDateZone(3) = Format$(rst_Lst!lstLogDate, sgShowDateForm)
                                    tmPostInfo(iUpper).sTimeZone(3) = sTime
                                    tmPostInfo(iUpper).lLstCodeZone(3) = rst_Lst!lstCode
                                    tmPostInfo(iUpper).lCifZone(3) = 0
                                    tmPostInfo(iUpper).sCartZone(3) = ""
                                End If
                        End Select
                    End If
                End If
            End If
            rst_Lst.MoveNext
        Wend
    Else
        sVehicle = ""
        For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            If tgVehicleInfo(iLoop).iCode = imVefCode Then
                sVehicle = Trim$(tgVehicleInfo(iLoop).sVehicle)
                Exit For
            End If
        Next iLoop
        gMsgBox "No Log Spots Generated for " & smFWkDate & "-" & smLWkDate & " " & sVehicle, vbOKOnly + vbInformation
    End If
    rst_Lst.Close
    If sgSpfUseCartNo <> "N" Then
        sCartOrISCI = "Cart"
    Else
        sCartOrISCI = "ISCI"
    End If
    imNoZones = 0
    For iLoop = 0 To 3 Step 1
        If iZoneFd(iLoop) Then
            imNoZones = imNoZones + 1
        End If
    Next iLoop
    If imNoZones = 0 Then
        grdPost.TextMatrix(0, CARTESTINDEX) = sCartOrISCI & " #"
        mSetColumnWidths True, False, False, False
        optShow(0).Enabled = False
        optShow(1).Enabled = False
        optShow(2).Enabled = False
        optShow(3).Enabled = False
        optShow(0).Value = False
        optShow(1).Value = False
        optShow(2).Value = False
        optShow(3).Value = False
        fraShow.Visible = True
    Else
        iZone = False
        For iLoop = 0 To 3 Step 1
            If iZoneFd(iLoop) Then
                If Not iZone Then
                    optShow(iLoop).Value = True
                    iZone = True
                End If
                optShow(iLoop).Enabled = True
                Select Case iLoop
                    Case 0
                        grdPost.TextMatrix(0, CARTESTINDEX) = sCartOrISCI & " -EST"
                    Case 1
                        grdPost.TextMatrix(0, CARTCSTINDEX) = sCartOrISCI & " -CST"
                    Case 2
                        grdPost.TextMatrix(0, CARTMSTINDEX) = sCartOrISCI & " -MST"
                    Case 3
                        grdPost.TextMatrix(0, CARTPSTINDEX) = sCartOrISCI & " -PST"
                End Select
            Else
                optShow(iLoop).Enabled = False
            End If
        Next iLoop
        mSetColumnWidths iZoneFd(0), iZoneFd(1), iZoneFd(2), iZoneFd(3)
        fraShow.Visible = True
    End If
    ReDim tmCntrInfo(0 To 0) As CNTRINFO
    'SQLQuery = "SELECT chf.chfCntrNo, chf.chfProduct, chf.chfAdfCode, chf.chfCode from CHF_Contract_Header chf"
    SQLQuery = "SELECT chfCntrNo, chfProduct, chfAdfCode, chfCode"
    SQLQuery = SQLQuery & " FROM CHF_Contract_Header"
    SQLQuery = SQLQuery + " WHERE (chfEndDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery + " AND chfStartDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery & " AND chfDelete <> 'Y'"
    SQLQuery = SQLQuery + " AND (chfStatus = 'O'"
    SQLQuery = SQLQuery + " OR chfStatus = 'H'))"
    SQLQuery = SQLQuery + " ORDER BY chfAdfCode, chfCntrNo"
    Set rst_chf = gSQLSelectCall(SQLQuery)
    While Not rst_chf.EOF
        iUpper = UBound(tmCntrInfo)
        tmCntrInfo(iUpper).lCntrNo = rst_chf!chfCntrNo
        tmCntrInfo(iUpper).sProd = rst_chf!chfProduct
        tmCntrInfo(iUpper).lChfCode = rst_chf!chfCode
        tmCntrInfo(iUpper).iAdfCode = rst_chf!chfAdfCode
        ReDim Preserve tmCntrInfo(0 To iUpper + 1) As CNTRINFO
        rst_chf.MoveNext
    Wend
    rst_chf.Close
    For iLoop = 0 To UBound(tmPostInfo) - 1 Step 1
        iFound = False
        iInsert = -1
        For iTest = 0 To UBound(tmCntrInfo) - 1 Step 1
            If tmPostInfo(iLoop).iAdfCode <= tmCntrInfo(iTest).iAdfCode Then
                If tmPostInfo(iLoop).iAdfCode < tmCntrInfo(iTest).iAdfCode Then
                   iInsert = iTest
                Else
                    If tmPostInfo(iLoop).lCntrNo < tmCntrInfo(iTest).lCntrNo Then
                        iInsert = iTest
                    End If
                End If
                If tmPostInfo(iLoop).lCntrNo = tmCntrInfo(iTest).lCntrNo Then
                    iFound = True
                    Exit For
                End If
            End If
        Next iTest
        If Not iFound Then
            iInsert = iInsert + 1
            'SQLQuery = "SELECT chf.chfCntrNo, chf.chfProduct, chf.chfAdfCode, chf.chfCode from CHF_Contract_Header chf"
            'SQLQuery = SQLQuery + " WHERE (chf.chfCntrNo = " & tmPostInfo(iLoop).lCntrNo
            'SQLQuery = SQLQuery & " AND chf.chfDelete <> 'Y'"
            'SQLQuery = SQLQuery + " AND (chf.chfStatus = 'O'"
            'SQLQuery = SQLQuery + " OR chf.chfStatus = 'H'))"
            'SQLQuery = SQLQuery + " ORDER BY chf.chfCntRevNo Desc"
            'Set rst = gSQLSelectCall(SQLQuery)
            'If Not rst_Lst.EOF Then
            '    For iTest = UBound(tmCntrInfo) - 1 To iInsert Step -1
            '        tmCntrInfo(iTest + 1) = tmCntrInfo(iTest)
            '    Next iTest
            '    tmCntrInfo(iInsert).lCntrNo = rst_Lst!chfCntrNo
            '    tmCntrInfo(iInsert).sProd = rst_Lst!chfProduct
            '    tmCntrInfo(iInsert).lChfCode = rst_Lst!chfCode
            '    tmCntrInfo(iInsert).iAdfCode = rst_Lst!chfAdfCode
            '    ReDim Preserve tmCntrInfo(0 To UBound(tmCntrInfo) + 1) As CNTRINFO
            'End If
            For iTest = UBound(tmCntrInfo) - 1 To iInsert Step -1
                tmCntrInfo(iTest + 1) = tmCntrInfo(iTest)
            Next iTest
            tmCntrInfo(iInsert).lCntrNo = tmPostInfo(iLoop).lCntrNo
            tmCntrInfo(iInsert).sProd = tmPostInfo(iLoop).sProd
            tmCntrInfo(iInsert).lChfCode = 0
            tmCntrInfo(iInsert).iAdfCode = tmPostInfo(iLoop).iAdfCode
            tmCntrInfo(iInsert).iAgfCode = tmPostInfo(iLoop).iAgfCode
            ReDim Preserve tmCntrInfo(0 To UBound(tmCntrInfo) + 1) As CNTRINFO
        End If
    Next iLoop
    'Select primary zone
    imFromGetLst = True
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        If tgVehicleInfo(iLoop).iCode = imVefCode Then
            Select Case Left$(tgVehicleInfo(iLoop).sPrimaryZone, 1)
                Case "E"
                    If optShow(0).Value Then
                        mGridPaint True
                    Else
                        optShow(0).Value = True
                    End If
                Case "C"
                    If optShow(1).Value Then
                        mGridPaint True
                    Else
                        optShow(1).Value = True
                    End If
                Case "M"
                    If optShow(2).Value Then
                        mGridPaint True
                    Else
                        optShow(2).Value = True
                    End If
                Case "P"
                    If optShow(3).Value Then
                        mGridPaint True
                    Else
                        optShow(3).Value = True
                    End If
                Case Else
                    mGridPaint True
            End Select
            Exit For
        End If
    Next iLoop
    imFromGetLst = False
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmPostLog-mGetLst"
    Exit Sub
End Sub

Private Sub txtWeek_LostFocus()
    If (imVefCode > 0) And (gIsDate(txtWeek.Text) = True) Then
        Screen.MousePointer = vbHourglass
        mGetLst False
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mFillCntr(llRow As Long)
    Dim iLoop As Integer
    Dim sAdvtName As String
    Dim iAdfCode As Integer
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim llRowIndex As Long
    
    If imInFillCntr Then
        Exit Sub
    End If
    imInFillCntr = True
    On Error GoTo ErrHand
    sAdvtName = grdPost.TextMatrix(llRow, ADVTINDEX)
    llRowIndex = SendMessageByString(lbcAdvt.hwnd, LB_FINDSTRING, -1, sAdvtName)
    If llRowIndex >= 0 Then
        iAdfCode = lbcAdvt.ItemData(llRowIndex)
    Else
        iAdfCode = -1
    End If
        
    If (iAdfCode = imFillCntrAdfCode) Then
        imInFillCntr = False
        Exit Sub
    End If
    lbcCntr.Clear
    If iAdfCode = -1 Then
        imInFillCntr = False
        Exit Sub
    End If
    imFillCntrAdfCode = iAdfCode
    iStart = 0
    iEnd = UBound(tmCntrInfo) - 1
    If iAdfCode > tmCntrInfo(iEnd \ 2).iAdfCode Then
        iStart = iEnd \ 2
    End If
    For iLoop = iStart To iEnd Step 1
        If iAdfCode < tmCntrInfo(iLoop).iAdfCode Then
            Exit For
        End If
        If iAdfCode = tmCntrInfo(iLoop).iAdfCode Then
            lbcCntr.AddItem Trim$(tmCntrInfo(iLoop).lCntrNo) & " " & Trim$(tmCntrInfo(iLoop).sProd)
            lbcCntr.ItemData(lbcCntr.NewIndex) = iLoop
        End If
    Next iLoop
    imInFillCntr = False
    Exit Sub
    
ErrHand:
    imInFillCntr = False
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmPostLog-mFillCntr"
End Sub

Private Sub mFillCart(llRow As Long)
    Dim iLoop As Integer
    Dim sAdvtName As String
    Dim iAdfCode As Integer
    Dim iLen As Integer
    Dim sName As String
    Dim llRowIndex As Long
    
    If imInFillCart Then
        Exit Sub
    End If
    imInFillCart = True
    On Error GoTo ErrHand
    sAdvtName = grdPost.TextMatrix(llRow, ADVTINDEX)
    llRowIndex = SendMessageByString(lbcAdvt.hwnd, LB_FINDSTRING, -1, sAdvtName)
    If llRowIndex >= 0 Then
        iAdfCode = lbcAdvt.ItemData(llRowIndex)
    Else
        iAdfCode = -1
    End If
    If iAdfCode = imFillCartAdfCode Then
        imInFillCart = False
        Exit Sub
    End If
    lbcCart(0).Clear
    lbcCart(1).Clear
    lbcCart(2).Clear
    lbcCart(3).Clear
    If iAdfCode = -1 Then
        imInFillCart = False
        Exit Sub
    End If
    imFillCartAdfCode = iAdfCode
    ReDim tmCopyInfo(0 To 0) As COPYINFO
    iLen = Val(grdPost.TextMatrix(llRow, LENINDEX))
    If sgSpfUseCartNo = "N" Then
        'SQLQuery = "SELECT cif.cifName, cif.cifCut, cif.cifLen, cif.cifCode, cpf.cpfName, cpf.cpfISCI, cpf.cpfCode from CIF_Copy_Inventory cif, CPF_Copy_Prodct_ISCI cpf"
        SQLQuery = "SELECT cifName, cifCut, cifLen, cifCode, cpfName, cpfISCI, cpfCode"
        SQLQuery = SQLQuery & " FROM CIF_Copy_Inventory, "
        SQLQuery = SQLQuery & "CPF_Copy_Prodct_ISCI"
        SQLQuery = SQLQuery + " WHERE (cifAdfCode = " & iAdfCode
        SQLQuery = SQLQuery & " AND cifLen = " & iLen
        SQLQuery = SQLQuery + " AND cifRotEndDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery + " AND cifRotStartDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND cpfCode = cifcpfCode" & ")"
        SQLQuery = SQLQuery + " ORDER BY cpfISCI"
    Else
        'SQLQuery = "SELECT cif.cifName, cif.cifCut, cif.cifLen, cif.cifCode, cpf.cpfName, cpf.cpfISCI, cpf.cpfCode, mcf.mcfName, mcf.mcfPrefix from  CIF_Copy_Inventory cif, CPF_Copy_Prodct_ISCI cpf, MCF_Media_Code mcf"
        SQLQuery = "SELECT cifName, cifCut, cifLen, cifCode, cpfName, cpfISCI, cpfCode, mcfName, mcfPrefix"
        SQLQuery = SQLQuery & " FROM CIF_Copy_Inventory, "
        SQLQuery = SQLQuery & "CPF_Copy_Prodct_ISCI, "
        SQLQuery = SQLQuery & "MCF_Media_Code"
        SQLQuery = SQLQuery + " WHERE (cifAdfCode = " & iAdfCode
        SQLQuery = SQLQuery & " AND cifLen = " & iLen
        SQLQuery = SQLQuery + " AND cifRotEndDate >= '" & Format$(smFWkDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery + " AND cifRotStartDate <= '" & Format$(smLWkDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " AND cpfCode = cifcpfCode"
        SQLQuery = SQLQuery & " AND mcfCode = cifmcfCode" & ")"
        SQLQuery = SQLQuery + " ORDER BY cifName"
    End If
    Set rst_Cif = gSQLSelectCall(SQLQuery)
    While Not rst_Cif.EOF
        If sgSpfUseCartNo = "N" Then
            sName = Trim$(rst_Cif!cpfISCI)
        Else
            sName = Trim$(rst_Cif!mcfName) & Trim$(rst_Cif!cifName) & " " & Trim$(rst_Cif!cpfISCI)
        End If
        For iLoop = 5 To 8 Step 1
            lbcCart(iLoop - 5).AddItem Trim$(sName)
            lbcCart(iLoop - 5).ItemData(lbcCart(iLoop - 5).NewIndex) = UBound(tmCopyInfo)
            tmCopyInfo(UBound(tmCopyInfo)).lCifCode = rst_Cif!cifCode
            tmCopyInfo(UBound(tmCopyInfo)).lCpfCode = rst_Cif!cpfCode
            If sgSpfUseCartNo = "N" Then
                tmCopyInfo(UBound(tmCopyInfo)).sCart = ""
                tmCopyInfo(UBound(tmCopyInfo)).sISCI = sName
            Else
                tmCopyInfo(UBound(tmCopyInfo)).sCart = Trim$(rst_Cif!mcfName) & Trim$(rst_Cif!cifName)
                tmCopyInfo(UBound(tmCopyInfo)).sISCI = Trim$(rst_Cif!cpfISCI)
            End If
            ReDim Preserve tmCopyInfo(0 To UBound(tmCopyInfo) + 1) As COPYINFO
        Next iLoop
        rst_Cif.MoveNext
    Wend
    rst_Cif.Close
    imInFillCart = False
    Exit Sub
    
ErrHand:
    imInFillCart = False
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmPostLog-mFillCart"
End Sub

Private Function mPutLst(iRowNo As Integer) As Integer
    '
    '   iRowNo(I)- Row Number to Save (-1=All)
    '
    Dim iLoop As Integer
    Dim iZone As Integer
    Dim iCopyMsg As Integer
    Dim iCntrMsg As Integer
    Dim iRowChgd As Integer
    Dim iAdvtChgd As Integer
    Dim iCntrChgd As Integer
    Dim iCartChgd As Integer
    Dim iLenChgd As Integer
    Dim iRet As Integer
    Dim lLstCode As Long
    Dim iAdf As Integer
    Dim iAdfCode As Integer
    Dim sAdvtName As String
    Dim iPos As Integer
    Dim sStr As String
    Dim lChfCode As Long
    Dim iChf As Integer
    Dim sProd As String
    Dim iAgfCode As Integer
    Dim lCntrNo As Long
    Dim iOk As Integer
    Dim iCart As Integer
    Dim lCifCode As Long
    Dim lCpfCode As Long
    Dim sCart As String
    Dim sISCI As String
    Dim iZCount As Integer
    Dim iStatus As Integer
    Dim sStatus As String
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim iChgType As Integer
    Dim iIndex As Integer
    Dim iCntrInfo As Integer
    Dim llRow As Long
    Dim llTRow As Long
    Dim ilDelAvail As Integer
    
       
    If igPreOrPost = 0 Then
        If sgUstWin(3) <> "I" Then
            mPutLst = False
            Exit Function
        End If
    Else
        If sgUstWin(5) <> "I" Then
            mPutLst = False
            Exit Function
        End If
    End If
    
    
    llTRow = grdPost.TopRow
    If iRowNo = -1 Then
        iCopyMsg = False
        iCntrMsg = False
    Else
        iCopyMsg = True
        iCntrMsg = True
    End If
    grdPost.Redraw = False
    iStart = 0
    iEnd = UBound(tmPostInfo) - 1
    On Error GoTo ErrHand
    For iLoop = iStart To iEnd Step 1
        iChgType = False
        If (iRowNo = -1) Or (iLoop = iRowNo) Then
            llRow = iLoop + grdPost.FixedRows
            iStatus = -1
            sStatus = Trim$(grdPost.TextMatrix(llRow, STATUSINDEX))
            For iIndex = 0 To UBound(tmStatusTypes) Step 1
                If StrComp(sStatus, Trim$(tmStatusTypes(iIndex).sName), 1) = 0 Then
                    iStatus = tmStatusTypes(iIndex).iStatus
                    Exit For
                End If
            Next iIndex
            If (tmPostInfo(iLoop).iType = 0) Then
                iRowChgd = False
                iAdvtChgd = False
                iCntrChgd = False
                iCartChgd = False
                iLenChgd = False
                If StrComp(Trim$(grdPost.TextMatrix(llRow, ADVTINDEX)), Trim$(tmPostInfo(iLoop).sAdfName), 1) <> 0 Then
                    iRowChgd = True
                    iAdvtChgd = True
                End If
                
                If tmPostInfo(iLoop).iType = 0 Then
                    If tmPostInfo(iLoop).lCntrNo > 0 Then
                        sStr = Trim$(Str(tmPostInfo(iLoop).lCntrNo)) & " " & Trim$(tmPostInfo(iLoop).sProd)
                    Else
                        sStr = Trim$(tmPostInfo(iLoop).sProd)
                    End If
                Else
                    sStr = ""
                End If
                If StrComp(Trim$(grdPost.TextMatrix(llRow, CNTRNOINDEX)), sStr, 1) <> 0 Then
                'If Val(Trim$(grdPost.TextMatrix(llRow, CNTRNOINDEX))) <> tmPostInfo(iLoop).lCntrNo Then
                    iRowChgd = True
                    iCntrChgd = True
                End If
                If imNoZones = 0 Then
                    iZCount = 0
                    If StrComp(Trim$(grdPost.TextMatrix(llRow, CARTESTINDEX)), Trim$(tmPostInfo(iLoop).sCartZone(0)), 1) <> 0 Then
                        iRowChgd = True
                        iCartChgd = True
                    End If
                Else
                    iZCount = 3
                    For iZone = 0 To 3 Step 1
                        If optShow(iZone).Enabled Then
                            Select Case iZone
                                Case 0
                                    If StrComp(Trim$(grdPost.TextMatrix(llRow, CARTESTINDEX + iZone)), Trim$(tmPostInfo(iLoop).sCartZone(0)), 1) <> 0 Then
                                        iRowChgd = True
                                        iCartChgd = True
                                    End If
                                Case 1
                                    If StrComp(Trim$(grdPost.TextMatrix(llRow, CARTESTINDEX + iZone)), Trim$(tmPostInfo(iLoop).sCartZone(1)), 1) <> 0 Then
                                        iRowChgd = True
                                        iCartChgd = True
                                    End If
                                Case 2
                                    If StrComp(Trim$(grdPost.TextMatrix(llRow, CARTESTINDEX + iZone)), Trim$(tmPostInfo(iLoop).sCartZone(2)), 1) <> 0 Then
                                        iRowChgd = True
                                        iCartChgd = True
                                    End If
                                Case 3
                                    If StrComp(Trim$(grdPost.TextMatrix(llRow, CARTESTINDEX + iZone)), Trim$(tmPostInfo(iLoop).sCartZone(3)), 1) <> 0 Then
                                        iRowChgd = True
                                        iCartChgd = True
                                    End If
                            End Select
                        End If
                    Next iZone
                End If
                If Val(Trim$(grdPost.TextMatrix(llRow, LENINDEX))) <> tmPostInfo(iLoop).iLen Then
                    iRowChgd = True
                    iLenChgd = True
                End If
                If tmPostInfo(iLoop).iStatus <> iStatus Then
                    iRowChgd = True
                End If
                If tmPostInfo(iLoop).iChgd = True Then
                    iRowChgd = True
                End If
                'Check Time in case spots swapped
                If optShow(1).Value Then
                    If DateValue(gAdjYear(Trim$(grdPost.TextMatrix(llRow, DATEINDEX)))) <> DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(1))) Then
                        iRowChgd = True
                    End If
                    If gTimeToLong(Trim$(grdPost.TextMatrix(llRow, TIMEINDEX)), False) <> gTimeToLong(tmPostInfo(iLoop).sTimeZone(1), False) Then
                        iRowChgd = True
                    End If
                ElseIf optShow(2).Value Then
                    If DateValue(gAdjYear(Trim$(grdPost.TextMatrix(llRow, DATEINDEX)))) <> DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(2))) Then
                        iRowChgd = True
                    End If
                    If gTimeToLong(Trim$(grdPost.TextMatrix(llRow, TIMEINDEX)), False) <> gTimeToLong(tmPostInfo(iLoop).sTimeZone(2), False) Then
                        iRowChgd = True
                    End If
                ElseIf optShow(3).Value Then
                    If DateValue(gAdjYear(Trim$(grdPost.TextMatrix(llRow, DATEINDEX)))) <> DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(3))) Then
                        iRowChgd = True
                    End If
                    If gTimeToLong(Trim$(grdPost.TextMatrix(llRow, TIMEINDEX)), False) <> gTimeToLong(tmPostInfo(iLoop).sTimeZone(3), False) Then
                        iRowChgd = True
                    End If
                Else
                    If DateValue(gAdjYear(Trim$(grdPost.TextMatrix(llRow, DATEINDEX)))) <> DateValue(gAdjYear(tmPostInfo(iLoop).sDateZone(0))) Then
                        iRowChgd = True
                    End If
                    If gTimeToLong(Trim$(grdPost.TextMatrix(llRow, TIMEINDEX)), False) <> gTimeToLong(tmPostInfo(iLoop).sTimeZone(0), False) Then
                        iRowChgd = True
                    End If
                End If
            ElseIf tmPostInfo(iLoop).iType = 1 Then
'                If tmPostInfo(iLoop).iStatus <> iStatus Then
'                    iRowChgd = True
'                End If
'                If tmPostInfo(iLoop).iChgd = True Then
'                    iRowChgd = True
'                End If
                ilDelAvail = False
                sStr = grdPost.TextMatrix(llRow, LENINDEX)
                iPos = InStr(sStr, "/")
                If iPos > 0 Then
                    If (Val(Trim$(Left$(sStr, iPos - 1))) <> tmPostInfo(iLoop).iUnits) Then
                        iRowChgd = True
                        iLenChgd = True
                        If Val(Trim$(Left$(sStr, iPos - 1))) = 0 Then
                            ilDelAvail = True
                        End If
                    End If
                    If (Val(Trim$(Mid$(sStr, iPos + 1))) <> tmPostInfo(iLoop).iLen) Then
                        iRowChgd = True
                        iLenChgd = True
                        If Val(Trim$(Mid$(sStr, iPos + 1))) = 0 Then
                            ilDelAvail = True
                        End If
                    End If
                Else
                    If (Val(Trim$(grdPost.TextMatrix(llRow, LENINDEX))) <> tmPostInfo(iLoop).iUnits) Then
                        iRowChgd = True
                        iLenChgd = True
                        If Val(Trim$(Left$(sStr, iPos - 1))) = 0 Then
                            ilDelAvail = True
                        End If
                    End If
                End If
                If ilDelAvail Then
                    grdPost.TextMatrix(llRow, LENINDEX) = ""
                End If
            ElseIf tmPostInfo(iLoop).iType = 2 Then
                iRowChgd = True
                iAdvtChgd = True
                iCntrChgd = True
                iCartChgd = True
                iLenChgd = True
                If iStatus = -1 Then
                    iStatus = 0
                End If
            ElseIf tmPostInfo(iLoop).iType = 3 Then
                iRowChgd = False
                iAdvtChgd = False
                iCntrChgd = False
                iCartChgd = False
                iLenChgd = False
            End If
            'Don't add row if Advertiser missing- This was added instead of testing that
            'fields are defined.
            If (tmPostInfo(iLoop).iType = 0) Or (tmPostInfo(iLoop).iType = 2) Then
                If Trim$(grdPost.TextMatrix(llRow, ADVTINDEX)) = "" Then
                    iRowChgd = False
                End If
            End If
            If ((iRowChgd) And (iStatus >= 0) And (tmPostInfo(iLoop).iType <> 1)) Or ((iRowChgd) And (tmPostInfo(iLoop).iType = 1)) Then
                For iZone = 0 To iZCount Step 1
                    If optShow(iZone).Enabled Or (imNoZones = 0) Then
                        lLstCode = 0
                        Select Case iZone
                            Case 0
                                lLstCode = tmPostInfo(iLoop).lLstCodeZone(0)
                            Case 1
                                lLstCode = tmPostInfo(iLoop).lLstCodeZone(1)
                            Case 2
                                lLstCode = tmPostInfo(iLoop).lLstCodeZone(2)
                            Case 3
                                lLstCode = tmPostInfo(iLoop).lLstCodeZone(3)
                        End Select
                        If ((lLstCode > 0) And (tmPostInfo(iLoop).iType = 0)) Or ((lLstCode > 0) And (tmPostInfo(iLoop).iType = 2)) Then
                            sAdvtName = grdPost.TextMatrix(llRow, ADVTINDEX)
                            For iAdf = 0 To lbcAdvt.ListCount - 1 Step 1
                                If StrComp(sAdvtName, lbcAdvt.List(iAdf), 1) = 0 Then
                                    iAdfCode = lbcAdvt.ItemData(iAdf)
                                    Exit For
                                End If
                            Next iAdf
                            If iCntrChgd Then
                                sStr = Trim$(grdPost.TextMatrix(llRow, CNTRNOINDEX))
                                If sStr <> "" Then
                                    mFillCntr llRow
                                    sStr = Trim$(lbcCntr.Text)
                                    For iChf = 0 To lbcCntr.ListCount - 1 Step 1
                                        If StrComp(sStr, lbcCntr.List(iChf), 1) = 0 Then
                                            iCntrInfo = lbcCntr.ItemData(iChf)
                                            Exit For
                                        End If
                                    Next iChf
                                    
                                    'SQLQuery = "SELECT chf.chfCntrNo, chf.chfAgfCode, chf.chfProduct from CHF_Contract_Header chf"
                                    'SQLQuery = SQLQuery + " WHERE (chf.chfCode = " & lChfCode & ")"
                                    'Set rst = gSQLSelectCall(SQLQuery)
                                    'If Not rst.EOF Then
                                    '    lCntrNo = rst!chfCntrNo
                                    '    sProd = rst!chfProduct
                                    '    iAgfCode = rst!chfAgfCode
                                    '    iOk = True
                                    'Else
                                    '    iOk = False
                                    'End If
                                    lCntrNo = tmCntrInfo(iCntrInfo).lCntrNo
                                    sProd = tmCntrInfo(iCntrInfo).sProd
                                    iAgfCode = tmCntrInfo(iCntrInfo).iAgfCode
                                    iOk = True
                                Else
                                    lCntrNo = 0
                                    sProd = ""
                                    iAgfCode = 0
                                    iOk = True
                                End If
                            Else
                                iOk = True
                            End If
                            If iCartChgd Then
                                mFillCart llRow
                            End If
                            If iOk Then
                                SQLQuery = "Update lst SET "
                                'SQLQuery = SQLQuery & "lstType = 0" & ", "
                                If iAdvtChgd Then
                                    SQLQuery = SQLQuery & "lstAdfCode = " & iAdfCode & ", "
                                End If
                                If iCntrChgd Then
                                    SQLQuery = SQLQuery & "lstCntrNo = " & lCntrNo & ", "
                                    SQLQuery = SQLQuery & "lstAgfCode = " & iAgfCode & ", "
                                    SQLQuery = SQLQuery & "lstProd = '" & gFixQuote(sProd) & "', "
                                End If
                                If iAdvtChgd Or iCntrChgd Then
                                    'Blank out fields that are line dependent
                                    SQLQuery = SQLQuery & "lstLineNo = " & 0 & ", "
                                    SQLQuery = SQLQuery & "lstLnVefCode = " & 0 & ", "
                                    SQLQuery = SQLQuery & "lstStartDate = '" & Format$("1/1/1970", sgSQLDateForm) & "'" & ", "
                                    SQLQuery = SQLQuery & "lstEndDate = '" & Format$("1/1/1970", sgSQLDateForm) & "'" & ", "
                                    SQLQuery = SQLQuery & "lstMon = " & 0 & ", "
                                    SQLQuery = SQLQuery & "lstTue = " & 0 & ", "
                                    SQLQuery = SQLQuery & "lstWed = " & 0 & ", "
                                    SQLQuery = SQLQuery & "lstThu = " & 0 & ", "
                                    SQLQuery = SQLQuery & "lstFri = " & 0 & ", "
                                    SQLQuery = SQLQuery & "lstSat = " & 0 & ", "
                                    SQLQuery = SQLQuery & "lstSun = " & 0 & ", "
                                    SQLQuery = SQLQuery & "lstSpotsWk = " & 0 & ", "
                                    SQLQuery = SQLQuery & "lstPriceType = " & 1 & ", "
                                    SQLQuery = SQLQuery & "lstPrice = " & 0 & ", "
                                    SQLQuery = SQLQuery & "lstSpotType = " & 4 & ", "
                                    SQLQuery = SQLQuery & "lstDemo = ' '" & ", "
                                    SQLQuery = SQLQuery & "lstAud = " & 0 & ", "
                                End If
                                If iCartChgd Then
                                    'mFillCart
                                    If sgSpfUseCartNo = "N" Then
                                        sCart = " "
                                        sStr = grdPost.TextMatrix(llRow, CARTESTINDEX + iZone)
                                        'iPos = InStr(1, sStr, " ")
                                        'If iPos > 0 Then
                                        '    sISCI = Left$(sStr, iPos - 1)
                                        'Else
                                        '    sISCI = sStr
                                        'End If
                                        lCifCode = 0
                                        lCpfCode = 0
                                        sISCI = " "
                                        For iCart = 0 To lbcCart(iZone).ListCount - 1 Step 1
                                            If StrComp(sStr, lbcCart(iZone).List(iCart), 1) = 0 Then
                                                lCifCode = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).lCifCode
                                                lCpfCode = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).lCpfCode
                                                sISCI = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).sISCI
                                                Exit For
                                            End If
                                        Next iCart
                                    Else
                                        sStr = grdPost.TextMatrix(llRow, CARTESTINDEX + iZone)
                                        'iPos = InStr(1, sStr, " ")
                                        'If iPos > 0 Then
                                        '    sCart = Left$(sStr, iPos - 1)
                                        '    sISCI = Mid$(sStr, iPos + 1)
                                        'Else
                                        '    sCart = sStr
                                        'End If
                                        lCifCode = 0
                                        lCpfCode = 0
                                        sCart = " "
                                        sISCI = " "
                                        For iCart = 0 To lbcCart(iZone).ListCount - 1 Step 1
                                            If StrComp(sStr, lbcCart(iZone).List(iCart), 1) = 0 Then
                                                lCifCode = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).lCifCode
                                                lCpfCode = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).lCpfCode
                                                sCart = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).sCart
                                                sISCI = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).sISCI
                                                Exit For
                                            End If
                                        Next iCart
                                    End If
                                    SQLQuery = SQLQuery & "lstISCI = '" & gFixQuote(sISCI) & "', "
                                    SQLQuery = SQLQuery & "lstCart = '" & gFixQuote(sCart) & "', "
                                    SQLQuery = SQLQuery & "lstCifCode = " & lCifCode & ", "
                                    SQLQuery = SQLQuery & "lstCpfCode = " & lCpfCode & ", "
                                    SQLQuery = SQLQuery & "lstCrfCsfCode = " & 0 & ", "
                                    If imNoZones <= 0 Then
                                        SQLQuery = SQLQuery & "lstZone = '" & "   " & "', "
                                    Else
                                        Select Case iZone
                                            Case 0
                                                SQLQuery = SQLQuery & "lstZone = '" & "EST" & "', "
                                            Case 1
                                                SQLQuery = SQLQuery & "lstZone = '" & "CST" & "', "
                                            Case 2
                                                SQLQuery = SQLQuery & "lstZone = '" & "MST" & "', "
                                            Case 3
                                                SQLQuery = SQLQuery & "lstZone = '" & "PST" & "', "
                                        End Select
                                    End If
                                End If
                                SQLQuery = SQLQuery & "lstLen = " & Val(grdPost.TextMatrix(llRow, LENINDEX)) & ", "
                                SQLQuery = SQLQuery & "lstType = " & 0 & ", "
                                SQLQuery = SQLQuery & "lstStatus = " & iStatus & ", "
                                Select Case iZone
                                    Case 0
                                        SQLQuery = SQLQuery & "lstWkNo = " & tmPostInfo(iLoop).iWkNoZone(0) & ", "
                                        SQLQuery = SQLQuery & "lstBreakNo = " & tmPostInfo(iLoop).iBreakNoZone(0) & ", "
                                        SQLQuery = SQLQuery & "lstPositionNo = " & tmPostInfo(iLoop).iPositionNoZone(0) & ", "
                                        SQLQuery = SQLQuery & "lstSeqNo = " & tmPostInfo(iLoop).iSeqNoZone(0) & ", "
                                    Case 1
                                        SQLQuery = SQLQuery & "lstWkNo = " & tmPostInfo(iLoop).iWkNoZone(1) & ", "
                                        SQLQuery = SQLQuery & "lstBreakNo = " & tmPostInfo(iLoop).iBreakNoZone(1) & ", "
                                        SQLQuery = SQLQuery & "lstPositionNo = " & tmPostInfo(iLoop).iPositionNoZone(1) & ", "
                                        SQLQuery = SQLQuery & "lstSeqNo = " & tmPostInfo(iLoop).iSeqNoZone(1) & ", "
                                    Case 2
                                        SQLQuery = SQLQuery & "lstWkNo = " & tmPostInfo(iLoop).iWkNoZone(2) & ", "
                                        SQLQuery = SQLQuery & "lstBreakNo = " & tmPostInfo(iLoop).iBreakNoZone(2) & ", "
                                        SQLQuery = SQLQuery & "lstPositionNo = " & tmPostInfo(iLoop).iPositionNoZone(2) & ", "
                                        SQLQuery = SQLQuery & "lstSeqNo = " & tmPostInfo(iLoop).iSeqNoZone(2) & ", "
                                    Case 3
                                        SQLQuery = SQLQuery & "lstWkNo = " & tmPostInfo(iLoop).iWkNoZone(3) & ", "
                                        SQLQuery = SQLQuery & "lstBreakNo = " & tmPostInfo(iLoop).iBreakNoZone(3) & ", "
                                        SQLQuery = SQLQuery & "lstPositionNo = " & tmPostInfo(iLoop).iPositionNoZone(3) & ", "
                                        SQLQuery = SQLQuery & "lstSeqNo = " & tmPostInfo(iLoop).iSeqNoZone(3) & ", "
                                End Select
                                'Update date/time in case spot swapped
                                sStr = grdPost.TextMatrix(llRow, DATEINDEX)
                                SQLQuery = SQLQuery & "lstLogDate = '" & Format$(sStr, sgSQLDateForm) & "', "
                                sStr = grdPost.TextMatrix(llRow, TIMEINDEX)
                                SQLQuery = SQLQuery & "lstLogTime = '" & Format$(sStr, sgSQLTimeForm) & "'"
                                SQLQuery = SQLQuery & " WHERE (lstCode = " & lLstCode & ")"
                                cnn.BeginTrans
                                'cnn.Execute SQLQuery, rdExecDirect
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/12/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    Screen.MousePointer = vbDefault
                                    gHandleError "AffErrorLog.txt", "PostLog-mPutLst"
                                    cnn.RollbackTrans
                                    mPutLst = False
                                    Exit Function
                                End If
                                cnn.CommitTrans
                                sStr = grdPost.TextMatrix(llRow, DATEINDEX)
                                If iCntrChgd Then
                                    If Not gAlertFound("F", "S", imVefCode, sStr) Then
                                        iRet = gAlertAdd("R", "S", imVefCode, sStr)
                                    End If
                                End If
                                If iCartChgd Then
                                    If Not gAlertFound("F", "I", imVefCode, sStr) Then
                                        iRet = gAlertAdd("R", "I", imVefCode, sStr)
                                    End If
                                End If
                                mAddAbfDate grdPost.TextMatrix(llRow, DATEINDEX)
                            Else
                            End If
                        ElseIf (lLstCode > 0) And (tmPostInfo(iLoop).iType = 1) Then
                            If Not ilDelAvail Then
                                SQLQuery = "Update lst SET "
                                sStr = Trim$(grdPost.TextMatrix(llRow, LENINDEX))
                                iPos = InStr(sStr, "/")
                                If iPos > 0 Then
                                    SQLQuery = SQLQuery & "lstUnits = " & Val(Left$(sStr, iPos - 1)) & ", "
                                    SQLQuery = SQLQuery & "lstLen = " & Val(Mid$(sStr, iPos + 1)) & ", "
                                Else
                                    SQLQuery = SQLQuery & "lstUnits = " & Val(grdPost.TextMatrix(llRow, LENINDEX)) & ", "
                                End If
                                SQLQuery = SQLQuery & "lstStatus = " & iStatus & ", "
                                'Update date/time in case spot swapped
                                sStr = grdPost.TextMatrix(llRow, DATEINDEX)
                                SQLQuery = SQLQuery & "lstLogDate = '" & Format$(sStr, sgSQLDateForm) & "', "
                                sStr = grdPost.TextMatrix(llRow, TIMEINDEX)
                                SQLQuery = SQLQuery & "lstLogTime = '" & Format$(sStr, sgSQLTimeForm) & "'"
                                SQLQuery = SQLQuery & " WHERE (lstCode = " & lLstCode & ")"
                            Else
                                SQLQuery = "DELETE From LST " & " WHERE (lstCode = " & lLstCode & ")"
                            End If
                            cnn.BeginTrans
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/12/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "PostLog-mPutLst"
                                cnn.RollbackTrans
                                mPutLst = False
                                Exit Function
                            End If
                            cnn.CommitTrans
                            mAddAbfDate grdPost.TextMatrix(llRow, DATEINDEX)
                        ElseIf (lLstCode = 0) And (tmPostInfo(iLoop).iType = 2) Then
                            iChgType = True
                            sAdvtName = grdPost.TextMatrix(llRow, ADVTINDEX)
                            For iAdf = 0 To lbcAdvt.ListCount - 1 Step 1
                                If StrComp(sAdvtName, lbcAdvt.List(iAdf), 1) = 0 Then
                                    iAdfCode = lbcAdvt.ItemData(iAdf)
                                    Exit For
                                End If
                            Next iAdf
                            sStr = Trim$(grdPost.TextMatrix(llRow, CNTRNOINDEX))
                            If sStr <> "" Then
                                mFillCntr llRow
                                sStr = Trim$(grdPost.TextMatrix(llRow, CNTRNOINDEX))
                                For iChf = 0 To lbcCntr.ListCount - 1 Step 1
                                    If StrComp(sStr, lbcCntr.List(iChf), 1) = 0 Then
                                        iCntrInfo = lbcCntr.ItemData(iChf)
                                        Exit For
                                    End If
                                Next iChf
                                'SQLQuery = "SELECT chf.chfCntrNo, chf.chfAgfCode, chf.chfProduct from CHF_Contract_Header chf"
                                'SQLQuery = SQLQuery + " WHERE (chf.chfCode = " & lChfCode & ")"
                                'Set rst = gSQLSelectCall(SQLQuery)
                                'If Not rst.EOF Then
                                '    lCntrNo = rst!chfCntrNo
                                '    sProd = rst!chfProduct
                                '    iAgfCode = rst!chfAgfCode
                                '    iOk = True
                                'Else
                                '    iOk = False
                                'End If
                                lCntrNo = tmCntrInfo(iCntrInfo).lCntrNo
                                sProd = tmCntrInfo(iCntrInfo).sProd
                                iAgfCode = tmCntrInfo(iCntrInfo).iAgfCode
                                iOk = True
                            Else
                                lCntrNo = 0
                                sProd = ""
                                iAgfCode = 0
                                iOk = True
                            End If
                            mFillCart llRow
                            If sgSpfUseCartNo = "N" Then
                                sCart = " "
                                sStr = grdPost.TextMatrix(llRow, CARTESTINDEX + iZone)
                                'iPos = InStr(1, sStr, " ")
                                'If iPos > 0 Then
                                '    sISCI = Left$(sStr, iPos - 1)
                                'Else
                                '    sISCI = sStr
                                'End If
                                lCifCode = 0
                                lCpfCode = 0
                                sISCI = " "
                                For iCart = 0 To lbcCart(iZone).ListCount - 1 Step 1
                                    If StrComp(sStr, lbcCart(iZone).List(iCart), 1) = 0 Then
                                        lCifCode = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).lCifCode
                                        lCpfCode = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).lCpfCode
                                        sISCI = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).sISCI
                                        Exit For
                                    End If
                                Next iCart
                            Else
                                sStr = grdPost.TextMatrix(llRow, CARTESTINDEX + iZone)
                                'iPos = InStr(1, sStr, " ")
                                'If iPos > 0 Then
                                '    sCart = Left$(sStr, iPos - 1)
                                '    sISCI = Mid$(sStr, iPos + 1)
                                'Else
                                '    sCart = sStr
                                'End If
                                lCifCode = 0
                                lCpfCode = 0
                                sCart = " "
                                sISCI = " "
                                For iCart = 0 To lbcCart(iZone).ListCount - 1 Step 1
                                    If StrComp(sStr, lbcCart(iZone).List(iCart), 1) = 0 Then
                                        lCifCode = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).lCifCode
                                        lCpfCode = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).lCpfCode
                                        sCart = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).sCart
                                        sISCI = tmCopyInfo(lbcCart(iZone).ItemData(iCart)).sISCI
                                        Exit For
                                    End If
                                Next iCart
                            End If
                            If iOk Then
                                SQLQuery = "INSERT INTO lst (lstType, lstSdfCode, lstCntrNo, "
                                SQLQuery = SQLQuery & "lstAdfCode, lstAgfCode, lstProd, lstLineNo, lstLnVefCode, "
                                SQLQuery = SQLQuery & "lstStartDate, lstEndDate, lstMon, lstTue, lstWed, lstThu, lstFri, lstSat, lstSun, "
                                SQLQuery = SQLQuery & "lstSpotsWk, lstPriceType, lstPrice, lstSpotType, lstLogVefCode, "
                                SQLQuery = SQLQuery & "lstLogDate, lstLogTime, lstDemo, lstAud, lstISCI, "
                                SQLQuery = SQLQuery & "lstWkNo, lstBreakNo, lstPositionNo, lstSeqNo, lstZone, "
                                '12/28/06
                                'SQLQuery = SQLQuery & "lstCart, lstCpfCode, lstCrfCsfCode, lstStatus, lstLen, lstUnits, lstCifCode, lstAnfCode)"
                                SQLQuery = SQLQuery & "lstCart, lstCpfCode, lstCrfCsfCode, lstStatus, lstLen, lstUnits, lstCifCode, "
                                SQLQuery = SQLQuery & "lstAnfCode, lstEvtIDCefCode, lstSplitNetwork, "
                                'SQLQuery = SQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, lstUnused)"
                                SQLQuery = SQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, "
                                SQLQuery = SQLQuery & "lstLnStartTime, lstLnEndTime, lstUnused)"
                                SQLQuery = SQLQuery & " VALUES (" & 0 & ", " & 0 & ", " & lCntrNo & ", "
                                SQLQuery = SQLQuery & iAdfCode & ", " & iAgfCode & ", '" & gFixQuote(sProd) & "', " & 0 & ", " & 0 & ", "
                                SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', '" & Format$("12/31/2069", sgSQLDateForm) & "', " & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ", "
                                SQLQuery = SQLQuery & 0 & ", " & 1 & ", " & 0 & ", " & 2 & ", " & imVefCode & ", "
                                SQLQuery = SQLQuery & "'" & Format$(grdPost.TextMatrix(llRow, DATEINDEX), sgSQLDateForm) & "', '" & Format$(grdPost.TextMatrix(llRow, TIMEINDEX), sgSQLTimeForm) & "', '" & "" & "', " & 0 & ", '" & gFixQuote(sISCI) & "', "
                                Select Case iZone
                                    Case 0
                                        SQLQuery = SQLQuery & tmPostInfo(iLoop).iWkNoZone(0) & ", " & tmPostInfo(iLoop).iBreakNoZone(0) & ", " & tmPostInfo(iLoop).iPositionNoZone(0) & ", " & 0 & ", "
                                        If imNoZones <= 0 Then
                                            SQLQuery = SQLQuery & "'" & "" & "', "
                                        Else
                                            SQLQuery = SQLQuery & "'" & "EST" & "', "
                                        End If
                                    Case 1
                                        SQLQuery = SQLQuery & tmPostInfo(iLoop).iWkNoZone(1) & ", " & tmPostInfo(iLoop).iBreakNoZone(1) & ", " & tmPostInfo(iLoop).iPositionNoZone(1) & ", " & 0 & ", "
                                        If imNoZones <= 0 Then
                                            SQLQuery = SQLQuery & "'" & "" & "', "
                                        Else
                                            SQLQuery = SQLQuery & "'" & "CST" & "', "
                                        End If
                                    Case 2
                                        SQLQuery = SQLQuery & tmPostInfo(iLoop).iWkNoZone(2) & ", " & tmPostInfo(iLoop).iBreakNoZone(2) & ", " & tmPostInfo(iLoop).iPositionNoZone(2) & ", " & 0 & ", "
                                        If imNoZones <= 0 Then
                                            SQLQuery = SQLQuery & "'" & "" & "', "
                                        Else
                                            SQLQuery = SQLQuery & "'" & "MST" & "', "
                                        End If
                                    Case 3
                                        SQLQuery = SQLQuery & tmPostInfo(iLoop).iWkNoZone(3) & ", " & tmPostInfo(iLoop).iBreakNoZone(3) & ", " & tmPostInfo(iLoop).iPositionNoZone(3) & ", " & 0 & ", "
                                        If imNoZones <= 0 Then
                                            SQLQuery = SQLQuery & "'" & "" & "', "
                                        Else
                                            SQLQuery = SQLQuery & "'" & "PST" & "', "
                                        End If
                                End Select
                                '12/28/06
                                'SQLQuery = SQLQuery & "'" & gFixQuote(sCart) & "', " & lCpfCode & ", " & 0 & ", " & iStatus & ", " & Val(grdPost.TextMatrix(llRow, LENINDEX)) & ", " & 0 & ", " & lCifCode & ", " & tmPostInfo(iLoop).iAnfCode & ")"
                                SQLQuery = SQLQuery & "'" & gFixQuote(sCart) & "', " & lCpfCode & ", " & 0 & ", " & iStatus & ", " & Val(grdPost.TextMatrix(llRow, LENINDEX)) & ", " & 0 & ", " & lCifCode & ", "
                                SQLQuery = SQLQuery & tmPostInfo(iLoop).iAnfCode & ", " & 0 & ", '" & "N" & "', "
                                'SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "N" & "', " & 0 & ", '" & "" & "'" & ")"
                                SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "N" & "', " & 0 & ", "
                                SQLQuery = SQLQuery & "'" & Format("12am", sgSQLTimeForm) & "', '" & Format("12am", sgSQLTimeForm) & "', '" & "" & "'" & ")"
                                cnn.BeginTrans
                                'cnn.Execute SQLQuery, rdExecDirect
                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                    '6/12/16: Replaced GoSub
                                    'GoSub ErrHand:
                                    Screen.MousePointer = vbDefault
                                    gHandleError "AffErrorLog.txt", "PostLog-mPutLst"
                                    cnn.RollbackTrans
                                    mPutLst = False
                                    Exit Function
                                End If
                                cnn.CommitTrans
                                SQLQuery = "Select MAX(lstCode) from lst"
                                Set rst_Lst = gSQLSelectCall(SQLQuery)
                                Select Case iZone
                                    Case 0
                                        tmPostInfo(iLoop).lLstCodeZone(0) = rst_Lst(0).Value
                                    Case 1
                                        tmPostInfo(iLoop).lLstCodeZone(1) = rst_Lst(0).Value
                                    Case 2
                                        tmPostInfo(iLoop).lLstCodeZone(2) = rst_Lst(0).Value
                                    Case 3
                                        tmPostInfo(iLoop).lLstCodeZone(3) = rst_Lst(0).Value
                                End Select
                                rst_Lst.Close
                                sStr = grdPost.TextMatrix(llRow, DATEINDEX)
                                If Not gAlertFound("F", "S", imVefCode, sStr) Then
                                    iRet = gAlertAdd("R", "S", imVefCode, sStr)
                                End If
                                If sISCI <> "" Then
                                    If Not gAlertFound("F", "I", imVefCode, sStr) Then
                                        iRet = gAlertAdd("R", "I", imVefCode, sStr)
                                    End If
                                End If
                                mAddAbfDate grdPost.TextMatrix(llRow, DATEINDEX)
                            End If
                        End If
                    End If
                Next iZone
            End If
            If iRowNo <> -1 Then
                Exit For
            End If
        End If
        If iChgType Then
            tmPostInfo(iLoop).iType = 0
        End If
        llRow = llRow + 1
    Next iLoop
'Call mGetLst instead of resetting tmPostInfo because of delete and not getting all changes
'    'Reset array
'    For iLoop = iStart To iEnd Step 1
'        If (iRowNo = -1) Or (iLoop = iRowNo) Then
'            llRow = iLoop + grdPost.FixedRows
'            sStatus = Trim$(grdPost.TextMatrix(llRow, STATUSINDEX))
'            For iIndex = 0 To UBound(tmStatusTypes) Step 1
'                If StrComp(sStatus, Trim$(tmStatusTypes(iIndex).sName), 1) = 0 Then
'                    iStatus = tmStatusTypes(iIndex).iStatus
'                    Exit For
'                End If
'            Next iIndex
'            If (tmPostInfo(iLoop).iType = 0) Then
'                tmPostInfo(iLoop).sAdfName = grdPost.TextMatrix(llRow, ADVTINDEX)
'                tmPostInfo(iLoop).lCntrNo = Val(grdPost.TextMatrix(llRow, CNTRNOINDEX))
'                For iZone = 0 To 3 Step 1
'                    tmPostInfo(iLoop).sCartZone(iZone) = grdPost.TextMatrix(llRow, CARTESTINDEX + iZone)
'                Next iZone
'                tmPostInfo(iLoop).iLen = Val(grdPost.TextMatrix(llRow, LENINDEX))
'                tmPostInfo(iLoop).iStatus = iStatus
'                tmPostInfo(iLoop).iChgd = False
'                ilPostIndex = ilPostIndex + 1
'            ElseIf tmPostInfo(iLoop).iType = 1 Then
'                sStr = grdPost.TextMatrix(llRow, LENINDEX)
'                iPos = InStr(sStr, "/")
'                If iPos > 0 Then
'                    tmPostInfo(iLoop).iUnits = Val(Trim$(Left$(sStr, iPos - 1)))
'                    tmPostInfo(iLoop).iLen = Val(Trim$(Mid$(sStr, iPos + 1)))
'
'                Else
'                    tmPostInfo(iLoop).iUnits = Val(Trim$(grdPost.TextMatrix(llRow, LENINDEX)))
'                End If
'                tmPostInfo(iLoop).iStatus = iStatus
'                tmPostInfo(iLoop).iChgd = False
'            End If
'            If iRowNo <> -1 Then
'                Exit For
'            End If
'        End If
'        llRow = llRow + 1
'    Next iLoop
'    mGridPaint True
'    grdPost.TopRow = llTRow
    grdPost.Redraw = True
    mPutLst = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmPostLog-mPutLst"
    mPutLst = False
    Exit Function
End Function

Public Sub mSwap()
    Dim sMsg As String
    Dim ilRow2 As Integer
    Dim ilRow1 As Integer
    Dim sDate As String
    Dim sTime As String
    Dim iTRow As Integer
    Dim iLoop As Integer
    Dim iRow As Integer
    
    If grdPost.Row < grdPost.FixedRows Then
        Exit Sub
    End If
    imIgnoreChg = True
    ilRow1 = lmSwapStartRow - grdPost.FixedRows
    'grdPostLog.Columns(0).CellStyleSet ("Swap"), iToRow
    'Ask Swap Question
    If optShow(1).Value Then
        sDate = tmPostInfo(ilRow1).sDateZone(1)
        sTime = tmPostInfo(ilRow1).sTimeZone(1)
    ElseIf optShow(2).Value Then
        sDate = tmPostInfo(ilRow1).sDateZone(2)
        sTime = tmPostInfo(ilRow1).sTimeZone(2)
    ElseIf optShow(3).Value Then
        sDate = tmPostInfo(ilRow1).sDateZone(3)
        sTime = tmPostInfo(ilRow1).sTimeZone(3)
    Else
        sDate = tmPostInfo(ilRow1).sDateZone(0)
        sTime = tmPostInfo(ilRow1).sTimeZone(0)
    End If
    sMsg = "Swap " & sDate & " " & sTime & " " & Trim$(grdPost.TextMatrix(lmSwapStartRow, ADVTINDEX)) & " " & Trim$(grdPost.TextMatrix(lmSwapStartRow, CNTRNOINDEX)) & sgCRLF
    ilRow2 = grdPost.Row - grdPost.FixedRows
    If optShow(1).Value Then
        sDate = tmPostInfo(ilRow2).sDateZone(1)
        sTime = tmPostInfo(ilRow2).sTimeZone(1)
    ElseIf optShow(2).Value Then
        sDate = tmPostInfo(ilRow2).sDateZone(2)
        sTime = tmPostInfo(ilRow2).sTimeZone(2)
    ElseIf optShow(3).Value Then
        sDate = tmPostInfo(ilRow2).sDateZone(3)
        sTime = tmPostInfo(ilRow2).sTimeZone(3)
    Else
        sDate = tmPostInfo(ilRow2).sDateZone(0)
        sTime = tmPostInfo(ilRow2).sTimeZone(0)
    End If
    sMsg = sMsg & "With  " & sDate & " " & sTime & " " & Trim$(grdPost.TextMatrix(grdPost.Row, ADVTINDEX)) & " " & Trim$(grdPost.TextMatrix(grdPost.Row, CNTRNOINDEX))
    If gMsgBox(sMsg, vbYesNo) = vbNo Then
        mResetSwapColor
        imIgnoreChg = False
        Exit Sub
    End If
    tmToPostInfo = tmPostInfo(ilRow2)
    For iLoop = DATEINDEX To ORIGTYPEINDEX Step 1
        smToCols(iLoop) = Trim$(grdPost.TextMatrix(grdPost.Row, iLoop))
    Next iLoop
    tmFromPostInfo = tmPostInfo(ilRow1)
    For iLoop = DATEINDEX To ORIGTYPEINDEX Step 1
        smFromCols(iLoop) = Trim$(grdPost.TextMatrix(lmSwapStartRow, iLoop))
    Next iLoop
    grdPost.Redraw = False
    tmPostInfo(ilRow2) = tmFromPostInfo
    tmPostInfo(ilRow1) = tmToPostInfo
    tmPostInfo(ilRow2).iWkNoZone(0) = tmToPostInfo.iWkNoZone(0)
    tmPostInfo(ilRow2).iWkNoZone(1) = tmToPostInfo.iWkNoZone(1)
    tmPostInfo(ilRow2).iWkNoZone(2) = tmToPostInfo.iWkNoZone(2)
    tmPostInfo(ilRow2).iWkNoZone(3) = tmToPostInfo.iWkNoZone(3)
    tmPostInfo(ilRow2).iBreakNoZone(0) = tmToPostInfo.iBreakNoZone(0)
    tmPostInfo(ilRow2).iBreakNoZone(1) = tmToPostInfo.iBreakNoZone(1)
    tmPostInfo(ilRow2).iBreakNoZone(2) = tmToPostInfo.iBreakNoZone(2)
    tmPostInfo(ilRow2).iBreakNoZone(3) = tmToPostInfo.iBreakNoZone(3)
    tmPostInfo(ilRow2).iPositionNoZone(0) = tmToPostInfo.iPositionNoZone(0)
    tmPostInfo(ilRow2).iPositionNoZone(1) = tmToPostInfo.iPositionNoZone(1)
    tmPostInfo(ilRow2).iPositionNoZone(2) = tmToPostInfo.iPositionNoZone(2)
    tmPostInfo(ilRow2).iPositionNoZone(3) = tmToPostInfo.iPositionNoZone(3)
    tmPostInfo(ilRow2).iSeqNoZone(0) = tmToPostInfo.iSeqNoZone(0)
    tmPostInfo(ilRow2).iSeqNoZone(1) = tmToPostInfo.iSeqNoZone(1)
    tmPostInfo(ilRow2).iSeqNoZone(2) = tmToPostInfo.iSeqNoZone(2)
    tmPostInfo(ilRow2).iSeqNoZone(3) = tmToPostInfo.iSeqNoZone(3)
    tmPostInfo(ilRow2).iChgd = True

    tmPostInfo(ilRow1).iWkNoZone(0) = tmFromPostInfo.iWkNoZone(0)
    tmPostInfo(ilRow1).iWkNoZone(1) = tmFromPostInfo.iWkNoZone(1)
    tmPostInfo(ilRow1).iWkNoZone(2) = tmFromPostInfo.iWkNoZone(2)
    tmPostInfo(ilRow1).iWkNoZone(3) = tmFromPostInfo.iWkNoZone(3)
    tmPostInfo(ilRow1).iBreakNoZone(0) = tmFromPostInfo.iBreakNoZone(0)
    tmPostInfo(ilRow1).iBreakNoZone(1) = tmFromPostInfo.iBreakNoZone(1)
    tmPostInfo(ilRow1).iBreakNoZone(2) = tmFromPostInfo.iBreakNoZone(2)
    tmPostInfo(ilRow1).iBreakNoZone(3) = tmFromPostInfo.iBreakNoZone(3)
    tmPostInfo(ilRow1).iPositionNoZone(0) = tmFromPostInfo.iPositionNoZone(0)
    tmPostInfo(ilRow1).iPositionNoZone(1) = tmFromPostInfo.iPositionNoZone(1)
    tmPostInfo(ilRow1).iPositionNoZone(2) = tmFromPostInfo.iPositionNoZone(2)
    tmPostInfo(ilRow1).iPositionNoZone(3) = tmFromPostInfo.iPositionNoZone(3)
    tmPostInfo(ilRow1).iSeqNoZone(0) = tmFromPostInfo.iSeqNoZone(0)
    tmPostInfo(ilRow1).iSeqNoZone(1) = tmFromPostInfo.iSeqNoZone(1)
    tmPostInfo(ilRow1).iSeqNoZone(2) = tmFromPostInfo.iSeqNoZone(2)
    tmPostInfo(ilRow1).iSeqNoZone(3) = tmFromPostInfo.iSeqNoZone(3)
    tmPostInfo(ilRow1).iChgd = True
    'Cycle to first row and set
    'For iLoop = ADVTINDEX To ORIGTYPEINDEX Step 1
    For iLoop = STATUSINDEX To ORIGTYPEINDEX Step 1
        grdPost.TextMatrix(grdPost.Row, iLoop) = smFromCols(iLoop)
    Next iLoop
    'For iLoop = ADVTINDEX To ORIGTYPEINDEX Step 1
    For iLoop = STATUSINDEX To ORIGTYPEINDEX Step 1
        grdPost.TextMatrix(lmSwapStartRow, iLoop) = smToCols(iLoop)
    Next iLoop
    mResetSwapColor
    grdPost.Redraw = True
    imFieldChgd = True
    'D.S. Correct timing issue
    DoEvents
    imIgnoreChg = False
    pbcClickFocus.SetFocus
End Sub

Private Function mInsert() As Integer
    Dim sDate As String
    Dim sTime As String
    Dim iLoop As Integer
    Dim iIndex As Integer
    Dim iTRow As Integer
    Dim sMsg As String
    Dim llTRow As Long
    Dim llRow As Long
    Dim ilType As Integer
    
    imIgnoreChg = True
    llTRow = grdPost.TopRow
    llRow = grdPost.Row
    If optShow(1).Value Then
        sDate = tmPostInfo(llRow - grdPost.FixedRows).sDateZone(1)
        sTime = tmPostInfo(llRow - grdPost.FixedRows).sTimeZone(1)
    ElseIf optShow(2).Value Then
        sDate = tmPostInfo(llRow - grdPost.FixedRows).sDateZone(2)
        sTime = tmPostInfo(llRow - grdPost.FixedRows).sTimeZone(2)
    ElseIf optShow(3).Value Then
        sDate = tmPostInfo(llRow - grdPost.FixedRows).sDateZone(3)
        sTime = tmPostInfo(llRow - grdPost.FixedRows).sTimeZone(3)
    Else
        sDate = tmPostInfo(llRow - grdPost.FixedRows).sDateZone(0)
        sTime = tmPostInfo(llRow - grdPost.FixedRows).sTimeZone(0)
    End If
    ilType = grdPost.TextMatrix(grdPost.Row, ORIGTYPEINDEX)
    If ilType <> 1 Then
        sMsg = "Insert on " & Trim$(sDate) & " at " & Trim$(sTime) & " after " & Trim$(grdPost.TextMatrix(llRow, ADVTINDEX)) & " " & Trim$(grdPost.TextMatrix(llRow, CNTRNOINDEX))
    Else
        sMsg = "Insert on " & Trim$(sDate) & " at " & Trim$(sTime) & " after Avail"
    End If
    If gMsgBox(sMsg, vbYesNo) = vbNo Then
        imIgnoreChg = False
        mInsert = False
        Exit Function
    End If
    grdPost.Redraw = False
    grdPost.AddItem sDate & vbTab & sTime, llRow + 1
    iIndex = llRow + 1 - grdPost.FixedRows
    grdPost.TextMatrix(llRow + 1, POSTINDEX) = iIndex
    grdPost.TextMatrix(llRow + 1, ORIGTYPEINDEX) = 0
    grdPost.Row = llRow + 1
    grdPost.Col = DATEINDEX
    grdPost.CellBackColor = LIGHTYELLOW
    grdPost.Col = TIMEINDEX
    grdPost.CellBackColor = LIGHTYELLOW
    For iLoop = UBound(tmPostInfo) To iIndex + 1 Step -1
        tmPostInfo(iLoop) = tmPostInfo(iLoop - 1)
    Next iLoop
    tmPostInfo(iIndex).iType = 3        'Temporary set to 2 which indicates that inserted but advertiser and PositionNo within break not set
    tmPostInfo(iIndex).lSdfCode = 0
    tmPostInfo(iIndex).sProd = ""
    tmPostInfo(iIndex).lCntrNo = 0
    tmPostInfo(iIndex).iLen = 0
    tmPostInfo(iIndex).iUnits = 0
    tmPostInfo(iIndex).iStatus = 0
    tmPostInfo(iIndex).sAdfName = ""
    tmPostInfo(iIndex).iAnfCode = 0
    tmPostInfo(iIndex).sDateZone(0) = sDate
    tmPostInfo(iIndex).sTimeZone(0) = sTime
    tmPostInfo(iIndex).lLstCodeZone(0) = 0
    tmPostInfo(iIndex).lCifZone(0) = 0
    tmPostInfo(iIndex).sCartZone(0) = ""
    tmPostInfo(iIndex).sDateZone(1) = ""
    tmPostInfo(iIndex).sTimeZone(1) = ""
    tmPostInfo(iIndex).lLstCodeZone(1) = 0
    tmPostInfo(iIndex).lCifZone(1) = 0
    tmPostInfo(iIndex).sCartZone(1) = ""
    tmPostInfo(iIndex).sDateZone(2) = ""
    tmPostInfo(iIndex).sTimeZone(2) = ""
    tmPostInfo(iIndex).lLstCodeZone(2) = 0
    tmPostInfo(iIndex).lCifZone(2) = 0
    tmPostInfo(iIndex).sCartZone(2) = ""
    tmPostInfo(iIndex).sDateZone(3) = ""
    tmPostInfo(iIndex).sTimeZone(3) = ""
    tmPostInfo(iIndex).lLstCodeZone(3) = 0
    tmPostInfo(iIndex).lCifZone(3) = 0
    tmPostInfo(iIndex).sCartZone(3) = ""
    If (iIndex > 0) And ((tmPostInfo(iIndex - 1).iType = 0) Or (tmPostInfo(iIndex - 1).iType = 2)) Then
        tmPostInfo(iIndex).iAnfCode = tmPostInfo(iIndex - 1).iAnfCode
        tmPostInfo(iIndex).iWkNoZone(0) = tmPostInfo(iIndex - 1).iWkNoZone(0)
        tmPostInfo(iIndex).iBreakNoZone(0) = tmPostInfo(iIndex - 1).iBreakNoZone(0)
        tmPostInfo(iIndex).iPositionNoZone(0) = tmPostInfo(iIndex - 1).iPositionNoZone(0) + 1
        tmPostInfo(iIndex).iWkNoZone(1) = tmPostInfo(iIndex - 1).iWkNoZone(1)
        tmPostInfo(iIndex).iBreakNoZone(1) = tmPostInfo(iIndex - 1).iBreakNoZone(1)
        tmPostInfo(iIndex).iPositionNoZone(1) = tmPostInfo(iIndex - 1).iPositionNoZone(1) + 1
        tmPostInfo(iIndex).iWkNoZone(2) = tmPostInfo(iIndex - 1).iWkNoZone(2)
        tmPostInfo(iIndex).iBreakNoZone(2) = tmPostInfo(iIndex - 1).iBreakNoZone(2)
        tmPostInfo(iIndex).iPositionNoZone(2) = tmPostInfo(iIndex - 1).iPositionNoZone(2) + 1
        tmPostInfo(iIndex).iWkNoZone(3) = tmPostInfo(iIndex - 1).iWkNoZone(3)
        tmPostInfo(iIndex).iBreakNoZone(3) = tmPostInfo(iIndex - 1).iBreakNoZone(3)
        tmPostInfo(iIndex).iPositionNoZone(3) = tmPostInfo(iIndex - 1).iPositionNoZone(3) + 1
    ElseIf (iIndex > 0) And (tmPostInfo(iIndex - 1).iType = 1) Then
        tmPostInfo(iIndex).iAnfCode = tmPostInfo(iIndex - 1).iAnfCode
        tmPostInfo(iIndex).iWkNoZone(0) = tmPostInfo(iIndex - 1).iWkNoZone(0)
        tmPostInfo(iIndex).iBreakNoZone(0) = 0
        tmPostInfo(iIndex).iPositionNoZone(0) = 0
        tmPostInfo(iIndex).iWkNoZone(1) = tmPostInfo(iIndex - 1).iWkNoZone(1)
        tmPostInfo(iIndex).iBreakNoZone(1) = 0
        tmPostInfo(iIndex).iPositionNoZone(1) = 0
        tmPostInfo(iIndex).iWkNoZone(2) = tmPostInfo(iIndex - 1).iWkNoZone(2)
        tmPostInfo(iIndex).iBreakNoZone(2) = 0
        tmPostInfo(iIndex).iPositionNoZone(2) = 0
        tmPostInfo(iIndex).iWkNoZone(3) = tmPostInfo(iIndex - 1).iWkNoZone(3)
        tmPostInfo(iIndex).iBreakNoZone(3) = 0
        tmPostInfo(iIndex).iPositionNoZone(3) = 0
    Else
        tmPostInfo(iIndex).iAnfCode = tmPostInfo(iIndex + 1).iAnfCode
        tmPostInfo(iIndex).iWkNoZone(0) = tmPostInfo(iIndex + 1).iWkNoZone(0)
        tmPostInfo(iIndex).iBreakNoZone(0) = 0
        tmPostInfo(iIndex).iPositionNoZone(0) = 0
        tmPostInfo(iIndex).iWkNoZone(1) = tmPostInfo(iIndex + 1).iWkNoZone(1)
        tmPostInfo(iIndex).iBreakNoZone(1) = 0
        tmPostInfo(iIndex).iPositionNoZone(1) = 0
        tmPostInfo(iIndex).iWkNoZone(2) = tmPostInfo(iIndex + 1).iWkNoZone(2)
        tmPostInfo(iIndex).iBreakNoZone(2) = 0
        tmPostInfo(iIndex).iPositionNoZone(2) = 0
        tmPostInfo(iIndex).iWkNoZone(3) = tmPostInfo(iIndex + 1).iWkNoZone(3)
        tmPostInfo(iIndex).iBreakNoZone(3) = 0
        tmPostInfo(iIndex).iPositionNoZone(3) = 0
    End If
    ReDim Preserve tmPostInfo(0 To UBound(tmPostInfo) + 1) As POSTINFO
    grdPost.Redraw = False
    For iLoop = 0 To UBound(tmPostInfo) - 1 Step 1
        If iLoop > iIndex Then
            grdPost.TextMatrix(iLoop + grdPost.FixedRows, POSTINDEX) = CInt(grdPost.TextMatrix(iLoop + grdPost.FixedRows, POSTINDEX)) + 1
        End If
    Next iLoop
    grdPost.TopRow = llTRow
    grdPost.Redraw = True
    'D.S. 7/23/01 Correct timing issue with DoEvents below
    DoEvents
    grdPost.Row = llRow + 1
    'grdPost.Col = ADVTINDEX
    grdPost.Col = STATUSINDEX
    mPostEnableBox
    
    imIgnoreChg = False

End Function

Private Sub mPopVehBox()

    Dim iLoop As Integer
    
    cboSort.Clear
    
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        If rbcVeh(0).Value = True Then
            If tgVehicleInfo(iLoop).sState = "A" Then
                cboSort.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
                cboSort.ItemData(cboSort.NewIndex) = tgVehicleInfo(iLoop).iCode
            End If
        Else
            cboSort.AddItem Trim$(tgVehicleInfo(iLoop).sVehicle)
            cboSort.ItemData(cboSort.NewIndex) = tgVehicleInfo(iLoop).iCode
        End If
    Next iLoop

End Sub

Private Sub mSetColumnWidths(ilESTZones As Integer, ilCSTZones As Integer, ilMSTZones As Integer, ilPSTZones As Integer)
    grdPost.ColWidth(LSTCODEESTINDEX) = 0
    grdPost.ColWidth(LSTCODECSTINDEX) = 0
    grdPost.ColWidth(LSTCODEMSTINDEX) = 0
    grdPost.ColWidth(LSTCODEPSTINDEX) = 0
    grdPost.ColWidth(POSTINDEX) = 0
    grdPost.ColWidth(ORIGTYPEINDEX) = 0
    grdPost.ColWidth(DATEINDEX) = grdPost.Width * 0.09    'Date
    grdPost.ColWidth(TIMEINDEX) = grdPost.Width * 0.09    'Time
    
    grdPost.ColWidth(LENINDEX) = grdPost.Width * 0.05    'Length
    lmMaxCol = LENINDEX
    If ilESTZones Then
        grdPost.ColWidth(CARTESTINDEX) = grdPost.Width * 0.12     'Cart-CST
        lmMaxCol = lmMaxCol + 1
    Else
        grdPost.ColWidth(CARTESTINDEX) = 0    'Cart-CST
        lmMaxCol = lmMaxCol + 1
    End If
    If ilCSTZones Then
        grdPost.ColWidth(CARTCSTINDEX) = grdPost.Width * 0.12     'Cart-CST
        lmMaxCol = lmMaxCol + 1
    Else
        grdPost.ColWidth(CARTCSTINDEX) = 0    'Cart-CST
    End If
    If ilMSTZones Then
        grdPost.ColWidth(CARTMSTINDEX) = grdPost.Width * 0.12     'Cart-MST
        lmMaxCol = lmMaxCol + 1
    Else
        grdPost.ColWidth(CARTMSTINDEX) = 0    'Cart-MST
    End If
    If ilPSTZones Then
        grdPost.ColWidth(CARTPSTINDEX) = grdPost.Width * 0.12     'Cart-PST
        lmMaxCol = lmMaxCol + 1
    Else
        grdPost.ColWidth(CARTPSTINDEX) = 0    'Cart-PST
    End If
    grdPost.ColWidth(STATUSINDEX) = grdPost.Width * 0.08    'Status
    grdPost.ColWidth(CNTRNOINDEX) = (grdPost.Width - grdPost.ColWidth(DATEINDEX) - grdPost.ColWidth(TIMEINDEX) - grdPost.ColWidth(LENINDEX) - grdPost.ColWidth(CARTESTINDEX) - grdPost.ColWidth(CARTCSTINDEX) - grdPost.ColWidth(CARTMSTINDEX) - grdPost.ColWidth(CARTPSTINDEX) - grdPost.ColWidth(STATUSINDEX) - GRIDSCROLLWIDTH) / 2
    grdPost.ColWidth(ADVTINDEX) = grdPost.Width - grdPost.ColWidth(DATEINDEX) - grdPost.ColWidth(TIMEINDEX) - grdPost.ColWidth(CNTRNOINDEX) - grdPost.ColWidth(LENINDEX) - grdPost.ColWidth(CARTESTINDEX) - grdPost.ColWidth(CARTCSTINDEX) - grdPost.ColWidth(CARTMSTINDEX) - grdPost.ColWidth(CARTPSTINDEX) - grdPost.ColWidth(STATUSINDEX) - GRIDSCROLLWIDTH

End Sub

Private Sub mAdjPositionNo()

    Dim ilIndex As Integer
    Dim ilLoop As Integer
    ReDim ilPositionNoZone(0 To 3) As Integer
    
'    ilIndex = lmEnableRow - grdPost.FixedRows
'    If tmPostInfo(ilIndex).iType = 1 Then
'        grdPost.TextMatrix(lmEnableRow, LENINDEX) = "" 'Remove Units/Length value
'    End If
    ilIndex = Val(grdPost.TextMatrix(lmEnableRow, POSTINDEX))
    tmPostInfo(ilIndex).iType = 2
    ilPositionNoZone(0) = tmPostInfo(ilIndex).iPositionNoZone(0)
    ilPositionNoZone(1) = tmPostInfo(ilIndex).iPositionNoZone(1)
    ilPositionNoZone(2) = tmPostInfo(ilIndex).iPositionNoZone(2)
    ilPositionNoZone(3) = tmPostInfo(ilIndex).iPositionNoZone(3)
    For ilLoop = ilIndex + 1 To UBound(tmPostInfo) - 1 Step 1
        If tmPostInfo(ilLoop).iType = 1 Then
            Exit For
        End If
        If optShow(1).Value Then
            If DateValue(gAdjYear(tmPostInfo(ilIndex).sDateZone(1))) = DateValue(gAdjYear(tmPostInfo(ilLoop).sDateZone(1))) Then
                If gTimeToLong(tmPostInfo(ilIndex).sTimeZone(1), False) <> gTimeToLong(tmPostInfo(ilLoop).sTimeZone(1), False) Then
                    Exit For
                Else
                    ilPositionNoZone(1) = ilPositionNoZone(1) + 1
                    tmPostInfo(ilLoop).iPositionNoZone(1) = ilPositionNoZone(1)
                End If
            Else
                Exit For
            End If
        ElseIf optShow(2).Value Then
            If DateValue(gAdjYear(tmPostInfo(ilIndex).sDateZone(2))) = DateValue(gAdjYear(tmPostInfo(ilLoop).sDateZone(2))) Then
                If gTimeToLong(tmPostInfo(ilIndex).sTimeZone(2), False) <> gTimeToLong(tmPostInfo(ilLoop).sTimeZone(2), False) Then
                    Exit For
                Else
                    ilPositionNoZone(2) = ilPositionNoZone(2) + 1
                    tmPostInfo(ilLoop).iPositionNoZone(2) = ilPositionNoZone(2)
                End If
            Else
                Exit For
            End If
        ElseIf optShow(3).Value Then
            If DateValue(gAdjYear(tmPostInfo(ilIndex).sDateZone(3))) = DateValue(gAdjYear(tmPostInfo(ilLoop).sDateZone(3))) Then
                If gTimeToLong(tmPostInfo(ilIndex).sTimeZone(3), False) <> gTimeToLong(tmPostInfo(ilLoop).sTimeZone(3), False) Then
                    Exit For
                Else
                    ilPositionNoZone(3) = ilPositionNoZone(3) + 1
                    tmPostInfo(ilLoop).iPositionNoZone(3) = ilPositionNoZone(3)
                End If
            Else
                Exit For
            End If
        Else
            If DateValue(gAdjYear(tmPostInfo(ilIndex).sDateZone(0))) = DateValue(gAdjYear(tmPostInfo(ilLoop).sDateZone(0))) Then
                If gTimeToLong(tmPostInfo(ilIndex).sTimeZone(0), False) <> gTimeToLong(tmPostInfo(ilLoop).sTimeZone(0), False) Then
                    Exit For
                Else
                    ilPositionNoZone(0) = ilPositionNoZone(0) + 1
                    tmPostInfo(ilLoop).iPositionNoZone(0) = ilPositionNoZone(0)
                End If
            Else
                Exit For
            End If
        End If
    Next ilLoop

End Sub

Private Sub mAdjUnit_Sec()

    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilAvailIndex As Integer
    Dim llAvailRow As Long
    Dim ilUnits As Integer
    Dim ilSec As Integer
    Dim llRow As Long
    Dim sStr As String
    Dim iPos As Integer
    
    ilIndex = Val(grdPost.TextMatrix(lmEnableRow, POSTINDEX))
    
    'Find avail, if does not exist, then exit
    ilAvailIndex = -1
    llRow = lmEnableRow
    For ilLoop = ilIndex - 1 To LBound(tmPostInfo) Step -1
        llRow = llRow - 1
        If optShow(1).Value Then
            If DateValue(gAdjYear(tmPostInfo(ilIndex).sDateZone(1))) = DateValue(gAdjYear(tmPostInfo(ilLoop).sDateZone(1))) Then
                If gTimeToLong(tmPostInfo(ilIndex).sTimeZone(1), False) <> gTimeToLong(tmPostInfo(ilLoop).sTimeZone(1), False) Then
                    Exit For
                End If
            End If
        ElseIf optShow(2).Value Then
            If DateValue(gAdjYear(tmPostInfo(ilIndex).sDateZone(2))) = DateValue(gAdjYear(tmPostInfo(ilLoop).sDateZone(2))) Then
                If gTimeToLong(tmPostInfo(ilIndex).sTimeZone(2), False) <> gTimeToLong(tmPostInfo(ilLoop).sTimeZone(2), False) Then
                    Exit For
                End If
            End If
        ElseIf optShow(3).Value Then
            If DateValue(gAdjYear(tmPostInfo(ilIndex).sDateZone(3))) = DateValue(gAdjYear(tmPostInfo(ilLoop).sDateZone(3))) Then
                If gTimeToLong(tmPostInfo(ilIndex).sTimeZone(3), False) <> gTimeToLong(tmPostInfo(ilLoop).sTimeZone(3), False) Then
                    Exit For
                End If
            End If
        Else
            If DateValue(gAdjYear(tmPostInfo(ilIndex).sDateZone(0))) = DateValue(gAdjYear(tmPostInfo(ilLoop).sDateZone(0))) Then
                If gTimeToLong(tmPostInfo(ilIndex).sTimeZone(0), False) <> gTimeToLong(tmPostInfo(ilLoop).sTimeZone(0), False) Then
                    Exit For
                End If
            End If
        End If
        If tmPostInfo(ilLoop).iType = 1 Then
            ilAvailIndex = ilLoop
            Exit For
        End If
    Next ilLoop
    If ilAvailIndex < 0 Then
        Exit Sub
    End If
    ilUnits = tmPostInfo(ilAvailIndex).iUnits
    ilSec = tmPostInfo(ilAvailIndex).iLen
    llAvailRow = llRow
    llRow = llRow + 1
    For ilLoop = ilAvailIndex + 1 To UBound(tmPostInfo) - 1 Step 1
        If tmPostInfo(ilLoop).iType = 1 Then
            Exit For
        End If
        If optShow(1).Value Then
            If DateValue(gAdjYear(tmPostInfo(ilIndex).sDateZone(1))) = DateValue(gAdjYear(tmPostInfo(ilLoop).sDateZone(1))) Then
                If gTimeToLong(tmPostInfo(ilIndex).sTimeZone(1), False) <> gTimeToLong(tmPostInfo(ilLoop).sTimeZone(1), False) Then
                    Exit For
                End If
            Else
                Exit For
            End If
        ElseIf optShow(2).Value Then
            If DateValue(gAdjYear(tmPostInfo(ilIndex).sDateZone(2))) = DateValue(gAdjYear(tmPostInfo(ilLoop).sDateZone(2))) Then
                If gTimeToLong(tmPostInfo(ilIndex).sTimeZone(2), False) <> gTimeToLong(tmPostInfo(ilLoop).sTimeZone(2), False) Then
                    Exit For
                End If
             Else
                Exit For
           End If
        ElseIf optShow(3).Value Then
            If DateValue(gAdjYear(tmPostInfo(ilIndex).sDateZone(3))) = DateValue(gAdjYear(tmPostInfo(ilLoop).sDateZone(3))) Then
                If gTimeToLong(tmPostInfo(ilIndex).sTimeZone(3), False) <> gTimeToLong(tmPostInfo(ilLoop).sTimeZone(3), False) Then
                    Exit For
                End If
            Else
                Exit For
            End If
        Else
            If DateValue(gAdjYear(tmPostInfo(ilIndex).sDateZone(0))) = DateValue(gAdjYear(tmPostInfo(ilLoop).sDateZone(0))) Then
                If gTimeToLong(tmPostInfo(ilIndex).sTimeZone(0), False) <> gTimeToLong(tmPostInfo(ilLoop).sTimeZone(0), False) Then
                    Exit For
                End If
            Else
                Exit For
            End If
        End If
        If tmPostInfo(ilLoop).iType = 2 Then
            ilUnits = ilUnits - 1
            ilSec = ilSec - Val(grdPost.TextMatrix(llRow, LENINDEX))
        End If
        llRow = llRow + 1
    Next ilLoop
    grdPost.TextMatrix(llAvailRow, LENINDEX) = Trim$(Str$(ilUnits)) & "/" & Trim$(Str$(ilSec))
End Sub

Private Sub mResetSwapColor()
    Dim llRow As Long
    Dim llCol As Long
    
    If lmSwapStartRow <> -1 Then
        llRow = grdPost.Row
        llCol = grdPost.Col
        grdPost.Row = lmSwapStartRow
        grdPost.Col = TIMEINDEX
        grdPost.CellForeColor = vbBlack
        grdPost.Row = llRow
        grdPost.Col = llCol
    End If
    lmSwapStartRow = -1
    imSwapClickCount = -1
End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    gGrid_Clear grdPost, True
    For llRow = grdPost.FixedRows To grdPost.Rows - 1 Step 1
        grdPost.Row = llRow
        grdPost.Col = DATEINDEX
        grdPost.CellBackColor = LIGHTYELLOW
        grdPost.Col = TIMEINDEX
        grdPost.CellBackColor = LIGHTYELLOW
    Next llRow

End Sub

Private Sub mAddAbfDate(slDate As String)
    Dim ilLoop As Integer
    Dim llDate As Long
    
    bmCreateAbfRecord = True
    imAbfVefCode = imVefCode
    llDate = gDateValue(slDate)
    For ilLoop = 0 To UBound(lmAbfDate) - 1 Step 1
        If llDate = lmAbfDate(ilLoop) Then
            Exit Sub
        End If
    Next ilLoop
    lmAbfDate(UBound(lmAbfDate)) = llDate
    ReDim Preserve lmAbfDate(0 To UBound(lmAbfDate) + 1) As Long
End Sub

Private Sub mAddAbfRecords()
    Dim ilLoop As Integer
    Dim ilNext As Integer
    
    ilLoop = 0
    If Not bmCreateAbfRecord Then
        Exit Sub
    End If
    Do While ilLoop < UBound(lmAbfDate) - 1
        ilNext = ilLoop + 1
        Do While ilNext < UBound(lmAbfDate) - 1
            If lmAbfDate(ilLoop) = lmAbfDate(ilNext) - (ilNext - ilLoop) Then
                ilNext = ilNext + 1
            Else
                Exit Do
            End If
        Loop
        gSetStationSpotBuilder "P", imAbfVefCode, 0, lmAbfDate(ilLoop), lmAbfDate(ilNext - 1)
        ilLoop = ilNext
    Loop
    bmCreateAbfRecord = False
    ReDim lmAbfDate(0 To 0) As Long
End Sub
