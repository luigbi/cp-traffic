VERSION 5.00
Begin VB.Form PostItem 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5085
   ClientLeft      =   630
   ClientTop       =   1680
   ClientWidth     =   9360
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
   ScaleHeight     =   5085
   ScaleWidth      =   9360
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
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
      Height          =   180
      Left            =   30
      Picture         =   "Postitem.frx":0000
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   105
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
      Left            =   8775
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4230
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmcUpdate 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4755
      TabIndex        =   26
      Top             =   4740
      Width           =   1050
   End
   Begin VB.Timer tmcDrag 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   4320
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   915
      Top             =   4290
   End
   Begin VB.PictureBox pbcYN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5040
      ScaleHeight     =   210
      ScaleWidth      =   180
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   975
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.ListBox lbcBItem 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2055
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2340
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox lbcBDate 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   4440
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2355
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.ListBox lbcBVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2400
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   1470
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
      Left            =   1320
      Picture         =   "Postitem.frx":030A
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   300
      MaxLength       =   20
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   1020
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
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   75
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4560
      Width           =   75
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1500
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   270
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox edcUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      HelpContextID   =   8
      Left            =   6645
      MaxLength       =   10
      TabIndex        =   14
      Top             =   1155
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox edcDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   1785
      MaxLength       =   20
      TabIndex        =   11
      Top             =   975
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.TextBox edcNoItems 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      HelpContextID   =   8
      Left            =   7260
      MaxLength       =   3
      TabIndex        =   16
      Top             =   1380
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.TextBox edcAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Height          =   210
      Left            =   5535
      MaxLength       =   12
      TabIndex        =   13
      Top             =   975
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmcUndo 
      Appearance      =   0  'Flat
      Caption         =   "U&ndo"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6075
      TabIndex        =   27
      Top             =   4740
      Width           =   945
   End
   Begin VB.PictureBox pbcIBTab 
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
      Height          =   90
      Left            =   45
      ScaleHeight     =   90
      ScaleWidth      =   60
      TabIndex        =   18
      Top             =   4200
      Width           =   60
   End
   Begin VB.PictureBox pbcIBSTab 
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
      Height          =   120
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   105
      TabIndex        =   4
      Top             =   525
      Width           =   105
   End
   Begin VB.VScrollBar vbcPostItem 
      Height          =   3480
      LargeChange     =   15
      Left            =   8970
      TabIndex        =   19
      Top             =   645
      Width           =   240
   End
   Begin VB.PictureBox plcSelect 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   2835
      ScaleHeight     =   345
      ScaleWidth      =   6345
      TabIndex        =   1
      Top             =   75
      Width           =   6405
      Begin VB.ComboBox cbcSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3765
         TabIndex        =   3
         Top             =   15
         Width           =   2565
      End
      Begin VB.ComboBox cbcAdvt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
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
         Height          =   315
         Left            =   45
         TabIndex        =   2
         Top             =   15
         Width           =   3435
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3435
      TabIndex        =   25
      Top             =   4740
      Width           =   1050
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   45
      ScaleHeight     =   270
      ScaleWidth      =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1680
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2055
      TabIndex        =   24
      Top             =   4740
      Width           =   1050
   End
   Begin VB.PictureBox pbcPostItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   180
      Picture         =   "Postitem.frx":0404
      ScaleHeight     =   3495
      ScaleWidth      =   8790
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   8790
      Begin VB.Label lacIBFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Height          =   240
         Left            =   -60
         TabIndex        =   23
         Top             =   390
         Visible         =   0   'False
         Width           =   8730
      End
   End
   Begin VB.PictureBox plcPostItem 
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
      Height          =   3615
      Left            =   210
      ScaleHeight     =   3555
      ScaleWidth      =   9015
      TabIndex        =   5
      Top             =   555
      Width           =   9075
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   4620
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8715
      Picture         =   "Postitem.frx":11316
      Top             =   4575
      Width           =   480
   End
   Begin VB.Label lacTotals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7905
      TabIndex        =   22
      Top             =   4275
      Width           =   210
   End
End
Attribute VB_Name = "PostItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Postitem.frm on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: CCancel.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract revision number increment screen code
Option Explicit
Option Compare Text
Dim tmItemCode() As SORTCODE
Dim smItemCodeTag As String
Dim tmContract() As SORTCODE
Dim smContractTag As String
Dim tmBVehicle() As SORTCODE
Dim smBVehicleTag As String
Dim imUpdateAllowed As Integer
Dim imFirstActivate As Integer

'Billing Items
Dim tmIBCtrls(0 To 12) As FIELDAREA
Dim imLBIBCtrls As Integer
Dim hmChf As Integer
Dim tmChf As CHF
Dim tmChfSrchKey As LONGKEY0
Dim imChfRecLen As Integer
Dim hmClf As Integer
Dim tmClf() As CLF
Dim tmClfSrchKey As CLFKEY0
Dim imClfRecLen As Integer
Dim hmSbf As Integer        'Special billing
Dim tmIBSbf() As SBFLIST    'SBF record image of billing items
Dim tmSbfSrchKey As SBFKEY0    'SBF key record image
Dim imSbfRecLen As Integer        'SBF record length
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imSelectedIndex As Integer  'Index of selected record (0 if new)
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imIBBoxNo As Integer
Dim imIBRowNo As Integer
Dim smIBSave() As String    'Values saved for Item bill(1=Transaction; 2=Description 3= Amount/item; 4= Units; 5=# Items; 6=Total Amount; 7=Billed)
Dim imIBSave() As Integer   'Values saved for Item Bill (1=Vehicle; 2=Date; 3=Item Billing type; 4=Agy Comm; 5= Salesperson comm (-1=new,0=Yes;1=No;2=No comm defined); 6= sales tax)
Dim smIBShow() As String    'Show values for Item Bill
Dim imIBChg As Integer
Dim imComboBoxIndex As Integer
Dim imSettingValue As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imTaxDefined As Integer
Dim imBypassFocus As Integer
Dim fmDragX As Single       'Start x location of drag
Dim fmDragY As Single       'Start y location
Dim imDragType As Integer   '0=Start Drag; 1=scroll up; 2= scroll down
'Help
Dim tmSbfHelp() As HLF
'Dim tmRec As LPOPREC
Const IBVEHICLEINDEX = 1   'Vehicle control/field
Const IBDATEINDEX = 2       'Bill date control/index
Const IBDESCRIPTINDEX = 3   'Description control/field
Const IBITEMTYPEINDEX = 4   'Item billing type control/field
Const IBACINDEX = 5         'Agency commission type control/field
Const IBSCINDEX = 6         'Salesperson commission control/field
Const IBTXINDEX = 7         'Taxable control/field
Const IBAMOUNTINDEX = 8     'Amount per item control/field
Const IBUNITSINDEX = 9      'Units control/field
Const IBNOITEMSINDEX = 10    'Number of items control/field
Const IBTAMOUNTINDEX = 11    'Total amount control/field
Const IBBILLINDEX = 12      'Billed flag control/field
'Help messages
Private Sub cbcSelect_Change()
    Dim ilIndex As Integer  'Current index selected from combo box
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        imChgMode = True    'Set change mode to avoid infinite loop
        Screen.MousePointer = vbHourglass  'Wait
        If cbcSelect.Text <> "" Then
            gManLookAhead cbcSelect, imBSMode, imComboBoxIndex
        End If
        If cbcSelect.ListIndex >= 0 Then
            ilIndex = cbcSelect.ListIndex
            If Not mReadChfRec(ilIndex) Then
                GoTo cbcSelectErr
            End If
            If Not mReadRec() Then  'Sbf
                GoTo cbcSelectErr
            End If
        Else
            mClearCtrlFields 'If coming from [New], select_change will not be generated
        End If
        Screen.MousePointer = vbDefault
        mSetCommands
        imChgMode = False
    End If
    Exit Sub
cbcSelectErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    Exit Sub
End Sub
Private Sub cbcSelect_DropDown()
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
End Sub
Private Sub cbcSelect_GotFocus()
    Dim slSvText As String   'Save so list box can be reset

    If imTerminate Then
        Exit Sub
    End If
    mIBSetShow imIBBoxNo
    imIBBoxNo = -1
    imIBRowNo = -1
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
'    gSetIndexFromText cbcSelect
    slSvText = cbcSelect.Text
'    ilSvIndex = cbcSelect.ListIndex
    mPopulate
    If imTerminate Then
        Exit Sub
    End If
    If cbcSelect.ListCount <= 1 Then
        cbcSelect.ListIndex = 0
        mClearCtrlFields 'Make sure all fields cleared
        mSetCommands
'        pbcHdSTab.SetFocus
        Exit Sub
    End If
'    gShowHelpMess tmChfHelp(), CHFCNTRSELECT
    gCtrlGotFocus ActiveControl
    If (slSvText = "") Then
        cbcSelect.ListIndex = 0
        cbcSelect_Change
    Else
        gFindMatch slSvText, 1, cbcSelect
        If gLastFound(cbcSelect) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcSelect)) Or (ilSvIndex <> cbcSelect.ListIndex) Then
            If (slSvText <> cbcSelect.List(gLastFound(cbcSelect))) Then
                cbcSelect.ListIndex = gLastFound(cbcSelect)
            End If
        Else
            cbcSelect.ListIndex = 0
            mClearCtrlFields
            cbcSelect_Change
        End If
    End If
    mSetCommands
End Sub
Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub cmcCancel_GotFocus()
    mIBSetShow imIBBoxNo
    imIBBoxNo = -1
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mIBEnableBox imIBBoxNo
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub cmcDone_GotFocus()
    mIBSetShow imIBBoxNo
    imIBBoxNo = -1
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcDropDown_Click()
    Select Case imIBBoxNo
        Case IBVEHICLEINDEX
            lbcBVehicle.Visible = Not lbcBVehicle.Visible
        Case IBDATEINDEX
            lbcBDate.Visible = Not lbcBDate.Visible
        Case IBITEMTYPEINDEX
            lbcBItem.Visible = Not lbcBItem.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUndo_Click()
    If Not mReadRec() Then
        imTerminate = True
        Exit Sub
    End If
    pbcPostItem.Cls
    mMoveIBRecToCtrl
    mInitIBShow
    pbcPostItem_Paint
    pbcIBSTab.SetFocus
End Sub
Private Sub cmcUndo_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub cmcUndo_GotFocus()
    mIBSetShow imIBBoxNo
    imIBBoxNo = -1
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcUndo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub cmcUpdate_Click()
    If Not imUpdateAllowed Then
        Exit Sub
    End If
    If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mIBEnableBox imIBBoxNo
        Exit Sub
    End If
    mIBEnableBox imIBBoxNo
End Sub
Private Sub cmcUpdate_GotFocus()
    mIBSetShow imIBBoxNo
    imIBBoxNo = -1
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcAmount_GotFocus()
    Dim slNameCode As String
    Dim slAmount As String
    Dim ilRet As Integer
    If (edcAmount.Text = "") Then
        If imIBSave(3, imIBRowNo) > 0 Then
            slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey    'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
            ilRet = gParseItem(slNameCode, 3, "\", slAmount)
            If ilRet = CP_MSG_NONE Then
                edcAmount.Text = slAmount
            End If
        End If
    End If
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcAmount_KeyPress(KeyAscii As Integer)
    Dim ilPos As Integer
    Dim slStr As String
    ilPos = InStr(ActiveControl.SelText, ".")
    If ilPos = 0 Then
        ilPos = InStr(ActiveControl.Text, ".")    'Disallow multi-decimal points
        If ilPos > 0 Then
            If KeyAscii = KEYDECPOINT Then
                Beep
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) And (KeyAscii <> KEYDECPOINT) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcAmount.Text
    slStr = Left$(slStr, edcAmount.SelStart) & Chr$(KeyAscii) & Right$(slStr, Len(slStr) - edcAmount.SelStart - edcAmount.SelLength)
    If gCompNumberStr(slStr, "9999999.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDescription_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim slDate As String
    Dim ilRet As Integer
    Select Case imIBBoxNo
        Case IBVEHICLEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcBVehicle, imBSMode, imComboBoxIndex
        Case IBDATEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcBDate, imBSMode, slStr)
            If ilRet > 1 Then
                'Reset dates
                If gValidDate(slStr) Then
                    If (tmChf.sBillCycle = "C") Or (tmChf.sBillCycle = "D") Then
                        slStr = gObtainEndCal(slStr)
                        gEndCalDatePop slStr, 24, lbcBDate
                    Else
                        slStr = gObtainEndStd(slStr)
                        gEndStdDatePop slStr, 24, lbcBDate
                    End If
                    slDate = Format$(slStr, "m/d/yy")
                    imChgMode = True
                    gFindMatch slDate, 0, lbcBDate
                    If gLastFound(lbcBDate) >= 0 Then
                        lbcBDate.ListIndex = gLastFound(lbcBDate)
                    Else
                        lbcBDate.ListIndex = -1
                        edcDropDown.Text = ""
                        imChgMode = False
                        Exit Sub
                    End If
                    imChgMode = False
                Else
                    Exit Sub
                End If
            ElseIf ilRet = 1 Then
                Exit Sub
            End If
        Case IBITEMTYPEINDEX
            imLbcArrowSetting = True
            ilRet = gOptionalLookAhead(edcDropDown, lbcBItem, imBSMode, slStr)
            If ilRet = 1 Then
                lbcBItem.ListIndex = 0
            End If
    End Select
    imLbcArrowSetting = False
End Sub
Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imIBBoxNo
        Case IBVEHICLEINDEX
            If lbcBVehicle.ListCount = 1 Then
                lbcBVehicle.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcIBSTab.SetFocus
                'Else
                '    pbcIBTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case IBDATEINDEX
            If lbcBDate.ListCount = 1 Then
                lbcBDate.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcIBSTab.SetFocus
                'Else
                '    pbcIBTab.SetFocus
                'End If
                'Exit Sub
            End If
        Case IBITEMTYPEINDEX
            If lbcBItem.ListCount = 1 Then
                lbcBItem.ListIndex = 0
                'If imTabDirection = -1 Then  'Right To Left
                '    pbcIBSTab.SetFocus
                'Else
                '    pbcIBTab.SetFocus
                'End If
                'Exit Sub
            End If
    End Select
End Sub
Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub
Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub
Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KeyUp) Or (KeyCode = KeyDown) Then
        Select Case imIBBoxNo
            Case IBVEHICLEINDEX
                gProcessArrowKey Shift, KeyCode, lbcBVehicle, imLbcArrowSetting
            Case IBDATEINDEX
                gProcessArrowKey Shift, KeyCode, lbcBDate, imLbcArrowSetting
            Case IBITEMTYPEINDEX
                gProcessArrowKey Shift, KeyCode, lbcBItem, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imIBBoxNo
            Case IBITEMTYPEINDEX
                If imTabDirection = -1 Then  'Right To Left
                    pbcIBSTab.SetFocus
                Else
                    pbcIBTab.SetFocus
                End If
                Exit Sub
        End Select
        imDoubleClickName = False
    End If
End Sub
Private Sub edcNoItems_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcNoItems_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcNoItems.Text
    slStr = Left$(slStr, edcNoItems.SelStart) & Chr$(KeyAscii) & Right$(slStr, Len(slStr) - edcNoItems.SelStart - edcNoItems.SelLength)
    If gCompNumberStr(slStr, "9999") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcUnits_GotFocus()
    Dim slNameCode As String
    Dim slUnits As String
    Dim ilRet As Integer
    If edcUnits.Text = "" Then
        If imIBSave(3, imIBRowNo) > 0 Then
            slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey    'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
            ilRet = gParseItem(slNameCode, 4, "\", slUnits)
            If ilRet = CP_MSG_NONE Then
                edcUnits.Text = slUnits
            End If
        End If
    End If
    gCtrlGotFocus ActiveControl
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
    Me.KeyPreview = True
    If (igWinStatus(INVOICESJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        If tgUrf(0).iSlfCode > 0 Then
            imUpdateAllowed = False
        Else
            imUpdateAllowed = True
        End If
    End If
    gShowBranner imUpdateAllowed
    Me.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Deactivate()
    Me.KeyPreview = False
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer

    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        If (cbcSelect.Enabled) And (imIBBoxNo > 0) Then
            cbcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imIBBoxNo > 0 Then
            mIBEnableBox imIBBoxNo
        End If
        If ilReSet Then
            cbcSelect.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not igManUnload Then
        mIBSetShow imIBBoxNo
        imIBBoxNo = -1
        pbcArrow.Visible = False
        lacIBFrame.Visible = False
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            If imIBBoxNo <> -1 Then
                mIBEnableBox imIBBoxNo
            End If
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub
Private Sub imcTrash_Click()
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilUpperBound As Integer
    Dim ilRowNo As Integer
    Dim slStr As String
    If (imIBRowNo < 1) Then
        Exit Sub
    End If
    If (smIBSave(7, imIBRowNo) = "B") Then
        Exit Sub
    End If
    ilRowNo = imIBRowNo
    mIBSetShow imIBBoxNo
    imIBBoxNo = -1   '
    imIBRowNo = -1
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
    gCtrlGotFocus ActiveControl
    ilUpperBound = UBound(smIBSave, 2)
    If ilRowNo = ilUpperBound Then
        For ilLoop = imLBIBCtrls To UBound(tmIBCtrls) Step 1
            slStr = ""
            gSetShow pbcPostItem, slStr, tmIBCtrls(ilLoop)
            smIBShow(ilLoop, ilRowNo) = tmIBCtrls(ilLoop).sShow
        Next ilLoop
        pbcPostItem_Paint
        mInitNewIB ilRowNo   'Set defaults for extra row
    Else
        For ilLoop = ilRowNo To ilUpperBound - 1 Step 1
            For ilIndex = 1 To UBound(smIBSave, 1) Step 1
                smIBSave(ilIndex, ilLoop) = smIBSave(ilIndex, ilLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(imIBSave, 1) Step 1
                imIBSave(ilIndex, ilLoop) = imIBSave(ilIndex, ilLoop + 1)
            Next ilIndex
            For ilIndex = 1 To UBound(smIBShow, 1) Step 1
                smIBShow(ilIndex, ilLoop) = smIBShow(ilIndex, ilLoop + 1)
            Next ilIndex
        Next ilLoop
        ilUpperBound = UBound(smIBSave, 2)
        ReDim Preserve smIBSave(1 To 7, 1 To ilUpperBound - 1) As String
        ReDim Preserve imIBSave(1 To 6, 1 To ilUpperBound - 1) As Integer
        ReDim Preserve smIBShow(1 To IBTAMOUNTINDEX, 1 To ilUpperBound - 1) As String
    End If
    imIBChg = True
    mSetCommands
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcPostItem.Cls
    pbcPostItem_Paint
    mIBTotals
End Sub
Private Sub imcTrash_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
    imcTrash_Click
End Sub
Private Sub imcTrash_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbEnter Then    'Enter drag over
        lacIBFrame.DragIcon = IconTraf!imcIconDwnArrow.DragIcon
'        lacEvtFrame.DragIcon = IconTraf!imcIconTrash.DragIcon
        imcTrash.Picture = IconTraf!imcTrashOpened.Picture
    ElseIf State = vbLeave Then
        lacIBFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
        imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    End If
End Sub
Private Sub imcTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub lacTotals_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub lacTotals_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub lbcBDate_Click()
    gProcessLbcClick lbcBDate, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcBDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcBItem_Click()
    If imLbcMouseDown Then
        tmcClick.Interval = 300 'Delay processing encase double click
        tmcClick.Enabled = True
        imLbcMouseDown = False
    Else
        gProcessLbcClick lbcBItem, edcDropDown, imChgMode, imLbcArrowSetting
    End If
End Sub
Private Sub lbcBItem_DblClick()
    tmcClick.Enabled = False
    imDoubleClickName = True    'Double click event is followed by a mouse up event
                                'Process the double click event in the mouse up event
                                'to avoid the mouse up event being in next form
End Sub
Private Sub lbcBItem_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub lbcBItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imLbcMouseDown = True
End Sub
Private Sub lbcBItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        imLbcArrowSetting = False
        gProcessLbcClick lbcBItem, edcDropDown, imChgMode, imLbcArrowSetting
        If imTabDirection = -1 Then  'Right To Left
            pbcIBSTab.SetFocus
        Else
            pbcIBTab.SetFocus
        End If
    End If
End Sub
Private Sub lbcBVehicle_Click()
    gProcessLbcClick lbcBVehicle, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcBVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mAdvtPop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Advertiser list box   *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mAdvtPop()
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = cbcAdvt.ListIndex
    If ilIndex >= 0 Then
        slName = cbcAdvt.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtBox(PostItem, cbcAdvt, Traffic!lbcAdvertiser)
    ilRet = gPopAdvtBox(PostItem, cbcAdvt, tgAdvertiser(), sgAdvertiserTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", PostItem
        On Error GoTo 0
'        cbcAdvt.AddItem "[New]", 0  'Force as first item on list
        If ilIndex >= 0 Then
            gFindMatch slName, 0, cbcAdvt
            If gLastFound(cbcAdvt) >= 0 Then
                cbcAdvt.ListIndex = gLastFound(cbcAdvt)
            Else
                cbcAdvt.ListIndex = -1
            End If
        Else
            cbcAdvt.ListIndex = ilIndex
        End If
    End If
    Exit Sub
mAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearCtrlFields                *
'*                                                     *
'*             Created:6/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Clear each control on the      *
'*                      screen                         *
'*                                                     *
'*******************************************************
Private Sub mClearCtrlFields()
'
'   mClearCtrlFields
'   Where:
'       ilOnlyAddr (I)- Clear only fields after address
'
    Dim ilLoop As Integer
    lbcBVehicle.ListIndex = -1
    lbcBDate.ListIndex = -1
    lbcBItem.ListIndex = -1
    edcAmount.Text = ""
    edcDescription.Text = ""
    For ilLoop = imLBIBCtrls To UBound(tmIBCtrls) Step 1
        tmIBCtrls(ilLoop).sShow = ""
        tmIBCtrls(ilLoop).iChg = False
    Next ilLoop
    ReDim smIBSave(1 To 7, 1 To 1)
    ReDim imIBSave(1 To 6, 1 To 1)
    ReDim smIBShow(1 To IBBILLINDEX, 1 To 1)
    imIBBoxNo = -1
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcMouseDown = False
    imChgMode = False
    imBSMode = False
    imIBChg = False
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mDatePop                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Date list box         *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Private Sub mDatePop()
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartdate As Long
    Dim slDate As String
    Dim llEndDate As Long
    Dim ilNoPds As Integer
    lbcBDate.Clear
    gUnpackDate tmChf.iStartDate(0), tmChf.iStartDate(1), slStartDate    'Week Start date
    llStartdate = gDateValue(slStartDate)
    gUnpackDate tmChf.iEndDate(0), tmChf.iEndDate(1), slEndDate    'Week Start date
    llEndDate = gDateValue(slEndDate)
    If (tmChf.sBillCycle = "C") Or (tmChf.sBillCycle = "D") Then
        slDate = gObtainEndCal(slStartDate)
        ilNoPds = (llEndDate - llStartdate + 28) \ 28 + 24
        gEndCalDatePop slDate, ilNoPds, lbcBDate
    Else
        slDate = gObtainEndStd(slStartDate)
        ilNoPds = (llEndDate - llStartdate + 28) \ 28 + 24
        gEndStdDatePop slDate, ilNoPds, lbcBDate
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mIBEnableBox                    *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mIBEnableBox(ilBoxNo As Integer)
'
'   mIBEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slName As String
    If (ilBoxNo < imLBIBCtrls) Or (ilBoxNo > UBound(tmIBCtrls)) Then
        Exit Sub
    End If

    If (imIBRowNo < vbcPostItem.Value) Or (imIBRowNo >= vbcPostItem.Value + vbcPostItem.LargeChange + 1) Then
        mIBSetShow ilBoxNo
        pbcArrow.Visible = False
        lacIBFrame.Visible = False
        Exit Sub
    End If
    lacIBFrame.Move 0, tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15) - 30
    lacIBFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcPostItem.Top + tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case IBVEHICLEINDEX 'Vehicle
            mVehPop
'            gShowHelpMess tmSbfHelp(), SBFVEHICLE
            lbcBVehicle.Height = gListBoxHeight(lbcBVehicle.ListCount, 8)
            edcDropDown.Width = tmIBCtrls(IBVEHICLEINDEX).fBoxW
            If tgSpf.iVehLen <= 40 Then
                edcDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcDropDown.MaxLength = 20
            End If
            gMoveTableCtrl pbcPostItem, edcDropDown, tmIBCtrls(IBVEHICLEINDEX).fBoxX, tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            imChgMode = True
            If imIBSave(1, imIBRowNo) >= 0 Then
                lbcBVehicle.ListIndex = imIBSave(1, imIBRowNo)
                imComboBoxIndex = lbcBVehicle.ListIndex
                edcDropDown.Text = lbcBVehicle.List(imIBSave(1, imIBRowNo))
            Else
                If imIBRowNo > 1 Then
                    lbcBVehicle.ListIndex = imIBSave(1, imIBRowNo - 1)
                    imComboBoxIndex = lbcBVehicle.ListIndex
                    edcDropDown.Text = lbcBVehicle.List(imIBSave(1, imIBRowNo - 1))
                Else
                    'slNameCode = Contract!lbcRateCard.List(Contract!lbcRateCard.ListIndex)
                    'ilRet = gParseItem(slNameCode, 2, "/", slName)
                    'If ilRet <> CP_MSG_NONE Then
                        slName = sgUserDefVehicleName
                    'Else
                    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    '    If InStr(slCode, "-") <> 0 Then
                    '        slName = sgUserDefVehicleName
                    '    End If
                    'End If
                    gFindMatch slName, 0, lbcBVehicle
                    If gLastFound(lbcBVehicle) >= 0 Then
                        lbcBVehicle.ListIndex = gLastFound(lbcBVehicle)
                        imComboBoxIndex = lbcBVehicle.ListIndex
                        edcDropDown.Text = lbcBVehicle.List(lbcBVehicle.ListIndex)
                    Else
                        lbcBVehicle.ListIndex = 0
                        imComboBoxIndex = lbcBVehicle.ListIndex
                        edcDropDown.Text = lbcBVehicle.List(0)
                    End If
                End If
            End If
            imChgMode = False
            If imIBRowNo - vbcPostItem.Value <= vbcPostItem.LargeChange \ 2 Then
                lbcBVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcBVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcBVehicle.Height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IBDATEINDEX 'Date
'            gShowHelpMess tmSbfHelp(), SBFITEMDATE
            lbcBDate.Height = gListBoxHeight(lbcBDate.ListCount, 8)
            edcDropDown.Width = tmIBCtrls(IBDATEINDEX).fBoxW
            edcDropDown.MaxLength = 10
            gMoveTableCtrl pbcPostItem, edcDropDown, tmIBCtrls(IBDATEINDEX).fBoxX, tmIBCtrls(IBDATEINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcBDate.ListIndex = imIBSave(2, imIBRowNo)
            imChgMode = True
            If lbcBDate.ListIndex < 0 Then
                If imIBRowNo <= 1 Then
                    lbcBDate.ListIndex = 0
                    edcDropDown.Text = lbcBDate.List(0)
                Else
                    lbcBDate.ListIndex = imIBSave(2, imIBRowNo - 1)
                    edcDropDown.Text = lbcBDate.List(lbcBDate.ListIndex)
                End If
            Else
                edcDropDown.Text = lbcBDate.List(lbcBDate.ListIndex)
            End If
            imChgMode = False
            If imIBRowNo - vbcPostItem.Value <= vbcPostItem.LargeChange \ 2 Then
                lbcBDate.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcBDate.Move edcDropDown.Left, edcDropDown.Top - lbcBDate.Height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IBDESCRIPTINDEX
'            gShowHelpMess tmSbfHelp(), SBFITEMDESC
            edcDescription.Width = tmIBCtrls(IBDESCRIPTINDEX).fBoxW
            gMoveTableCtrl pbcPostItem, edcDescription, tmIBCtrls(IBDESCRIPTINDEX).fBoxX, tmIBCtrls(IBDESCRIPTINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15)
            edcDescription.Text = smIBSave(2, imIBRowNo)
            edcDescription.Visible = True  'Set visibility
            edcDescription.SetFocus
        Case IBITEMTYPEINDEX 'Item bill type
            lbcBItem.Height = gListBoxHeight(lbcBItem.ListCount, 8)
            edcDropDown.Width = tmIBCtrls(IBITEMTYPEINDEX).fBoxW
            edcDropDown.MaxLength = 20
            gMoveTableCtrl pbcPostItem, edcDropDown, tmIBCtrls(IBITEMTYPEINDEX).fBoxX, tmIBCtrls(IBITEMTYPEINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            lbcBItem.ListIndex = imIBSave(3, imIBRowNo)
            imChgMode = True
            If lbcBItem.ListIndex < 0 Then
                lbcBItem.ListIndex = 0
                edcDropDown.Text = lbcBItem.List(0)
            Else
                edcDropDown.Text = lbcBItem.List(lbcBItem.ListIndex)
            End If
            imChgMode = False
            If imIBRowNo - vbcPostItem.Value <= vbcPostItem.LargeChange \ 2 Then
                lbcBItem.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcBItem.Move edcDropDown.Left, edcDropDown.Top - lbcBItem.Height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case IBACINDEX
'            gShowHelpMess tmSbfHelp(), SBFITEMAC
            gMoveTableCtrl pbcPostItem, pbcYN, tmIBCtrls(ilBoxNo).fBoxX, tmIBCtrls(ilBoxNo).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            If imIBSave(4, imIBRowNo) = -1 Then
                imIBSave(4, imIBRowNo) = 1  'Default to No
            End If
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case IBSCINDEX
            If imIBSave(5, imIBRowNo) = 2 Then
                Exit Sub
            End If
'            gShowHelpMess tmSbfHelp(), SBFITEMSC
            gMoveTableCtrl pbcPostItem, pbcYN, tmIBCtrls(ilBoxNo).fBoxX, tmIBCtrls(ilBoxNo).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            If imIBSave(5, imIBRowNo) = -1 Then
                imIBSave(5, imIBRowNo) = 1  'Default to No
            End If
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case IBTXINDEX
            If imIBSave(6, imIBRowNo) = 2 Then
                Exit Sub
            End If
'            gShowHelpMess tmSbfHelp(), SBFITEMTX
            gMoveTableCtrl pbcPostItem, pbcYN, tmIBCtrls(ilBoxNo).fBoxX, tmIBCtrls(ilBoxNo).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15) ' - fgBoxGridH / 2
            If imIBSave(6, imIBRowNo) = -1 Then
                If imTaxDefined Then
                    imIBSave(6, imIBRowNo) = 0  'Yes
                Else
                    imIBSave(6, imIBRowNo) = 2  'Default to No
                End If
            End If
            pbcYN_Paint
            pbcYN.Visible = True
            pbcYN.SetFocus
        Case IBAMOUNTINDEX
'            gShowHelpMess tmSbfHelp(), SBFITEMCOST
            edcAmount.Width = tmIBCtrls(IBAMOUNTINDEX).fBoxW
            gMoveTableCtrl pbcPostItem, edcAmount, tmIBCtrls(IBAMOUNTINDEX).fBoxX, tmIBCtrls(IBAMOUNTINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15)
            edcAmount.Text = smIBSave(3, imIBRowNo)
            edcAmount.Visible = True  'Set visibility
            edcAmount.SetFocus
        Case IBUNITSINDEX
'            gShowHelpMess tmSbfHelp(), SBFITEMUNITS
'            edcUnits.Width = tmIBCtrls(IBUNITSINDEX).fBoxW
'            gMoveTableCtrl pbcPostItem, edcUnits, tmIBCtrls(IBUNITSINDEX).fBoxX, tmIBCtrls(IBUNITSINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15)
'            edcUnits.Text = smIBSave(4, imIBRowNo)
'            edcUnits.Visible = True  'Set visibility
'            edcUnits.SetFocus
        Case IBNOITEMSINDEX
'            gShowHelpMess tmSbfHelp(), SBFITEMNO
            edcNoItems.Width = tmIBCtrls(IBNOITEMSINDEX).fBoxW
            gMoveTableCtrl pbcPostItem, edcNoItems, tmIBCtrls(IBNOITEMSINDEX).fBoxX, tmIBCtrls(IBNOITEMSINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15)
            edcNoItems.Text = smIBSave(5, imIBRowNo)
            edcNoItems.Visible = True  'Set visibility
            edcNoItems.SetFocus
        Case IBTAMOUNTINDEX
'            gShowHelpMess tmSbfHelp(), SBFITEMTOT
        Case IBBILLINDEX
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mIBSetShow                      *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mIBSetShow(ilBoxNo As Integer)
'
'   mIBSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim slNameCode As String
    Dim slAmount As String
    Dim slUnits As String
    Dim ilRet As Integer
    Dim slSlspComm As String
    Dim slTax As String
    If (ilBoxNo < imLBIBCtrls) Or (ilBoxNo > UBound(tmIBCtrls)) Then
        Exit Sub
    End If

    lacIBFrame.Visible = False
    pbcArrow.Visible = False
    Select Case ilBoxNo 'Branch on box type (control)
        Case IBVEHICLEINDEX 'Vehicle
            lbcBVehicle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcPostItem, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(IBVEHICLEINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If imIBSave(1, imIBRowNo) <> lbcBVehicle.ListIndex Then
                imIBSave(1, imIBRowNo) = lbcBVehicle.ListIndex
                If imIBRowNo < UBound(imIBSave, 2) Then  'New lines set after all fields entered
                    imIBChg = True
                End If
                If smIBSave(1, imIBRowNo) = "" Then
                    smIBSave(1, imIBRowNo) = "C"
                End If
            End If
        Case IBDATEINDEX 'Date index
            lbcBDate.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            slStr = gFormatDate(slStr)
            gSetShow pbcPostItem, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(IBDATEINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If imIBSave(2, imIBRowNo) <> lbcBDate.ListIndex Then
                imIBSave(2, imIBRowNo) = lbcBDate.ListIndex
                If imIBRowNo < UBound(tmIBSbf) + 1 Then   'New lines set after all fields entered
                    imIBChg = True
                End If
            End If
        Case IBDESCRIPTINDEX
            edcDescription.Visible = False
            slStr = edcDescription.Text
            gSetShow pbcPostItem, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(IBDESCRIPTINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If smIBSave(2, imIBRowNo) <> edcDescription.Text Then
                If imIBRowNo < UBound(tmIBSbf) + 1 Then   'New lines set after all fields entered
                    imIBChg = True
                End If
                smIBSave(2, imIBRowNo) = edcDescription.Text
            End If
        Case IBITEMTYPEINDEX 'Item type index
            lbcBItem.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            If lbcBItem.ListIndex < 0 Then
                slStr = ""
            Else
                slStr = lbcBItem.List(lbcBItem.ListIndex)
            End If
            gSetShow pbcPostItem, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(IBITEMTYPEINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If imIBSave(3, imIBRowNo) <> lbcBItem.ListIndex Then
                imIBSave(3, imIBRowNo) = lbcBItem.ListIndex
                If imIBRowNo < UBound(tmIBSbf) + 1 Then   'New lines set after all fields entered
                    imIBChg = True
                End If
                If imIBSave(3, imIBRowNo) > 0 Then
                    slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey    'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                    ilRet = gParseItem(slNameCode, 3, "\", slAmount)
                    If ilRet <> CP_MSG_NONE Then
                        slAmount = ""
                    End If
                    If Val(smIBSave(3, imIBRowNo)) <> Val(slAmount) Then
                        smIBSave(3, imIBRowNo) = "" 'Amount/unit
                        slStr = ""
                        gSetShow pbcPostItem, slStr, tmIBCtrls(IBAMOUNTINDEX)
                        smIBShow(IBAMOUNTINDEX, imIBRowNo) = tmIBCtrls(IBAMOUNTINDEX).sShow
                        smIBSave(6, imIBRowNo) = "" 'Total
                        gSetShow pbcPostItem, slStr, tmIBCtrls(IBTAMOUNTINDEX)
                        smIBShow(IBTAMOUNTINDEX, imIBRowNo) = tmIBCtrls(IBTAMOUNTINDEX).sShow
                    End If
                    slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey    'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                    ilRet = gParseItem(slNameCode, 4, "\", slUnits)
                    If ilRet <> CP_MSG_NONE Then
                        slUnits = ""
                    End If
                    If slUnits <> smIBSave(4, imIBRowNo) Then
                        smIBSave(4, imIBRowNo) = slUnits ' "" 'Unit definition
                        slStr = slUnits ' ""
                        gSetShow pbcPostItem, slStr, tmIBCtrls(IBUNITSINDEX)
                        smIBShow(IBUNITSINDEX, imIBRowNo) = tmIBCtrls(IBUNITSINDEX).sShow
                    End If
                    slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey    'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                    ilRet = gParseItem(slNameCode, 5, "\", slSlspComm)
                    If ilRet <> CP_MSG_NONE Then
                        imIBSave(5, imIBRowNo) = 2  'No
                    Else
                        If Val(slSlspComm) = 0 Then
                            imIBSave(5, imIBRowNo) = 2  'No
                        Else
                            imIBSave(5, imIBRowNo) = 0  'Yes
                        End If
                    End If
                    If Not imTaxDefined Then
                        imIBSave(6, imIBRowNo) = 2
                    Else
                        slNameCode = tmItemCode(imIBSave(3, imIBRowNo) - 1).sKey    'lbcItemCode.List(imIBSave(3, imIBRowNo) - 1)
                        ilRet = gParseItem(slNameCode, 6, "\", slTax)
                        If ilRet <> CP_MSG_NONE Then
                            imIBSave(6, imIBRowNo) = 2  'No
                        Else
                            If slTax = "N" Then
                                imIBSave(6, imIBRowNo) = 1  'No
                            Else
                                imIBSave(6, imIBRowNo) = 0  'Yes
                            End If
                        End If
                    End If
                End If
                mIBTotals
            End If
        Case IBACINDEX
            pbcYN.Visible = False
            If imIBSave(4, imIBRowNo) = 0 Then
                slStr = "Yes"
            ElseIf imIBSave(4, imIBRowNo) = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcPostItem, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(ilBoxNo, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
        Case IBSCINDEX
            pbcYN.Visible = False
            If imIBSave(5, imIBRowNo) = 0 Then
                slStr = "Yes"
            ElseIf imIBSave(5, imIBRowNo) = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcPostItem, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(ilBoxNo, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
        Case IBTXINDEX
            pbcYN.Visible = False
            If imIBSave(6, imIBRowNo) = 0 Then
                slStr = "Yes"
            ElseIf imIBSave(6, imIBRowNo) = 1 Then
                slStr = "No"
            Else
                slStr = ""
            End If
            gSetShow pbcPostItem, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(ilBoxNo, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
        Case IBAMOUNTINDEX
            edcAmount.Visible = False
            slStr = edcAmount.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            gSetShow pbcPostItem, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(IBAMOUNTINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If smIBSave(3, imIBRowNo) <> edcAmount.Text Then
                If imIBRowNo < UBound(tmIBSbf) + 1 Then   'New lines set after all fields entered
                    imIBChg = True
                End If
                smIBSave(3, imIBRowNo) = edcAmount.Text
                If (smIBSave(3, imIBRowNo) = "") Or (smIBSave(5, imIBRowNo) = "") Then
                    smIBSave(6, imIBRowNo) = ""
                    slStr = ""
                Else
                    smIBSave(6, imIBRowNo) = gMulStr(smIBSave(3, imIBRowNo), smIBSave(5, imIBRowNo))
                    slStr = smIBSave(6, imIBRowNo)
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                End If
                gSetShow pbcPostItem, slStr, tmIBCtrls(IBTAMOUNTINDEX)
                smIBShow(IBTAMOUNTINDEX, imIBRowNo) = tmIBCtrls(IBTAMOUNTINDEX).sShow
                mIBTotals
            End If
        Case IBUNITSINDEX
'            edcUnits.Visible = False  'Set visibility
'            slstr = edcUnits.Text
'            gSetShow pbcPostItem, slstr, tmIBCtrls(ilBoxNo)
'            smIBShow(IBUNITSINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
'            If smIBSave(4, imIBRowNo) <> edcUnits.Text Then
'                If imIBRowNo < UBound(tmIBSbf) + 1 Then   'New lines set after all fields entered
'                    imIBChg = True
'                End If
'                smIBSave(4, imIBRowNo) = edcUnits.Text
'            End If
        Case IBNOITEMSINDEX
            edcNoItems.Visible = False  'Set visibility
            slStr = edcNoItems.Text
            gSetShow pbcPostItem, slStr, tmIBCtrls(ilBoxNo)
            smIBShow(IBNOITEMSINDEX, imIBRowNo) = tmIBCtrls(ilBoxNo).sShow
            If smIBSave(5, imIBRowNo) <> edcNoItems.Text Then
                If imIBRowNo < UBound(tmIBSbf) + 1 Then   'New lines set after all fields entered
                    imIBChg = True
                End If
                smIBSave(5, imIBRowNo) = edcNoItems.Text
                If (smIBSave(3, imIBRowNo) = "") Or (smIBSave(5, imIBRowNo) = "") Then
                    smIBSave(6, imIBRowNo) = ""
                    slStr = ""
                Else
                    smIBSave(6, imIBRowNo) = gMulStr(smIBSave(3, imIBRowNo), smIBSave(5, imIBRowNo))
                    slStr = smIBSave(6, imIBRowNo)
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                End If
                gSetShow pbcPostItem, slStr, tmIBCtrls(IBTAMOUNTINDEX)
                smIBShow(IBTAMOUNTINDEX, imIBRowNo) = tmIBCtrls(IBTAMOUNTINDEX).sShow
                mIBTotals
            End If
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mIBTestFields                   *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mIBTestFields() As Integer
'
'   iRet = mIBTestFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilSbf As Integer
    For ilSbf = LBound(tmIBSbf) To UBound(tmIBSbf) - 1 Step 1
        If (tmIBSbf(ilSbf).iStatus = 0) Or (tmIBSbf(ilSbf).iStatus = 1) Then
            If tmIBSbf(ilSbf).SbfRec.iBillVefCode <= 0 Then
                ilRes = MsgBox("Vehicle must be specified", vbOkOnly + vbExclamation, "Incomplete")
                imIBRowNo = ilSbf + 1
                imIBBoxNo = IBVEHICLEINDEX
                mIBTestFields = NO
                Exit Function
            End If
            If (tmIBSbf(ilSbf).SbfRec.iDate(0) = 0) And (tmIBSbf(ilSbf).SbfRec.iDate(1) = 0) Then
                ilRes = MsgBox("Date must be specified", vbOkOnly + vbExclamation, "Incomplete")
                imIBRowNo = ilSbf + 1
                imIBBoxNo = IBDATEINDEX
                mIBTestFields = NO
                Exit Function
            End If
            If tmIBSbf(ilSbf).SbfRec.sDescr = "" Then
                ilRes = MsgBox("Description must be specified", vbOkOnly + vbExclamation, "Incomplete")
                imIBRowNo = ilSbf + 1
                imIBBoxNo = IBDESCRIPTINDEX
                mIBTestFields = NO
                Exit Function
            End If
            'gStrToPDN "", 2, 5, slStr
            'If StrComp(tmIBSbf(ilSbf).SbfRec.sItemAmount, slStr, 0) = 0 Then
            '    ilRes = MsgBox("Price must be specified", vbOkOnly + vbExclamation, "Incomplete")
            '    imIBRowNo = ilSbf + 1
            '    imIBBoxNo = IBAMOUNTINDEX
            '    mIBTestFields = NO
            '    Exit Function
            'End If
            'If tmIBSbf(ilSbf).SbfRec.sUnitName = "" Then
            '    ilRes = MsgBox("Units must be specified", vbOkOnly + vbExclamation, "Incomplete")
            '    imIBRowNo = ilSbf + 1
            '    imIBBoxNo = IBUNITSINDEX
            '    mIBTestFields = NO
            '    Exit Function
            'End If
            If tmIBSbf(ilSbf).SbfRec.iNoItems <= 0 Then
                ilRes = MsgBox("Number of Items must be specified", vbOkOnly + vbExclamation, "Incomplete")
                imIBRowNo = ilSbf + 1
                imIBBoxNo = IBNOITEMSINDEX
                mIBTestFields = NO
                Exit Function
            End If
        End If
    Next ilSbf
    mIBTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mIBTestSaveFields               *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mIBTestSaveFields() As Integer
'
'   iRet = mIBTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    If imIBSave(1, imIBRowNo) < 0 Then
        ilRes = MsgBox("Vehicle must be specified", vbOkOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBVEHICLEINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    If imIBSave(2, imIBRowNo) < 0 Then
        ilRes = MsgBox("Date must be specified", vbOkOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBDATEINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    If smIBSave(2, imIBRowNo) = "" Then
        ilRes = MsgBox("Description must be specified", vbOkOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBDESCRIPTINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    If imIBSave(3, imIBRowNo) <= 0 Then
        ilRes = MsgBox("Item Billing Type must be specified", vbOkOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBITEMTYPEINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    If smIBSave(3, imIBRowNo) = "" Then
        ilRes = MsgBox("Amount/Item must be specified", vbOkOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBAMOUNTINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    If smIBSave(4, imIBRowNo) = "" Then
        ilRes = MsgBox("Units must be specified", vbOkOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBUNITSINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    If smIBSave(5, imIBRowNo) = "" Then
        ilRes = MsgBox("# Items must be specified", vbOkOnly + vbExclamation, "Incomplete")
        imIBBoxNo = IBNOITEMSINDEX
        mIBTestSaveFields = NO
        Exit Function
    End If
    mIBTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mIBTotals                       *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute totals                  *
'*                                                     *
'*******************************************************
Private Sub mIBTotals()
    Dim slCTotal As String
    Dim slPTotal As String
    Dim slBTotal As String
    Dim ilLoop As Integer
    Dim slStr As String

    slCTotal = "0"
    slPTotal = sgIBPTotal
    slBTotal = sgIBBTotal
    For ilLoop = LBound(smIBSave, 2) To UBound(smIBSave, 2) - 1 Step 1
        If smIBSave(1, ilLoop) = "C" Then
            slCTotal = gAddStr(slCTotal, smIBSave(6, ilLoop))
        End If
    Next ilLoop
'    If (slCTotal = "0") And (slPTotal = "0") And (slBTotal = "0") Then
'        lacTotals.Visible = False
'    End If
    lacTotals.BackColor = WHITE
    lacTotals.Visible = True
    slStr = "Totals:"
    If slCTotal <> "0" Then
        gFormatStr slCTotal, FMTLEAVEBLANK + FMTCOMMA, 2, slCTotal
        slStr = slStr & " Contracted $" & slCTotal
    End If
    If slPTotal <> "0" Then
        gFormatStr slPTotal, FMTLEAVEBLANK + FMTCOMMA, 2, slPTotal
        slStr = slStr & "  Posted $" & slPTotal
    End If
    If slBTotal <> "0" Then
        gFormatStr slBTotal, FMTLEAVEBLANK + FMTCOMMA, 2, slBTotal
        slStr = slStr & "  Billed $" & slBTotal
    End If
    lacTotals.Caption = slStr
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
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
    Dim ilRet As Integer    'Return Status
    Dim tlClf As CLF    'Only used to get size of CLF
    Dim tlSbf As SBF    'Only used to get size of SBF
    ReDim smIBSave(1 To 7, 1 To 1)
    ReDim imIBSave(1 To 6, 1 To 1)
    ReDim smIBShow(1 To IBBILLINDEX, 1 To 1)
    ReDim tmClf(0 To 0)
    Screen.MousePointer = vbHourglass
    imLBIBCtrls = 1
    imFirstActivate = True
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    vbcPostItem.Min = LBound(smIBShow, 2)
    vbcPostItem.Max = LBound(smIBShow, 2)
    vbcPostItem.Value = vbcPostItem.Min
    gHlfRead "SBF", tmSbfHelp()
    'gPDNToStr tgSpf.sBTax(0), 2, slStr1
    'gPDNToStr tgSpf.sBTax(1), 2, slStr2
    'If (Val(slStr1) = 0) And (Val(slStr2) = 0) Then
    '12/17/06-Change to tax by agency or vehicle
    'If (tgSpf.iBTax(0) = 0) Or (tgSpf.iBTax(1) = 0) Then
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
        imTaxDefined = True
    Else
        imTaxDefined = False
    End If
    imTerminate = False
    imFirstActivate = True
    imBypassFocus = False
    imIBBoxNo = -1 'Initialize current Box to N/A
    imIBRowNo = -1
    imIBChg = False
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imSettingValue = False
    imChgMode = False
    imBSMode = False
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    mAdvtPop
    If imTerminate Then
        Exit Sub
    End If
    mItemPop
    If imTerminate Then
        Exit Sub
    End If
    hmChf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", PostItem
    On Error GoTo 0
    imChfRecLen = Len(tmChf) 'btrRecordLength(hmChf)    'Get Chf size
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", PostItem
    On Error GoTo 0
    imClfRecLen = Len(tlClf) 'btrRecordLength(hmClf)    'Get Clf size
    hmSbf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sbf.Btr)", PostItem
    On Error GoTo 0
    imSbfRecLen = Len(tlSbf) 'btrRecordLength(hmSbf)    'Get Sbf size
    PostItem.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    'gCenterModalForm PostItem
    gCenterStdAlone PostItem
    'Traffic!plcHelp.Caption = ""
    mInitBox
    lacTotals.Visible = False
    Screen.MousePointer = vbDefault
    'imcHelp.Picture = IconTraf!imcHelp.Picture
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim flTextHeight As Single  'Standard text height
    flTextHeight = pbcPostItem.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcPostItem.Move 135, 555, pbcPostItem.Width + vbcPostItem.Width + fgPanelAdj - 15, pbcPostItem.Height + fgPanelAdj
    pbcPostItem.Move plcPostItem.Left + fgBevelX, plcPostItem.Top + fgBevelY
    vbcPostItem.Move pbcPostItem.Left + pbcPostItem.Width, pbcPostItem.Top + 15
    pbcArrow.Move plcPostItem.Left - pbcArrow.Width - 15    'Vehicle
    gSetCtrl tmIBCtrls(IBVEHICLEINDEX), 30, 375, 690, fgBoxGridH
    'Date
    gSetCtrl tmIBCtrls(IBDATEINDEX), 735, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 825, fgBoxGridH
    'Description
    gSetCtrl tmIBCtrls(IBDESCRIPTINDEX), 1575, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 1905, fgBoxGridH
    'Item Billing
    gSetCtrl tmIBCtrls(IBITEMTYPEINDEX), 3495, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 1110, fgBoxGridH
    'Agency Commission
    gSetCtrl tmIBCtrls(IBACINDEX), 4620, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 180, fgBoxGridH
    'Salesperson Commission
    gSetCtrl tmIBCtrls(IBSCINDEX), 4815, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 180, fgBoxGridH
    'Taxable
    gSetCtrl tmIBCtrls(IBTXINDEX), 5010, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 180, fgBoxGridH
    'Amount/Item
    gSetCtrl tmIBCtrls(IBAMOUNTINDEX), 5205, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 1005, fgBoxGridH
    'Units
    gSetCtrl tmIBCtrls(IBUNITSINDEX), 6225, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 600, fgBoxGridH
    '# Items
    gSetCtrl tmIBCtrls(IBNOITEMSINDEX), 6840, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 600, fgBoxGridH
    'Total Amount
    gSetCtrl tmIBCtrls(IBTAMOUNTINDEX), 7455, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 1005, fgBoxGridH
    'Billed
    gSetCtrl tmIBCtrls(IBBILLINDEX), 8475, tmIBCtrls(IBVEHICLEINDEX).fBoxY, 255, fgBoxGridH
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitIBShow                     *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize show values         *
'*                                                     *
'*******************************************************
Private Sub mInitIBShow()
    Dim ilRowNo As Integer
    Dim ilBoxNo As Integer
    Dim slStr As String
    Dim ilSvIBRowNo As Integer
    Dim ilSvIBBoxNo As Integer
    ilSvIBRowNo = imIBRowNo
    ilSvIBBoxNo = imIBBoxNo
    For ilRowNo = LBound(smIBSave, 2) To UBound(smIBSave, 2) - 1 Step 1
        For ilBoxNo = IBVEHICLEINDEX To IBTAMOUNTINDEX Step 1
            Select Case ilBoxNo 'Branch on box type (control)
                Case IBVEHICLEINDEX 'Vehicle
                    slStr = lbcBVehicle.List(imIBSave(1, ilRowNo))
                    gSetShow pbcPostItem, slStr, tmIBCtrls(IBVEHICLEINDEX)
                    smIBShow(IBVEHICLEINDEX, ilRowNo) = tmIBCtrls(IBVEHICLEINDEX).sShow
                Case IBDATEINDEX 'Date
                    slStr = lbcBDate.List(imIBSave(2, ilRowNo))
                    slStr = gFormatDate(slStr)
                    gSetShow pbcPostItem, slStr, tmIBCtrls(IBDATEINDEX)
                    smIBShow(IBDATEINDEX, ilRowNo) = tmIBCtrls(IBDATEINDEX).sShow
                Case IBDESCRIPTINDEX
                    slStr = smIBSave(2, ilRowNo)
                    gSetShow pbcPostItem, slStr, tmIBCtrls(IBDESCRIPTINDEX)
                    smIBShow(IBDESCRIPTINDEX, ilRowNo) = tmIBCtrls(IBDESCRIPTINDEX).sShow
                Case IBITEMTYPEINDEX 'Date
                    slStr = lbcBItem.List(imIBSave(3, ilRowNo))
                    gSetShow pbcPostItem, slStr, tmIBCtrls(IBITEMTYPEINDEX)
                    smIBShow(IBITEMTYPEINDEX, ilRowNo) = tmIBCtrls(IBITEMTYPEINDEX).sShow
                Case IBACINDEX
                    If imIBSave(4, ilRowNo) = 0 Then
                        slStr = "Yes"
                    ElseIf imIBSave(4, ilRowNo) = 1 Then
                        slStr = "No"
                    Else
                        slStr = ""
                    End If
                    gSetShow pbcPostItem, slStr, tmIBCtrls(ilBoxNo)
                    smIBShow(ilBoxNo, ilRowNo) = tmIBCtrls(ilBoxNo).sShow
                Case IBSCINDEX
                    If imIBSave(5, ilRowNo) = 0 Then
                        slStr = "Yes"
                    ElseIf imIBSave(5, ilRowNo) = 1 Then
                        slStr = "No"
                    Else
                        slStr = ""
                    End If
                    gSetShow pbcPostItem, slStr, tmIBCtrls(ilBoxNo)
                    smIBShow(ilBoxNo, ilRowNo) = tmIBCtrls(ilBoxNo).sShow
                Case IBTXINDEX
                    If imIBSave(6, ilRowNo) = 0 Then
                        slStr = "Yes"
                    ElseIf imIBSave(6, ilRowNo) = 1 Then
                        slStr = "No"
                    Else
                        slStr = ""
                    End If
                    gSetShow pbcPostItem, slStr, tmIBCtrls(ilBoxNo)
                    smIBShow(ilBoxNo, ilRowNo) = tmIBCtrls(ilBoxNo).sShow
                Case IBAMOUNTINDEX
                    slStr = smIBSave(3, ilRowNo)
                    gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    gSetShow pbcPostItem, slStr, tmIBCtrls(IBAMOUNTINDEX)
                    smIBShow(IBAMOUNTINDEX, ilRowNo) = tmIBCtrls(IBAMOUNTINDEX).sShow
                Case IBUNITSINDEX
                    slStr = smIBSave(4, ilRowNo)
                    gSetShow pbcPostItem, slStr, tmIBCtrls(IBUNITSINDEX)
                    smIBShow(IBUNITSINDEX, ilRowNo) = tmIBCtrls(IBUNITSINDEX).sShow
                Case IBNOITEMSINDEX
                    slStr = smIBSave(5, ilRowNo)
                    gSetShow pbcPostItem, slStr, tmIBCtrls(IBNOITEMSINDEX)
                    smIBShow(IBNOITEMSINDEX, ilRowNo) = tmIBCtrls(IBNOITEMSINDEX).sShow
                Case IBTAMOUNTINDEX
                    If (smIBSave(3, ilRowNo) = "") And (smIBSave(5, ilRowNo) = "") Then
                        smIBSave(6, ilRowNo) = ""
                        slStr = ""
                    Else
                        smIBSave(6, ilRowNo) = gMulStr(smIBSave(3, ilRowNo), smIBSave(5, ilRowNo))
                        slStr = smIBSave(6, ilRowNo)
                        gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    End If
                    gSetShow pbcPostItem, slStr, tmIBCtrls(IBTAMOUNTINDEX)
                    smIBShow(IBTAMOUNTINDEX, ilRowNo) = tmIBCtrls(IBTAMOUNTINDEX).sShow
                Case IBBILLINDEX
            End Select
        Next ilBoxNo
    Next ilRowNo
    imIBBoxNo = ilSvIBBoxNo
    imIBRowNo = ilSvIBRowNo
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNewIB                      *
'*                                                     *
'*             Created:9/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize Item billing        *
'*                                                     *
'*******************************************************
Private Sub mInitNewIB(ilRowNo As Integer)
    Dim ilLoop As Integer
    smIBSave(1, ilRowNo) = ""   'Transaction
    smIBSave(2, ilRowNo) = ""   'Description
    smIBSave(3, ilRowNo) = ""   'Amount/item
    smIBSave(4, ilRowNo) = ""   'Units
    smIBSave(5, ilRowNo) = ""   '# items
    smIBSave(6, ilRowNo) = ""   'Total
    smIBSave(7, ilRowNo) = "R"  'Ready to be invoiced
    imIBSave(1, ilRowNo) = -1   'Vehicle
    imIBSave(2, ilRowNo) = -1   'Date
    imIBSave(3, ilRowNo) = -1   'Item billing
    imIBSave(4, ilRowNo) = -1    'Agency commission-default = N
    imIBSave(5, ilRowNo) = -1   'Salesperson commission- Test item bill
    imIBSave(6, ilRowNo) = -1
    If imTaxDefined Then
        imIBSave(6, ilRowNo) = -1   'Taxable- test item bill if taxes defined
    Else
        imIBSave(6, ilRowNo) = 1    'Set to No
    End If
    For ilLoop = IBVEHICLEINDEX To IBBILLINDEX Step 1
        smIBShow(ilLoop, ilRowNo) = ""
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mItemPop                       *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Item Billing Types    *
'*                      list box if required           *
'*                                                     *
'*******************************************************
Private Sub mItemPop()
'
'   mAgyDPPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = lbcBItem.ListIndex
    If ilIndex > 0 Then
        slName = lbcBItem.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopMnfPlusFieldsBox(PostItem, lbcBItem, lbcItemCode, "I")
    ilRet = gPopMnfPlusFieldsBox(PostItem, lbcBItem, tmItemCode(), smItemCodeTag, "I")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mItemPopErr
        gCPErrorMsg ilRet, "mItemPop (gPopMnfPlusFieldsBox)", PostItem
        On Error GoTo 0
        lbcBItem.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, lbcBItem
            If gLastFound(lbcBItem) > 0 Then
                lbcBItem.ListIndex = gLastFound(lbcBItem)
            Else
                lbcBItem.ListIndex = -1
            End If
        Else
            lbcBItem.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mItemPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mItemTypeBranch                 *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set up communication to Item   *
'*                      Billing Type and process       *
'*                      communication back from item   *
'*                      Billing Type                   *
'*                                                     *
'*                                                     *
'*  General flow: pbc--Tab calls this function which   *
'*                initiates a task as a MODAL form.    *
'*                This form and the control loss focus *
'*                When the called task terminates two  *
'*                events are generated (Form activated;*
'*                GotFocus to pbc-Tab).  Also, control *
'*                is sent back to this function (the   *
'*                GotFocus event is initiated after    *
'*                this function finishes processing)   *
'*                                                     *
'*******************************************************
Private Function mItemTypeBranch()
'
'   ilRet = mItemTypeBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    ilRet = gOptionalLookAhead(edcDropDown, lbcBItem, imBSMode, slStr)
    If (ilRet = 0) And (Not imDoubleClickName) Then
        mItemTypeBranch = False
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(ITEMBILLINGTYPESLIST)) Then
    '    mItemTypeBranch = True
    '    mIBEnableBox imIBBoxNo
    '    Exit Function
    'End If
    'MousePointer = vbHourGlass  'Wait
    sgMnfCallType = "I"
    igMNmCallSource = CALLSOURCEPOSTITEM
    If edcDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'Invoice!edcLinkSrceHelpMsg.Text = ""
    'If (Not igStdAloneMode) And (igShowHelpMsg) Then
        If igTestSystem Then
            slStr = "Invoice^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        Else
            slStr = "Invoice^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
        End If
    'Else
    '    If igTestSystem Then
    '        slStr = "Invoice^Test^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    Else
    '        slStr = "Invoice^Prod^NOHELP\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(Str$(igMNmCallSource)) & "\" & sgMNmName
    '    End If
    'End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'PostItem.Enabled = False
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'PostItem.Enabled = True
    'Invoice!edcLinkSrceHelpMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mItemTypeBranch = True
    imUpdateAllowed = ilUpdateAllowed
    'gShowBranner
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        lbcBItem.Clear
        smItemCodeTag = ""
        mItemPop
        If imTerminate Then
            mItemTypeBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, lbcBItem
'        mSetChg AGYDPINDEX
        sgMNmName = ""
        If gLastFound(lbcBItem) > 0 Then
            imChgMode = True
            lbcBItem.ListIndex = gLastFound(lbcBItem)
            edcDropDown.Text = lbcBItem.List(lbcBItem.ListIndex)
            imChgMode = False
            mItemTypeBranch = False
        Else
            imChgMode = True
            lbcBItem.ListIndex = 0
            edcDropDown.Text = lbcBItem.List(0)
            imChgMode = False
            edcDropDown.SetFocus
            Exit Function
        End If
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mIBEnableBox imIBBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        mIBEnableBox imIBBoxNo
        Exit Function
    End If
    Exit Function

    On Error GoTo 0
    imTerminate = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveIBCtrlToRec                *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move controls values to record *
'*                                                     *
'*******************************************************
Private Sub mMoveIBCtrlToRec()
'
'   mMoveIBCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer
    ilIndex = LBound(tmIBSbf)
    For ilLoop = LBound(smIBSave, 2) To UBound(smIBSave, 2) - 1 Step 1
        tmIBSbf(ilIndex).SbfRec.lChfCode = tmChf.lCode
        tmIBSbf(ilIndex).SbfRec.sTranType = smIBSave(1, ilLoop)
        slNameCode = lbcVehicle.List(imIBSave(1, ilLoop))
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveIBCtrlToRecErr
        gCPErrorMsg ilRet, "mMoveIBCtrlToRec (gParseItem field 2)", PostItem
        tmIBSbf(ilIndex).SbfRec.iBillVefCode = CInt(slCode)
        slStr = lbcBDate.List(imIBSave(2, ilLoop))
        gPackDate slStr, tmIBSbf(ilIndex).SbfRec.iDate(0), tmIBSbf(ilIndex).SbfRec.iDate(1)
        tmIBSbf(ilIndex).SbfRec.sDescr = smIBSave(2, ilLoop)
        slNameCode = tmItemCode(imIBSave(3, ilLoop) - 1).sKey   'lbcItemCode.List(imIBSave(3, ilLoop) - 1)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveIBCtrlToRecErr
        gCPErrorMsg ilRet, "mMoveIBCtrlToRec (gParseItem field 2)", PostItem
        tmIBSbf(ilIndex).SbfRec.iMnfItem = CInt(slCode)
        'gStrToPDN smIBSave(3, ilLoop), 2, 5, tmIBSbf(ilIndex).SbfRec.sItemAmount
        tmIBSbf(ilIndex).SbfRec.lGross = gStrDecToLong(smIBSave(3, ilLoop), 2)
        slStr = smIBSave(4, ilLoop)
        ilPos = InStr(1, slStr, "per ", 1)
        If ilPos = 1 Then
            slStr = Right$(slStr, Len(slStr) - 4)
        End If
        If Len(slStr) > 6 Then
            slStr = Left$(slStr, 6)
        End If
        'tmIBSbf(ilIndex).SbfRec.sUnitName = slStr
        tmIBSbf(ilIndex).SbfRec.iNoItems = Val(smIBSave(5, ilLoop))
        If imIBSave(4, ilLoop) = 0 Then
            tmIBSbf(ilIndex).SbfRec.sAgyComm = "Y"
        Else
            tmIBSbf(ilIndex).SbfRec.sAgyComm = "N"
        End If
        If imIBSave(5, ilLoop) = 0 Then
            tmIBSbf(ilIndex).SbfRec.sSlsComm = "Y"
        Else    'Map -1 or 1 or 2 as No
            tmIBSbf(ilIndex).SbfRec.sSlsComm = "N"
        End If
        '12/17/06-Change to tax by agency or vehicle
        'If imIBSave(6, ilLoop) = 0 Then
        '    tmIBSbf(ilIndex).SbfRec.sSlsTax = "Y"
        'Else    'Map -1 or 1 or 2 as No
        '    tmIBSbf(ilIndex).SbfRec.sSlsTax = "N"
        'End If
        tmIBSbf(ilIndex).SbfRec.sBilled = smIBSave(7, ilLoop)
        If tmIBSbf(ilIndex).iStatus = -1 Then
            tmIBSbf(ilIndex).iStatus = 0
        ElseIf tmIBSbf(ilIndex).iStatus = 2 Then
            tmIBSbf(ilIndex).iStatus = 1
        End If
        ilIndex = ilIndex + 1
        If ilIndex > UBound(tmIBSbf) Then
            ReDim Preserve tmIBSbf(0 To ilIndex)
        End If
    Next ilLoop
    For ilLoop = ilIndex To UBound(tmIBSbf) Step 1
        If tmIBSbf(ilIndex).iStatus = 0 Then
            tmIBSbf(ilIndex).iStatus = -1
        ElseIf tmIBSbf(ilIndex).iStatus = 1 Then
            tmIBSbf(ilIndex).iStatus = 2
        End If
        tmIBSbf(ilIndex).SbfRec.sTranType = ""
        tmIBSbf(ilIndex).SbfRec.iBillVefCode = 0
        tmIBSbf(ilIndex).SbfRec.iDate(0) = 0
        tmIBSbf(ilIndex).SbfRec.iDate(1) = 0
        'gStrToPDN "", 2, 5, tmIBSbf(ilIndex).SbfRec.sItemAmount
        tmIBSbf(ilIndex).SbfRec.lGross = 0
        tmIBSbf(ilIndex).SbfRec.iMnfItem = 0
        tmIBSbf(ilIndex).SbfRec.iNoItems = 0
        'tmIBSbf(ilIndex).SbfRec.sUnitName = ""
        tmIBSbf(ilIndex).SbfRec.sDescr = ""
        tmIBSbf(ilIndex).SbfRec.sBilled = "N"
        tmIBSbf(ilIndex).SbfRec.iPrintInvDate(0) = 0
        tmIBSbf(ilIndex).SbfRec.iPrintInvDate(1) = 0
    Next ilLoop
    Exit Sub
mMoveIBCtrlToRecErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveIBRecToCtrl                *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveIBRecToCtrl()
'
'   mMoveIBRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilIndex As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slDate As String
    For ilLoop = 0 To UBound(tmIBSbf) - 1 Step 1
        imIBSave(1, ilLoop + 1) = -1
        slRecCode = Trim$(Str$(tmIBSbf(ilLoop).SbfRec.iBillVefCode))
        For ilTest = 0 To lbcVehicle.ListCount - 1 Step 1
            slNameCode = lbcVehicle.List(ilTest)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveIBRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveIBRecToCtrl (gParseItem field 2)", PostItem
            On Error GoTo 0
            If slRecCode = slCode Then
                imIBSave(1, ilLoop + 1) = ilTest
                Exit For
            End If
        Next ilTest
        If imIBSave(1, ilLoop + 1) = -1 Then
            gFindMatch sgUserDefVehicleName, 0, lbcBVehicle
            If gLastFound(lbcBVehicle) >= 0 Then
                imIBSave(1, ilLoop + 1) = gLastFound(lbcBVehicle)
            Else
                imIBSave(1, ilLoop + 1) = 0
            End If
        End If
        If tmIBSbf(ilLoop).SbfRec.sTranType = "C" Then
            smIBSave(1, ilLoop + 1) = "C"
        Else
            smIBSave(1, ilLoop + 1) = "I"
        End If
        gUnpackDate tmIBSbf(ilLoop).SbfRec.iDate(0), tmIBSbf(ilLoop).SbfRec.iDate(1), slDate
        gFindMatch slDate, 0, lbcBDate
        If gLastFound(lbcBDate) >= 0 Then
            imIBSave(2, ilLoop + 1) = gLastFound(lbcBDate)
        Else
        End If
        smIBSave(2, ilLoop + 1) = tmIBSbf(ilLoop).SbfRec.sDescr
        slRecCode = Trim$(Str$(tmIBSbf(ilLoop).SbfRec.iMnfItem))
        For ilIndex = 0 To UBound(tmItemCode) - 1 Step 1 'lbcItemCode.ListCount - 1 Step 1
            slNameCode = tmItemCode(ilIndex).sKey    'lbcItemCode.List(ilIndex)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveIBRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveIBRecToCtrl (gParseItem field 2)", PostItem
            On Error GoTo 0
            If slRecCode = slCode Then
                imIBSave(3, ilLoop + 1) = ilIndex + 1
                ilRet = gParseItem(slNameCode, 5, "\", slCode)
                On Error GoTo mMoveIBRecToCtrlErr
                gCPErrorMsg ilRet, "mMoveIBRecToCtrl (gParseItem field 2)", PostItem
                On Error GoTo 0
                If Val(slCode) = 0 Then
                    imIBSave(5, ilLoop + 1) = 2
                Else
                    If tmIBSbf(ilLoop).SbfRec.sSlsComm = "Y" Then
                        imIBSave(5, ilLoop + 1) = 0
                    Else
                        imIBSave(5, ilLoop + 1) = 1
                    End If
                End If
                Exit For
            End If
        Next ilIndex
        'gPDNToStr tmIBSbf(ilLoop).SbfRec.sItemAmount, 2, smIBSave(3, ilLoop + 1)
        smIBSave(3, ilLoop + 1) = gLongToStrDec(tmIBSbf(ilLoop).SbfRec.lGross, 2)
        'smIBSave(4, ilLoop + 1) = tmIBSbf(ilLoop).SbfRec.sUnitName
        smIBSave(5, ilLoop + 1) = Trim$(Str$(tmIBSbf(ilLoop).SbfRec.iNoItems))
        If tmIBSbf(ilLoop).SbfRec.sAgyComm = "Y" Then
            imIBSave(4, ilLoop + 1) = 0
        Else
            imIBSave(4, ilLoop + 1) = 1
        End If
        If Not imTaxDefined Then
            imIBSave(6, ilLoop + 1) = 2
        Else
            '12/17/06-Change to tax by agency or vehicle
            'If tmIBSbf(ilLoop).SbfRec.sSlsTax = "Y" Then
            '    imIBSave(6, ilLoop + 1) = 0
            'Else
            '    imIBSave(6, ilLoop + 1) = 1
            'End If
        End If
        smIBSave(6, ilLoop + 1) = gMulStr(smIBSave(3, ilLoop + 1), smIBSave(5, ilLoop + 1))
        smIBSave(7, ilLoop + 1) = tmIBSbf(ilLoop).SbfRec.sBilled
    Next ilLoop
    mInitNewIB UBound(smIBSave, 2)
    Exit Sub
mMoveIBRecToCtrlErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mPopulate                       *
'*                                                     *
'*             Created:6/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box                            *
'*                                                     *
'*******************************************************
Private Sub mPopulate()
'
'   mPopulate
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    Dim slNameCode As String  'Name and code
    Dim slCode As String    'Code number
    Dim ilCurrent As Integer
    Dim ilAAS As Integer
    Dim ilShow As Integer
    Dim ilAdfCode As Integer
    Dim slCntrStatus  As String
    Dim slCntrType As String
    Dim ilState As Integer
    Screen.MousePointer = vbHourglass  'Wait
    slNameCode = tgAdvertiser(cbcAdvt.ListIndex).sKey  'Traffic!lbcAdvertiser.List(cbcAdvt.ListIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mPopulateErr
    gCPErrorMsg ilRet, "mPopulate (gParseItem field 2)", PostItem
    On Error GoTo 0
    slCode = Trim$(slCode)
    ilAdfCode = Val(slCode)
    ilCurrent = 0
    ilAAS = 0
    slCntrStatus = "HO"
    slCntrType = "CRQ"
    ilShow = 3
    ilState = 1
    'ilRet = gPopCntrBox(PostItem, 1, ilFilter, ilCurrent, 0, -1, cbcSelect, lbcContract, True, False, False, False)
    'ilRet = gPopCntrForAASBox(PostItem, ilAAS, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, cbcSelect, lbcContract)
    ilRet = gPopCntrForAASBox(PostItem, ilAAS, ilAdfCode, slCntrStatus, slCntrType, ilCurrent, ilState, ilShow, cbcSelect, tmContract(), smContractTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mPopulateErr
        gCPErrorMsg ilRet, "mPopulate (gPopCntrForAASBox)", PostItem
        On Error GoTo 0
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mPopulateErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadChfRec                     *
'*                                                     *
'*             Created:7/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadChfRec(ilSelectIndex As Integer) As Integer
'
'   iRet = mReadChfRec(ilSelectIndex)
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim slNameCode As String  'Code and Code strings from Traffic!lbcLockBox or Traffic!lbcAgencyDP
    Dim slCode As String    'Code number- so record can be found
    Dim ilRet As Integer    'Return status

    slNameCode = tmContract(ilSelectIndex).sKey 'lbcContract.List(ilSelectIndex)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    On Error GoTo mReadChfRecErr
    gCPErrorMsg ilRet, "mReadChfRecErr (gParseItem field 2)", PostItem
    On Error GoTo 0
    slCode = Trim$(slCode)
    tmChfSrchKey.lCode = CLng(slCode)
    ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    On Error GoTo mReadChfRecErr
    gBtrvErrorMsg ilRet, "mReadChfRecErr (btrGetEqual: Contract)", PostItem
    On Error GoTo 0
    mReadChfRec = True
    If Not mReadClfRec() Then
        mReadChfRec = False
    End If
    mDatePop
    Exit Function
mReadChfRecErr:
    On Error GoTo 0
    mReadChfRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadClfRec                     *
'*                                                     *
'*             Created:8/01/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadClfRec() As Integer
'
'   iRet = mReadClfRec
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpperBound As Integer
    Dim tlClf As CLF
    Dim tlClfExt As CLFEXT    'Contract line extract record
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffset As Integer

    ReDim tmClf(0 To 0)
    ilUpperBound = UBound(tmClf)
    btrExtClear hmClf   'Clear any previous extend operation
    ilExtLen = Len(tlClfExt)  'Extract operation record size
    tmClfSrchKey.lChfCode = tmChf.lCode
    tmClfSrchKey.iLine = 0
    tmClfSrchKey.iCntRevNo = 32000
    tmClfSrchKey.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hmClf, tlClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If (tlClf.lChfCode = tmChf.lCode) And (tlClf.sDelete <> "Y") Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        'tmClf(ilUpperBound) = tlClf
        'ilUpperBound = ilUpperBound + 1
        'ReDim Preserve tmClf(0 To ilUpperBound)
        Call btrExtSetBounds(hmClf, llNoRec, -1, "UC", "CLFEXTPK", CLFEXTPK) 'Set extract limits (all records)
        ilOffset = gFieldOffset("Clf", "ClfChfCode")
        ilRet = btrExtAddLogicConst(hmClf, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tmChf.lCode, 4)
        On Error GoTo mReadClfRecErr
        gBtrvErrorMsg ilRet, "mReadClfRec (btrExtAddLogicConst):" & "Clf.Btr", PostItem
        On Error GoTo 0
        ilOffset = gFieldOffset("Clf", "ClfDelete")
        ilRet = btrExtAddLogicConst(hmClf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, ByVal "Y", 1)
        On Error GoTo mReadClfRecErr
        gBtrvErrorMsg ilRet, "mReadClfRec (btrExtAddLogicConst):" & "Clf.Btr", PostItem
        On Error GoTo 0
        ilOffset = gFieldOffset("Clf", "ClfChfCode")
        ilRet = btrExtAddField(hmClf, ilOffset, ilExtLen - 1) 'Extract start/end time, and days
        On Error GoTo mReadClfRecErr
        gBtrvErrorMsg ilRet, "mReadCLFRec (btrExtAddField):" & "Clf.Btr", PostItem
        On Error GoTo 0
        ilOffset = gFieldOffset("Clf", "ClfSchStatus")
        ilRet = btrExtAddField(hmClf, ilOffset, 1) 'Extract start/end time, and days
        On Error GoTo mReadClfRecErr
        gBtrvErrorMsg ilRet, "mReadCLFRec (btrExtAddField):" & "Clf.Btr", PostItem
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hmClf)    'Extract record
        ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mReadClfRecErr
            gBtrvErrorMsg ilRet, "mReadClfRec (btrExtGetNextExt):" & "Clf.Btr", PostItem
            On Error GoTo 0
            'ilRet = btrExtGetFirst(hmClf, tlClfExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilRet = btrGetDirect(hmClf, tmClf(ilUpperBound), imClfRecLen, llRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                On Error GoTo mReadClfRecErr
                gBtrvErrorMsg ilRet, "ReadClfRec (btrGetDirect):" & "Clf.Btr", PostItem
                On Error GoTo 0
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tmClf(0 To ilUpperBound)
                ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmClf, tlClfExt, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    mReadClfRec = True
    Exit Function
mReadClfRecErr:
    On Error GoTo 0
    mReadClfRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read item bill records         *
'*                                                     *
'*******************************************************
Private Function mReadRec() As Integer
'
'   iRet = mReadRec
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim tlSbf As SBF
    Dim ilRet As Integer    'Return status
    Dim ilIBUpperBound As Integer
    Dim llRecPos As Long
    imSbfRecLen = Len(tlSbf)
    ReDim tmIBSbf(0 To 0) As SBFLIST
    ilIBUpperBound = UBound(tmIBSbf)
    tmIBSbf(ilIBUpperBound).iStatus = -1 'Not Used
    tmIBSbf(ilIBUpperBound).lRecPos = 0
    If imSelectedIndex = 0 Then
        mReadRec = True
        Exit Function
    End If
    tmSbfSrchKey.lChfCode = tmChf.lCode
    tmSbfSrchKey.iDate(0) = 0
    tmSbfSrchKey.iDate(1) = 0
    tmSbfSrchKey.sTranType = " "
    ilRet = btrGetGreaterOrEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tlSbf.lChfCode = tmChf.lCode)
        ilRet = btrGetPosition(hmSbf, llRecPos)
        If (tlSbf.sTranType = "C") Or (tlSbf.sTranType = "I") Then
            tmIBSbf(ilIBUpperBound).SbfRec = tlSbf
            tmIBSbf(ilIBUpperBound).lRecPos = llRecPos
            tmIBSbf(ilIBUpperBound).iStatus = 1
            ilIBUpperBound = ilIBUpperBound + 1
            ReDim Preserve tmIBSbf(0 To ilIBUpperBound) As SBFLIST
            tmIBSbf(ilIBUpperBound).iStatus = -1 'Not Used
            tmIBSbf(ilIBUpperBound).lRecPos = 0
        End If
        ilRet = btrGetNext(hmSbf, tlSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_KEY_NOT_FOUND) Then
        On Error GoTo mReadRecErr
        gBtrvErrorMsg ilRet, "mReadRec (btrGetEqual)", PostItem
        On Error GoTo 0
    End If
    vbcPostItem.Min = LBound(tmIBSbf) + 1
    If UBound(tmIBSbf) <= vbcPostItem.LargeChange Then
        vbcPostItem.Max = LBound(tmIBSbf) + 1
    Else
        vbcPostItem.Max = UBound(tmIBSbf) - vbcPostItem.LargeChange
    End If
    vbcPostItem.Value = vbcPostItem.Min
    ReDim smIBSave(1 To 7, 1 To UBound(tmIBSbf) + 1) As String
    ReDim imIBSave(1 To 6, 1 To UBound(tmIBSbf) + 1) As Integer
    ReDim smIBShow(1 To IBBILLINDEX, 1 To UBound(tmIBSbf) + 1) As String
    mReadRec = True
    Exit Function
mReadRecErr:
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRec                        *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Update or added record          *
'*                                                     *
'*******************************************************
Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilLoop As Integer   'For loop control
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim llSbfRecPos As Long
    Dim slMsg As String
    Dim tlSbf As SBF
    Dim tlSbf1 As MOVEREC
    Dim tlSbf2 As MOVEREC
    mIBSetShow imIBBoxNo
    mMoveIBCtrlToRec
    If mIBTestFields() = NO Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    For ilLoop = LBound(tmIBSbf) To UBound(tmIBSbf) - 1 Step 1
        If tmIBSbf(ilLoop).iStatus >= 0 Then
            Do  'Loop until record updated or added
                If tmIBSbf(ilLoop).iStatus = 0 Then 'New selected
                    tmIBSbf(ilLoop).SbfRec.lCode = 0
                    tmIBSbf(ilLoop).SbfRec.iCalCarryBonus = 0
                    tmIBSbf(ilLoop).SbfRec.lChfCode = tmChf.lCode
                    tmIBSbf(ilLoop).SbfRec.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                    tmIBSbf(ilLoop).SbfRec.sTranType = "I"
                    ilRet = btrInsert(hmSbf, tmIBSbf(ilLoop).SbfRec, imSbfRecLen, INDEXKEY0)
                    slMsg = "mSaveRec (btrInsert: Item Billing)"
                Else 'Old record-Update
                    slMsg = "mSaveRec (btrGetDirect: Item Billing)"
                    ilRet = btrGetDirect(hmSbf, tlSbf, imSbfRecLen, tmIBSbf(ilLoop).lRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, slMsg, PostItem
                    On Error GoTo 0
                    If tmIBSbf(ilLoop).iStatus = 2 Then
                        ilRet = btrGetPosition(hmSbf, llSbfRecPos)
                        Do
                            'tmRec = tlSbf
                            'ilRet = gGetByKeyForUpdate("SBF", hmSbf, tmRec)
                            'tlSbf = tmRec
                            ilRet = btrDelete(hmSbf)
                            If ilRet = BTRV_ERR_CONFLICT Then
                                ilCRet = btrGetDirect(hmSbf, tlSbf, imSbfRecLen, llSbfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        slMsg = "mSaveRec (btrDelete: Item Billing)"
                    Else
                        tlSbf1 = tlSbf
                        tlSbf2 = tmIBSbf(ilLoop).SbfRec
                        If StrComp(tlSbf1.sChar, tlSbf2.sChar, 0) <> 0 Then
                            tmIBSbf(ilLoop).SbfRec.lCode = 0
                            tmIBSbf(ilLoop).SbfRec.iCalCarryBonus = 0
                            tmIBSbf(ilLoop).SbfRec.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                            ilRet = btrGetPosition(hmSbf, llSbfRecPos)
                            Do
                                'tmRec = tlSbf
                                'ilRet = gGetByKeyForUpdate("SBF", hmSbf, tmRec)
                                'tlSbf = tmRec
                                ilRet = btrDelete(hmSbf)
                                If ilRet = BTRV_ERR_CONFLICT Then
                                    ilCRet = btrGetDirect(hmSbf, tlSbf, imSbfRecLen, llSbfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            ilRet = btrInsert(hmSbf, tmIBSbf(ilLoop).SbfRec, imSbfRecLen, INDEXKEY0)
                            slMsg = "mSaveRec (btrUpdate: Item Billing)"
                        Else
                            ilRet = BTRV_ERR_NONE
                        End If
                    End If
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, PostItem
            On Error GoTo 0
        End If
    Next ilLoop
    mSaveRec = True
    Screen.MousePointer = vbDefault
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if record altered and*
'*                      requires updating              *
'*                                                     *
'*******************************************************
Private Function mSaveRecChg(ilAsk As Integer) As Integer
'
'   iAsk = True
'   iRet = mSaveRecChg(iAsk)
'   Where:
'       iAsk (I)- True = Ask if changed records should be updated;
'                 False= Update record if required without asking user
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRes As Integer
    Dim slMess As String
    If imIBChg Then
        If ilAsk Then
            slMess = "Save Changes"
            ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
            If ilRes = vbCancel Then
                mSaveRecChg = False
                Exit Function
            End If
            If ilRes = vbYes Then
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
            If ilRes = vbNo Then
            End If
        Else
            ilRes = mSaveRec()
            mSaveRecChg = ilRes
            Exit Function
        End If
    End If
    mSaveRecChg = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set command buttons (enable or *
'*                      disabled)                      *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    'Update button set if all mandatory fields have data and any field altered
    'Revert button set if any field changed
    If imIBChg Then
        cmcUndo.Enabled = True
    Else
        cmcUndo.Enabled = False
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
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
    Erase tmItemCode
    Erase tmContract
    Erase tmBVehicle
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    ilRet = btrClose(hmChf)
    btrDestroy hmChf
    Erase tmClf
    Erase tmIBSbf
    Erase smIBSave
    Erase imIBSave
    Erase smIBShow
    Erase tmSbfHelp
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload PostItem
    Set PostItem = Nothing   'Remove data segment
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the Vehicle box      *
'*                                                     *
'*******************************************************
Private Sub mVehPop()
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilIndex As Integer
    Dim tlVsf As VSF
    Dim hlVsf As Integer
    Dim ilRecLen As Integer     'Vsf record length
    Dim tlSrchKey As INTKEY0
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String


    hlVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mVehPopErr
    gBtrvErrorMsg ilRet, "mVehPop (btrOpen)" & "Vsf.Btr", PostItem
    On Error GoTo 0
    ilRecLen = Len(tlVsf)  'btrRecordLength(hlVpf)  'Get and save record length
    lbcVehicle.Clear
    smBVehicleTag = ""
    'lbcBVehicle.Clear
    'lbcBVehicle.Tag = ""
    'ilRet = gPopUserVehicleBox(PostItem, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcVehicle, lbcBVehicle)
    ilRet = gPopUserVehicleBox(PostItem, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH, lbcVehicle, tmBVehicle(), smBVehicleTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gPopUserVehicleBox: Vehicle/Combo)", PostItem
        On Error GoTo 0
    End If
    lbcVehicle.Clear
    For ilClf = LBound(tmClf) To UBound(tmClf) - 1 Step 1
            If tmClf(ilClf).iVefCode > 0 Then
                slRecCode = Trim$(Str$(tmClf(ilLoop).iVefCode))
                For ilTest = 0 To UBound(tmBVehicle) - 1 Step 1 'lbcBVehicle.ListCount - 1 Step 1
                    slNameCode = tmBVehicle(ilTest).sKey    'lbcBVehicle.List(ilTest)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    On Error GoTo mVehPopErr
                    gCPErrorMsg ilRet, "mVehPop (gParseItem field 2)", PostItem
                    On Error GoTo 0
                    If slRecCode = slCode Then
                        gFindMatch slNameCode, 0, lbcVehicle
                        If gLastFound(lbcVehicle) < 0 Then
                            lbcVehicle.AddItem slNameCode
                        End If
                        Exit For
                    End If
                Next ilTest
            Else
                tlSrchKey.iCode = -tmClf(ilClf).iVefCode
                ilRet = btrGetEqual(hlVsf, tlVsf, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                On Error GoTo mVehPopErr
                gBtrvErrorMsg ilRet, "mVehPop (btrGetEqual)", PostItem
                On Error GoTo 0
                For ilIndex = LBound(tlVsf.iFSCode) To UBound(tlVsf.iFSCode) Step 1
                    If tlVsf.iFSCode(ilIndex) <= 0 Then
                        Exit For
                    End If
                    slRecCode = Trim$(Str$(tlVsf.iFSCode(ilIndex)))
                    For ilTest = 0 To UBound(tmBVehicle) - 1 Step 1 'lbcBVehicle.ListCount - 1 Step 1
                        slNameCode = tmBVehicle(ilTest).sKey    'lbcBVehicle.List(ilTest)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        On Error GoTo mVehPopErr
                        gCPErrorMsg ilRet, "mVehPop (gParseItem field 2)", PostItem
                        On Error GoTo 0
                        If slRecCode = slCode Then
                            gFindMatch slNameCode, 0, lbcVehicle
                            If gLastFound(lbcVehicle) < 0 Then
                                lbcVehicle.AddItem slNameCode
                            End If
                            Exit For
                        End If
                    Next ilTest
                Next ilIndex
            End If
    Next ilClf
    lbcBVehicle.Clear
    For ilLoop = 0 To lbcVehicle.ListCount - 1 Step 1
        slNameCode = lbcVehicle.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gParseItem field 1)", PostItem
        On Error GoTo 0
        ilRet = gParseItem(slName, 3, "|", slName)
        On Error GoTo mVehPopErr
        gCPErrorMsg ilRet, "mVehPop (gParseItem field 1)", PostItem
        On Error GoTo 0
        lbcBVehicle.AddItem Trim$(slName)
    Next ilLoop
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    imTerminate = True
    Exit Sub
End Sub
Private Sub pbcArrow_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub pbcArrow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcClickFocus_GotFocus()
    mIBSetShow imIBBoxNo
    imIBBoxNo = -1
    pbcArrow.Visible = False
    lacIBFrame.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcIBSTab_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub pbcIBSTab_GotFocus()
    Dim ilBox As Integer
    Dim slStr As String
    If GetFocus() <> pbcIBSTab.hWnd Then
        Exit Sub
    End If
    If imIBBoxNo = IBITEMTYPEINDEX Then
        If mItemTypeBranch() Then
            Exit Sub
        End If
    End If
    imTabDirection = -1  'Set-right to left
    Select Case imIBBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = 0  'Set-Left to right
            imSettingValue = True
            vbcPostItem.Value = vbcPostItem.Min
            If UBound(smIBSave, 2) <= vbcPostItem.LargeChange + 1 Then 'was <=
                vbcPostItem.Max = LBound(smIBSave, 2)
            Else
                vbcPostItem.Max = UBound(smIBSave, 2) - vbcPostItem.LargeChange ' - 1
            End If
            imIBRowNo = 1
            Do While (imIBRowNo < UBound(smIBSave, 2)) And (smIBSave(7, imIBRowNo) = "B")
                imIBRowNo = imIBRowNo + 1
                If imIBRowNo > vbcPostItem.Value + vbcPostItem.LargeChange Then
                    imSettingValue = True
                    vbcPostItem.Value = vbcPostItem.Value + 1
                End If
            Loop
            If (imIBRowNo = UBound(smIBSave, 2)) And (imIBSave(1, 1) = -1) Then
                mInitNewIB imIBRowNo
            End If
            ilBox = 1
            imIBBoxNo = ilBox
            mIBEnableBox ilBox
            Exit Sub
        Case IBVEHICLEINDEX 'Name (first control within header)
            mIBSetShow imIBBoxNo
            ilBox = IBNOITEMSINDEX
            Do
                If imIBRowNo <= 1 Then
                    imIBBoxNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imIBRowNo = imIBRowNo - 1
                If imIBRowNo < vbcPostItem.Value Then
                    imSettingValue = True
                    vbcPostItem.Value = vbcPostItem.Value - 1
                End If
            Loop While (smIBSave(7, imIBRowNo) = "B")
            imIBBoxNo = ilBox
            mIBEnableBox ilBox
            Exit Sub
        Case IBNOITEMSINDEX
            ilBox = IBAMOUNTINDEX
        Case IBAMOUNTINDEX
            If imIBSave(6, imIBRowNo) = 2 Then
                If imIBSave(5, imIBRowNo) = 2 Then
                    ilBox = IBACINDEX
                Else
                    ilBox = IBSCINDEX
                End If
            Else
                ilBox = IBTXINDEX
            End If
        Case IBTXINDEX
            If imIBSave(5, imIBRowNo) = 2 Then
                ilBox = IBACINDEX
            Else
                ilBox = IBSCINDEX
            End If
        Case IBDATEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = IBVEHICLEINDEX
        Case Else
            ilBox = imIBBoxNo - 1
    End Select
    mIBSetShow imIBBoxNo
    imIBBoxNo = ilBox
    mIBEnableBox ilBox
End Sub
Private Sub pbcIBSTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcIBTab_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub pbcIBTab_GotFocus()
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim slStr As String

    If GetFocus() <> pbcIBTab.hWnd Then
        Exit Sub
    End If
    If imIBBoxNo = IBITEMTYPEINDEX Then
        If mItemTypeBranch() Then
            Exit Sub
        End If
    End If
    imTabDirection = 0  'Set-Left to right
    Select Case imIBBoxNo
        Case -1 'Tab from control prior to form area
            imTabDirection = -1  'Set-Right to left
            imIBRowNo = UBound(smIBSave, 2)
            imSettingValue = True
            If imIBRowNo <= vbcPostItem.LargeChange + 1 Then
                vbcPostItem.Value = vbcPostItem.Min
            Else
                vbcPostItem.Value = imIBRowNo - vbcPostItem.LargeChange
            End If
            ilBox = 1
        Case IBVEHICLEINDEX
            If (imIBRowNo >= UBound(smIBSave, 2)) And (lbcBVehicle.ListIndex < 0) Then
                mIBSetShow imIBBoxNo
                For ilLoop = IBVEHICLEINDEX To IBTAMOUNTINDEX Step 1
                    slStr = ""
                    gSetShow pbcPostItem, slStr, tmIBCtrls(ilLoop)
                    smIBShow(ilLoop, imIBRowNo) = tmIBCtrls(ilLoop).sShow
                Next ilLoop
                imIBSave(1, imIBRowNo) = -1
                imIBBoxNo = -1
                pbcPostItem_Paint
                If cmcDone.Enabled Then
                    cmcDone.SetFocus
                Else
                    cmcCancel.SetFocus
                End If
                Exit Sub
            End If
            ilBox = IBDATEINDEX
        Case IBDATEINDEX
            slStr = edcDropDown.Text
            If slStr <> "" Then
                If Not gValidDate(slStr) Then
                    Beep
                    edcDropDown.SetFocus
                    Exit Sub
                End If
            Else
                Beep
                edcDropDown.SetFocus
                Exit Sub
            End If
            ilBox = imIBBoxNo + 1
        Case IBACINDEX
            If imIBSave(5, imIBRowNo) = 2 Then  '2=no salesperson commission defined
                If imIBSave(6, imIBRowNo) = 2 Then  '2=no tax defined
                    ilBox = IBAMOUNTINDEX
                Else
                    ilBox = IBTXINDEX
                End If
            Else
                ilBox = IBSCINDEX
            End If
        Case IBSCINDEX
            If imIBSave(6, imIBRowNo) = 2 Then
                ilBox = IBAMOUNTINDEX
            Else
                ilBox = IBTXINDEX
            End If
        Case IBAMOUNTINDEX
            ilBox = IBNOITEMSINDEX
        Case IBNOITEMSINDEX 'Last control
            mIBSetShow imIBBoxNo
            If mIBTestSaveFields() = NO Then
                mIBEnableBox imIBBoxNo
                Exit Sub
            End If
            If imIBRowNo >= UBound(smIBSave, 2) Then
                gSetShow pbcPostItem, "R", tmIBCtrls(IBBILLINDEX)
                smIBShow(IBBILLINDEX, imIBRowNo) = tmIBCtrls(IBBILLINDEX).sShow
                imIBChg = True
                ReDim Preserve smIBSave(1 To 7, 1 To imIBRowNo + 1) As String
                ReDim Preserve imIBSave(1 To 6, 1 To imIBRowNo + 1) As Integer
                ReDim Preserve smIBShow(1 To IBBILLINDEX, 1 To imIBRowNo + 1) As String
                mInitNewIB imIBRowNo + 1
                If UBound(smIBSave, 2) <= vbcPostItem.LargeChange + 1 Then 'was <=
                    vbcPostItem.Max = LBound(smIBSave, 2)
                Else
                    vbcPostItem.Max = UBound(smIBSave, 2) - vbcPostItem.LargeChange ' - 1
                End If
                mIBTotals
            End If
            Do
                imIBRowNo = imIBRowNo + 1
                If imIBRowNo > vbcPostItem.Value + vbcPostItem.LargeChange Then
                    imSettingValue = True
                    vbcPostItem.Value = vbcPostItem.Value + 1
                End If
            Loop While (smIBSave(7, imIBRowNo) = "B")
            If imIBRowNo >= UBound(smIBSave, 2) Then
                mSetCommands
                imIBBoxNo = 0
                lacIBFrame.Move 0, tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15) - 30
                lacIBFrame.Visible = True
                pbcArrow.Move pbcArrow.Left, plcPostItem.Top + tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15) + 45
                pbcArrow.Visible = True
                pbcArrow.SetFocus
                Exit Sub
            Else
                ilBox = 1
                imIBBoxNo = ilBox
                mIBEnableBox ilBox
            End If
            Exit Sub
        Case Else
            ilBox = imIBBoxNo + 1
    End Select
    mIBSetShow imIBBoxNo
    imIBBoxNo = ilBox
    mIBEnableBox ilBox
End Sub
Private Sub pbcIBTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub pbcPostItem_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub pbcPostItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fmDragX = X
    fmDragY = Y
    imDragType = 0
    tmcDrag.Enabled = True  'Start timer to see if drag or click
End Sub
Private Sub pbcPostItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
    ilCompRow = vbcPostItem.LargeChange + 1
    If UBound(smIBSave, 2) > ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smIBSave, 2)
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBIBCtrls To UBound(tmIBCtrls) Step 1
            If (X >= tmIBCtrls(ilBox).fBoxX) And (X <= (tmIBCtrls(ilBox).fBoxX + tmIBCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmIBCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmIBCtrls(ilBox).fBoxY + tmIBCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcPostItem.Value - 1
                    mIBSetShow imIBBoxNo
                    If (smIBSave(1, ilRowNo) = "I") And (smIBSave(7, ilRowNo) = "B") Then
                        Beep
                        Exit Sub
                    End If
                    If ilBox = IBUNITSINDEX Then
                        Beep
'                        Exit Sub
                    End If
                    If ilBox = IBTAMOUNTINDEX Then
                        Beep
                        Exit Sub
                    End If
                    imIBRowNo = ilRowNo
                    imIBBoxNo = ilBox
                    mIBEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcPostItem_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer

    ilStartRow = vbcPostItem.Value  'Top location
    ilEndRow = vbcPostItem.Value + vbcPostItem.LargeChange
    If ilEndRow > UBound(smIBSave, 2) Then
        ilEndRow = UBound(smIBSave, 2) 'include blank row as it might have data
    End If
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBIBCtrls To UBound(tmIBCtrls) Step 1
            If imIBRowNo = ilRow Then
                If (ilBox = IBUNITSINDEX) Or (ilBox = IBTAMOUNTINDEX) Then
                    gPaintArea pbcPostItem, tmIBCtrls(ilBox).fBoxX, tmIBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmIBCtrls(ilBox).fBoxW - 15, tmIBCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                Else
                    gPaintArea pbcPostItem, tmIBCtrls(ilBox).fBoxX, tmIBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmIBCtrls(ilBox).fBoxW - 15, tmIBCtrls(ilBox).fBoxH - 15, WHITE
                End If
            End If
            pbcPostItem.CurrentX = tmIBCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcPostItem.CurrentY = tmIBCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            pbcPostItem.Print smIBShow(ilBox, ilRow)
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcYN_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub pbcYN_KeyPress(KeyAscii As Integer)
    If imIBBoxNo = IBACINDEX Then
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            If imIBSave(4, imIBRowNo) <> 0 Then
                imIBChg = True
            End If
            imIBSave(4, imIBRowNo) = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imIBSave(4, imIBRowNo) <> 1 Then
                imIBChg = True
            End If
            imIBSave(4, imIBRowNo) = 1
            pbcYN_Paint
        End If
    ElseIf imIBBoxNo = IBSCINDEX Then
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            If imIBSave(5, imIBRowNo) <> 0 Then
                imIBChg = True
            End If
            imIBSave(5, imIBRowNo) = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imIBSave(5, imIBRowNo) <> 1 Then
                imIBChg = True
            End If
            imIBSave(5, imIBRowNo) = 1
            pbcYN_Paint
        End If
    ElseIf imIBBoxNo = IBTXINDEX Then
        If (KeyAscii = Asc("Y")) Or (KeyAscii = Asc("y")) Then
            If imIBSave(6, imIBRowNo) <> 0 Then
                imIBChg = True
            End If
            imIBSave(6, imIBRowNo) = 0
            pbcYN_Paint
        ElseIf KeyAscii = Asc("N") Or (KeyAscii = Asc("n")) Then
            If imIBSave(6, imIBRowNo) <> 1 Then
                imIBChg = True
            End If
            imIBSave(6, imIBRowNo) = 1
            pbcYN_Paint
        End If
    End If
    If KeyAscii = Asc(" ") Then
        If imIBBoxNo = IBACINDEX Then
            If imIBSave(4, imIBRowNo) = 0 Then
                imIBChg = True
                imIBSave(4, imIBRowNo) = 1
            Else
                imIBChg = True
                imIBSave(4, imIBRowNo) = 0
            End If
        ElseIf imIBBoxNo = IBSCINDEX Then
            If imIBSave(5, imIBRowNo) = 0 Then
                imIBChg = True
                imIBSave(5, imIBRowNo) = 1
            Else
                imIBChg = True
                imIBSave(5, imIBRowNo) = 0
            End If
        ElseIf imIBBoxNo = IBTXINDEX Then
            If imIBSave(6, imIBRowNo) = 0 Then
                imIBChg = True
                imIBSave(6, imIBRowNo) = 1
            Else
                imIBChg = True
                imIBSave(6, imIBRowNo) = 0
            End If
        End If
        pbcYN_Paint
    End If
End Sub
Private Sub pbcYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imIBBoxNo = IBACINDEX Then
        If imIBSave(4, imIBRowNo) = 0 Then
            imIBChg = True
            imIBSave(4, imIBRowNo) = 1
        Else
            imIBChg = True
            imIBSave(4, imIBRowNo) = 0
        End If
    ElseIf imIBBoxNo = IBSCINDEX Then
        If imIBSave(5, imIBRowNo) = 0 Then
            imIBChg = True
            imIBSave(5, imIBRowNo) = 1
        Else
            imIBChg = True
            imIBSave(5, imIBRowNo) = 0
        End If
    ElseIf imIBBoxNo = IBTXINDEX Then
        If imIBSave(6, imIBRowNo) = 0 Then
            imIBChg = True
            imIBSave(6, imIBRowNo) = 1
        Else
            imIBChg = True
            imIBSave(6, imIBRowNo) = 0
        End If
    End If
    pbcYN_Paint
End Sub
Private Sub pbcYN_Paint()
    pbcYN.Cls
    pbcYN.CurrentX = fgBoxInsetX
    pbcYN.CurrentY = 0 'fgBoxInsetY
    If imIBBoxNo = IBACINDEX Then
        If imIBSave(4, imIBRowNo) = 0 Then
            pbcYN.Print "Yes"
        ElseIf imIBSave(4, imIBRowNo) = 1 Then
            pbcYN.Print "No"
        End If
    ElseIf imIBBoxNo = IBSCINDEX Then
        If imIBSave(5, imIBRowNo) = 0 Then
            pbcYN.Print "Yes"
        ElseIf imIBSave(5, imIBRowNo) = 1 Then
            pbcYN.Print "No"
        End If
    ElseIf imIBBoxNo = IBTXINDEX Then
        If imIBSave(6, imIBRowNo) = 0 Then
            pbcYN.Print "Yes"
        ElseIf imIBSave(6, imIBRowNo) = 1 Then
            pbcYN.Print "No"
        End If
    End If
End Sub
Private Sub plcPostItem_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcPostItem_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub plcPostItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcScreen_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcScreen_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub plcScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub plcSelect_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSelect_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub plcSelect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmcDrag.Enabled Then
        imDragType = -1
        tmcDrag.Enabled = False
    End If
End Sub
Private Sub tmcClick_Timer()
    tmcClick.Enabled = False
    Select Case imIBBoxNo
        Case IBITEMTYPEINDEX
            imLbcArrowSetting = False
            gProcessLbcClick lbcBItem, edcDropDown, imChgMode, imLbcArrowSetting
    End Select
End Sub
Private Sub tmcDrag_Timer()
    Dim ilCompRow As Integer
    Dim ilMaxRow As Integer
    Dim ilRow As Integer
    Select Case imDragType
        Case 0  'Start Drag
            imDragType = -1
            tmcDrag.Enabled = False
            ilCompRow = vbcPostItem.LargeChange + 1
            If UBound(smIBSave, 2) > ilCompRow Then
                ilMaxRow = ilCompRow
            Else
                ilMaxRow = UBound(smIBSave, 2)
            End If
            For ilRow = 1 To ilMaxRow Step 1
                If (fmDragY >= ((ilRow - 1) * (fgBoxGridH + 15) + tmIBCtrls(IBVEHICLEINDEX).fBoxY)) And (fmDragY <= ((ilRow - 1) * (fgBoxGridH + 15) + tmIBCtrls(IBVEHICLEINDEX).fBoxY + tmIBCtrls(IBVEHICLEINDEX).fBoxH)) Then
                    'Only allow deletion of new- might want to be able to delete unbilled
                    If (smIBSave(7, ilRow + vbcPostItem.Value - 1) = "B") Then
                        Beep
                        Exit Sub
                    End If
                    mIBSetShow imIBBoxNo
                    imIBBoxNo = -1
                    imIBRowNo = -1
                    imIBRowNo = ilRow + vbcPostItem.Value - 1
                    lacIBFrame.Move 0, tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15) - 30
                    lacIBFrame.Visible = True
                    pbcArrow.Move pbcArrow.Left, plcPostItem.Top + tmIBCtrls(IBVEHICLEINDEX).fBoxY + (imIBRowNo - vbcPostItem.Value) * (fgBoxGridH + 15) + 45
                    pbcArrow.Visible = True
                    imcTrash.Enabled = True
                    lacIBFrame.Drag vbBeginDrag
                    lacIBFrame.DragIcon = IconTraf!imcIconDrag.DragIcon
                    Exit Sub
                End If
            Next ilRow
        Case 1  'scroll up
        Case 2  'Scroll down
    End Select
End Sub
Private Sub vbcPostItem_Change()
    If imSettingValue Then
        pbcPostItem.Cls
        pbcPostItem_Paint
        imSettingValue = False
    Else
        mIBSetShow imIBBoxNo
        pbcPostItem.Cls
        pbcPostItem_Paint
        mIBEnableBox imIBBoxNo
    End If
End Sub
Private Sub vbcPostItem_DragDrop(Source As Control, X As Single, Y As Single)
    lacIBFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub vbcPostItem_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Invoice Item Posting"
End Sub
