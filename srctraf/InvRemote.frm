VERSION 5.00
Begin VB.Form InvRemote 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5715
   ClientLeft      =   630
   ClientTop       =   1680
   ClientWidth     =   9435
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   9435
   Begin VB.PictureBox plcInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
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
      Height          =   630
      Left            =   315
      ScaleHeight     =   600
      ScaleWidth      =   8550
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5025
      Visible         =   0   'False
      Width           =   8580
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Invoice #: xxxxxx  Vehicle Name:"
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
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   26
         Top             =   45
         Width           =   8400
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Caption         =   "Contract #: xxxxxx  Check #  Transaction: Date xx/xx/xx  Type xx  Action  xx"
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
         Height          =   225
         Index           =   1
         Left            =   60
         TabIndex        =   25
         Top             =   315
         Width           =   8355
      End
   End
   Begin VB.PictureBox pbcKey 
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
      Height          =   1410
      Left            =   135
      Picture         =   "InvRemote.frx":0000
      ScaleHeight     =   1380
      ScaleWidth      =   5115
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   615
      Visible         =   0   'False
      Width           =   5145
   End
   Begin VB.Timer tmcClick 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   7380
      Top             =   5190
   End
   Begin VB.ListBox lbcVehicle 
      Appearance      =   0  'Flat
      Height          =   240
      ItemData        =   "InvRemote.frx":17042
      Left            =   2400
      List            =   "InvRemote.frx":17049
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   1470
   End
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
      Picture         =   "InvRemote.frx":17059
      ScaleHeight     =   180
      ScaleWidth      =   90
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   90
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
      TabIndex        =   21
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
      TabIndex        =   19
      Top             =   5175
      Width           =   1050
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
      Picture         =   "InvRemote.frx":17363
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4560
      Width           =   75
   End
   Begin VB.TextBox edcNoSpots 
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
      MaxLength       =   4
      TabIndex        =   10
      Top             =   1155
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox edcGross 
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
      TabIndex        =   9
      Top             =   975
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmcImport 
      Appearance      =   0  'Flat
      Caption         =   "&Import"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6075
      TabIndex        =   20
      Top             =   5175
      Width           =   945
   End
   Begin VB.PictureBox pbcTab 
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
      TabIndex        =   12
      Top             =   4200
      Width           =   60
   End
   Begin VB.PictureBox pbcSTab 
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
   Begin VB.VScrollBar vbcInvRemote 
      Height          =   4275
      LargeChange     =   19
      Left            =   9015
      TabIndex        =   13
      Top             =   630
      Width           =   270
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
      Left            =   2895
      ScaleHeight     =   345
      ScaleWidth      =   6345
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   75
      Width           =   6405
      Begin VB.ComboBox cbcInvDate 
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
         Left            =   30
         TabIndex        =   2
         Top             =   15
         Width           =   2565
      End
      Begin VB.ComboBox cbcMarket 
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
         Left            =   2895
         TabIndex        =   3
         Top             =   15
         Width           =   3435
      End
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   3435
      TabIndex        =   18
      Top             =   5175
      Width           =   1050
   End
   Begin VB.CommandButton cmcDone 
      Appearance      =   0  'Flat
      Caption         =   "&Done"
      Height          =   285
      Left            =   2055
      TabIndex        =   17
      Top             =   5175
      Width           =   1050
   End
   Begin VB.PictureBox pbcInvRemote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4290
      Left            =   165
      Picture         =   "InvRemote.frx":1745D
      ScaleHeight     =   4290
      ScaleWidth      =   8835
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   615
      Width           =   8835
      Begin VB.Label lacFrame 
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
         TabIndex        =   16
         Top             =   390
         Visible         =   0   'False
         Width           =   8730
      End
   End
   Begin VB.PictureBox plcInvRemote 
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
      Height          =   4440
      Left            =   120
      ScaleHeight     =   4380
      ScaleWidth      =   9150
      TabIndex        =   5
      Top             =   570
      Width           =   9210
   End
   Begin VB.Label lacScreen 
      Caption         =   "Remote Invoice Posting"
      Height          =   225
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   2130
   End
   Begin VB.Image imcInsert 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   615
      Picture         =   "InvRemote.frx":93047
      Top             =   5085
      Width           =   480
   End
   Begin VB.Image imcTrash 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8850
      Picture         =   "InvRemote.frx":93911
      Top             =   5100
      Width           =   480
   End
   Begin VB.Image imcKey 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   195
      Picture         =   "InvRemote.frx":93C1B
      Top             =   360
      Width           =   480
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   150
      Top             =   5070
      Visible         =   0   'False
      Width           =   360
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
      TabIndex        =   15
      Top             =   4275
      Width           =   210
   End
End
Attribute VB_Name = "InvRemote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim hmMsg As Integer   'From file hanle
Dim hmFrom As Integer

Dim tmMktCode() As SORTCODE
Dim imMktCode As Integer
Dim tmMktVefCode() As Integer

Dim tmInvVehicle() As SORTCODE
Dim smInvVehicleTag As String

'Library calendar
Dim hmLcf As Integer        'Library calendar file handle
Dim tmLcf As LCF
Dim imLcfRecLen As Integer
Dim tmLcfSrchKey As LCFKEY0

Dim tmContract() As SORTCODE
Dim smContractTag As String
Dim tmChfAdvtExt() As CHFADVTEXT

Dim tmVehicle() As SORTCODE
Dim smVehicleTag As String
Dim imUpdateAllowed As Integer
Dim imFirstActivate As Integer

'Billing Items
Dim tmCtrls(0 To 11) As FIELDAREA
Dim imLBCtrls as Integer

Dim hmChf As Integer
Dim tmChf As CHF
Dim tmChfSrchKey As LONGKEY0
Dim tmChfSrchKey1 As CHFKEY1
Dim imChfRecLen As Integer

Dim hmClf As Integer
Dim tmClf As CLF
Dim tmClfSrchKey As CLFKEY0
Dim imClfRecLen As Integer

Dim hmCff As Integer            'Contract line flight file handle
Dim tmCffSrchKey As CFFKEY0            'CFF record image
Dim imCffRecLen As Integer        'CFF record length
Dim tmCff As CFF

Dim hmSbf As Integer        'Special billing
Dim tmSbf As SBF    'SBF record image of billing items
Dim tmSbfSrchKey1 As LONGKEY0            'SBF record image
Dim tmSbfSrchKey2 As SBFKEY2    'SBF key record image
Dim imSbfRecLen As Integer        'SBF record length
Dim lmSbfDel() As Long      'SbfCode of records to be deleted

Dim hmVsf As Integer            'Virtual Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim tmVsfSrchKey As LONGKEY0            'VSF record image
Dim imVsfRecLen As Integer        'VSF record length


Dim imMarketIndex As Integer
Dim imInvDateIndex As Integer

Dim imButton As Integer 'Value 1= Left button; 2=Right button; 4=Middle button
Dim imButtonRow As Integer
Dim imIgnoreRightMove As Integer

Dim imChgMode As Integer    'Change mode status (so change not entered when in change)
Dim imBSMode As Integer     'Backspace flag
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imBoxNo As Integer
Dim imRowNo As Integer
Dim smSave() As String      'Values saved (1=C or T; 2=PctTrade; 3=Ordered No Spots; 4=Ordered Gross; 5=Aired No Spots; 6=Aired Gross; 7= Bonus No Spots; 8=Billed)
Dim imSave() As Integer     'Values saved (1=AdfCode; 2=VefCode; 3=lbcVehicle Index; 4=Valid vehicle)
Dim lmSave() As Long        'Values saved(1=ChfCode; 2=SbfCode)
Dim smShow() As String * 40     'Show values (1=Contract;2=Cash/Trade;3=Advertiser;4=Vehicle;5=Order Spots;6=Ordered Gross;7=Aired Spots;8=Aired Gross;9=Bonus)
Dim smInfo() As String * 12     'Import Info for Right Mouse (1=Source[I=Import;F=Sbf;C=Contract;T=Total;S=Insert]; 2=Export Date; 3=Import Date;
                            '4=Invoice Date; 5=Combine ID; 6=Ref Inv #; 7=Tax1; 8=Tax2; 9=Ordered Spots, 10=Ordered Gross; 11=Comm Pct)
Dim imChg As Integer
Dim smStartStd As String    'Starting date for standard billing
Dim smEndStd As String      'Ending date for standard billing
Dim smStartCal As String    'Starting date for standard billing
Dim smEndCal As String      'Ending date for standard billing
Dim lmStartStd As Long    'Starting date for standard billing
Dim lmEndStd As Long      'Ending date for standard billing
Dim lmStartCal As Long    'Starting date for standard billing
Dim lmEndCal As Long      'Ending date for standard billing
Dim smNowDate As String

Dim smFieldValues(1 To 17) As String
Dim imComboBoxIndex As Integer
Dim imSettingValue As Integer
Dim imLbcMouseDown As Integer  'True=List box mouse down
Dim imDoubleClickName As Integer    'Name from a list was selected by double clicking
Dim imLbcArrowSetting As Integer
Dim imDirProcess As Integer
Dim imTabDirection As Integer   '0=left to right (Tab); -1=right to left (Shift tab)
Dim imTotColor As Integer   '0=Red, 1=Green (used for package only)
Dim lmCntrStartDate As Long
Dim lmCntrEndDate As Long
Dim imTaxDefined As Integer
Dim imBypassFocus As Integer

Dim tmRec As LPOPREC

'Const VEHICLEINDEX = 1   'Vehicle control/field
'Const IBDATEINDEX = 2       'Bill date control/index
'Const IBDESCRIPTINDEX = 3   'Description control/field
'Const IBITEMTYPEINDEX = 4   'Item billing type control/field
'Const IBACINDEX = 5         'Agency commission type control/field
'Const IBSCINDEX = 6         'Salesperson commission control/field
'Const IBTXINDEX = 7         'Taxable control/field
'Const AGROSSINDEX = 8     'Amount per item control/field
'Const IBUNITSINDEX = 9      'Units control/field
'Const IBNOITEMSINDEX = 10    'Number of items control/field
'Const IBTAMOUNTINDEX = 11    'Total amount control/field
'Const IBBILLINDEX = 12      'Billed flag control/field

Const CONTRACTINDEX = 1
Const CASHTRADEINDEX = 2
Const ADVTINDEX = 3
Const VEHICLEINDEX = 4
Const ONOSPOTSINDEX = 5
Const OGROSSINDEX = 6
Const ANOSPOTSINDEX = 7
Const AGROSSINDEX = 8
Const ABONUSINDEX = 9
Const DNOSPOTSINDEX = 10
Const DGROSSINDEX = 11

'*******************************************************
'*                                                     *
'*      Procedure Name:mShowInfo                       *
'*                                                     *
'*             Created:5/13/94       By:D. Hannifan    *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Show Sdf information           *
'*                                                     *
'*******************************************************
Sub mShowInfo()
    If (imButtonRow >= LBound(smSave, 2)) And (imButtonRow <= UBound(smSave, 2)) Then
        If Trim$(smInfo(1, imButtonRow)) = "T" Then
            plcInfo.Visible = False
        Else
            lacInfo(0).Caption = "Export Date: " & Trim$(smInfo(2, imButtonRow)) & " Import Date: " & Trim$(smInfo(3, imButtonRow)) & " Invoice Date: " & Trim$(smInfo(4, imButtonRow)) & " Ref Invoice #: " & Trim$(smInfo(6, imButtonRow))
            lacInfo(1).Caption = "Remote Order Totals: Spots " & Trim$(smInfo(9, imButtonRow)) & " Gross " & Trim$(smInfo(10, imButtonRow)) & " Agency Comm % " & Trim$(smInfo(11, imButtonRow))
            plcInfo.Visible = True
        End If
    Else
        plcInfo.Visible = False
    End If
End Sub


'Help messages
Private Sub cbcInvDate_Change()
    Dim slStr As String     'Text entered

    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        imChgMode = True    'Set change mode to avoid infinite loop
        slStr = cbcInvDate.Text
        If slStr <> "" Then
            gManLookAhead cbcInvDate, imBSMode, imComboBoxIndex
            If cbcInvDate.ListIndex >= 0 Then
                tmcClick.Enabled = False
                tmcClick.Interval = 2000    '2 seconds
                tmcClick.Enabled = True
            End If
        End If
        imInvDateIndex = cbcInvDate.ListIndex
        imChgMode = False
    End If
    cmcImport.Enabled = False
    Exit Sub
cbcInvDateErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault    'Default
    imTerminate = True
    Exit Sub
End Sub

Private Sub cbcInvDate_Click()
    imComboBoxIndex = cbcInvDate.ListIndex
    cbcInvDate_Change
End Sub

Private Sub cbcInvDate_GotFocus()
    Dim slSvText As String   'Save so list box can be reset
    Dim ilSvIndex As Integer
    
    If imTerminate Then
        Exit Sub
    End If
    tmcClick.Enabled = False
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
'    gSetIndexFromText cbcInvDate
    slSvText = cbcInvDate.Text
    If cbcInvDate.ListCount <= 1 Then
        cbcInvDate.ListIndex = 0
        mClearCtrlFields 'Make sure all fields cleared
        mSetCommands
'        pbcHdSTab.SetFocus
        Exit Sub
    End If
'    gShowHelpMess tmChfHelp(), CHFCNTRSELECT
    gCtrlGotFocus ActiveControl
    If (slSvText = "") Then
        cbcInvDate.ListIndex = 0
        cbcInvDate_Change
    Else
        gFindMatch slSvText, 1, cbcInvDate
        If gLastFound(cbcInvDate) > 0 Then
'            If (ilSvIndex <> gLastFound(cbcInvDate)) Or (ilSvIndex <> cbcInvDate.ListIndex) Then
            If (slSvText <> cbcInvDate.List(gLastFound(cbcInvDate))) Then
                cbcInvDate.ListIndex = gLastFound(cbcInvDate)
            End If
        Else
            cbcInvDate.ListIndex = 0
            mClearCtrlFields
            cbcInvDate_Change
        End If
    End If
    mSetCommands
End Sub

Sub cbcInvDate_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Sub cbcInvDate_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcInvDate.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cbcMarket_Change()
    Dim ilRet As Integer
    Dim slDate As String
    Dim ilRes As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim slStr As String
    
    If imChgMode = False Then 'If currently in change mode- bypass any other changes (avoid infinite loop)
        imChgMode = True
        tmcClick.Enabled = False
        slStr = Trim$(cbcMarket.Text)
        If slStr <> "" Then
            gManLookAhead cbcMarket, imBSMode, imComboBoxIndex
            If cbcMarket.ListIndex >= 0 Then
                tmcClick.Interval = 2000    '2 seconds
                tmcClick.Enabled = True
            End If
        End If
        imMarketIndex = cbcMarket.ListIndex
        imChgMode = False
    End If
    cmcImport.Enabled = False
    Exit Sub
End Sub

Private Sub cbcMarket_Click()
    imComboBoxIndex = cbcMarket.ListIndex
    cbcMarket_Change
End Sub

Private Sub cbcMarket_GotFocus()
    If imTerminate Then
        Exit Sub
    End If
    mSetShow imBoxNo
    imBoxNo = -1
    imRowNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus cbcMarket
End Sub

Private Sub cbcMarket_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub cbcMarket_KeyPress(KeyAscii As Integer)
    'Backspace character cause selected test to be deleted or
    'the first character to the lEtf of the cursor if no text selected
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If cbcMarket.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
End Sub

Private Sub cmcCancel_Click()
    mTerminate
End Sub
Private Sub cmcCancel_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub cmcCancel_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    If mSaveRecChg(True) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        mEnableBox imBoxNo
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcDone_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub cmcDone_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub cmcDropDown_Click()
    Select Case imBoxNo
        Case VEHICLEINDEX
            lbcVehicle.Visible = Not lbcVehicle.Visible
    End Select
    edcDropDown.SelStart = 0
    edcDropDown.SelLength = Len(edcDropDown.Text)
    edcDropDown.SetFocus
End Sub
Private Sub cmcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub cmcImport_Click()
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim ilRet As Integer
    Dim slMsgFile As String
    
    ilRet = MsgBox("Current Airing values will be replaced, Continue?", vbYesNo + vbQuestion, "Warning")
    If ilRet = vbNo Then
        Exit Sub
    End If
    mBuildDate
    'gObtainYearMonthDayStr smStartStd, True, slFYear, slFMonth, slFDay
    gObtainYearMonthDayStr smEndStd, True, slFYear, slFMonth, slFDay
    If imInvDateIndex >= 0 Then
        slFMonth = Left$(cbcInvDate.List(imInvDateIndex), 3)
    Else
        slFMonth = "???"
    End If
    igBrowserType = 7  'Mask
    ''sgBrowseMaskFile = "F" & Right$(slFYear, 2) & slFMonth & slFDay & "?.I??"
    'sgBrowseMaskFile = "?" & right$(slFYear, 2) & slFMonth & slFDay & "?.I??"
    sgBrowseMaskFile = slFMonth & right$(slFYear, 2) & "In?.??"
    sgBrowserTitle = "Import for " & cbcMarket.List(imMarketIndex)
    Browser.Show vbModal
    sgBrowserTitle = ""
    If igBrowserReturn = 1 Then
        Screen.MousePointer = vbHourglass
        slMsgFile = sgBrowserFile
        If InStr(slMsgFile, ":") = 0 Then
            slMsgFile = sgImportPath & slMsgFile
        End If
        ilRet = mOpenMsgFile(slMsgFile)
        If Not ilRet Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Print #hmMsg, "Import " & sgBrowserFile & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        'Remove Aired count if not previously defined and not imported
        mRemoveAirCount True
        pbcInvRemote.Cls
        ilRet = mReadImportFile(sgBrowserFile)
        If ilRet Then
            Print #hmMsg, "Import Finish Successfully"
            Close #hmMsg
        Else
            Print #hmMsg, "** Import Failed or Terminated " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            MsgBox "See " & slMsgFile & " for errors related to Rejected Records"
        End If
        'pbcInvRemote.Cls
        'Compute totals
        mRecomputeTotals
        vbcInvRemote.Min = LBound(smSave, 2)
        If UBound(smSave, 2) <= vbcInvRemote.LargeChange Then
            vbcInvRemote.Max = LBound(smSave, 2)
        Else
            vbcInvRemote.Max = UBound(smSave, 2) - vbcInvRemote.LargeChange
        End If
        If vbcInvRemote.Value = vbcInvRemote.Min Then
            pbcInvRemote_Paint
        Else
            vbcInvRemote.Value = vbcInvRemote.Min
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmcImport_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub cmcImport_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
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
        mEnableBox imBoxNo
        Exit Sub
    End If
    mEnableBox imBoxNo
    mSetCommands
End Sub
Private Sub cmcUpdate_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcGross_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcGross_KeyPress(KeyAscii As Integer)
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
    slStr = edcGross.Text
    slStr = Left$(slStr, edcGross.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcGross.SelStart - edcGross.SelLength)
    If gCompNumberStr(slStr, "9999999.99") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
End Sub
Private Sub edcDropDown_Change()
    Dim slStr As String
    Dim slDate As String
    Dim ilRet As Integer
    Select Case imBoxNo
        Case VEHICLEINDEX
            imLbcArrowSetting = True
            gMatchLookAhead edcDropDown, lbcVehicle, imBSMode, imComboBoxIndex
    End Select
    imLbcArrowSetting = False
End Sub
Private Sub edcDropDown_DblClick()
    imDoubleClickName = True    'Double click event foolowed by mouse up
End Sub
Private Sub edcDropDown_GotFocus()
    Select Case imBoxNo
        Case VEHICLEINDEX
            If lbcVehicle.ListCount = 1 Then
                lbcVehicle.ListIndex = 0
                If imTabDirection = -1 Then  'Right To Left
                    pbcSTab.SetFocus
                Else
                    pbcTab.SetFocus
                End If
                Exit Sub
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
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case imBoxNo
            Case VEHICLEINDEX
                gProcessArrowKey Shift, KeyCode, lbcVehicle, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
End Sub
Private Sub edcDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imDoubleClickName Then
        Select Case imBoxNo
        End Select
        imDoubleClickName = False
    End If
End Sub
Private Sub edcNoSpots_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub edcNoSpots_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    'Filter characters (allow only BackSpace, numbers 0 thru 9
    If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
    slStr = edcNoSpots.Text
    slStr = Left$(slStr, edcNoSpots.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcNoSpots.SelStart - edcNoSpots.SelLength)
    If gCompNumberStr(slStr, "9999") > 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If
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
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ilReSet As Integer
   
    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
        If (cbcMarket.Enabled) And (imBoxNo > 0) Then
            plcSelect.Enabled = False
            ilReSet = True
        Else
            ilReSet = False
        End If
        gFunctionKeyBranch KeyCode
        If imBoxNo > 0 Then
            mEnableBox imBoxNo
        End If
        If ilReSet Then
            plcSelect.Enabled = True
        End If
    Else
        If KeyCode = KEYINSERT Then    'Insert Row
            imcInsert_Click
        ElseIf KeyCode = KEYDELETE Then
            imcTrash_Click
        End If

    End If
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not igManUnload Then
        mSetShow imBoxNo
        imBoxNo = -1
        pbcArrow.Visible = False
        lacFrame.Visible = False
        If mSaveRecChg(True) = False Then
            If imTerminate Then
                Exit Sub
            End If
            If imBoxNo <> -1 Then
                mEnableBox imBoxNo
            End If
            Cancel = 1
            igStopCancel = True
            Exit Sub
        End If
    End If
    Me.KeyPreview = False
    igJobShowing(POSTLOGSJOB) = False
End Sub
Private Sub imcHelp_Click()
    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
    'Traffic!cdcSetup.Action = 6
End Sub

Private Sub imcInsert_Click()
    Dim ilNewRow As Integer
    Dim ilIndex As Integer
    Dim ilCol As Integer
    
    mSetShow imBoxNo
    pbcArrow.Visible = False
    lacFrame.Visible = False
    If (imRowNo >= LBound(smSave, 2)) And (imRowNo <= UBound(smSave, 2)) Then
        If Trim$(smInfo(1, imRowNo)) <> "T" Then
            ilNewRow = imRowNo + 1
            'Move rows down and duplicate current row
            For ilIndex = UBound(smSave, 2) To ilNewRow Step -1
                For ilCol = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                    smSave(ilCol, ilIndex) = smSave(ilCol, ilIndex - 1)
                Next ilCol
                For ilCol = LBound(imSave, 1) To UBound(imSave, 1) Step 1
                    imSave(ilCol, ilIndex) = imSave(ilCol, ilIndex - 1)
                Next ilCol
                For ilCol = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
                    lmSave(ilCol, ilIndex) = lmSave(ilCol, ilIndex - 1)
                Next ilCol
                For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                    smShow(ilCol, ilIndex) = smShow(ilCol, ilIndex - 1)
                Next ilCol
                For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                    smInfo(ilCol, ilIndex) = smInfo(ilCol, ilIndex - 1)
                Next ilCol
            Next ilIndex
            ReDim Preserve smSave(1 To 8, 1 To UBound(smSave, 2) + 1) As String
            ReDim Preserve imSave(1 To 4, 1 To UBound(imSave, 2) + 1) As Integer
            ReDim Preserve lmSave(1 To 2, 1 To UBound(lmSave, 2) + 1) As Long
            ReDim Preserve smShow(1 To 9, 1 To UBound(smShow, 2) + 1) As String * 40
            ReDim Preserve smInfo(1 To 11, 1 To UBound(smInfo, 2) + 1) As String * 12
            smSave(3, ilNewRow) = ""
            smSave(4, ilNewRow) = ""
            smSave(5, ilNewRow) = "0"
            smSave(6, ilNewRow) = "0.00"
            smSave(7, ilNewRow) = "0"
            smSave(8, ilNewRow) = "N"
            gSetShow pbcInvRemote, smSave(3, ilNewRow), tmCtrls(ONOSPOTSINDEX)
            smShow(ONOSPOTSINDEX, ilNewRow) = tmCtrls(ONOSPOTSINDEX).sShow
            gSetShow pbcInvRemote, smSave(4, ilNewRow), tmCtrls(OGROSSINDEX)
            smShow(OGROSSINDEX, ilNewRow) = tmCtrls(OGROSSINDEX).sShow
            gSetShow pbcInvRemote, smSave(5, ilNewRow), tmCtrls(ANOSPOTSINDEX)
            smShow(ANOSPOTSINDEX, ilNewRow) = tmCtrls(ANOSPOTSINDEX).sShow
            gSetShow pbcInvRemote, smSave(6, ilNewRow), tmCtrls(AGROSSINDEX)
            smShow(AGROSSINDEX, ilNewRow) = tmCtrls(AGROSSINDEX).sShow
            gSetShow pbcInvRemote, smSave(7, ilNewRow), tmCtrls(ABONUSINDEX)
            smShow(ABONUSINDEX, ilNewRow) = tmCtrls(ABONUSINDEX).sShow
            imSave(2, ilNewRow) = -1
            imSave(3, ilNewRow) = -1
            imSave(4, ilNewRow) = True
            lmSave(2, ilNewRow) = 0
            smInfo(1, ilNewRow) = "S"   'Insert
            smInfo(2, ilNewRow) = ""
            smInfo(3, ilNewRow) = ""
            smInfo(4, ilNewRow) = ""
            smInfo(5, ilNewRow) = "0"
            smInfo(6, ilNewRow) = "0"
            smInfo(7, ilNewRow) = "0.00"
            smInfo(8, ilNewRow) = "0.00"
            smInfo(9, ilNewRow) = ""
            smInfo(10, ilNewRow) = ""
            smInfo(11, ilNewRow) = ""
            pbcInvRemote.Cls
            vbcInvRemote.Min = LBound(smSave, 2)
            If UBound(smSave, 2) <= vbcInvRemote.LargeChange Then
                vbcInvRemote.Max = LBound(smSave, 2)
            Else
                vbcInvRemote.Max = UBound(smSave, 2) - vbcInvRemote.LargeChange
            End If
            If vbcInvRemote.Value = vbcInvRemote.Min Then
                pbcInvRemote_Paint
            Else
                vbcInvRemote.Value = vbcInvRemote.Min
            End If
            imBoxNo = VEHICLEINDEX
            imRowNo = ilNewRow
            If imRowNo <= vbcInvRemote.LargeChange + 1 Then
                vbcInvRemote.Value = vbcInvRemote.Min
            Else
                vbcInvRemote.Value = imRowNo - vbcInvRemote.LargeChange
            End If
            mEnableBox imBoxNo
        End If
    End If

End Sub

Private Sub imcKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pbcKey.Visible = True
End Sub

Private Sub imcKey_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pbcKey.Visible = False
End Sub

Sub imcTrash_Click()
    Dim ilRet As Integer
    Dim ilDelete As Integer
    Dim ilIndex As Integer
    Dim ilCol As Integer

    If (imRowNo < LBound(smSave, 2)) Or (imRowNo > UBound(smSave, 2)) Then
        Exit Sub
    End If
    If InStr(1, smShow(ADVTINDEX, imRowNo), "Total:", 1) > 0 Then
        Beep
        Exit Sub
    End If
    ilDelete = False
    If (Trim$(smInfo(1, imRowNo)) = "S") Then
        ilDelete = True
    ElseIf (Val(smSave(3, imRowNo)) = 0) And (gStrDecToLong(smSave(4, imRowNo), 2) = 0) Then
        ilRet = MsgBox("Ok to Delete Row", vbYesNo + vbQuestion, "Update")
        If ilRet = vbYes Then
            ilDelete = True
        End If
    End If
    If ilDelete Then
        mSetShow imBoxNo
        imChg = True
        lacFrame.Visible = False
        pbcArrow.Visible = False
        If lmSave(2, imRowNo) > 0 Then
            lmSbfDel(UBound(lmSbfDel)) = lmSave(2, imRowNo)
            ReDim Preserve lmSbfDel(0 To UBound(lmSbfDel) + 1) As Long
        End If
        Screen.MousePointer = vbHourglass
        'Move rows up
        For ilIndex = imRowNo To UBound(smSave, 2) - 1 Step 1
            For ilCol = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                smSave(ilCol, ilIndex) = smSave(ilCol, ilIndex + 1)
            Next ilCol
            For ilCol = LBound(imSave, 1) To UBound(imSave, 1) Step 1
                imSave(ilCol, ilIndex) = imSave(ilCol, ilIndex + 1)
            Next ilCol
            For ilCol = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
                lmSave(ilCol, ilIndex) = lmSave(ilCol, ilIndex + 1)
            Next ilCol
            For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                smShow(ilCol, ilIndex) = smShow(ilCol, ilIndex + 1)
            Next ilCol
            For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                smInfo(ilCol, ilIndex) = smInfo(ilCol, ilIndex + 1)
            Next ilCol
        Next ilIndex
        ReDim Preserve smSave(1 To 8, 1 To UBound(smSave, 2) - 1) As String
        ReDim Preserve imSave(1 To 4, 1 To UBound(imSave, 2) - 1) As Integer
        ReDim Preserve lmSave(1 To 2, 1 To UBound(lmSave, 2) - 1) As Long
        ReDim Preserve smShow(1 To 9, 1 To UBound(smShow, 2) - 1) As String * 40
        ReDim Preserve smInfo(1 To 11, 1 To UBound(smInfo, 2) - 1) As String * 12
        pbcInvRemote.Cls
        mRecomputeTotals
        vbcInvRemote.Min = LBound(smSave, 2)
        If UBound(smSave, 2) <= vbcInvRemote.LargeChange Then
            vbcInvRemote.Max = LBound(smSave, 2)
        Else
            vbcInvRemote.Max = UBound(smSave, 2) - vbcInvRemote.LargeChange
        End If
        If vbcInvRemote.Value = vbcInvRemote.Min Then
            pbcInvRemote_Paint
        Else
            vbcInvRemote.Value = vbcInvRemote.Min
        End If
        mSetCommands
        If InStr(1, smShow(ADVTINDEX, imRowNo), "Total:", 1) > 0 Then
            If imRowNo + 1 >= UBound(smSave, 2) Then
                imRowNo = -1
                cmcCancel.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            imRowNo = imRowNo + 1
        End If
        If imRowNo <= vbcInvRemote.LargeChange + 1 Then
            vbcInvRemote.Value = vbcInvRemote.Min
        Else
            vbcInvRemote.Value = imRowNo - vbcInvRemote.LargeChange
        End If
        imBoxNo = 0
        lacFrame.Move 0, tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcInvRemote.Value) * (fgBoxGridH + 15) - 30
        lacFrame.Visible = True
        pbcArrow.Move pbcArrow.Left, plcInvRemote.Top + tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcInvRemote.Value) * (fgBoxGridH + 15) + 45
        pbcArrow.Visible = True
        pbcArrow.SetFocus
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub lacScreen_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub lacScreen_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub

Private Sub lacTotals_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub lbcVehicle_Click()
    gProcessLbcClick lbcVehicle, edcDropDown, imChgMode, imLbcArrowSetting
End Sub
Private Sub lbcVehicle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mBuildDate                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Converted selected date        *
'*                                                     *
'*******************************************************
Sub mAddGrandTotalLine()
    Dim ilRowNo As Integer
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilCol As Integer

    'Add Total line
    ilRowNo = UBound(smSave, 2)
    slStr = ""
    smSave(1, ilRowNo) = ""
    smSave(2, ilRowNo) = ""
    smSave(3, ilRowNo) = ""
    smSave(4, ilRowNo) = ""
    smSave(5, ilRowNo) = ""
    smSave(6, ilRowNo) = ""
    smSave(7, ilRowNo) = ""
    smSave(8, ilRowNo) = ""
    gSetShow pbcInvRemote, slStr, tmCtrls(ONOSPOTSINDEX)
    smShow(ONOSPOTSINDEX, ilRowNo) = tmCtrls(ONOSPOTSINDEX).sShow
    gSetShow pbcInvRemote, slStr, tmCtrls(OGROSSINDEX)
    smShow(OGROSSINDEX, ilRowNo) = tmCtrls(OGROSSINDEX).sShow
    gSetShow pbcInvRemote, slStr, tmCtrls(ANOSPOTSINDEX)
    smShow(ANOSPOTSINDEX, ilRowNo) = tmCtrls(ANOSPOTSINDEX).sShow
    gSetShow pbcInvRemote, slStr, tmCtrls(AGROSSINDEX)
    smShow(AGROSSINDEX, ilRowNo) = tmCtrls(AGROSSINDEX).sShow
    imSave(2, ilRowNo) = 0
    imSave(3, ilRowNo) = -1
    For ilLoop = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        smShow(ilLoop, ilRowNo) = ""
    Next ilLoop
    slStr = "Grand Total:"
    gSetShow pbcInvRemote, slStr, tmCtrls(ADVTINDEX)
    smShow(ADVTINDEX, ilRowNo) = tmCtrls(ADVTINDEX).sShow
    For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
        smInfo(ilCol, ilRowNo) = ""
    Next ilCol
    smInfo(1, ilRowNo) = "T"
    ReDim Preserve smSave(1 To 8, 1 To ilRowNo + 1) As String
    ReDim Preserve imSave(1 To 4, 1 To ilRowNo + 1) As Integer
    ReDim Preserve lmSave(1 To 2, 1 To ilRowNo + 1) As Long
    ReDim Preserve smShow(1 To 9, 1 To ilRowNo + 1) As String * 40
    ReDim Preserve smInfo(1 To 11, 1 To ilRowNo + 1) As String * 12
    mGrandTotal
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mBuildDate                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Converted selected date        *
'*                                                     *
'*******************************************************
Sub mBuildDate()
    Dim slName As String
    Dim slMonth As String
    Dim slYear As String
    Dim ilRet As Integer
    Dim slDate As String

    smStartCal = ""
    smEndCal = ""
    lmStartCal = 0
    lmEndCal = 0
    smStartStd = ""
    smEndStd = ""
    lmStartStd = 0
    lmEndStd = 0
    If imInvDateIndex < 0 Then
        Exit Sub
    End If
    'Build Dates
    slName = cbcInvDate.List(imInvDateIndex)
    ilRet = gParseItem(slName, 1, ",", slMonth)
    ilRet = gParseItem(slName, 2, ",", slYear)
    Select Case UCase$(slMonth)
        Case "JAN"
            slDate = "1/15/" & slYear
        Case "FEB"
            slDate = "2/15/" & slYear
        Case "MAR"
            slDate = "3/15/" & slYear
        Case "APR"
            slDate = "4/15/" & slYear
        Case "MAY"
            slDate = "5/15/" & slYear
        Case "June"
            slDate = "6/15/" & slYear
        Case "July"
            slDate = "7/15/" & slYear
        Case "AUG"
            slDate = "8/15/" & slYear
        Case "SEPT"
            slDate = "9/15/" & slYear
        Case "OCT"
            slDate = "10/15/" & slYear
        Case "NOV"
            slDate = "11/15/" & slYear
        Case "DEC"
            slDate = "12/15/" & slYear
    End Select
    smStartCal = gObtainStartCal(slDate)
    smEndCal = gObtainEndCal(smStartCal)
    smStartStd = gObtainStartStd(slDate)
    smEndStd = gObtainEndStd(smStartStd)
    lmStartStd = gDateValue(smStartStd)
    lmEndStd = gDateValue(smEndStd)
    lmStartCal = gDateValue(smStartCal)
    lmEndCal = gDateValue(smEndCal)
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
    Dim ilIndex As Integer
    Dim ilCol As Integer

    lbcVehicle.ListIndex = -1
    edcGross.Text = ""
    edcNoSpots.Text = ""
    For ilLoop = imLBCtrls To UBound(tmCtrls) Step 1
        tmCtrls(ilLoop).sShow = ""
        tmCtrls(ilLoop).iChg = False
    Next ilLoop
    ReDim smSave(1 To 8, 1 To 1) As String
    ReDim imSave(1 To 4, 1 To 1) As Integer
    ReDim lmSave(1 To 2, 1 To 1) As Long
    ReDim smShow(1 To 9, 1 To 1) As String * 40
    ReDim smInfo(1 To 11, 1 To 1) As String * 12
    ReDim lmSbfDel(0 To 0) As Long
    For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
            smShow(ilCol, 1) = ""
    Next ilCol
    For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
            smInfo(ilCol, 1) = ""
    Next ilCol
    imBoxNo = -1
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcMouseDown = False
    imChgMode = False
    imBSMode = False
    imChg = False
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
    Dim llLastStd As Long
    Dim llLastCal As Long
    Dim llEarliestDate As Long
    Dim slDate As String
    Dim llDate As Long
    'Dim ilMonth As Integer
    'Dim ilYear As Integer
    Dim slName As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilFound As Integer
    Dim slStr As String

    llEarliestDate = -1
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slDate
    llLastStd = gDateValue(slDate)
    gUnpackDate tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), slDate
    llLastCal = gDateValue(slDate)
    ilRet = gPopUserVehicleBox(InvRemote, VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH + DORMANTVEH, lbcVehicle, tmInvVehicle(), smInvVehicleTag)
    'For ilLoop = 0 To UBound(tmInvVehicle) - 1 Step 1 'Traffic!lbcVehicle.ListCount - 1 To 0 Step -1
    '    slNameCode = tmInvVehicle(ilLoop).sKey    'Traffic!lbcVehicle.List(ilLoop)
    '    ilRet = gParseItem(slNameCode, 1, "\", slName)    'Get application name
    '    ilRet = gParseItem(slName, 3, "|", slName)    'Get application name
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
    '    ilVefCode = Val(slCode)
    '    tmLcfSrchKey.sType = "O"  ' On Air code
    '    tmLcfSrchKey.sStatus = "C"  ' Current
    '    tmLcfSrchKey.iVefCode = ilVefCode
    '    slDate = Format$("1/1/95", "m/d/yy")
    '    gPackDate slDate, tmLcfSrchKey.iLogDate(0), tmLcfSrchKey.iLogDate(1)
    '    tmLcfSrchKey.iSeqNo = 0
    '    ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    '    If (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = ilVefCode) And (tmLcf.sStatus = "C") And (tmLcf.sType = "O") Then
    '        gUnpackDateLong tmLcf.iLogDate(0), tmLcf.iLogDate(1), llDate
    '        If llEarliestDate = -1 Then
    '            llEarliestDate = llDate
    '        ElseIf llDate < llEarliestDate Then
    '            llEarliestDate = llDate
    '        End If
    '    End If
    'Next ilLoop
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slDate
    'If (slDate <> "") And (llEarliestDate > 0) Then
    If (slDate <> "") Then
        llDate = gDateValue(slDate)
        llEarliestDate = llDate - 90
        If gDateValue(Format$(gNow(), "m/d/yy")) > llDate Then
            llDate = gDateValue(Format$(gNow(), "m/d/yy"))
        End If
        slDate = Format$(llDate, "m/d/yy")
        slDate = gObtainStartStd(slDate)
        llDate = gDateValue(slDate) - 1
        Do While llDate > llEarliestDate
            slDate = Format$(llDate, "m/d/yy")
            slName = gMonthYearFormat(slDate)
            cbcInvDate.AddItem slName
            slDate = gObtainStartStd(slDate)
            llDate = gDateValue(slDate) - 1
        Loop
    End If
    ''Check previous month
    'slDate = Format$(gDateValue(Format$(gNow(), "m/d/yy")) - 20, "m/d/yy")
    'slName = gMonthYearFormat(slDate)
    'ilFound = False
    'For ilLoop = 0 To cbcInvDate.ListCount - 1 Step 1
    '    slStr = cbcInvDate.List(ilLoop)
    '    If StrComp(slName, slStr, 1) = 0 Then
    '        ilFound = True
    '        Exit For
    '    End If
    'Next ilLoop
    'If Not ilFound Then
    '    cbcInvDate.AddItem slName, 0
    'End If
    'slDate = Format$(gNow(), "m/d/yy")
    'slName = gMonthYearFormat(slDate)
    'ilFound = False
    'For ilLoop = 0 To cbcInvDate.ListCount - 1 Step 1
    '    slStr = cbcInvDate.List(ilLoop)
    '    If StrComp(slName, slStr, 1) = 0 Then
    '        ilFound = True
    '        Exit For
    '    End If
    'Next ilLoop
    'If Not ilFound Then
    '    cbcInvDate.AddItem slName, 0
    'End If
    'Set to next month
    slDate = Format$(llLastStd + 15, "m/d/yy")
    slName = gMonthYearFormat(slDate)
    For ilLoop = 0 To cbcInvDate.ListCount - 1 Step 1
        slStr = cbcInvDate.List(ilLoop)
        If StrComp(slName, slStr, 1) = 0 Then
            cbcInvDate.ListIndex = ilLoop
            Exit For
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer

    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    
    If (imRowNo < vbcInvRemote.Value) Or (imRowNo >= vbcInvRemote.Value + vbcInvRemote.LargeChange + 1) Then
        mSetShow ilBoxNo
        pbcArrow.Visible = False
        lacFrame.Visible = False
        Exit Sub
    End If
    lacFrame.Move 0, tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcInvRemote.Value) * (fgBoxGridH + 15) - 30
    lacFrame.Visible = True
    pbcArrow.Move pbcArrow.Left, plcInvRemote.Top + tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcInvRemote.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHICLEINDEX 'Vehicle
            'mVehPop
'            gShowHelpMess tmSbfHelp(), SBFVEHICLE
            lbcVehicle.Height = gListBoxHeight(lbcVehicle.ListCount, 8)
            edcDropDown.Width = tmCtrls(VEHICLEINDEX).fBoxW
            If tgSpf.iVehLen <= 40 Then
                edcDropDown.MaxLength = tgSpf.iVehLen
            Else
                edcDropDown.MaxLength = 20
            End If
            gMoveTableCtrl pbcInvRemote, edcDropDown, tmCtrls(VEHICLEINDEX).fBoxX, tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcInvRemote.Value) * (fgBoxGridH + 15)
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top
            imChgMode = True
            If imSave(3, imRowNo) >= 0 Then
                lbcVehicle.ListIndex = imSave(3, imRowNo)
                imComboBoxIndex = lbcVehicle.ListIndex
                edcDropDown.Text = lbcVehicle.List(imSave(3, imRowNo))
            Else
                lbcVehicle.ListIndex = 0
                imComboBoxIndex = lbcVehicle.ListIndex
                edcDropDown.Text = lbcVehicle.List(0)
            End If
            imChgMode = False
            If imRowNo - vbcInvRemote.Value <= vbcInvRemote.LargeChange \ 2 Then
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height
            Else
                lbcVehicle.Move edcDropDown.Left, edcDropDown.Top - lbcVehicle.Height
            End If
            edcDropDown.SelStart = 0
            edcDropDown.SelLength = Len(edcDropDown.Text)
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            edcDropDown.SetFocus
        Case ANOSPOTSINDEX
'            gShowHelpMess tmSbfHelp(), SBFITEMDESC
            edcNoSpots.Width = tmCtrls(ANOSPOTSINDEX).fBoxW
            gMoveTableCtrl pbcInvRemote, edcNoSpots, tmCtrls(ANOSPOTSINDEX).fBoxX, tmCtrls(ANOSPOTSINDEX).fBoxY + (imRowNo - vbcInvRemote.Value) * (fgBoxGridH + 15)
            edcNoSpots.Text = smSave(5, imRowNo)
            edcNoSpots.Visible = True  'Set visibility
            edcNoSpots.SetFocus
        Case AGROSSINDEX
'            gShowHelpMess tmSbfHelp(), SBFITEMCOST
            edcGross.Width = tmCtrls(AGROSSINDEX).fBoxW
            gMoveTableCtrl pbcInvRemote, edcGross, tmCtrls(AGROSSINDEX).fBoxX, tmCtrls(AGROSSINDEX).fBoxY + (imRowNo - vbcInvRemote.Value) * (fgBoxGridH + 15)
            edcGross.Text = smSave(6, imRowNo)
            edcGross.Visible = True  'Set visibility
            edcGross.SetFocus
        Case ABONUSINDEX
            edcNoSpots.Width = tmCtrls(ABONUSINDEX).fBoxW
            gMoveTableCtrl pbcInvRemote, edcNoSpots, tmCtrls(ABONUSINDEX).fBoxX, tmCtrls(ABONUSINDEX).fBoxY + (imRowNo - vbcInvRemote.Value) * (fgBoxGridH + 15)
            edcNoSpots.Text = smSave(7, imRowNo)
            edcNoSpots.Visible = True  'Set visibility
            edcNoSpots.SetFocus
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetCntr                        *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Contracts for market       *
'*                                                     *
'*******************************************************
Sub mGetCntr()
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStatus As String
    Dim slCntrType As String
    Dim ilHOType As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVeh As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim llChfCode As Long
    Dim ilVefCode As Integer
    Dim ilFound As Integer
    Dim ilLoop1 As Integer
    Dim ilChf As Integer
    Dim slName As String
    Dim slNameSort As String
    Dim slSort As String
    Dim slKey As String
    Dim slStr As String
    Dim ilAdf As Integer

    'Moved to tmcClick
    'mBuildDate
    'ReDim tmMktVefCode(0 To 0) As Integer
    'If imMarketIndex >= 0 Then
    '    slNameCode = tmMktCode(imMarketIndex).sKey    'lbcMster.List(ilLoop)
    '    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    '    imMktCode = Val(slCode)
    '    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
    '        If tgMVef(ilLoop).iMnfVehGp3Mkt = imMktCode Then
    '            tmMktVefCode(UBound(tmMktVefCode)) = tgMVef(ilLoop).iCode
    '            ReDim Preserve tmMktVefCode(0 To UBound(tmMktVefCode) + 1) As Integer
    '        End If
    '    Next ilLoop
    'Else
    '    imMktCode = -1
    'End If
    If ((lmStartStd > 0) Or (lmStartCal > 0)) And (imMktCode > 0) Then
        If (lmStartStd > 0) And (lmStartCal > 0) Then
            If lmStartStd < lmStartCal Then
                slStartDate = Format$(lmStartStd, "m/d/yy")
            Else
                slStartDate = Format$(lmStartCal, "m/d/yy")
            End If
        ElseIf lmStartStd > 0 Then
            slStartDate = Format$(lmStartStd, "m/d/yy")
        Else
            slStartDate = Format$(lmStartCal, "m/d/yy")
        End If
        If (lmEndStd > 0) And (lmEndCal > 0) Then
            If lmEndStd > lmEndCal Then
                slEndDate = Format$(lmEndStd, "m/d/yy")
            Else
                slEndDate = Format$(lmEndCal, "m/d/yy")
            End If
        ElseIf lmEndStd > 0 Then
            slEndDate = Format$(lmEndStd, "m/d/yy")
        Else
            slEndDate = Format$(lmEndCal, "m/d/yy")
        End If
        slStatus = "HO"
        slCntrType = ""
        ilHOType = 1
        sgCntrForDateStamp = ""
        ilRet = gObtainCntrForDate(Invoice, slStartDate, slEndDate, slStatus, slCntrType, ilHOType, tmChfAdvtExt())
    Else
        ReDim tmChfAdvtExt(1 To 1) As CHFADVTEXT
    End If
    ReDim tmContract(0 To 0) As SORTCODE
    For ilChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
        For ilVeh = 0 To UBound(tmMktVefCode) - 1 Step 1
            ilVefCode = tmMktVefCode(ilVeh)
            ilFound = False
            If tmChfAdvtExt(ilChf).lVefCode > 0 Then
                If tmChfAdvtExt(ilChf).lVefCode = ilVefCode Then
                    ilFound = True
                End If
            ElseIf tmChfAdvtExt(ilChf).lVefCode < 0 Then
                tmVsfSrchKey.lCode = -tmChfAdvtExt(ilChf).lVefCode
                ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                    If tmVsf.iFSCode(ilLoop) > 0 Then
                        If tmVsf.iFSCode(ilLoop) = ilVefCode Then
                            ilFound = True
                            Exit For
                        End If
                    End If
                Next ilLoop
            End If
            If ilFound Then
                slStr = "Missing"
                For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                    If tgCommAdf(ilAdf).iCode = tmChfAdvtExt(ilChf).iAdfCode Then
                        slStr = Trim$(tgCommAdf(ilAdf).sName)
                        Exit For
                    End If
                Next ilAdf
                Do While Len(slStr) < 30
                    slStr = slStr & " "
                Loop
                slKey = slStr
                'Contract number
                slStr = Trim$(Str$(tmChfAdvtExt(ilChf).lCntrNo))
                Do While Len(slStr) < 8
                    slStr = "0" & slStr
                Loop
                slKey = slKey & slStr
                tmContract(UBound(tmContract)).sKey = slKey & "\" & Trim$(Str$(tmChfAdvtExt(ilChf).lCode))
                ReDim Preserve tmContract(0 To UBound(tmContract) + 1) As SORTCODE
                Exit For
            End If
        Next ilVeh
    Next ilChf
    If UBound(tmContract) > 0 Then
        'ArraySortTyp tgSort(), tgSort(0), ilUpper, 0, Len(tgSort(0)), 0, -9, 0
        ArraySortTyp fnAV(tmContract(), 0), UBound(tmContract), 0, LenB(tmContract(0)), 0, LenB(tmContract(0).sKey), 0
    End If
    'Build images
    For ilChf = 0 To UBound(tmContract) - 1 Step 1
        slNameCode = tmContract(ilChf).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        llChfCode = Val(slCode)
        mObtainCntrInfo llChfCode
    Next ilChf
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String

    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmCtrls)) Then
        Exit Sub
    End If
    If (imRowNo < LBound(smSave, 2)) Or (imRowNo > UBound(smSave, 2)) Then
        Exit Sub
    End If
    
    lacFrame.Visible = False
    pbcArrow.Visible = False
    Select Case ilBoxNo 'Branch on box type (control)
        Case VEHICLEINDEX 'Vehicle
            lbcVehicle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            slStr = edcDropDown.Text
            gSetShow pbcInvRemote, slStr, tmCtrls(ilBoxNo)
            smShow(VEHICLEINDEX, imRowNo) = tmCtrls(ilBoxNo).sShow
            If imSave(3, imRowNo) <> lbcVehicle.ListIndex Then
                imSave(3, imRowNo) = lbcVehicle.ListIndex
                slNameCode = tmVehicle(imSave(3, imRowNo)).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                imSave(2, imRowNo) = Val(slCode)
                imChg = True
            End If
        Case ANOSPOTSINDEX
            edcNoSpots.Visible = False
            slStr = edcNoSpots.Text
            gSetShow pbcInvRemote, slStr, tmCtrls(ilBoxNo)
            smShow(ANOSPOTSINDEX, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(5, imRowNo) <> edcNoSpots.Text Then
                imChg = True
                smSave(5, imRowNo) = edcNoSpots.Text
                mRecomputeTotals
                pbcInvRemote.Cls
                pbcInvRemote_Paint
            End If
        Case AGROSSINDEX
            edcGross.Visible = False
            slStr = edcGross.Text
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
            gSetShow pbcInvRemote, slStr, tmCtrls(ilBoxNo)
            smShow(AGROSSINDEX, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(6, imRowNo) <> edcGross.Text Then
                imChg = True
                smSave(6, imRowNo) = edcGross.Text
                slStr = smSave(6, imRowNo)
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                gSetShow pbcInvRemote, slStr, tmCtrls(AGROSSINDEX)
                smShow(AGROSSINDEX, imRowNo) = tmCtrls(AGROSSINDEX).sShow
                mRecomputeTotals
                pbcInvRemote.Cls
                pbcInvRemote_Paint
            End If
        Case ABONUSINDEX
            edcNoSpots.Visible = False
            slStr = edcNoSpots.Text
            gSetShow pbcInvRemote, slStr, tmCtrls(ilBoxNo)
            smShow(ABONUSINDEX, imRowNo) = tmCtrls(ilBoxNo).sShow
            If smSave(7, imRowNo) <> edcNoSpots.Text Then
                imChg = True
                smSave(7, imRowNo) = edcNoSpots.Text
                mRecomputeTotals
                pbcInvRemote.Cls
                pbcInvRemote_Paint
            End If
    End Select
    mSetCommands
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFields                     *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestFields() As Integer
'
'   iRet = mTestFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRowNo As Integer

    For ilRowNo = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
        If mTestSaveFields(ilRowNo) = NO Then
            imRowNo = ilRowNo
            imSettingValue = True
            If imRowNo <= vbcInvRemote.LargeChange + 1 Then
                vbcInvRemote.Value = vbcInvRemote.Min
            Else
                vbcInvRemote.Value = imRowNo - vbcInvRemote.LargeChange
            End If
            mTestFields = NO
            Exit Function
        End If
    Next ilRowNo
    mTestFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestSaveFields                 *
'*                                                     *
'*             Created:9/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test fields for mandatory and   *
'*                     blanks                          *
'*                                                     *
'*******************************************************
Private Function mTestSaveFields(ilRowNo As Integer) As Integer
'
'   iRet = mTestSaveFields()
'   Where:
'       iRet (O)- True if all mandatory fields answered
'
'
    Dim ilRes As Integer    'Result of MsgBox
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    If (imSave(3, ilRowNo) < 0) And (Trim$(smInfo(1, ilRowNo)) = "S") Then
        ilRes = MsgBox("Vehicle must be specified", vbOkOnly + vbExclamation, "Incomplete")
        imBoxNo = VEHICLEINDEX
        mTestSaveFields = NO
        Exit Function
    End If
    mTestSaveFields = YES
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mGrandTotal                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Compute Grand Total            *
'*                                                     *
'*******************************************************
Sub mGrandTotal()
    Dim ilLoop As Integer
    Dim ilRowNo As Integer
    Dim slOGTotalNoPerWk As String
    Dim slOGTotalRate As String
    Dim slAGTotalNoPerWk As String
    Dim slAGTotalRate As String
    Dim slABonus As String
    Dim ilCol As Integer

    slOGTotalNoPerWk = "0"
    slOGTotalRate = "0"
    slAGTotalNoPerWk = "0"
    slAGTotalRate = "0"
    slABonus = "0"
    For ilLoop = LBound(smSave, 2) To UBound(smSave, 2) - 2 Step 1
        If Trim$(smInfo(1, ilLoop)) = "T" Then
            slOGTotalNoPerWk = gAddStr(slOGTotalNoPerWk, smSave(3, ilLoop))
            If InStr(RTrim$(smSave(4, ilLoop)), ".") > 0 Then
                slOGTotalRate = gAddStr(slOGTotalRate, smSave(4, ilLoop))
            End If
            slAGTotalNoPerWk = gAddStr(slAGTotalNoPerWk, smSave(5, ilLoop))
            If InStr(RTrim$(smSave(6, ilLoop)), ".") > 0 Then
                slAGTotalRate = gAddStr(slAGTotalRate, smSave(6, ilLoop))
            End If
            slABonus = gAddStr(slABonus, smSave(7, ilLoop))
        End If
    Next ilLoop
    ilRowNo = UBound(smSave, 2) - 1
    'Set save values so that difference will be set
    smSave(3, ilRowNo) = slOGTotalNoPerWk
    smSave(4, ilRowNo) = slOGTotalRate
    smSave(5, ilRowNo) = slAGTotalNoPerWk
    smSave(6, ilRowNo) = slAGTotalRate
    smSave(7, ilRowNo) = slABonus
    For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
        If ilCol <> ADVTINDEX Then
            smShow(ilCol, ilRowNo) = ""
        Else
            smShow(ilCol, ilRowNo) = "Grand Total:"
        End If
    Next ilCol
    gSetShow pbcInvRemote, slOGTotalNoPerWk, tmCtrls(ONOSPOTSINDEX)
    smShow(ONOSPOTSINDEX, ilRowNo) = tmCtrls(ONOSPOTSINDEX).sShow
    gSetShow pbcInvRemote, slOGTotalRate, tmCtrls(OGROSSINDEX)
    smShow(OGROSSINDEX, ilRowNo) = tmCtrls(OGROSSINDEX).sShow
    gSetShow pbcInvRemote, slAGTotalNoPerWk, tmCtrls(ANOSPOTSINDEX)
    smShow(ANOSPOTSINDEX, ilRowNo) = tmCtrls(ANOSPOTSINDEX).sShow
    gSetShow pbcInvRemote, slAGTotalRate, tmCtrls(AGROSSINDEX)
    smShow(AGROSSINDEX, ilRowNo) = tmCtrls(AGROSSINDEX).sShow
    gSetShow pbcInvRemote, slABonus, tmCtrls(ABONUSINDEX)
    smShow(ABONUSINDEX, ilRowNo) = tmCtrls(ABONUSINDEX).sShow
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
    Dim ilLoop As Integer
    Dim ilCol As Integer
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim slDate As String
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim ilNoPds As Integer
    Dim slStr1 As String
    Dim slStr2 As String
    'Dim tlSbf As SBF    'Only used to get size of SBF
    ReDim smSave(1 To 8, 1 To 1) As String
    ReDim imSave(1 To 4, 1 To 1) As Integer
    ReDim lmSave(1 To 2, 1 To 1) As Long
    ReDim smShow(1 To 9, 1 To 1) As String * 40
    ReDim smInfo(1 To 11, 1 To 1) As String * 12
    For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
            smShow(ilCol, 1) = ""
    Next ilCol
    For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
            smInfo(ilCol, 1) = ""
    Next ilCol
    ReDim lmSbfDel(0 To 0) As Long
    Screen.MousePointer = vbHourglass
    imLBCtrls = 1
    igJobShowing(POSTLOGSJOB) = True
    imFirstActivate = True
    imcKey.Picture = IconTraf!imcKey.Picture
    imcInsert.Picture = IconTraf!imcInsert.Picture
    imcTrash.Picture = IconTraf!imcTrashClosed.Picture
    pbcArrow.Picture = IconTraf!imcArrow.Picture
    pbcArrow.Width = 90
    pbcArrow.Height = 165
    vbcInvRemote.Min = LBound(smShow, 2)
    vbcInvRemote.Max = LBound(smShow, 2)
    vbcInvRemote.Value = vbcInvRemote.Min
    'gPDNToStr tgSpf.sBTax(0), 2, slStr1
    'gPDNToStr tgSpf.sBTax(1), 2, slStr2
    'If (Val(slStr1) = 0) And (Val(slStr2) = 0) Then
    If (tgSpf.iBTax(0) = 0) Or (tgSpf.iBTax(1) = 0) Then
        imTaxDefined = False
    Else
        imTaxDefined = True
    End If
    imMarketIndex = -1
    imInvDateIndex = -1
    sgCntrForDateStamp = ""
    imIgnoreRightMove = False
    imTerminate = False
    imFirstActivate = True
    imBypassFocus = False
    imBoxNo = -1 'Initialize current Box to N/A
    imRowNo = -1
    imChg = False
    imDirProcess = -1
    imTabDirection = 0  'Left to right movement
    imLbcArrowSetting = False
    imDoubleClickName = False
    imLbcMouseDown = False
    imSettingValue = False
    imChgMode = False
    imBSMode = False
    mMarketPop
    If imTerminate Then
        Exit Sub
    End If
    If imTerminate Then
        Exit Sub
    End If
    hmChf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", InvRemote
    On Error GoTo 0
    imChfRecLen = Len(tmChf) 'btrRecordLength(hmChf)    'Get Chf size
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", InvRemote
    On Error GoTo 0
    imClfRecLen = Len(tmClf) 'btrRecordLength(hmClf)    'Get Clf size
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", InvRemote
    On Error GoTo 0
    imCffRecLen = Len(tmCff) 'btrRecordLength(hmCff)    'Get Cff size
    hmSbf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sbf.Btr)", InvRemote
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf) 'btrRecordLength(hmSbf)    'Get Sbf size
    hmLcf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Lcf.Btr)", InvRemote
    On Error GoTo 0
    imLcfRecLen = Len(tmLcf) 'btrRecordLength(hmLcf)    'Get Lcf size
    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Vsf.Btr)", InvRemote
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)
    mDatePop
    ilRet = gObtainVef()
    ilRet = gObtainAdvt()
    'mVehPop
    InvRemote.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    'gCenterModalForm InvRemote
    gCenterStdAlone InvRemote
    'Traffic!plcHelp.Caption = ""
    mInitBox
    lacTotals.Visible = False
    Screen.MousePointer = vbDefault
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
    Dim ilLoop As Integer
    flTextHeight = pbcInvRemote.TextHeight("1") - 35
    'Position panel and picture areas with panel
    plcInvRemote.Move 120, 555, pbcInvRemote.Width + vbcInvRemote.Width + fgPanelAdj, pbcInvRemote.Height + fgPanelAdj
    pbcInvRemote.Move plcInvRemote.Left + fgBevelX, plcInvRemote.Top + fgBevelY
    vbcInvRemote.Move pbcInvRemote.Left + pbcInvRemote.Width, pbcInvRemote.Top + 15
    pbcArrow.Move plcInvRemote.Left - pbcArrow.Width - 15    'Vehicle
    plcInfo.Move plcInvRemote.Left + (plcInvRemote.Left + plcInvRemote.Width - plcInfo.Width) / 2, plcInvRemote.Top + plcInvRemote.Height - 60
    pbcKey.Move plcInvRemote.Left, plcInvRemote.Top
    'Contract
    gSetCtrl tmCtrls(CONTRACTINDEX), 30, 375, 915, fgBoxGridH
    'Cash/Trade
    gSetCtrl tmCtrls(CASHTRADEINDEX), 960, tmCtrls(CONTRACTINDEX).fBoxY, 240, fgBoxGridH
    'Advertiser
    gSetCtrl tmCtrls(ADVTINDEX), 1215, tmCtrls(CONTRACTINDEX).fBoxY, 1455, fgBoxGridH
    'Vehicle
    gSetCtrl tmCtrls(VEHICLEINDEX), 2685, tmCtrls(CONTRACTINDEX).fBoxY, 1455, fgBoxGridH
    'Ordered Number of Spots
    gSetCtrl tmCtrls(ONOSPOTSINDEX), 4155, tmCtrls(CONTRACTINDEX).fBoxY, 495, fgBoxGridH
    'Ordered Gross
    gSetCtrl tmCtrls(OGROSSINDEX), 4665, tmCtrls(CONTRACTINDEX).fBoxY, 825, fgBoxGridH
    'Aired Number of Spots
    gSetCtrl tmCtrls(ANOSPOTSINDEX), 5505, tmCtrls(CONTRACTINDEX).fBoxY, 495, fgBoxGridH
    'Aired Gross
    gSetCtrl tmCtrls(AGROSSINDEX), 6015, tmCtrls(CONTRACTINDEX).fBoxY, 825, fgBoxGridH
    'Aired Bonus Number of Spots
    gSetCtrl tmCtrls(ABONUSINDEX), 6855, tmCtrls(CONTRACTINDEX).fBoxY, 495, fgBoxGridH
    'Difference Number of Spots
    gSetCtrl tmCtrls(DNOSPOTSINDEX), 7365, tmCtrls(VEHICLEINDEX).fBoxY, 540, fgBoxGridH
    'Difference Gross
    gSetCtrl tmCtrls(DGROSSINDEX), 7920, tmCtrls(CONTRACTINDEX).fBoxY, 900, fgBoxGridH
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitNew                        *
'*                                                     *
'*             Created:9/06/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize Item billing        *
'*                                                     *
'*******************************************************
Private Sub mInitNew(ilRowNo As Integer)
    Dim ilLoop As Integer

    smSave(1, ilRowNo) = smSave(1, ilRowNo + 1)
    smSave(2, ilRowNo) = "0"
    smSave(3, ilRowNo) = ""
    smSave(4, ilRowNo) = ""
    smSave(5, ilRowNo) = ""
    smSave(6, ilRowNo) = ""
    smSave(7, ilRowNo) = ""
    smSave(8, ilRowNo) = "N"
    imSave(1, ilRowNo) = imSave(1, ilRowNo + 1)
    imSave(2, ilRowNo) = 0      'Vehicle
    imSave(3, ilRowNo) = -1   'Vehicle
    lmSave(1, ilRowNo) = lmSave(1, ilRowNo + 1)
    lmSave(2, ilRowNo) = 0
    imSave(6, ilRowNo) = -1
    For ilLoop = CONTRACTINDEX To ADVTINDEX Step 1
        smShow(ilLoop, ilRowNo) = smShow(ilLoop, ilRowNo + 1)
    Next ilLoop
    For ilLoop = VEHICLEINDEX To ABONUSINDEX Step 1
        smShow(ilLoop, ilRowNo) = ""
    Next ilLoop
    For ilLoop = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
            smInfo(ilLoop, ilRowNo) = ""
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMarketPop                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate Market combobox       *
'*                      if requireds                   *
'*                                                     *
'*******************************************************
Sub mMarketPop()
'
'   mMarketPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim slNameCode As String
    Dim ilSortCode As Integer
    Dim ilLoop As Integer
    Dim llLen As Long
    Dim slStr As String

    cbcMarket.Clear
    ilSortCode = 0
    ReDim tmMktCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    ilRet = gObtainMnfForType("H3", slStr, tgMkMnf())
    For ilLoop = LBound(tgMkMnf) To UBound(tgMkMnf) - 1 Step 1
        If Trim$(tgMkMnf(ilLoop).sRPU) = "Y" Then
            slName = Trim$(tgMkMnf(ilLoop).sName)
            slName = slName & "\" & Trim$(Str$(tgMkMnf(ilLoop).iCode))
            tmMktCode(ilSortCode).sKey = slName
            If ilSortCode >= UBound(tmMktCode) Then
                ReDim Preserve tmMktCode(0 To UBound(tmMktCode) + 100) As SORTCODE
            End If
            ilSortCode = ilSortCode + 1
        End If
    Next ilLoop
    ReDim Preserve tmMktCode(0 To ilSortCode) As SORTCODE
    If UBound(tmMktCode) - 1 > 0 Then
        ArraySortTyp fnAV(tmMktCode(), 0), UBound(tmMktCode), 0, LenB(tmMktCode(0)), 0, LenB(tmMktCode(0).sKey), 0
    End If
    llLen = 0
    For ilLoop = 0 To UBound(tmMktCode) - 1 Step 1
        slNameCode = tmMktCode(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet = CP_MSG_NONE Then
            slName = Trim$(slName)
            If Not gOkAddStrToListBox(slName, llLen, True) Then
                Exit For
            End If
            cbcMarket.AddItem slName  'Add ID to list box
        End If
    Next ilLoop
    Exit Sub
mMarketPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMerge                          *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Merge SBF record into Save     *
'*                      images                         *
'*                                                     *
'*******************************************************
Function mMerge(slSource As String) As Integer
    Dim slStr As String
    Dim ilRet As Integer    'Return status
    Dim ilLoop As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilIndex As Integer
    Dim ilNewRow As Integer
    Dim ilPass As Integer
    Dim ilFound As Integer
    Dim ilAdf As Integer
    Dim ilVef As Integer

    ilFound = False
    For ilRow = LBound(smSave, 2) To UBound(smSave, 2) - 2 Step 1
        If (tmSbf.lChfCode = lmSave(1, ilRow)) And (tmSbf.iAirVefCode = imSave(2, ilRow)) And (tmSbf.sCashTrade = smSave(1, ilRow)) Then
            ilFound = True
            If slSource = "F" Then
                lmSave(2, ilRow) = tmSbf.lCode
            End If
            If Trim$(smInfo(1, ilRow)) = "C" Then
                smSave(5, ilRow) = Trim$(Str$(tmSbf.iAirNoSpots))
                smSave(6, ilRow) = gLongToStrDec(tmSbf.lGross, 2)
                smSave(7, ilRow) = Trim$(Str$(tmSbf.iBonusNoSpots))
            Else
                smSave(5, ilRow) = gAddStr(smSave(5, ilRow), Trim$(Str$(tmSbf.iAirNoSpots)))
                smSave(6, ilRow) = gAddStr(smSave(6, ilRow), gLongToStrDec(tmSbf.lGross, 2))
                smSave(7, ilRow) = gAddStr(smSave(7, ilRow), Trim$(Str$(tmSbf.iBonusNoSpots)))
            End If
            smSave(8, ilRow) = tmSbf.sBilled
            gSetShow pbcInvRemote, smSave(5, ilRow), tmCtrls(ANOSPOTSINDEX)
            smShow(ANOSPOTSINDEX, ilRow) = tmCtrls(ANOSPOTSINDEX).sShow
             gSetShow pbcInvRemote, smSave(6, ilRow), tmCtrls(AGROSSINDEX)
            smShow(AGROSSINDEX, ilRow) = tmCtrls(AGROSSINDEX).sShow
            gSetShow pbcInvRemote, smSave(7, ilRow), tmCtrls(ABONUSINDEX)
            smShow(ABONUSINDEX, ilRow) = tmCtrls(ABONUSINDEX).sShow
            'If tmSbf.lCode = 0 Then
            '    smInfo(1, ilRow) = "I"
            'Else
            '    smInfo(1, ilRow) = "F"
            'End If
            smInfo(1, ilRow) = slSource
            gUnpackDate tmSbf.iExportDate(0), tmSbf.iExportDate(1), smInfo(2, ilRow)
            gUnpackDate tmSbf.iImportDate(0), tmSbf.iImportDate(1), smInfo(3, ilRow)
            gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), smInfo(4, ilRow)
            smInfo(5, ilRow) = Trim$(Str$(tmSbf.iCombineID))
            smInfo(6, ilRow) = Trim$(Str$(tmSbf.lRefInvNo))
            smInfo(7, ilRow) = gLongToStrDec(tmSbf.lTax1, 2)
            smInfo(8, ilRow) = gLongToStrDec(tmSbf.lTax2, 2)
            smInfo(9, ilRow) = Trim$(Str$(tmSbf.iNoItems))
            smInfo(10, ilRow) = gLongToStrDec(tmSbf.lOGross, 2)
            smInfo(11, ilRow) = gIntToStrDec(tmSbf.iCommPct, 2)
            Exit For
        End If
    Next ilRow
    If Not ilFound Then
        'Add in above total record for contract
        ilFound = False
        For ilPass = 0 To 2 Step 1
            For ilRow = LBound(smSave, 2) To UBound(smSave, 2) - 2 Step 1
                If ((tmSbf.lChfCode = lmSave(1, ilRow)) And (tmSbf.sCashTrade = smSave(1, ilRow)) And (ilPass = 0)) Or ((tmSbf.lChfCode = lmSave(1, ilRow)) And (ilPass = 1)) Or ((ilRow = UBound(smSave, 2) - 2) And (ilPass = 2)) Then
                    ilFound = True
                    'Continue search until total record
                    For ilLoop = ilRow + 1 To UBound(smSave, 2) - 1 Step 1
                        If InStr(1, smShow(ADVTINDEX, ilLoop), "Total:", 1) > 0 Then
                            If ilPass = 0 Then
                                ilNewRow = ilLoop
                            Else
                                ilNewRow = ilLoop + 1
                                If InStr(1, smShow(ADVTINDEX, ilLoop), "Grand Total:", 1) > 0 Then
                                    ilNewRow = ilNewRow - 1
                                End If
                            End If
                            'Move all records from and including ilLoop dowm one
                            For ilIndex = UBound(smSave, 2) To ilNewRow Step -1
                                For ilCol = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                                    smSave(ilCol, ilIndex) = smSave(ilCol, ilIndex - 1)
                                Next ilCol
                                For ilCol = LBound(imSave, 1) To UBound(imSave, 1) Step 1
                                    imSave(ilCol, ilIndex) = imSave(ilCol, ilIndex - 1)
                                Next ilCol
                                For ilCol = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
                                    lmSave(ilCol, ilIndex) = lmSave(ilCol, ilIndex - 1)
                                Next ilCol
                                For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                                    smShow(ilCol, ilIndex) = smShow(ilCol, ilIndex - 1)
                                Next ilCol
                                For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                                    smInfo(ilCol, ilIndex) = smInfo(ilCol, ilIndex - 1)
                                Next ilCol
                            Next ilIndex
                            'Add row
                            If ilPass = 0 Then
                                ReDim Preserve smSave(1 To 8, 1 To UBound(smSave, 2) + 1) As String
                                ReDim Preserve imSave(1 To 4, 1 To UBound(imSave, 2) + 1) As Integer
                                ReDim Preserve lmSave(1 To 2, 1 To UBound(lmSave, 2) + 1) As Long
                                ReDim Preserve smShow(1 To 9, 1 To UBound(smShow, 2) + 1) As String * 40
                                ReDim Preserve smInfo(1 To 11, 1 To UBound(smInfo, 2) + 1) As String * 12
                            Else
                                ReDim Preserve smSave(1 To 8, 1 To UBound(smSave, 2) + 2) As String
                                ReDim Preserve imSave(1 To 4, 1 To UBound(imSave, 2) + 2) As Integer
                                ReDim Preserve lmSave(1 To 2, 1 To UBound(lmSave, 2) + 2) As Long
                                ReDim Preserve smShow(1 To 9, 1 To UBound(smShow, 2) + 2) As String * 40
                                ReDim Preserve smInfo(1 To 11, 1 To UBound(smInfo, 2) + 2) As String * 12
                            End If
                            'Set values into new at ilLoop
                            smSave(1, ilNewRow) = tmSbf.sCashTrade
                            If ilPass = 0 Then
                                smSave(2, ilNewRow) = smSave(2, ilLoop - 1)
                            End If
                            smSave(3, ilNewRow) = ""   'Trim$(Str$(tmSbf.iNoItems))
                            smSave(4, ilNewRow) = ""
                            smSave(5, ilNewRow) = Trim$(Str$(tmSbf.iAirNoSpots))
                            smSave(6, ilNewRow) = gLongToStrDec(tmSbf.lGross, 2)
                            smSave(7, ilNewRow) = Trim$(Str$(tmSbf.iBonusNoSpots))
                            smSave(8, ilNewRow) = tmSbf.sBilled
                            tmChfSrchKey.lCode = tmSbf.lChfCode
                            ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                imSave(1, ilNewRow) = tmChf.iAdfCode
                                slStr = Trim$(Str$(tmChf.lCntrNo))
                                gSetShow pbcInvRemote, slStr, tmCtrls(CONTRACTINDEX)
                                smShow(CONTRACTINDEX, ilNewRow) = tmCtrls(CONTRACTINDEX).sShow
                                If ilPass <> 0 Then
                                    smSave(2, ilNewRow) = gIntToStrDec(tmChf.iPctTrade, 0)
                                End If
                            Else
                                imSave(1, ilNewRow) = -1
                                slStr = "Missing:" & Trim$(Str$(tmSbf.lChfCode))
                                gSetShow pbcInvRemote, slStr, tmCtrls(CONTRACTINDEX)
                                smShow(CONTRACTINDEX, ilNewRow) = tmCtrls(CONTRACTINDEX).sShow
                                If ilPass <> 0 Then
                                    smSave(2, ilNewRow) = "0"
                                End If
                            End If
                            imSave(2, ilNewRow) = tmSbf.iAirVefCode
                            imSave(3, ilNewRow) = 0
                            lmSave(1, ilNewRow) = tmSbf.lChfCode
                            'lmSave(2, ilNewRow) = tmSbf.lCode
                            If slSource = "F" Then
                                lmSave(2, ilNewRow) = tmSbf.lCode
                            Else
                                lmSave(2, ilNewRow) = 0
                            End If
                            slStr = smSave(1, ilNewRow)
                            gSetShow pbcInvRemote, slStr, tmCtrls(CASHTRADEINDEX)
                            smShow(CASHTRADEINDEX, ilNewRow) = tmCtrls(CASHTRADEINDEX).sShow
                            'Advertiser
                            slStr = "Missing"
                            For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                                If tgCommAdf(ilAdf).iCode = imSave(1, ilNewRow) Then
                                    slStr = Trim$(tgCommAdf(ilAdf).sName)
                                    Exit For
                                End If
                            Next ilAdf
                            gSetShow pbcInvRemote, slStr, tmCtrls(ADVTINDEX)
                            smShow(ADVTINDEX, ilNewRow) = tmCtrls(ADVTINDEX).sShow
                            'Vehicle
                            slStr = "Missing"
                            For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                If tgMVef(ilVef).iCode = imSave(2, ilNewRow) Then
                                    slStr = Trim$(tgMVef(ilVef).sName)
                                    Exit For
                                End If
                            Next ilVef
                            gSetShow pbcInvRemote, slStr, tmCtrls(VEHICLEINDEX)
                            smShow(VEHICLEINDEX, ilNewRow) = tmCtrls(VEHICLEINDEX).sShow
                            gSetShow pbcInvRemote, smSave(3, ilNewRow), tmCtrls(ONOSPOTSINDEX)
                            smShow(ONOSPOTSINDEX, ilNewRow) = tmCtrls(ONOSPOTSINDEX).sShow
                            gSetShow pbcInvRemote, smSave(4, ilNewRow), tmCtrls(OGROSSINDEX)
                            smShow(OGROSSINDEX, ilNewRow) = tmCtrls(OGROSSINDEX).sShow
                            gSetShow pbcInvRemote, smSave(5, ilNewRow), tmCtrls(ANOSPOTSINDEX)
                            smShow(ANOSPOTSINDEX, ilNewRow) = tmCtrls(ANOSPOTSINDEX).sShow
                            gSetShow pbcInvRemote, smSave(6, ilNewRow), tmCtrls(AGROSSINDEX)
                            smShow(AGROSSINDEX, ilNewRow) = tmCtrls(AGROSSINDEX).sShow
                            gSetShow pbcInvRemote, smSave(7, ilNewRow), tmCtrls(ABONUSINDEX)
                            smShow(ABONUSINDEX, ilNewRow) = tmCtrls(ABONUSINDEX).sShow
                            imSave(4, ilNewRow) = False
                            For ilVef = 0 To UBound(tmMktVefCode) - 1 Step 1
                                If imSave(2, ilNewRow) = tmMktVefCode(ilVef) Then
                                    imSave(4, ilNewRow) = True
                                    Exit For
                                End If
                            Next ilVef
                            If ilPass <> 0 Then
                                For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                                    smShow(ilCol, ilNewRow + 1) = ""
                                Next ilCol
                                For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                                    smInfo(ilCol, ilNewRow + 1) = ""
                                Next ilCol
                                smSave(1, ilNewRow + 1) = ""
                                smSave(2, ilNewRow + 1) = ""
                                smSave(3, ilNewRow + 1) = ""
                                smSave(4, ilNewRow + 1) = ""
                                smSave(5, ilNewRow + 1) = ""
                                smSave(6, ilNewRow + 1) = ""
                                smSave(7, ilNewRow + 1) = ""
                                smSave(8, ilNewRow + 1) = ""
                                gSetShow pbcInvRemote, smSave(3, ilNewRow + 1), tmCtrls(ONOSPOTSINDEX)
                                smShow(ONOSPOTSINDEX, ilNewRow + 1) = tmCtrls(ONOSPOTSINDEX).sShow
                                gSetShow pbcInvRemote, smSave(4, ilNewRow + 1), tmCtrls(OGROSSINDEX)
                                smShow(OGROSSINDEX, ilNewRow + 1) = tmCtrls(OGROSSINDEX).sShow
                                gSetShow pbcInvRemote, smSave(5, ilNewRow + 1), tmCtrls(ANOSPOTSINDEX)
                                smShow(ANOSPOTSINDEX, ilNewRow + 1) = tmCtrls(ANOSPOTSINDEX).sShow
                                gSetShow pbcInvRemote, smSave(6, ilNewRow + 1), tmCtrls(AGROSSINDEX)
                                smShow(AGROSSINDEX, ilNewRow + 1) = tmCtrls(AGROSSINDEX).sShow
                                gSetShow pbcInvRemote, smSave(7, ilNewRow + 1), tmCtrls(ABONUSINDEX)
                                smShow(ABONUSINDEX, ilNewRow + 1) = tmCtrls(ABONUSINDEX).sShow
                                imSave(2, ilNewRow + 1) = 0
                                imSave(3, ilNewRow + 1) = -1
                                imSave(4, ilNewRow + 1) = True
                                If imSave(1, ilNewRow) = -1 Then
                                    slStr = "# Missing"
                                Else
                                    slStr = "Total: " & Trim$(Str$(tmChf.lCntrNo))
                                End If
                                gSetShow pbcInvRemote, slStr, tmCtrls(ADVTINDEX)
                                smShow(ADVTINDEX, ilNewRow + 1) = tmCtrls(ADVTINDEX).sShow
                                smInfo(1, ilNewRow + 1) = "T"
                            End If
                            'If tmSbf.lCode = 0 Then
                            '    smInfo(1, ilNewRow) = "I"
                            'Else
                            '    smInfo(1, ilNewRow) = "F"
                            'End If
                            smInfo(1, ilNewRow) = slSource
                            gUnpackDate tmSbf.iExportDate(0), tmSbf.iExportDate(1), smInfo(2, ilNewRow)
                            gUnpackDate tmSbf.iImportDate(0), tmSbf.iImportDate(1), smInfo(3, ilNewRow)
                            gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), smInfo(4, ilNewRow)
                            smInfo(5, ilNewRow) = Trim$(Str$(tmSbf.iCombineID))
                            smInfo(6, ilNewRow) = Trim$(Str$(tmSbf.lRefInvNo))
                            smInfo(7, ilNewRow) = gLongToStrDec(tmSbf.lTax1, 2)
                            smInfo(8, ilNewRow) = gLongToStrDec(tmSbf.lTax2, 2)
                            smInfo(9, ilNewRow) = Trim$(Str$(tmSbf.iNoItems))
                            smInfo(10, ilNewRow) = gLongToStrDec(tmSbf.lOGross, 2)
                            smInfo(11, ilNewRow) = gIntToStrDec(tmSbf.iCommPct, 2)
                            Exit For
                        End If
                    Next ilLoop
                End If
                If ilFound Then
                    Exit For
                End If
            Next ilRow
            If ilFound Then
                Exit For
            End If
        Next ilPass
    End If
    mMerge = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveCtrlToRec                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move controls values to record *
'*                                                     *
'*******************************************************
Private Sub mMoveCtrlToRec(ilRowNo As Integer)
'
'   mMoveCtrlToRec
'   Where:
'
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer

    tmSbf.lCode = lmSave(2, ilRowNo)
    tmSbf.lChfCode = lmSave(1, ilRowNo)
    If Trim$(smInfo(4, ilRowNo)) <> "" Then
        gPackDate Trim$(smInfo(4, ilRowNo)), tmSbf.iDate(0), tmSbf.iDate(1)
    Else
        gPackDate smEndStd, tmSbf.iDate(0), tmSbf.iDate(1)
    End If
    tmSbf.sTranType = "T"
    tmSbf.iBillVefCode = imSave(2, ilRowNo)
    tmSbf.iMnfItem = 0
    If Trim$(smInfo(9, ilRowNo)) <> "" Then
        tmSbf.iNoItems = Val(Trim$(smInfo(9, ilRowNo)))
    Else
        tmSbf.iNoItems = 0  'Val(smSave(3, ilRowNo))
    End If
    tmSbf.lGross = gStrDecToLong(smSave(6, ilRowNo), 2)
    tmSbf.sUnitName = ""
    tmSbf.sDescr = ""
    tmSbf.sAgyComm = "Y"
    tmSbf.sSlsComm = "Y"
    tmSbf.sSlsTax = "Y"
    tmSbf.sBilled = smSave(8, ilRowNo)
    tmSbf.sCashTrade = smSave(1, ilRowNo)
    tmSbf.iAirVefCode = imSave(2, ilRowNo)
    tmSbf.iAirNoSpots = Val(smSave(5, ilRowNo))
    tmSbf.iBonusNoSpots = Val(smSave(7, ilRowNo))
    tmSbf.lTax1 = gStrDecToLong(Trim$(smInfo(7, ilRowNo)), 2)
    tmSbf.lTax2 = gStrDecToLong(Trim$(smInfo(8, ilRowNo)), 2)
    If Trim$(smInfo(3, ilRowNo)) <> "" Then
        gPackDate Trim$(smInfo(3, ilRowNo)), tmSbf.iImportDate(0), tmSbf.iImportDate(1)
    Else
        gPackDate smNowDate, tmSbf.iImportDate(0), tmSbf.iImportDate(1)
    End If
    If Trim$(smInfo(2, ilRowNo)) <> "" Then
        gPackDate Trim$(smInfo(2, ilRowNo)), tmSbf.iExportDate(0), tmSbf.iExportDate(1)
    Else
        gPackDate smNowDate, tmSbf.iExportDate(0), tmSbf.iExportDate(1)
    End If
    tmSbf.lRefInvNo = Val(Trim$(smInfo(6, ilRowNo)))
    tmSbf.iCombineID = Val(Trim$(smInfo(5, ilRowNo)))
    tmSbf.lOGross = gStrDecToLong(Trim$(smInfo(10, ilRowNo)), 2)
    tmSbf.iCommPct = gStrDecToInt(Trim$(smInfo(11, ilRowNo)), 2)
    Exit Sub
mMoveCtrlToRecErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Private Sub mMoveRecToCtrl()
'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilIndex As Integer
    Dim slStr As String
    Dim slRecCode As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slDate As String
    'For ilLoop = 0 To UBound(tmSbfList) - 1 Step 1
    'Next ilLoop
    Exit Sub
mMoveRecToCtrlErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mProcFlight                     *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Sub mObtainCntrInfo(llChfCode As Long)
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim ilPass As Integer
    Dim ilCTSplit As Integer
    Dim ilAdf As Integer
    Dim ilClf As Integer
    Dim ilVef As Integer
    Dim ilInsertLine As Integer
    Dim slSFlightDate As String
    Dim slEFlightDate As String
    Dim ilIncludeFlight As Integer
    Dim slTotalNoPerWk As String
    Dim slTotalRate As String
    Dim slCTotalNoPerWk As String
    Dim slCTotalRate As String
    Dim slPctTrade As String
    Dim ilAddTo As Integer
    Dim ilLoop As Integer
    Dim ilCff As Integer
    Dim ilStartRowNo As Integer
    Dim ilCol As Integer

    ilRet = gObtainCntr(hmChf, hmClf, hmCff, llChfCode, False, tgChfInv, tgClfInv(), tgCffInv())
    If ilRet Then
        For ilPass = 0 To 1 Step 1
            ilStartRowNo = UBound(smSave, 2)
            ilInsertLine = True
            ilRowNo = UBound(smSave, 2)
            'Contract Number
            lmSave(1, ilRowNo) = tgChfInv.lCode
            lmSave(2, ilRowNo) = 0
            slStr = Trim$(Str$(tgChfInv.lCntrNo))
            gSetShow pbcInvRemote, slStr, tmCtrls(CONTRACTINDEX)
            smShow(CONTRACTINDEX, ilRowNo) = tmCtrls(CONTRACTINDEX).sShow
            'Cash/Trade flag
            slPctTrade = gIntToStrDec(tgChfInv.iPctTrade, 0)
            If (ilPass = 0) And (tgChfInv.iPctTrade <> 100) Then
                slStr = "C"
            ElseIf (ilPass = 1) And (tgChfInv.iPctTrade <> 0) Then
                slStr = "T"
            Else
                ilInsertLine = False
            End If
            If ilInsertLine Then
                smSave(1, ilRowNo) = slStr
                smSave(2, ilRowNo) = slPctTrade
                gSetShow pbcInvRemote, slStr, tmCtrls(CASHTRADEINDEX)
                smShow(CASHTRADEINDEX, ilRowNo) = tmCtrls(CASHTRADEINDEX).sShow
                'Advertiser
                slStr = "Missing"
                imSave(1, ilRowNo) = tgChfInv.iAdfCode
                For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                    If tgCommAdf(ilAdf).iCode = tgChfInv.iAdfCode Then
                        slStr = Trim$(tgCommAdf(ilAdf).sName)
                        Exit For
                    End If
                Next ilAdf
                gSetShow pbcInvRemote, slStr, tmCtrls(ADVTINDEX)
                smShow(ADVTINDEX, ilRowNo) = tmCtrls(ADVTINDEX).sShow
                For ilClf = LBound(tgClfInv) To UBound(tgClfInv) - 1 Step 1
                    ilRowNo = UBound(smSave, 2)
                    If (tgClfInv(ilClf).ClfRec.sType = "S") Or (tgClfInv(ilClf).ClfRec.sType = "H") Then
                        'Vehicle
                        slStr = "Missing"
                        imSave(2, ilRowNo) = tgClfInv(ilClf).ClfRec.iVefCode
                        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            If tgMVef(ilVef).iCode = tgClfInv(ilClf).ClfRec.iVefCode Then
                                slStr = Trim$(tgMVef(ilVef).sName)
                                Exit For
                            End If
                        Next ilVef
                        gSetShow pbcInvRemote, slStr, tmCtrls(VEHICLEINDEX)
                        smShow(VEHICLEINDEX, ilRowNo) = tmCtrls(VEHICLEINDEX).sShow
                        imSave(4, ilRowNo) = False
                        For ilVef = 0 To UBound(tmMktVefCode) - 1 Step 1
                            If imSave(2, ilRowNo) = tmMktVefCode(ilVef) Then
                                imSave(4, ilRowNo) = True
                                Exit For
                            End If
                        Next ilVef
                        slTotalNoPerWk = "0"
                        slTotalRate = "0"
                        ilAddTo = False
                        For ilLoop = ilStartRowNo To ilRowNo - 1 Step 1
                            If imSave(2, ilLoop) = imSave(2, ilRowNo) Then
                                ilAddTo = True
                                ilRowNo = ilLoop
                                slTotalNoPerWk = smSave(3, ilRowNo)
                                slTotalRate = smSave(4, ilRowNo)
                                Exit For
                            End If
                        Next ilLoop
                        ilCff = tgClfInv(ilClf).iFirstCff
                        Do While ilCff <> -1
                            gUnpackDate tgCffInv(ilCff).CffRec.iStartDate(0), tgCffInv(ilCff).CffRec.iStartDate(1), slSFlightDate
                            gUnpackDate tgCffInv(ilCff).CffRec.iEndDate(0), tgCffInv(ilCff).CffRec.iEndDate(1), slEFlightDate
                            ilIncludeFlight = True
                            If tgChfInv.sBillCycle = "C" Then
                                If (gDateValue(slSFlightDate) > lmEndCal) Or (gDateValue(slEFlightDate) < lmStartCal) Then
                                    ilIncludeFlight = False
                                End If
                            Else
                                If (gDateValue(slSFlightDate) > lmEndStd) Or (gDateValue(slEFlightDate) < lmStartStd) Then
                                    ilIncludeFlight = False
                                End If
                            End If
                            'Test if CBS
                            If gDateValue(slEFlightDate) < gDateValue(slSFlightDate) Then
                                ilIncludeFlight = False
                            End If
                            If ilIncludeFlight Then
                                mProcFlight ilCff, slSFlightDate, slEFlightDate, ilPass, slPctTrade, slTotalNoPerWk, slTotalRate
                                If slTotalNoPerWk <> "0" Then
                                    smSave(3, ilRowNo) = slTotalNoPerWk
                                    smSave(4, ilRowNo) = slTotalRate
                                    smSave(5, ilRowNo) = slTotalNoPerWk
                                    smSave(6, ilRowNo) = slTotalRate
                                    smSave(7, ilRowNo) = ""
                                    smSave(8, ilRowNo) = ""
                                    gSetShow pbcInvRemote, slTotalNoPerWk, tmCtrls(ONOSPOTSINDEX)
                                    smShow(ONOSPOTSINDEX, ilRowNo) = tmCtrls(ONOSPOTSINDEX).sShow
                                    gSetShow pbcInvRemote, slTotalRate, tmCtrls(OGROSSINDEX)
                                    smShow(OGROSSINDEX, ilRowNo) = tmCtrls(OGROSSINDEX).sShow
                                    gSetShow pbcInvRemote, slTotalNoPerWk, tmCtrls(ANOSPOTSINDEX)
                                    smShow(ANOSPOTSINDEX, ilRowNo) = tmCtrls(ANOSPOTSINDEX).sShow
                                    gSetShow pbcInvRemote, slTotalRate, tmCtrls(AGROSSINDEX)
                                    smShow(AGROSSINDEX, ilRowNo) = tmCtrls(AGROSSINDEX).sShow
                                    If Not ilAddTo Then
                                        smShow(ABONUSINDEX, ilRowNo) = ""
                                        'smShow(DNOSPOTSINDEX, ilRowNo) = ""
                                        'smShow(DGROSSINDEX, ilRowNo) = ""
                                        For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                                            smInfo(ilCol, ilRowNo) = ""
                                        Next ilCol
                                        smInfo(1, ilRowNo) = "C"
                                        ReDim Preserve smSave(1 To 8, 1 To ilRowNo + 1) As String
                                        ReDim Preserve imSave(1 To 4, 1 To ilRowNo + 1) As Integer
                                        ReDim Preserve lmSave(1 To 2, 1 To ilRowNo + 1) As Long
                                        ReDim Preserve smShow(1 To 9, 1 To ilRowNo + 1) As String * 40
                                        ReDim Preserve smInfo(1 To 11, 1 To ilRowNo + 1) As String * 12
                                        imSave(1, ilRowNo + 1) = imSave(1, ilRowNo)
                                        lmSave(1, ilRowNo + 1) = lmSave(1, ilRowNo)
                                        lmSave(2, ilRowNo + 1) = lmSave(2, ilRowNo)
                                        smSave(1, ilRowNo + 1) = smSave(1, ilRowNo)
                                        smSave(2, ilRowNo + 1) = smSave(2, ilRowNo)
                                        smShow(CONTRACTINDEX, ilRowNo + 1) = smShow(CONTRACTINDEX, ilRowNo)
                                        smShow(CASHTRADEINDEX, ilRowNo + 1) = smShow(CASHTRADEINDEX, ilRowNo)
                                        smShow(ADVTINDEX, ilRowNo + 1) = smShow(ADVTINDEX, ilRowNo)
                                    End If
                                End If
                            End If
                            ilCff = tgCffInv(ilCff).iNextCff
                        Loop
                        
                    End If
                Next ilClf
                If ilStartRowNo < UBound(smSave, 2) Then
                    'Add Total line
                    slCTotalNoPerWk = "0"
                    slCTotalRate = "0"
                    For ilLoop = ilStartRowNo To UBound(smSave, 2) - 1 Step 1
                        slCTotalNoPerWk = gAddStr(slCTotalNoPerWk, smSave(3, ilLoop))
                        If InStr(RTrim$(smSave(4, ilLoop)), ".") > 0 Then
                            slCTotalRate = gAddStr(slCTotalRate, smSave(4, ilLoop))
                        End If
                    Next ilLoop
                    ilRowNo = UBound(smSave, 2)
                    For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                        smShow(ilCol, ilRowNo) = ""
                    Next ilCol
                    For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                        smInfo(ilCol, ilRowNo) = ""
                    Next ilCol
                    smSave(1, ilRowNo) = ""
                    smSave(2, ilRowNo) = ""
                    smSave(3, ilRowNo) = slCTotalNoPerWk
                    smSave(4, ilRowNo) = slCTotalRate
                    smSave(5, ilRowNo) = slCTotalNoPerWk
                    smSave(6, ilRowNo) = slCTotalRate
                    smSave(7, ilRowNo) = ""
                    smSave(8, ilRowNo) = ""
                    gSetShow pbcInvRemote, slCTotalNoPerWk, tmCtrls(ONOSPOTSINDEX)
                    smShow(ONOSPOTSINDEX, ilRowNo) = tmCtrls(ONOSPOTSINDEX).sShow
                    gSetShow pbcInvRemote, slCTotalRate, tmCtrls(OGROSSINDEX)
                    smShow(OGROSSINDEX, ilRowNo) = tmCtrls(OGROSSINDEX).sShow
                    gSetShow pbcInvRemote, slCTotalNoPerWk, tmCtrls(ANOSPOTSINDEX)
                    smShow(ANOSPOTSINDEX, ilRowNo) = tmCtrls(ANOSPOTSINDEX).sShow
                    gSetShow pbcInvRemote, slCTotalRate, tmCtrls(AGROSSINDEX)
                    smShow(AGROSSINDEX, ilRowNo) = tmCtrls(AGROSSINDEX).sShow
                    slStr = Trim$(Str$(UBound(smSave, 2) - ilStartRowNo))
                    gSetShow pbcInvRemote, slStr, tmCtrls(VEHICLEINDEX)
                    smShow(VEHICLEINDEX, ilRowNo) = tmCtrls(VEHICLEINDEX).sShow
                    imSave(2, ilRowNo) = 0
                    imSave(3, ilRowNo) = -1
                    slStr = "Total: " & Trim$(Str$(tgChfInv.lCntrNo))
                    gSetShow pbcInvRemote, slStr, tmCtrls(ADVTINDEX)
                    smShow(ADVTINDEX, ilRowNo) = tmCtrls(ADVTINDEX).sShow
                    smInfo(1, ilRowNo) = "T"
                    ReDim Preserve smSave(1 To 8, 1 To ilRowNo + 1) As String
                    ReDim Preserve imSave(1 To 4, 1 To ilRowNo + 1) As Integer
                    ReDim Preserve lmSave(1 To 2, 1 To ilRowNo + 1) As Long
                    ReDim Preserve smShow(1 To 9, 1 To ilRowNo + 1) As String * 40
                    ReDim Preserve smInfo(1 To 11, 1 To ilRowNo + 1) As String * 12
                End If
            End If
        Next ilPass
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:D. Smith       *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Function mOpenMsgFile(slMsgFile As String) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim ilPos As Integer

    ilRet = 0
    On Error GoTo mOpenMsgFileErr:
    ilPos = InStr(1, slMsgFile, ".")
    If ilPos > 0 Then
        'slToFile = Left$(slMsgFile, ilPos) & "T" & Mid$(slMsgFile, ilPos + 2)
        If InStr(1, slMsgFile, "Inv", 1) > 0 Then
            slToFile = Left$(slMsgFile, ilPos - 4) & Mid$(slMsgFile, ilPos + 1, 2) & ".Txt"
        Else
            slToFile = Left$(slMsgFile, ilPos - 4) & Mid$(slMsgFile, ilPos + 1, 2) & Mid$(slMsgFile, ilPos - 1, 1) & ".Txt"
        End If
    Else
        slToFile = slMsgFile & ".Txt"
    End If
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        Kill slToFile
        On Error GoTo 0
        ilRet = 0
        On Error GoTo mOpenMsgFileErr:
        hmMsg = FreeFile
        Open slToFile For Output As hmMsg
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        On Error GoTo mOpenMsgFileErr:
        hmMsg = FreeFile
        Open slToFile For Output As hmMsg
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile, vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    Print #hmMsg, ""
    slMsgFile = slToFile
    mOpenMsgFile = True
    Exit Function
mOpenMsgFileErr:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mProcFlight                     *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Sub mProcFlight(ilCff As Integer, slSFlightDate As String, slEFlightDate As String, ilPass As Integer, slPctTrade As String, slTotalNoPerWk As String, slTotalRate As String)
'
'   Where
'       ilCff(I)- Flight record index
'       slSFlightDate(I)- Flight Start date
'       slEFlightDate(I)- Flight End Date
'       slTotalNoPerWk(O)- Running Total number of spots per week
'       slTotalRate(I/O)- Ordered Total $'s
'
    
    Dim llDate As Long
    Dim ilDay As Integer
    Dim ilLoop As Integer
    Dim llSDate As Long
    Dim slRate As String

    'Get flight rate
    Select Case tgCffInv(ilCff).CffRec.sPriceType
        Case "T"    'True
            slRate = gLongToStrDec(tgCffInv(ilCff).CffRec.lActPrice, 2)
            If (ilPass = 0) And (Val(slPctTrade) <> 0) Then
                slRate = gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100")
            ElseIf (ilPass = 1) And (Val(slPctTrade) <> 100) Then
                slRate = gSubStr(RTrim$(slRate), gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100"))
            End If
        Case "N"    'No Charge
            slRate = "N/C"
        Case "M"    'MG Line
            slRate = "MG"
        Case "B"    'Bonus
            slRate = "Bonus"
        Case "S"    'Spinoff
            slRate = "Spinoff"
        Case "P"    'Package
            slRate = gLongToStrDec(tgCffInv(ilCff).CffRec.lActPrice, 2)
            If (ilPass = 0) And (Val(slPctTrade) <> 0) Then
                slRate = gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100")
            ElseIf (ilPass = 1) And (Val(slPctTrade) <> 100) Then
                slRate = gSubStr(RTrim$(slRate), gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100"))
            End If
        Case "R"    'Recapturable
            slRate = "Recapturable"
        Case "A"    'ADU
            slRate = "ADU"
    End Select

    If (tgCffInv(ilCff).CffRec.sDyWk <> "D") Then    'Weekly
        If tgChfInv.sBillCycle = "C" Then
            llDate = gDateValue(slSFlightDate)
            Do While llDate <= gDateValue(slEFlightDate)
                If llDate < lmStartCal Then
                    If llDate + 6 >= lmStartCal Then
                        slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(Str$(tgCffInv(ilCff).CffRec.iXSpotsWk)))
                        If InStr(RTrim$(slRate), ".") > 0 Then
                            slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(Str$(tgCffInv(ilCff).CffRec.iXSpotsWk))))
                        End If
                    End If
                ElseIf (llDate <= lmEndCal) Then
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(Str$(tgCffInv(ilCff).CffRec.iSpotsWk)))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(Str$(tgCffInv(ilCff).CffRec.iSpotsWk))))
                    End If
                Else
                    Exit Do
                End If
                llDate = gDateValue(gObtainNextMonday(Format$(llDate + 1, "m/d/yy")))
            Loop
        Else
            llDate = gDateValue(slSFlightDate)
            Do While llDate <= gDateValue(slEFlightDate)
                If (llDate >= lmStartStd) And (llDate <= lmEndStd) Then
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(Str$(tgCffInv(ilCff).CffRec.iSpotsWk)))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(Str$(tgCffInv(ilCff).CffRec.iSpotsWk))))
                    End If
                End If
                If llDate > lmEndStd Then
                    Exit Do
                End If
                llDate = gDateValue(gObtainNextMonday(Format$(llDate + 1, "m/d/yy")))
            Loop
        End If
    Else    'Daily
        If tgChfInv.sBillCycle = "C" Then
            If gDateValue(slSFlightDate) >= lmStartCal Then
                llSDate = gDateValue(slSFlightDate)
            Else
                llSDate = lmStartCal
            End If
            For llDate = llSDate To gDateValue(slEFlightDate) Step 1
                If (llDate >= lmStartCal) And (llDate <= lmEndCal) Then
                    ilDay = gWeekDayLong(llDate)
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(Str$(tgCffInv(ilCff).CffRec.iDay(ilDay))))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(Str$(tgCffInv(ilCff).CffRec.iDay(ilDay)))))
                    End If
                End If
                If llDate >= lmEndCal Then
                    Exit For
                End If
            Next llDate
        Else
            If gDateValue(slSFlightDate) >= lmStartStd Then
                llSDate = gDateValue(slSFlightDate)
            Else
                llSDate = lmStartStd
            End If
            For llDate = llSDate To gDateValue(slEFlightDate) Step 1
                If (llDate >= lmStartStd) And (llDate <= lmEndStd) Then
                    ilDay = gWeekDayLong(llDate)
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(Str$(tgCffInv(ilCff).CffRec.iDay(ilDay))))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(Str$(tgCffInv(ilCff).CffRec.iDay(ilDay)))))
                    End If
                End If
                If llDate >= lmEndStd Then
                    Exit For
                End If
            Next llDate
        End If
    End If
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mReadImportFile                 *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read File                      *
'*                                                     *
'*******************************************************
Function mReadImportFile(slFromFile As String) As Integer
    Dim ilRet As Integer
    Dim slLine As String
    Dim ilEof As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilMatch As Integer
    Dim slCode As String
    Dim ilError As Integer
    Dim ilErrorLogged As Integer
    Dim slMsg As String
    Dim ilPos As Integer

    ilRet = 0
    On Error GoTo mReadImportFileErr:
    hmFrom = FreeFile
    Open slFromFile For Input Access Read As hmFrom
    If ilRet <> 0 Then
        Close hmFrom
        mReadImportFile = False
        Exit Function
    End If
    ilErrorLogged = False
    smNowDate = Format$(gNow(), "m/d/yy")
    Do
        ilRet = 0
        On Error GoTo mReadImportFileErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        DoEvents
        If imTerminate Then
            Close hmFrom
            mReadImportFile = False
            Exit Function
        End If
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                ilEof = True
            Else
                gParseCDFields slLine, False, smFieldValues()    'Change case
                For ilLoop = LBound(smFieldValues) To UBound(smFieldValues) Step 1
                    smFieldValues(ilLoop) = Trim$(smFieldValues(ilLoop))
                Next ilLoop
                If Trim$(smFieldValues(2)) <> "" Then
                    'Test if fields 5 and 6 are still enclosed in quotes- if so remove
                    ilPos = InStr(1, smFieldValues(5), """", 1)
                    If ilPos = 1 Then
                        smFieldValues(5) = right$(smFieldValues(5), Len(smFieldValues(5)) - 1)
                        smFieldValues(5) = Left$(smFieldValues(5), Len(smFieldValues(5)) - 1)
                    End If
                    ilPos = InStr(1, smFieldValues(6), """", 1)
                    If ilPos = 1 Then
                        smFieldValues(6) = right$(smFieldValues(6), Len(smFieldValues(6)) - 1)
                        smFieldValues(6) = Left$(smFieldValues(6), Len(smFieldValues(6)) - 1)
                    End If
                    'Make SBF record, then merge
                    ilError = False
                    slMsg = "Contract # " & smFieldValues(1)
                    'tmSbf.lCode = 0
                    tmChfSrchKey1.lCntrNo = Val(smFieldValues(1))
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = Val(smFieldValues(1))) Then
                        tmSbf.lChfCode = tmChf.lCode
                    Else
                        ilError = True
                        slMsg = slMsg & " Missing"
                    End If
                        'Test that date date request date
                    If gDateValue(smFieldValues(2)) <> lmEndStd Then
                        slMsg = slMsg & ", Bill Date " & smFieldValues(2) & " not matching requested date"
                        ilError = True
                    End If
                    gPackDate smFieldValues(2), tmSbf.iDate(0), tmSbf.iDate(1)
                    tmSbf.sTranType = "T"
                    ilMatch = False
                    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        If StrComp(Trim$(tgMVef(ilLoop).sName), Trim$(smFieldValues(5)), 1) = 0 Then
                            tmSbf.iBillVefCode = tgMVef(ilLoop).iCode
                            ilMatch = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilMatch Then
                        slMsg = slMsg & ", " & smFieldValues(5) & " Bill Vehicle Missing"
                        ilError = True
                    Else
                        ilMatch = False
                        For ilLoop = 0 To UBound(tmMktVefCode) - 1 Step 1
                            If tmSbf.iBillVefCode = tmMktVefCode(ilLoop) Then
                                ilMatch = True
                                Exit For
                            End If
                        Next ilLoop
                        If Not ilMatch Then
                            slMsg = slMsg & ", " & smFieldValues(5) & " Bill Vehicle not in Market"
                            ilError = True
                        End If
                    End If
                    tmSbf.iNoItems = Val(smFieldValues(11))
                    tmSbf.lGross = gStrDecToLong(smFieldValues(7), 2)
                    tmSbf.sBilled = "N"
                    tmSbf.sCashTrade = smFieldValues(4)
                    ilMatch = False
                    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        If StrComp(Trim$(tgMVef(ilLoop).sName), Trim$(smFieldValues(6)), 1) = 0 Then
                            tmSbf.iAirVefCode = tgMVef(ilLoop).iCode
                            ilMatch = True
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilMatch Then
                        slMsg = slMsg & ", " & smFieldValues(6) & " Air Vehicle Missing"
                        ilError = True
                    Else
                        ilMatch = False
                        For ilLoop = 0 To UBound(tmMktVefCode) - 1 Step 1
                            If tmSbf.iAirVefCode = tmMktVefCode(ilLoop) Then
                                ilMatch = True
                                Exit For
                            End If
                        Next ilLoop
                        If Not ilMatch Then
                            slMsg = slMsg & ", " & smFieldValues(5) & " Air Vehicle not in Market"
                            ilError = True
                        End If
                    End If
                    tmSbf.iAirNoSpots = Val(smFieldValues(12))
                    tmSbf.iBonusNoSpots = Val(smFieldValues(13))
                    tmSbf.lTax1 = gStrDecToLong(smFieldValues(9), 2)
                    tmSbf.lTax2 = gStrDecToLong(smFieldValues(10), 2)
                    gPackDate smNowDate, tmSbf.iImportDate(0), tmSbf.iImportDate(1)
                    gPackDate smFieldValues(15), tmSbf.iExportDate(0), tmSbf.iExportDate(1)
                    tmSbf.lRefInvNo = Val(smFieldValues(3))
                    tmSbf.iCombineID = Val(smFieldValues(14))
                    tmSbf.lOGross = gStrDecToLong(smFieldValues(16), 2)
                    tmSbf.iCommPct = gStrDecToInt(smFieldValues(17), 2)
                    If Not ilError Then
                        ilRet = mMerge("I")
                        If ilRet Then
                            imChg = True
                        End If
                    Else
                        Print #hmMsg, slMsg
                        ilErrorLogged = True
                    End If
                End If
            End If
        End If
    Loop Until ilEof
    Close hmFrom
    If ilErrorLogged Then
        mReadImportFile = False
    Else
        mReadImportFile = True
    End If
    mSetCommands
    Exit Function
mReadImportFileErr:
    ilRet = Err
    Resume Next

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadSbfRec                     *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read item bill records         *
'*                                                     *
'*******************************************************
Private Function mReadSbfRec(ilTestForSbf As Integer) As Integer
'
'   iRet = mReadSbfRec
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilVeh As Integer
    Dim llDate As Long

    If (imMarketIndex >= 0) And (imInvDateIndex >= 0) Then
        imSbfRecLen = Len(tmSbf)
        tmSbfSrchKey2.sTranType = "T"
        'tmSbfSrchKey2.iDate(0) = 0
        'tmSbfSrchKey2.iDate(1) = 0
        gPackDate smStartStd, tmSbfSrchKey2.iDate(0), tmSbfSrchKey2.iDate(1)
        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.sTranType = "T")
            gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
            If llDate > lmEndStd Then
                Exit Do
            End If
            For ilVeh = 0 To UBound(tmMktVefCode) - 1 Step 1
                If tmSbf.iAirVefCode = tmMktVefCode(ilVeh) Then
                    If ilTestForSbf Then
                        mReadSbfRec = True
                        Exit Function
                    End If
                    imChg = False
                    ilRet = mMerge("F")
                    Exit For
                End If
            Next ilVeh
            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    If ilTestForSbf Then
        mReadSbfRec = False
    Else
        mReadSbfRec = True
    End If
    Exit Function
mReadSbfRecErr:
    On Error GoTo 0
    mReadSbfRec = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mRecomputeTotals                *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Recompute all totals           *
'*                                                     *
'*******************************************************
Sub mRecomputeTotals()
    Dim slONoSpots As String
    Dim slOGross As String
    Dim slANoSpots As String
    Dim slAGross As String
    Dim slBonusNoSpots As String
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String

    ilStartRow = LBound(smSave, 2)
    ilEndRow = ilStartRow + 1
    If ilEndRow >= UBound(smSave, 2) Then
        Exit Sub
    End If
    Do
        If InStr(1, smShow(ADVTINDEX, ilEndRow), "Total:", 1) > 0 Then
            slONoSpots = "0"
            slOGross = "0.00"
            slANoSpots = "0"
            slAGross = "0.00"
            slBonusNoSpots = "0"
            For ilRow = ilStartRow To ilEndRow - 1 Step 1
                slONoSpots = gAddStr(slONoSpots, smSave(3, ilRow))
                slOGross = gAddStr(slOGross, smSave(4, ilRow))
                slANoSpots = gAddStr(slANoSpots, smSave(5, ilRow))
                slAGross = gAddStr(slAGross, smSave(6, ilRow))
                slBonusNoSpots = gAddStr(slBonusNoSpots, smSave(7, ilRow))
            Next ilRow
            smSave(3, ilEndRow) = slONoSpots
            smSave(4, ilEndRow) = slOGross
            smSave(5, ilEndRow) = slANoSpots
            smSave(6, ilEndRow) = slAGross
            smSave(7, ilEndRow) = slBonusNoSpots
            gSetShow pbcInvRemote, slONoSpots, tmCtrls(ONOSPOTSINDEX)
            smShow(ONOSPOTSINDEX, ilEndRow) = tmCtrls(ONOSPOTSINDEX).sShow
            gSetShow pbcInvRemote, slOGross, tmCtrls(OGROSSINDEX)
            smShow(OGROSSINDEX, ilEndRow) = tmCtrls(OGROSSINDEX).sShow
            gSetShow pbcInvRemote, slANoSpots, tmCtrls(ANOSPOTSINDEX)
            smShow(ANOSPOTSINDEX, ilEndRow) = tmCtrls(ANOSPOTSINDEX).sShow
            gSetShow pbcInvRemote, slAGross, tmCtrls(AGROSSINDEX)
            smShow(AGROSSINDEX, ilEndRow) = tmCtrls(AGROSSINDEX).sShow
            gSetShow pbcInvRemote, slBonusNoSpots, tmCtrls(ABONUSINDEX)
            smShow(ABONUSINDEX, ilEndRow) = tmCtrls(ABONUSINDEX).sShow
            slStr = Trim$(Str$(ilEndRow - ilStartRow))
            gSetShow pbcInvRemote, slStr, tmCtrls(VEHICLEINDEX)
            smShow(VEHICLEINDEX, ilEndRow) = tmCtrls(VEHICLEINDEX).sShow
            ilStartRow = ilEndRow + 1
            ilEndRow = ilStartRow + 1
        Else
            ilEndRow = ilEndRow + 1
        End If
    Loop While ilEndRow < UBound(smSave, 2) - 1
    mGrandTotal
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mRemoveAirCount                 *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: If no sbf and no import set    *
'*                      set air counts to zero         *
'*                                                     *
'*******************************************************
Sub mRemoveAirCount(ilRemoveAll As Integer)
    Dim ilLoop As Integer
    Dim slStr As String

    For ilLoop = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
        If ((lmSave(2, ilLoop) = 0) And (Trim$(smInfo(1, ilLoop)) = "C")) Or (ilRemoveAll) Then
            smSave(5, ilLoop) = "0"
            smSave(6, ilLoop) = "0.00"
            smSave(7, ilLoop) = "0"
            gSetShow pbcInvRemote, smSave(5, ilLoop), tmCtrls(ANOSPOTSINDEX)
            smShow(ANOSPOTSINDEX, ilLoop) = tmCtrls(ANOSPOTSINDEX).sShow
            gSetShow pbcInvRemote, smSave(6, ilLoop), tmCtrls(AGROSSINDEX)
            smShow(AGROSSINDEX, ilLoop) = tmCtrls(AGROSSINDEX).sShow
            gSetShow pbcInvRemote, smSave(7, ilLoop), tmCtrls(ABONUSINDEX)
            smShow(ABONUSINDEX, ilLoop) = tmCtrls(ABONUSINDEX).sShow
        End If
    Next ilLoop
End Sub

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
    Dim slStamp As String   'Date/Time stamp for file
    Dim ilSbf As Integer
    Dim tlSbf As SBF

    Dim tlSbf1 As MOVEREC
    Dim tlSbf2 As MOVEREC


    mSetShow imBoxNo
    If mTestFields() = NO Then
        mSaveRec = False
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    For ilLoop = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
        If (smSave(8, ilLoop) <> "Y") And (InStr(1, smShow(ADVTINDEX, ilLoop), "Total:", 1) = 0) Then
            mMoveCtrlToRec ilLoop
            Do  'Loop until record updated or added
                If tmSbf.lCode = 0 Then 'New selected
                    tmSbf.lCode = 0
                    tmSbf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                    ilRet = btrInsert(hmSbf, tmSbf, imSbfRecLen, INDEXKEY0)
                    slMsg = "mSaveRec (btrInsert: Remote Posting)"
                Else 'Old record-Update
                    slMsg = "mSaveRec (btrGetEqual: Remote Posting)"
                    tmSbfSrchKey1.lCode = tmSbf.lCode
                    ilRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    On Error GoTo mSaveRecErr
                    gBtrvErrorMsg ilRet, slMsg, InvRemote
                    On Error GoTo 0
                    tlSbf1 = tlSbf
                    tlSbf2 = tmSbf
                    If StrComp(tlSbf1.sChar, tlSbf2.sChar, 0) <> 0 Then
                        Do
                            tmSbf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                            ilRet = btrUpdate(hmSbf, tmSbf, imSbfRecLen)
                            If ilRet = BTRV_ERR_CONFLICT Then
                                tmSbfSrchKey1.lCode = tmSbf.lCode
                                ilCRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        slMsg = "mSaveRec (btrUpdate: Remote Posting)"
                    Else
                        ilRet = BTRV_ERR_NONE
                    End If
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            On Error GoTo mSaveRecErr
            gBtrvErrorMsg ilRet, slMsg, InvRemote
            On Error GoTo 0
            lmSave(2, ilLoop) = tmSbf.lCode
        End If
    Next ilLoop
    For ilLoop = 0 To UBound(lmSbfDel) - 1 Step 1
        slMsg = "mSaveRec (btrGetEqual for Delete: Remote Posting)"
        tmSbfSrchKey1.lCode = lmSbfDel(ilLoop)
        ilRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, InvRemote
        On Error GoTo 0
        Do
            ilRet = btrDelete(hmSbf)
            If ilRet = BTRV_ERR_CONFLICT Then
                tmSbfSrchKey1.lCode = lmSbfDel(ilLoop)
                ilCRet = btrGetEqual(hmSbf, tlSbf, imSbfRecLen, tmSbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        slMsg = "mSaveRec (btrDelete: Remote Posting)"
        On Error GoTo mSaveRecErr
        gBtrvErrorMsg ilRet, slMsg, InvRemote
        On Error GoTo 0
    Next ilLoop
    ReDim lmSbfDel(0 To 0) As Long
    imChg = False
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
    Dim ilAltered As Integer
    If imChg Then
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
    Dim ilAltered As Integer
    'Update button set if all mandatory fields have data and any field altered
    If imChg Then
        cmcUpdate.Enabled = True
        cbcInvDate.Enabled = False
        cbcMarket.Enabled = False
    Else
        cbcInvDate.Enabled = True
        cbcMarket.Enabled = True
        cmcUpdate.Enabled = False
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

    Erase tmMktCode
    Erase tmMktVefCode

    Erase tmContract
    Erase tmChfAdvtExt
    Erase tmVehicle

    Erase tmInvVehicle
    smInvVehicleTag = ""
    Erase lmSbfDel

    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    ilRet = btrClose(hmChf)
    btrDestroy hmChf
    ilRet = btrClose(hmLcf)
    btrDestroy hmLcf
    ilRet = btrClose(hmVsf)
    btrDestroy hmVsf
    Erase smSave
    Erase imSave
    Erase lmSave
    Erase smShow
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload InvRemote
    Set InvRemote = Nothing   'Remove data segment
    igManUnload = NO
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestVehType                    *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if vehicle requested       *
'*                     Taken from PopSubs.Bas          *
'*                                                     *
'*******************************************************
Private Function mTestVehType(ilVehType As Integer, tlVef As VEF) As Integer
    Dim ilOk As Integer
    Dim ilVpfIndex As Integer
    Dim ilLoop As Integer

    ilOk = True
    Select Case Trim$(tlVef.sType)
        Case "C"
            If ((ilVehType And VEHCONV_WO_FEED) <> VEHCONV_WO_FEED) Or ((ilVehType And VEHCONV_W_FEED) <> VEHCONV_W_FEED) Then
                'ilVpfIndex = gVpfFind(Frm, tlVef.iCode)
                ilVpfIndex = -1
                For ilLoop = 0 To UBound(tgVpf) Step 1
                    If tlVef.iCode = tgVpf(ilLoop).iVefKCode Then
                        ilVpfIndex = ilLoop
                        Exit For
                    End If
                Next ilLoop
                If ilVpfIndex = -1 Then
                    mTestVehType = False
                    Exit Function
                End If
                If (ilVehType And VEHCONV_WO_FEED) <> VEHCONV_WO_FEED Then
                    If tgVpf(ilVpfIndex).iGMnfNCode(1) = 0 Then
                        ilOk = False
                    End If
                End If
                If (ilVehType And VEHCONV_W_FEED) <> VEHCONV_W_FEED Then
                    If tgVpf(ilVpfIndex).iGMnfNCode(1) > 0 Then
                        ilOk = False
                    End If
                End If
            End If
            'If Log vehicle requested, then exclude vehicles that have log vehicle
            If (ilVehType And VEHLOG) = VEHLOG Then
                If (ilVehType And VEHLOGVEHICLE) <> VEHLOGVEHICLE Then
                    If tlVef.iVefCode > 0 Then
                        ilOk = False
                    End If
                End If
            End If
        Case "S"
            If (ilVehType And VEHSELLING) <> VEHSELLING Then
                ilOk = False
            End If
        Case "A"
            If (ilVehType And VEHAIRING) <> VEHAIRING Then
                ilOk = False
            End If
        Case "L"
            If (ilVehType And VEHLOG) <> VEHLOG Then
                ilOk = False
            End If
        Case "V"
            If (ilVehType And VEHVIRTUAL) <> VEHVIRTUAL Then
                ilOk = False
            End If
        Case "T"
            If (ilVehType And VEHSIMUL) <> VEHSIMUL Then
                ilOk = False
            End If
        Case "P"
            If ((ilVehType And VEHPACKAGE) <> VEHPACKAGE) And ((ilVehType And VEHSTDPKG) <> VEHSTDPKG) Then
                ilOk = False
            Else
                If ((ilVehType And VEHSTDPKG) = VEHSTDPKG) And ((ilVehType And VEHPACKAGE) <> VEHPACKAGE) Then
                    If tlVef.lPvfCode <= 0 Then
                        ilOk = False
                    End If
                End If
            End If
    End Select
    Select Case tlVef.sState
        Case "A"
            If (ilVehType And ACTIVEVEH) <> ACTIVEVEH Then
                ilOk = False
            End If
        Case "D"
            If (ilVehType And DORMANTVEH) <> DORMANTVEH Then
                ilOk = False
            End If
    End Select
    'If ilOk Then
    '    'Market vehicle selection
    '    If ((ilVehType And VEHBYMKT) = VEHBYMKT) And (tgSpf.sMktBase = "Y") Then
    '        ilOk = False
    '        For ilLoop = 0 To UBound(igMktCode) - 1 Step 1
    '            If tlVef.iMnfVehGp3Mkt = igMktCode(ilLoop) Then
    '                ilOk = True
    '                Exit For
    '            End If
    '        Next ilLoop
    '    End If
    'End If
    If ilOk Then
        'Market vehicle selection
        If ((ilVehType And VEHBYPASSNOLOG) = VEHBYPASSNOLOG) Then
            For ilLoop = 0 To UBound(tgVpf) Step 1
                If tlVef.iCode = tgVpf(ilLoop).iVefKCode Then
                    If tgVpf(ilLoop).sGenLog = "N" Then
                        ilOk = False
                    End If
                    Exit For
                End If
            Next ilLoop
        End If
    End If
    mTestVehType = ilOk
End Function

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
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim ilLoop As Integer
    Dim ilVehType As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    
    ilVehType = VEHCONV_WO_FEED + VEHCONV_W_FEED + VEHSELLING + VEHVIRTUAL + ACTIVEVEH
    lbcVehicle.Clear
    smVehicleTag = ""
    ReDim tmVehicle(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If tgMVef(ilVef).iMnfVehGp3Mkt = imMktCode Then
            If mTestVehType(ilVehType, tgMVef(ilVef)) Then
                tmVehicle(UBound(tmVehicle)).sKey = tgMVef(ilVef).sName & "\" & Trim$(Str$(tgMVef(ilVef).iCode))
                ReDim Preserve tmVehicle(0 To UBound(tmVehicle) + 1) As SORTCODE 'VB list box clear (list box used to retain code number so record can be found)
            End If
        End If
    Next ilVef
    If UBound(tmVehicle) - 1 > 0 Then
        'ArraySortTyp tmVehicle(), tmVehicle(0), UBound(tmVehicle), 0, Len(tmVehicle(0)), 0, Len(tmVehicle(0).sKey), 0
        ArraySortTyp fnAV(tmVehicle(), 0), UBound(tmVehicle), 0, LenB(tmVehicle(0)), 0, LenB(tmVehicle(0).sKey), 0
    End If
    For ilLoop = 0 To UBound(tmVehicle) - 1 Step 1
        slNameCode = tmVehicle(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        slName = Trim$(slName)
        lbcVehicle.AddItem slName  'Add ID to list box
    Next ilLoop
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
Private Sub pbcArrow_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub

Sub pbcArrow_KeyUp(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
    imBoxNo = -1
    pbcArrow.Visible = False
    lacFrame.Visible = False
End Sub
Private Sub pbcClickFocus_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYF1 Then    'Functio key 1= Help
        'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
        'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
        'Traffic!cdcSetup.Action = 6
    End If
End Sub
Private Sub pbcSTab_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub pbcSTab_GotFocus()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim slStr As String
    Dim ilFound As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    imTabDirection = -1  'Set-right to left
    ilBox = imBoxNo
    ilRow = imRowNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imTabDirection = 0  'Set-Left to right
                imSettingValue = True
                vbcInvRemote.Value = vbcInvRemote.Min
                If UBound(smSave, 2) <= vbcInvRemote.LargeChange + 1 Then 'was <=
                    vbcInvRemote.Max = LBound(smSave, 2)
                Else
                    vbcInvRemote.Max = UBound(smSave, 2) - vbcInvRemote.LargeChange ' - 1
                End If
                imRowNo = 1
                Do While (imRowNo < UBound(smSave, 2)) And (smSave(8, imRowNo) = "Y")
                    imRowNo = imRowNo + 1
                    If imRowNo > vbcInvRemote.Value + vbcInvRemote.LargeChange Then
                        imSettingValue = True
                        vbcInvRemote.Value = vbcInvRemote.Value + 1
                    End If
                Loop
                If (imRowNo = UBound(smSave, 2)) Then
                    cmcCancel.SetFocus
                    Exit Sub
                End If
                ilBox = ANOSPOTSINDEX
                imBoxNo = ilBox
                mEnableBox ilBox
                Exit Sub
            Case VEHICLEINDEX 'Name (first control within header)
                If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                    mSetShow imBoxNo
                End If
                ilBox = ABONUSINDEX
                If imRowNo <= 1 Then
                    imBoxNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imRowNo = imRowNo - 1
                If imRowNo < vbcInvRemote.Value Then
                    imSettingValue = True
                    vbcInvRemote.Value = vbcInvRemote.Value - 1
                End If
                'imBoxNo = ilBox
                'mEnableBox ilBox
                'Exit Sub
            Case ANOSPOTSINDEX 'Name (first control within header)
                If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                    mSetShow imBoxNo
                End If
                ilBox = ABONUSINDEX
                If imRowNo <= 1 Then
                    imBoxNo = -1
                    cmcDone.SetFocus
                    Exit Sub
                End If
                imRowNo = imRowNo - 1
                If imRowNo < vbcInvRemote.Value Then
                    imSettingValue = True
                    vbcInvRemote.Value = vbcInvRemote.Value - 1
                End If
                'imBoxNo = ilBox
                'mEnableBox ilBox
                'Exit Sub
            Case Else
                ilBox = ilBox - 1
        End Select
        If (smSave(8, imRowNo) = "Y") Or (InStr(1, smShow(ADVTINDEX, imRowNo), "Total:", 1) > 0) Then
            ilFound = False
        End If
    Loop While Not ilFound
    If (imRowNo = ilRow) Then
        mSetShow imBoxNo
    End If
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcTab_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilFound As Integer
    
    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    imTabDirection = 0  'Set-Left to right
    ilBox = imBoxNo
    ilRow = imRowNo
    Do
        ilFound = True
        Select Case ilBox
            Case -1 'Tab from control prior to form area
                imTabDirection = -1  'Set-Right to left
                imRowNo = UBound(smSave, 2)
                If imRowNo = LBound(smSave, 2) Then
                    cmcCancel.SetFocus
                    Exit Sub
                End If
                imSettingValue = True
                If imRowNo <= vbcInvRemote.LargeChange + 1 Then
                    vbcInvRemote.Value = vbcInvRemote.Min
                Else
                    vbcInvRemote.Value = imRowNo - vbcInvRemote.LargeChange
                End If
                ilBox = ANOSPOTSINDEX
            Case 0
                ilBox = ANOSPOTSINDEX
            Case VEHICLEINDEX
                ilBox = ANOSPOTSINDEX
            Case ABONUSINDEX 'Last control
                If (imBoxNo = ilBox) And (imRowNo = ilRow) Then
                    mSetShow imBoxNo
                    If mTestSaveFields(imRowNo) = NO Then
                        mEnableBox imBoxNo
                        Exit Sub
                    End If
                End If
                imRowNo = imRowNo + 1
                If imRowNo > vbcInvRemote.Value + vbcInvRemote.LargeChange Then
                    imSettingValue = True
                    vbcInvRemote.Value = vbcInvRemote.Value + 1
                End If
                If imRowNo >= UBound(smSave, 2) Then
                    mSetCommands
                    imBoxNo = 0
                    cmcCancel.SetFocus
                    Exit Sub
                End If
                ilBox = ANOSPOTSINDEX
            Case Else
                ilBox = ilBox + 1
        End Select
        If (smSave(8, imRowNo) = "Y") Or (InStr(1, smShow(ADVTINDEX, imRowNo), "Total:", 1) > 0) Then
            ilFound = False
        End If
    Loop While Not ilFound
    If (imRowNo = ilRow) Then
        mSetShow imBoxNo
    End If
    imBoxNo = ilBox
    mEnableBox ilBox
End Sub
Private Sub pbcInvRemote_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Sub pbcInvRemote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    imButton = Button
    If Button = 2 Then  'Right Mouse
        ilCompRow = vbcInvRemote.LargeChange + 1
        If UBound(smSave, 2) > ilCompRow - 1 Then
            ilMaxRow = ilCompRow
        Else
            ilMaxRow = UBound(smSave, 2) - 1
        End If
        For ilRow = 1 To ilMaxRow Step 1
            For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
                If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                    If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                        imButtonRow = ilRow + vbcInvRemote.Value - 1
                        mShowInfo
                        Exit Sub
                    End If
                End If
            Next ilBox
        Next ilRow
    End If
End Sub
Sub pbcInvRemote_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer

    If imIgnoreRightMove Then
        Exit Sub
    End If
    imButton = Button
    If Button <> 2 Then  'Right Mouse
        Exit Sub
    End If
    imButton = Button
    imIgnoreRightMove = True
    ilCompRow = vbcInvRemote.LargeChange + 1
    If UBound(smSave, 2) > ilCompRow - 1 Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smSave, 2) - 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    imButtonRow = ilRow + vbcInvRemote.Value - 1
                    mShowInfo
                    imIgnoreRightMove = False
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
    imIgnoreRightMove = False
End Sub
Private Sub pbcInvRemote_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim ilLoop As Integer

    If Button = 2 Then
        plcInfo.Visible = False
        Exit Sub
    End If
    ilCompRow = vbcInvRemote.LargeChange + 1
    If UBound(smSave, 2) > ilCompRow - 1 Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(smSave, 2) - 1
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            If (X >= tmCtrls(ilBox).fBoxX) And (X <= (tmCtrls(ilBox).fBoxX + tmCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmCtrls(ilBox).fBoxY + tmCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcInvRemote.Value - 1
                    mSetShow imBoxNo
                    If ilRowNo > UBound(smSave, 2) - 1 Then
                        Beep
                        Exit Sub
                    End If
                    If smSave(8, ilRowNo) = "Y" Then    'If billed disallow change
                        Beep
                        Exit Sub
                    End If
                    If smShow(CONTRACTINDEX, ilRowNo) = "" Then
                        Beep
                        Exit Sub
                    End If
                    If (ilBox <> AGROSSINDEX) And (ilBox <> ANOSPOTSINDEX) And (ilBox <> ABONUSINDEX) Then
                        If (Trim$(smInfo(1, ilRowNo)) <> "S") Or (lmSave(2, ilRowNo) <> 0) Or (ilBox <> VEHICLEINDEX) Then
                            imRowNo = ilRowNo
                            imBoxNo = 0
                            lacFrame.Move 0, tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcInvRemote.Value) * (fgBoxGridH + 15) - 30
                            lacFrame.Visible = True
                            pbcArrow.Move pbcArrow.Left, plcInvRemote.Top + tmCtrls(VEHICLEINDEX).fBoxY + (imRowNo - vbcInvRemote.Value) * (fgBoxGridH + 15) + 45
                            pbcArrow.Visible = True
                            pbcArrow.SetFocus
                            Exit Sub
                        End If
                    End If
                    imRowNo = ilRowNo
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            End If
        Next ilBox
    Next ilRow
End Sub
Private Sub pbcInvRemote_Paint()
    Dim ilBox As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim llColor As Long
    
    llColor = pbcInvRemote.ForeColor
    ilStartRow = vbcInvRemote.Value  'Top location
    ilEndRow = vbcInvRemote.Value + vbcInvRemote.LargeChange
    If ilEndRow > UBound(smSave, 2) - 1 Then
        ilEndRow = UBound(smSave, 2) - 1 'Don't include blank row
    End If
    For ilRow = ilStartRow To ilEndRow Step 1
        For ilBox = imLBCtrls To UBound(tmCtrls) Step 1
            'If (ilBox = VEHICLEINDEX) Then
            '    If (lmSave(1, ilRow) > 0) Or (lmSave(2, ilRow) > 0) Then
            '        gPaintArea pbcInvRemote, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
            '    Else
            '        gPaintArea pbcInvRemote, tmCtrls(ilBox).fBoxX, tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmCtrls(ilBox).fBoxW - 15, tmCtrls(ilBox).fBoxH - 15, WHITE
            '    End If
            'End If
            If smSave(8, ilRow) = "Y" Then    'If billed- override any other color
                pbcInvRemote.ForeColor = DARKGREEN
            End If
            If (imSave(4, ilRow) = False) And (ilBox = VEHICLEINDEX) And ((Trim$(smInfo(1, ilRow)) = "I") Or (Trim$(smInfo(1, ilRow)) = "F") Or (Trim$(smInfo(1, ilRow)) = "C")) Then
                pbcInvRemote.ForeColor = RED
            End If
            pbcInvRemote.CurrentX = tmCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcInvRemote.CurrentY = tmCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
            If (ilBox = DNOSPOTSINDEX) Then
                slStr = ""
                'If smSave(5, ilRow) <> "" Then
                '    slStr = gSubStr(smSave(5, ilRow), smSave(3, ilRow))
                'End If
                If (smSave(5, ilRow) <> "") And (smSave(7, ilRow) <> "") Then
                    slStr = gAddStr(smSave(5, ilRow), smSave(7, ilRow))
                ElseIf (smSave(5, ilRow) <> "") Then
                    slStr = smSave(5, ilRow)
                ElseIf (smSave(7, ilRow) <> "") Then
                    slStr = smSave(7, ilRow)
                End If
                If (slStr <> "") Then
                    slStr = gSubStr(slStr, smSave(3, ilRow))
                End If
                gSetShow pbcInvRemote, slStr, tmCtrls(DNOSPOTSINDEX)
                slStr = tmCtrls(DNOSPOTSINDEX).sShow
                pbcInvRemote.Print slStr
            ElseIf (ilBox = DGROSSINDEX) Then
                slStr = ""
                If smSave(6, ilRow) <> "" Then
                    slStr = gSubStr(smSave(6, ilRow), smSave(4, ilRow))
                End If
                gSetShow pbcInvRemote, slStr, tmCtrls(DGROSSINDEX)
                slStr = tmCtrls(DGROSSINDEX).sShow
                pbcInvRemote.Print slStr
            Else
                pbcInvRemote.Print Trim$(smShow(ilBox, ilRow))
            End If
            If (ilBox = VEHICLEINDEX) Then
                pbcInvRemote.ForeColor = llColor
            End If
        Next ilBox
        pbcInvRemote.ForeColor = llColor
    Next ilRow
End Sub
Private Sub plcInvRemote_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcInvRemote_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub plcSelect_Click()
    pbcClickFocus.SetFocus
End Sub
Private Sub plcSelect_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub

Private Sub tmcClick_Timer()
    Dim ilRet As Integer
    Dim ilSbfFound As Integer
    Dim slMsgFile As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilLoop As Integer
    Dim slFYear As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim ilSbfExist As Integer

    Screen.MousePointer = vbHourglass
    tmcClick.Enabled = False
    pbcInvRemote.Cls
    mClearCtrlFields

    ilSbfFound = False
    mBuildDate
    ReDim tmMktVefCode(0 To 0) As Integer
    pbcInvRemote.Cls
    If imMarketIndex >= 0 Then
        slNameCode = tmMktCode(imMarketIndex).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        imMktCode = Val(slCode)
        For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If tgMVef(ilLoop).iMnfVehGp3Mkt = imMktCode Then
                tmMktVefCode(UBound(tmMktVefCode)) = tgMVef(ilLoop).iCode
                ReDim Preserve tmMktVefCode(0 To UBound(tmMktVefCode) + 1) As Integer
            End If
        Next ilLoop
    Else
        imMktCode = -1
    End If
    igBrowserReturn = 0
    If (imInvDateIndex >= 0) And (imMarketIndex >= 0) Then
        If Not mReadSbfRec(True) Then
            ilSbfExist = False
            'Remove automatic import- Jim request on 4/26/02
            'Screen.MousePointer = vbDefault
            ''gObtainYearMonthDayStr smStartStd, True, slFYear, slFMonth, slFDay
            'gObtainYearMonthDayStr smEndStd, True, slFYear, slFMonth, slFDay
            'slFMonth = Left$(cbcInvDate.List(imInvDateIndex), 3)
            'igBrowserType = 7  'Mask
            '''sgBrowseMaskFile = "F" & Right$(slFYear, 2) & slFMonth & slFDay & "?.I??"
            ''sgBrowseMaskFile = "?" & right$(slFYear, 2) & slFMonth & slFDay & "?.I??"
            'sgBrowseMaskFile = slFMonth & Right$(slFYear, 2) & "In?.??"
            'sgBrowserTitle = "Import for " & cbcMarket.List(imMarketIndex)
            'Browser.Show vbModal
            'sgBrowserTitle = ""
        Else
            igBrowserReturn = 0
            ilSbfExist = True
        End If
    Else
        igBrowserReturn = 0
    End If
    Screen.MousePointer = vbHourglass
    'Get contracts for market
    mGetCntr
    If UBound(smSave, 2) > LBound(smSave, 2) Then
        imChg = True    'Set as changed so that contracts can be saved
    End If
    'Populate vehicle list box
    mVehPop
    'Get previously entered SBF for market
    ilRet = mReadSbfRec(False)
    pbcInvRemote.Cls
    If (imInvDateIndex >= 0) And (imMarketIndex >= 0) Then
        mAddGrandTotalLine
    End If
    If igBrowserReturn = 1 Then
        slMsgFile = sgBrowserFile
        If InStr(slMsgFile, ":") = 0 Then
            slMsgFile = sgImportPath & slMsgFile
        End If
        ilRet = mOpenMsgFile(slMsgFile)
        If Not ilRet Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Print #hmMsg, "Import " & sgBrowserFile & " " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        mRemoveAirCount True
        pbcInvRemote.Cls
        ilRet = mReadImportFile(sgBrowserFile)
        If ilRet Then
            Print #hmMsg, "Import Finish Successfully"
            Close #hmMsg
        Else
            Print #hmMsg, "** Import Errors or Terminated " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM") & " **"
            Close #hmMsg
            MsgBox "See " & slMsgFile & " for errors related to Rejected Records"
        End If
    Else
        'Retain air count Jim request on 4/24/02 along with not importing
        ''Remove Aired count if not previously defined and not imported
        'mRemoveAirCount False
        pbcInvRemote.Cls
    End If
    'pbcInvRemote.Cls
    'Compute totals
    mRecomputeTotals
    vbcInvRemote.Min = LBound(smSave, 2)
    If UBound(smSave, 2) <= vbcInvRemote.LargeChange Then
        vbcInvRemote.Max = LBound(smSave, 2)
    Else
        vbcInvRemote.Max = UBound(smSave, 2) - vbcInvRemote.LargeChange
    End If
    If vbcInvRemote.Value = vbcInvRemote.Min Then
        pbcInvRemote_Paint
    Else
        vbcInvRemote.Value = vbcInvRemote.Min
    End If
    If (imInvDateIndex >= 0) And (imMarketIndex >= 0) Then
        cmcImport.Enabled = True
    Else
        cmcImport.Enabled = False
    End If
    If imChg Then
        cmcUpdate.Enabled = True
    Else
        cmcUpdate.Enabled = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub vbcInvRemote_Change()
    If imSettingValue Then
        pbcInvRemote.Cls
        pbcInvRemote_Paint
        imSettingValue = False
    Else
        mSetShow imBoxNo
        pbcInvRemote.Cls
        pbcInvRemote_Paint
        mEnableBox imBoxNo
    End If
End Sub
Private Sub vbcInvRemote_DragDrop(Source As Control, X As Single, Y As Single)
    lacFrame.DragIcon = IconTraf!imcIconStd.DragIcon
End Sub
Private Sub vbcInvRemote_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub
