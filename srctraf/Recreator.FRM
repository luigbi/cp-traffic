VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Recreator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Re-Creation Utility"
   ClientHeight    =   6075
   ClientLeft      =   4245
   ClientTop       =   3630
   ClientWidth     =   7560
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
   Icon            =   "Recreator.frx":0000
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.PictureBox plcSignon 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6075
      Left            =   0
      Picture         =   "Recreator.frx":030A
      ScaleHeight     =   6015
      ScaleWidth      =   7485
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   7545
      Begin VB.CheckBox ckcPackage 
         BackColor       =   &H00FFFFC0&
         Caption         =   "All Vehicles"
         Height          =   210
         Left            =   2160
         TabIndex        =   17
         Top             =   2205
         Width           =   195
      End
      Begin vb6projectRecreator.CSI_Calendar cccResearch 
         Height          =   270
         Left            =   5610
         TabIndex        =   6
         Top             =   855
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   476
         Text            =   "9/14/11"
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BorderStyle     =   1
         CSI_ShowDropDownOnFocus=   -1  'True
         CSI_InputBoxBoxAlignment=   0
         CSI_CalBackColor=   16777130
         CSI_CalDateFormat=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CSI_CurDayBackColor=   16777215
         CSI_CurDayForeColor=   51200
         CSI_ForceMondaySelectionOnly=   0   'False
         CSI_AllowBlankDate=   0   'False
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   3
      End
      Begin vb6projectRecreator.CSI_Calendar cccPA 
         Height          =   270
         Left            =   2385
         TabIndex        =   4
         Top             =   855
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   476
         BorderStyle     =   1
         CSI_ShowDropDownOnFocus=   -1  'True
         CSI_InputBoxBoxAlignment=   0
         CSI_CalBackColor=   16777130
         CSI_CalDateFormat=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CSI_CurDayBackColor=   16777215
         CSI_CurDayForeColor=   0
         CSI_ForceMondaySelectionOnly=   0   'False
         CSI_AllowBlankDate=   0   'False
         CSI_AllowTFN    =   0   'False
         CSI_DefaultDateType=   3
      End
      Begin VB.PictureBox pbcTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   150
         ScaleHeight     =   30
         ScaleWidth      =   45
         TabIndex        =   25
         Top             =   5745
         Width           =   45
      End
      Begin VB.PictureBox pbcSTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   105
         ScaleHeight     =   30
         ScaleWidth      =   15
         TabIndex        =   20
         Top             =   5475
         Width           =   15
      End
      Begin VB.ListBox lbcAirVehicle 
         Appearance      =   0  'Flat
         Height          =   240
         ItemData        =   "Recreator.frx":9C57
         Left            =   2745
         List            =   "Recreator.frx":9C59
         Sorted          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   1560
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
         Left            =   5400
         Picture         =   "Recreator.frx":9C5B
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox edcDropDown 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4380
         MaxLength       =   20
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.PictureBox pbcToggle 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1440
         ScaleHeight     =   210
         ScaleWidth      =   795
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3225
         Visible         =   0   'False
         Width           =   795
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdVehicle 
         Height          =   2085
         Left            =   135
         TabIndex        =   19
         Top             =   2460
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   3678
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin VB.CommandButton cmcTo 
         Appearance      =   0  'Flat
         Caption         =   "&Browse..."
         Height          =   285
         Left            =   5715
         TabIndex        =   14
         Top             =   1800
         Width           =   1485
      End
      Begin VB.PictureBox plcTo 
         Height          =   375
         Left            =   1020
         ScaleHeight     =   315
         ScaleWidth      =   4245
         TabIndex        =   12
         Top             =   1740
         Width           =   4305
         Begin VB.TextBox edcTo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   15
            TabIndex        =   13
            Top             =   30
            Width           =   4230
         End
      End
      Begin VB.PictureBox plcFrom 
         Height          =   375
         Left            =   1020
         ScaleHeight     =   315
         ScaleWidth      =   4245
         TabIndex        =   8
         Top             =   1245
         Width           =   4305
         Begin VB.TextBox edcFrom 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   15
            TabIndex        =   9
            Top             =   30
            Width           =   4230
         End
      End
      Begin VB.CommandButton cmcFrom 
         Appearance      =   0  'Flat
         Caption         =   "&Browse..."
         Height          =   285
         Left            =   5715
         TabIndex        =   10
         Top             =   1365
         Width           =   1485
      End
      Begin VB.Timer tmcClick 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   1680
         Top             =   5085
      End
      Begin VB.CheckBox ckcDormant 
         BackColor       =   &H00FFFFC0&
         Caption         =   "All Vehicles"
         Height          =   210
         Left            =   135
         TabIndex        =   15
         Top             =   2205
         Width           =   195
      End
      Begin VB.PictureBox pbcClickFocus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   3495
         ScaleHeight     =   165
         ScaleWidth      =   105
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   105
      End
      Begin MSComDlg.CommonDialog CMDialogBox 
         Left            =   135
         Top             =   5040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   4100
         FontSize        =   0
         MaxFileSize     =   256
      End
      Begin VB.Label lacPackage 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Include Standard Packages"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2505
         TabIndex        =   18
         Top             =   2205
         Width           =   2295
      End
      Begin VB.Label lacRetentionDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Retention Dates (Active on or after Specified Date)"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   2
         Top             =   615
         Width           =   6930
      End
      Begin VB.Label lacTo 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   75
         TabIndex        =   11
         Top             =   1815
         Width           =   810
      End
      Begin VB.Label lacInfo 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   30
         Top             =   4590
         Width           =   1875
         WordWrap        =   -1  'True
      End
      Begin VB.Label lacInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   1950
         TabIndex        =   29
         Top             =   4575
         Width           =   3615
      End
      Begin VB.Label lacDormant 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Include Dormant"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   480
         TabIndex        =   16
         Top             =   2205
         Width           =   1635
      End
      Begin VB.Label lacFrom 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   75
         TabIndex        =   7
         Top             =   1365
         Width           =   810
      End
      Begin VB.Label lacPA 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Program and Agreements"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   75
         TabIndex        =   3
         Top             =   870
         Width           =   2340
      End
      Begin VB.Label lacEnd 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Research Book"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4185
         TabIndex        =   5
         Top             =   900
         Width           =   1305
      End
      Begin VB.Label lacStop 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3075
         TabIndex        =   28
         Top             =   5355
         Width           =   1335
      End
      Begin VB.Label lacMin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Minimize"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3075
         TabIndex        =   27
         Top             =   4845
         Width           =   1350
      End
      Begin VB.Image cmcCSLogo 
         Height          =   510
         Left            =   60
         Top             =   75
         Width           =   3210
      End
      Begin VB.Label imcOutline 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   420
         Left            =   4785
         TabIndex        =   26
         Top             =   4875
         Visible         =   0   'False
         Width           =   1410
      End
   End
End
Attribute VB_Name = "Recreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Recreator.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Set Advertiser/Agency Credit code
Option Explicit
Option Compare Text

Private Const LB_GETITEMHEIGHT = &H1A1

'Dim hmMsg As Integer

Dim imFirstActivate As Integer

Dim imLastColSorted As Integer
Dim imLastSort As Integer

Dim hmFrom As Integer
Dim hmTo As Integer

Dim tmRecBuff As POPREC

'Dim hmVsf As Integer            'Vehicle file handle
'Dim tmVsf As VSF                'VSF record image
'Dim imVsfRecLen As Integer       'VSF record length

'Dim hmSdf As Integer            'Spot file handle
'Dim tmSdf As SDF                'Spot detail record image
'Dim imSdfRecLen As Integer        'SDF record length
'Dim tmSdfSrchKey3 As LONGKEY0

'Dim hmSmf As Integer            'Spot MG file handle
'Dim tmSmf As SMF                'Spot MG record image
'Dim imSmfRecLen As Integer        'Spt MG record length

'Dim hmSsf As Integer
'Dim tmSsf As SSF                'Spot summary record image
'Dim imSsfRecLen As Integer        'SSF record length
'Dim tmSsfSrchKey As SSFKEY0 'SSF key record image
'Dim tmSsfSrchKey2 As SSFKEY2 'SSF key record image
'Dim tmAvail As AVAILSS
'Dim tmSpot As CSPOTSS

Dim hmFromUrf As Integer            'Spot MG file handle
Dim hmToUrf As Integer            'Spot MG file handle
Dim tmUrf As URF                'Spot MG record image
Dim imUrfRecLen As Integer        'Spt MG record length

Dim hmFromDpf As Integer            'Spot MG file handle
Dim hmToDpf As Integer            'Spot MG file handle
Dim tmDpf As DPF                'Spot MG record image
Dim imDpfRecLen As Integer        'Spt MG record length
Dim tmDpfSrchKey1 As DPFKEY1

Dim tmVef As VEF                'Vehicle record image
Dim imVefRecLen As Integer      'Vehicle record length

Dim tmPvf As PVF                'Vehicle record image
Dim hmPvf As Integer
Dim imPvfRecLen As Integer      'Vehicle record length
Dim tmPvfSrchKey0 As LONGKEY0

Dim smFromDBPath As String
Dim smToDBPath As String

Dim tmVpf As VPF
Dim tmVaf As VAF
Dim tmVbf As VBF
Dim tmVff As VFF
Dim tmVof As VOF
Dim tmPif As PIF
Dim tmLcf As LCF
Dim tmLvf As LVF
Dim tmLtf As LTF
Dim tmLef As LEF
Dim tmPaf As PAF
Dim tmPff As PFF
Dim tmVcf As VCF
Dim tmVlf As VLF
Dim tmIhf As IHF
Dim tmItf As ITF
Dim tmIif As IIF
Dim tmGhf As GHF
Dim tmGsf As GSF
Dim tmEnf As ENF
Dim tmAnf As ANF
Dim tmEtf As ETF
Dim tmDnf As DNF
Dim tmDrf As DRF
Dim tmDef As DEF
'Dim tmDpf As DPF
Dim tmRdf As RDF
Dim tmSaf As SAF
Dim tmAdf As ADF
Dim tmAgf As AGF
Dim tmSlf As SLF
Dim tmPrf As PRF
Dim tmArf As ARF
Dim tmTrf As TRF
Dim tmPnf As PNF
Dim tmCxf As CXF
Dim tmSof As SOF
Dim tmSnf As SNF
Dim tmSrf As SRF

Dim tmAtt As ATT
Dim tmDat As DAT
Dim tmEpt As EPT
Dim tmPft As PFT

Dim tmRht As RHT
Dim tmRet As RET

Dim tmCct As CCT
Dim tmCst As CST

Dim tmUst As UST
Dim tmDnt As DNT

Dim tmCef As CEF

Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imConverting As Integer
Dim lmNowDate As Long
Dim imAutoRun As Integer            '1 = auto run flag, 0 = manual
Dim imClosed As Integer
Dim imFileNo As Integer
Dim imTotalFiles As Integer
Dim lmTotalRecords As Long
Dim lmRecordNo As Long
Dim lmAdjCount As Long

Dim tmRecreateSort() As RECREATESORT
Dim imDnfCode() As Integer
Dim lmCefCode() As Long
Dim imArfCode() As Integer
Dim lmPvfCode() As Long
Dim tmLcfDateInfo() As LCFDATEINFO
Dim lmLvfCode() As Long
Dim tmLibTitleCount() As LIBTITLECOUNT
Dim imLtfCode() As Integer
Dim imIhfCode() As Integer
Dim imItfCode() As Integer
Dim imIifCode() As Integer
Dim lmGhfCode() As Long
Dim imEtfCode() As Integer
Dim imCheckDnfCode() As Integer
'Dim lmDrfCode() As Long
Dim imRdfCode() As Integer
Dim imAdfCode() As Integer
Dim imSlfCode() As Integer
Dim imAgfCode() As Integer
Dim imTrfCode() As Integer
Dim imPnfCode() As Integer
Dim imUrfCode() As Integer
Dim imSofCode() As Integer
Dim lmCxfCode() As Long
Dim imSnfCode() As Integer
Dim lmAttCode() As Long
Dim imCstCode() As Integer
Dim imUstCode() As Integer
Dim lmRhtCode() As Long

Dim imLbcArrowSetting As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imComboBoxIndex As Integer
Dim imChgMode As Integer

Dim lmPAStartDate As Long
Dim lmResearchStartDate As Long
Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim lmTopRow As Long
Dim bmIgnoreScroll As Boolean
Dim bmCtrlVisible As Boolean
Dim imLastTlfCode As Integer
Dim imLvfVersion As Integer

Dim imPassNo As Integer

Const NAMEINDEX = 0
Const STATUSINDEX = 1
Const TYPEINDEX = 2
Const AIRNAMEINDEX = 3
Const ORIGTYPEINDEX = 4
Const SORTINDEX = 5
Const VEFCODEINDEX = 6


Private Sub ckcPackage_Click()
    Dim ilVef As Integer
    Dim llRow As Long
    Dim llLoop As Long
    
    grdVehicle.Redraw = False
    If ckcPackage.Value = vbChecked Then
        'Add Package
        llRow = grdVehicle.FixedRows
        For llLoop = grdVehicle.Rows - 1 To grdVehicle.FixedRows Step -1
            If grdVehicle.TextMatrix(llLoop, NAMEINDEX) <> "" Then
                llRow = llLoop + 1
                Exit For
            End If
        Next llLoop
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If (tgMVef(ilVef).sType = "P") And (tgMVef(ilVef).lPvfCode > 0) Then
                If (tgMVef(ilVef).sState <> "D") Or ((tgMVef(ilVef).sState = "D") And (ckcDormant = vbChecked)) Then
                    mVehToGrid ilVef, llRow
                End If
            End If
        Next ilVef
        imLastColSorted = -1
        imLastSort = -1
        mSortCol NAMEINDEX
        grdVehicle.Row = 0
        grdVehicle.Col = VEFCODEINDEX
    Else
        'Remove Packages
        llRow = grdVehicle.Rows - 1
        Do While llRow >= grdVehicle.FixedRows
            If grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Package" Then
                grdVehicle.RemoveItem llRow
            Else
                llRow = llRow - 1
            End If
        Loop
    End If
    grdVehicle.Redraw = True
    mSetCommands
End Sub

Private Sub cmcDropDown_Click()
    Select Case grdVehicle.Col
        Case NAMEINDEX
        Case STATUSINDEX
        Case TYPEINDEX
        Case AIRNAMEINDEX
            lbcAirVehicle.Visible = Not lbcAirVehicle.Visible
    End Select

End Sub

Private Sub cmcFrom_Click()
    Dim slCurDir As String
    
    slCurDir = CurDir
    igPathType = 0
    sgGetPath = edcFrom.Text
    lgCallTop = Recreator.Top
    lgCallLeft = Recreator.Left
    lgCallHeight = Recreator.Height
    lgCallWidth = Recreator.Width
    GetPath.Show vbModal
    If igGetPath = 0 Then
        edcFrom.Text = sgGetPath
    End If
    
    ChDir slCurDir
    
    Exit Sub
End Sub

Private Sub edcDropDown_Change()
    Select Case grdVehicle.Col
        Case NAMEINDEX
        Case STATUSINDEX
        Case TYPEINDEX
        Case AIRNAMEINDEX
            imLbcArrowSetting = True
            mMatchLookAhead edcDropDown, lbcAirVehicle, imBSMode, imComboBoxIndex
    End Select
    imLbcArrowSetting = False
End Sub

Private Sub edcDropDown_GotFocus()
    Select Case grdVehicle.Col
        Case NAMEINDEX
        Case STATUSINDEX
        Case TYPEINDEX
        Case AIRNAMEINDEX
    End Select
    If Not imBypassFocus Then
        gCtrlGotFocus ActiveControl
    End If
    imBypassFocus = False
End Sub

Private Sub edcDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key causes the charact to the right of the cursor to be deleted
    imBSMode = False
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim ilFound As Integer
    Dim ilLoop As Integer
    If KeyAscii = 8 Then    'Process backspace key (delete key handled as a KeyDown Event)
        If edcDropDown.SelLength <> 0 Then    'avoid deleting two characters
            imBSMode = True 'Force deletion of character prior to selected text
        End If
    End If
    Select Case grdVehicle.Col
        Case NAMEINDEX
        Case STATUSINDEX
        Case TYPEINDEX
        Case AIRNAMEINDEX
    End Select

End Sub

Private Sub edcDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = KEYUP) Or (KeyCode = KEYDOWN) Then
        Select Case grdVehicle.Col
            Case NAMEINDEX
            Case STATUSINDEX
            Case TYPEINDEX
            Case AIRNAMEINDEX
                mProcessArrowKey Shift, KeyCode, lbcAirVehicle, imLbcArrowSetting
        End Select
        edcDropDown.SelStart = 0
        edcDropDown.SelLength = Len(edcDropDown.Text)
    End If
    If (KeyCode = KEYLEFT) Or (KeyCode = KEYRIGHT) Then
        Select Case grdVehicle.Col
            Case NAMEINDEX
            Case STATUSINDEX
            Case TYPEINDEX
            Case AIRNAMEINDEX
        End Select
    End If
End Sub

Private Sub edcFrom_Change()
    mSetCommands
End Sub

Private Sub edcFrom_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcTo_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
'    gShowBranner
End Sub

Private Sub Form_GotFocus()
    mSetShow
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        If imAutoRun = 0 Then
            MsgBox "Only one copy of Recreator can be run at a time, sorry", vbInformation + vbOKOnly, "Counterpoint"
        End If
        gLogMsg "Second copy of Recreator path: " & App.Path & " from " & Trim$(gGetComputerName()), "Recreator.txt", False
        End
    End If
    gSetBkgdMode
    'igBkgdProg = 6
    mInit
    If imAutoRun = 0 Then
        Me.WindowState = vbNormal
    Else
        Me.WindowState = vbMinimized
    End If
    tmcClick.Interval = 2000    '2 seconds
    tmcClick.Enabled = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ilRes As Integer
    Dim ilRet As Integer
    If imClosed = True Then
        Exit Sub
    End If
    Cancel = 0
    tmcClick.Enabled = False
    
    'If imAutoRun = 0 Then
    '    ilRet = MsgBox("Stop the Export Projection", vbQuestion + vbYesNo, "Stop Service")
    '    If ilRet = vbNo Then
    '        Cancel = 1
    '        imTerminate = False
    '        Exit Sub
    '    End If
    'End If
    imClosed = True
    imTerminate = True
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        Me.Left = Screen.Width / 2 - Me.Width / 2
        Me.Top = Screen.Height / 2 - Me.Height / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    Erase tmRecreateSort
    Erase imDnfCode
    Erase lmCefCode
    Erase imArfCode
    Erase lmPvfCode
    Erase tmLcfDateInfo
    Erase lmLvfCode
    Erase tmLibTitleCount
    Erase imLtfCode
    Erase imIhfCode
    Erase imItfCode
    Erase imIifCode
    Erase lmGhfCode
    Erase imEtfCode
    Erase imCheckDnfCode
    'Erase lmDrfCode
    Erase imRdfCode
    Erase imAdfCode
    Erase imSlfCode
    Erase imAgfCode
    Erase imTrfCode
    Erase imPnfCode
    Erase imUrfCode
    Erase imSofCode
    Erase lmCxfCode
    Erase imSnfCode
    
    Erase lmAttCode
    Erase imCstCode
    Erase imUstCode
    Erase lmRhtCode
    
    btrStopAppl
    Set Recreator = Nothing   'Remove data segment
    End
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    slCommand = Command$
    sgCommandStr = "debug"
    igDirectCall = 0
    sgIniPath = ""
    'slCommand = sgCommandStr  'Command$
    'If (StrComp(slCommand, "Debug", 1) = 0) Or (StrComp(slCommand, "CS", 1) = 0) Then
        igStdAloneMode = True 'Switch from/to stand alone mode
        sgCallAppName = ""
        slStr = "Guide"
        ilTestSystem = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
    '    ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
    '    If Trim$(slStr) = "" Then
    '        MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
    '        End
    '    End If
    '    ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
    '    ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
    '    If StrComp(slTestSystem, "Test", 1) = 0 Then
    '        ilTestSystem = True
    '    Else
    '        ilTestSystem = False
    '    End If
    '    ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    ilPos = InStr(1, slCommand, "/IniLoc:", vbTextCompare)
    If ilPos > 0 Then
        igDirectCall = -1
        ilSpace = InStr(ilPos, slCommand, " ")
        If ilSpace = 0 Then
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8))
        Else
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8, ilSpace - ilPos - 8))
        End If
        If right$(sgIniPath, 1) <> "\" Then
            sgIniPath = sgIniPath & "\"
        End If
    End If
    'imAutoRun = 1
    'ilPos = InStr(1, slCommand, "/UserInput", 1)
    'If ilPos > 0 Then
        imAutoRun = 0
        igBkgdProg = 0
        gLogMsg "Current Mode: UserInput.", "Recreator.txt", False
    'Else
    '    igBkgdProg = 8
    '    gLogMsg "Current Mode: Background.", "Recreator.txt", False
    'End If
    gInitStdAlone Recreator, slStr, ilTestSystem
End Sub



Private Function mCheckRecLength() As Integer
    'If Not gRecLengthOk("Sdf.Btr", Len(tmSdf)) Then
    '    mCheckRecLength = False
    '    Exit Function
    'End If
    If Not gRecLengthOk("Lcf.Btr", Len(tmLcf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    'If Not gRecLengthOk("Smf.Btr", Len(tmSmf)) Then
    '    mCheckRecLength = False
    '    Exit Function
    'End If
    If Not gRecLengthOk("Vef.Btr", Len(tmVef)) Then
        mCheckRecLength = False
        Exit Function
    End If
    'If Not gRecLengthOk("Vsf.Btr", Len(tmVsf)) Then
    '    mCheckRecLength = False
    '    Exit Function
    'End If
    mCheckRecLength = True
End Function

Private Sub grdVehicle_EnterCell()
    mSetShow
End Sub

Private Sub grdVehicle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lmTopRow = grdVehicle.TopRow
    grdVehicle.Redraw = False
End Sub

Private Sub grdVehicle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer

    bmIgnoreScroll = False
    If Y < grdVehicle.RowHeight(0) Then
        grdVehicle.Col = grdVehicle.MouseCol
        mSortCol grdVehicle.Col
        Exit Sub
    End If
    ilCol = grdVehicle.MouseCol
    ilRow = grdVehicle.MouseRow
    If ilCol < grdVehicle.FixedCols Then
        grdVehicle.Redraw = True
        Exit Sub
    End If
    If ilRow < grdVehicle.FixedRows Then
        grdVehicle.Redraw = True
        Exit Sub
    End If
    If grdVehicle.TextMatrix(ilRow, NAMEINDEX) = "" Then
        grdVehicle.Redraw = False
        Do
            ilRow = ilRow - 1
        Loop While grdVehicle.TextMatrix(ilRow, NAMEINDEX) = ""
        grdVehicle.Row = ilRow + 1
        grdVehicle.Col = STATUSINDEX
        grdVehicle.Redraw = True
    Else
        grdVehicle.Row = ilRow
        grdVehicle.Col = ilCol
    End If
    grdVehicle.Redraw = True
    lmTopRow = grdVehicle.TopRow
    If Not mColOk() Then
        Exit Sub
    End If
    mEnableBox
End Sub

Private Sub grdVehicle_Scroll()
    If bmIgnoreScroll Then  'Or igGridIgnoreScroll Then
        bmIgnoreScroll = False
        Exit Sub
    End If
    If grdVehicle.Redraw = False Then
        grdVehicle.Redraw = True
        If lmTopRow < grdVehicle.FixedRows Then
            grdVehicle.TopRow = grdVehicle.FixedRows
        Else
            grdVehicle.TopRow = lmTopRow
        End If
        grdVehicle.Refresh
        grdVehicle.Redraw = False
    End If
    If (bmCtrlVisible) And (grdVehicle.Row >= grdVehicle.FixedRows) And (grdVehicle.Col >= 0) And (grdVehicle.Col < grdVehicle.Cols - 1) Then
        If grdVehicle.RowIsVisible(grdVehicle.Row) Then
            'pbcArrow.Move grdVehicle.Left - pbcArrow.Width - 30, grdVehicle.Top + grdVehicle.RowPos(grdVehicle.Row) + (grdVehicle.RowHeight(grdVehicle.Row) - pbcArrow.Height) / 2
            'pbcArrow.Visible = True
            mSetFocus
        Else
            'pbcSetFocus.SetFocus
            pbcToggle.Visible = False
            edcDropDown.Visible = False
            cmcDropDown.Visible = False
            lbcAirVehicle.Visible = False
            'pbcArrow.Visible = False
        End If
    Else
        pbcClickFocus.SetFocus
    End If
End Sub

Private Sub lacDormant_Click()
    If ckcDormant.Value = vbChecked Then
        ckcDormant.Value = vbUnchecked
    Else
        ckcDormant.Value = vbChecked
    End If
End Sub

Private Sub lacMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move lacMin.Left - 30, lacMin.Top - 45
    imcOutline.Visible = True
End Sub

Private Sub lacPackage_Click()
    If ckcPackage.Value = vbChecked Then
        ckcPackage.Value = vbUnchecked
    Else
        ckcPackage.Value = vbChecked
    End If
End Sub

Private Sub lacStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imcOutline.Move lacStop.Left - 30, lacStop.Top - 45
    imcOutline.Visible = True
End Sub

Private Sub lbcAirVehicle_Click()
    mProcessLbcClick lbcAirVehicle, edcDropDown, imChgMode, imLbcArrowSetting
End Sub

Private Sub pbcClickFocus_Click()
    imcOutline.Visible = False
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If bmCtrlVisible Then
        mSetShow
        Do
            ilPrev = False
            If grdVehicle.Col = NAMEINDEX Then
                If grdVehicle.Row > grdVehicle.FixedRows Then
                    lmTopRow = -1
                    grdVehicle.Row = grdVehicle.Row - 1
                    If Not grdVehicle.RowIsVisible(grdVehicle.Row) Then
                        grdVehicle.TopRow = grdVehicle.TopRow - 1
                    End If
                    grdVehicle.Col = TYPEINDEX
                    mEnableBox
                Else
                    pbcClickFocus.SetFocus
                End If
            Else
                grdVehicle.Col = grdVehicle.Col - 1
                If mColOk() Then
                    mEnableBox
                Else
                    ilPrev = True
                End If
            End If
        Loop While ilPrev
    Else
        lmTopRow = -1
        grdVehicle.TopRow = grdVehicle.FixedRows
        grdVehicle.Col = STATUSINDEX
        grdVehicle.Row = grdVehicle.FixedRows
        If mColOk() Then
            mEnableBox
        Else
            pbcClickFocus.SetFocus
        End If
    End If
End Sub

Private Sub pbcTab_GotFocus()
    Dim llRow As Long
    Dim ilNext As Integer
    Dim llEnableRow As Long

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If bmCtrlVisible Then
        llEnableRow = lmEnableRow
        mSetShow
        Do
            ilNext = False
            If grdVehicle.Col = AIRNAMEINDEX Then
                llRow = grdVehicle.Rows
                Do
                    llRow = llRow - 1
                Loop While grdVehicle.TextMatrix(llRow, NAMEINDEX) = ""
                llRow = llRow + 1
                If (grdVehicle.Row + 1 < llRow) Then
                    lmTopRow = -1
                    grdVehicle.Row = grdVehicle.Row + 1
                    If Not grdVehicle.RowIsVisible(grdVehicle.Row) Then 'Or (grdVehicle.Row - (grdVehicle.TopRow - grdVehicle.FixedRows) >= imInitNoRows) Then
                        bmIgnoreScroll = True
                        grdVehicle.TopRow = grdVehicle.TopRow + 1
                    End If
                    grdVehicle.Col = STATUSINDEX
                    'grdVehicle.TextMatrix(grdVehicle.Row, CODEINDEX) = 0
                    If Trim$(grdVehicle.TextMatrix(grdVehicle.Row, NAMEINDEX)) <> "" Then
                        If mColOk() Then
                            mEnableBox
                        Else
                            pbcClickFocus.SetFocus
                        End If
                    Else
                        pbcClickFocus.SetFocus
                    End If
                Else
                    If Trim$(grdVehicle.TextMatrix(llEnableRow, NAMEINDEX)) <> "" Then
                        lmTopRow = -1
                        If grdVehicle.Row + 1 >= grdVehicle.Rows Then
                            pbcClickFocus.SetFocus
                            Exit Sub
                        End If
                        grdVehicle.Row = grdVehicle.Row + 1
                        If (Not grdVehicle.RowIsVisible(grdVehicle.Row)) Then   'Or (grdVehicle.Row - (grdVehicle.TopRow - grdVehicle.FixedRows) >= imInitNoRows) Then
                            bmIgnoreScroll = True
                            grdVehicle.TopRow = grdVehicle.TopRow + 1
                        End If
                        grdVehicle.Col = STATUSINDEX
                        If mColOk() Then
                            mEnableBox
                        Else
                            pbcClickFocus.SetFocus
                        End If
                    Else
                        pbcClickFocus.SetFocus
                    End If
                End If
            Else
                grdVehicle.Col = grdVehicle.Col + 1
                If mColOk() Then
                    mEnableBox
                Else
                    ilNext = True
                End If
            End If
        Loop While ilNext
    Else
        lmTopRow = -1
        grdVehicle.TopRow = grdVehicle.FixedRows
        grdVehicle.Col = STATUSINDEX
        grdVehicle.Row = grdVehicle.FixedRows
        If mColOk() Then
            mEnableBox
        Else
            pbcClickFocus.SetFocus
        End If
    End If
End Sub

Private Sub pbcToggle_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub pbcToggle_KeyPress(KeyAscii As Integer)
    If lmEnableCol = STATUSINDEX Then
        If KeyAscii = Asc("A") Or (KeyAscii = Asc("a")) Then
            grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Active"
            pbcToggle_Paint
        ElseIf KeyAscii = Asc("D") Or (KeyAscii = Asc("d")) Then
            grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Dormant"
            pbcToggle_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Active" Then
                grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Dormant"
                pbcToggle_Paint
            ElseIf grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Dormant" Then
                grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Active"
                pbcToggle_Paint
            End If
        End If
    ElseIf lmEnableCol = TYPEINDEX Then
        If KeyAscii = Asc("C") Or (KeyAscii = Asc("c")) Then
            grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Conventional"
            pbcToggle_Paint
        ElseIf KeyAscii = Asc("S") Or (KeyAscii = Asc("s")) Then
            grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Selling"
            pbcToggle_Paint
        End If
        If KeyAscii = Asc(" ") Then
            If grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Conventional" Then
                grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Selling"
                pbcToggle_Paint
            ElseIf grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Selling" Then
                grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Conventional"
                pbcToggle_Paint
            End If
        End If
    End If
    mSetCommands
End Sub

Private Sub pbcToggle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lmEnableCol = STATUSINDEX Then
        If grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Active" Then
            grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Dormant"
            pbcToggle_Paint
        ElseIf grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Dormant" Then
            grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Active"
            pbcToggle_Paint
        End If
    ElseIf lmEnableCol = TYPEINDEX Then
        If grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Selling" Then
            grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Conventional"
            pbcToggle_Paint
        ElseIf grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Conventional" Then
            grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Selling"
            pbcToggle_Paint
        End If
    End If
    pbcToggle_Paint
    mSetCommands
End Sub

Private Sub pbcToggle_Paint()
    pbcToggle.Cls
    pbcToggle.CurrentX = fgBoxInsetX
    pbcToggle.CurrentY = 0 'fgBoxInsetY
    pbcToggle.Print grdVehicle.TextMatrix(lmEnableRow, lmEnableCol)
End Sub

Private Sub plcSignon_Click()
    imcOutline.Visible = False
End Sub


Private Sub lacMin_Click()
    Dim ilRet As Integer
    If imAutoRun = 0 Then
        If Not imConverting Then
            mSetShow
            ilRet = mCheckSetting()
            If Not ilRet Then
                Exit Sub
            End If
            imConverting = True
            Screen.MousePointer = vbHourglass
            gSetMousePointer grdVehicle, grdVehicle, vbHourglass
            pbcClickFocus.SetFocus
            ilRet = mConvert()
            Screen.MousePointer = vbDefault
            gSetMousePointer grdVehicle, grdVehicle, vbDefault
            imConverting = False
            If ilRet Then
                lacStop.Caption = "Done"
                lacMin.Enabled = False
            ElseIf imTerminate = True Then
                lacStop.Caption = "Exit"
                lacMin.Enabled = False
            End If
        End If
    Else
        Me.WindowState = vbMinimized
    End If
End Sub

Private Sub lacStop_Click()
    If Not imConverting Then
        imTerminate = True
        mTerminate
    Else
        imTerminate = True
    End If
End Sub


Private Sub mBtrErrorMsg(slMsg As String, ilInRet As Integer)
    Dim ilRet As Integer
    
    ilRet = ilInRet
    If ilRet >= 30000 Then
        ilRet = csiHandleValue(0, 7)
    End If
    gLogMsg slMsg & Trim$(Str$(ilRet)), "Recreator.txt", False
End Sub



'*******************************************************
'*                                                     *
'*      Procedure Name:mVehPop                         *
'*                                                     *
'*             Created:8/17/05       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box for conventional           *
'*                      and selling vehicles           *
'*******************************************************
Private Sub mVehPop()
    Dim ilRet As Integer
    Dim ilVef As Integer
    Dim llRow As Long
    Dim llCol As Long
    
    grdVehicle.Redraw = False
    grdVehicle.Rows = 2
    llRow = grdVehicle.FixedRows
    For llCol = NAMEINDEX To VEFCODEINDEX Step 1
        grdVehicle.TextMatrix(llRow, llCol) = ""
    Next llCol
    
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If (tgMVef(ilVef).sType <> "P") And (tgMVef(ilVef).sType <> "A") Then
            If (tgMVef(ilVef).sState <> "D") Then
                mVehToGrid ilVef, llRow
            End If
        End If
    Next ilVef
    mSortCol NAMEINDEX
    grdVehicle.Row = 0
    grdVehicle.Col = VEFCODEINDEX
    grdVehicle.Redraw = True
    Exit Sub
mVehPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub



Private Sub ckcDormant_Click()
    Dim ilVef As Integer
    Dim llRow As Long
    Dim llLoop As Long
    
    grdVehicle.Redraw = False
    If ckcDormant.Value = vbChecked Then
        'Add Dormant
        llRow = grdVehicle.FixedRows
        For llLoop = grdVehicle.Rows - 1 To grdVehicle.FixedRows Step -1
            If grdVehicle.TextMatrix(llLoop, NAMEINDEX) <> "" Then
                llRow = llLoop + 1
                Exit For
            End If
        Next llLoop
        For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            If ((tgMVef(ilVef).sType <> "P") And (tgMVef(ilVef).sType <> "A")) Or ((tgMVef(ilVef).sType = "P") And (tgMVef(ilVef).lPvfCode > 0) And (ckcPackage = vbChecked)) Then
                If (tgMVef(ilVef).sState = "D") Then
                    mVehToGrid ilVef, llRow
                End If
            End If
        Next ilVef
        imLastColSorted = -1
        imLastSort = -1
        mSortCol NAMEINDEX
        grdVehicle.Row = 0
        grdVehicle.Col = VEFCODEINDEX
    Else
        'Remove dormant
        llRow = grdVehicle.Rows - 1
        Do While llRow >= grdVehicle.FixedRows
            If grdVehicle.TextMatrix(llRow, STATUSINDEX) = "Dormant" Then
                grdVehicle.RemoveItem llRow
            Else
                llRow = llRow - 1
            End If
        Loop
    End If
    grdVehicle.Redraw = True

    mSetCommands
End Sub


Private Sub cmcTo_Click()
    Dim slCurDir As String
    
    slCurDir = CurDir
    igPathType = 0
    sgGetPath = edcTo.Text
    lgCallTop = Recreator.Top
    lgCallLeft = Recreator.Left
    lgCallHeight = Recreator.Height
    lgCallWidth = Recreator.Width
    GetPath.Show vbModal
    If igGetPath = 0 Then
        edcTo.Text = sgGetPath
    End If
    
    ChDir slCurDir
    
    Exit Sub
End Sub

Private Sub edcTo_Change()
    mSetCommands
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
    Dim slTodayDate As String
    Dim slLastYearDate As String
    Dim slDay As String
    Dim slMonth As String
    Dim slYear As String
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim slEndDate As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVpf As Integer
    Dim ilVefCode As Integer
    Dim ilVehicle As Integer

    mParseCmmdLine
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If
    If imAutoRun = 0 Then
        lacMin.Caption = "Create"
    Else
        lacMin.Caption = "Minimize"
    End If
    imClosed = False
    imTerminate = False
    imFirstActivate = True
    imChgMode = False
    'mParseCmmdLine
    Screen.MousePointer = vbHourglass
    gSetMousePointer grdVehicle, grdVehicle, vbHourglass
    imLastColSorted = -1
    imLastSort = -1
    imConverting = False
    imBypassFocus = False
    lmNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    
    mSetGridColumns
    mSetGridTitles
    ilRet = gObtainVef() 'Build into tgMVef
    If ilRet = False Then
        imTerminate = True
    End If
    mVehPop
        
    slTodayDate = Format$(lmNowDate, "m/d/yy")
    gObtainYearMonthDayStr slTodayDate, True, slYear, slMonth, slDay
  
    edcFrom.Text = Left(sgDBPath, Len(sgDBPath) - 1)
    gSetMousePointer grdVehicle, grdVehicle, vbDefault

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
    Dim ilRet As Integer
  
    
    Screen.MousePointer = vbDefault
    gSetMousePointer grdVehicle, grdVehicle, vbDefault
    'igParentRestarted = False
    'If Not igStdAloneMode Then
    '    If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
    '        edcLinkDestHelpMsg.LinkExecute "@" & "Done"
    '    Else
    '        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
    '        edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
    '        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
    '        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
    '        edcLinkDestHelpMsg.LinkExecute "Done"
    '    End If
    '    Do While Not igParentRestarted
    '        DoEvents
    '    Loop
    'End If
    Screen.MousePointer = vbDefault
    gSetMousePointer grdVehicle, grdVehicle, vbDefault
    igManUnload = YES
    Unload Traffic
    Unload Recreator
    Set Recreator = Nothing   'Remve data segment
    igManUnload = NO
End Sub

Private Sub plcSignon_GotFocus()
    mSetShow
End Sub

Private Sub tmcClick_Timer()
    Dim slRepeat As String
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slDefaultFileName As String

    tmcClick.Enabled = False
    If imTerminate Then
        lacStop_Click
        Exit Sub
    End If
    
    Exit Sub
cmcExportDupNameErr:
    ilRet = 1
    Resume Next
End Sub

Public Sub mSetCommands()

'    If lbcVehicle.SelCount > 0 And edcTo.Text <> "" Then
'        'cmcExport.Enabled = True
'        lacMin.Enabled = True
'    Else
'        'cmcExport.Enabled = False
'        lacMin.Enabled = False
'    End If
    
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
    grdVehicle.ColWidth(VEFCODEINDEX) = 0
    grdVehicle.ColWidth(SORTINDEX) = 0
    grdVehicle.ColWidth(ORIGTYPEINDEX) = 0
    grdVehicle.ColWidth(STATUSINDEX) = grdVehicle.Width * 0.1
    grdVehicle.ColWidth(TYPEINDEX) = grdVehicle.Width * 0.2
    grdVehicle.ColWidth(AIRNAMEINDEX) = grdVehicle.Width * 0.25
    
    grdVehicle.ColWidth(NAMEINDEX) = grdVehicle.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To AIRNAMEINDEX Step 1
        If ilCol <> NAMEINDEX Then
            grdVehicle.ColWidth(NAMEINDEX) = grdVehicle.ColWidth(NAMEINDEX) - grdVehicle.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdVehicle
    

End Sub

Private Sub mSetGridTitles()
    Dim llCol As Long
    
    'Set column titles
    grdVehicle.TextMatrix(0, NAMEINDEX) = "Vehicle Name"
    grdVehicle.TextMatrix(0, STATUSINDEX) = "Status"
    grdVehicle.TextMatrix(0, TYPEINDEX) = "Type"
    grdVehicle.TextMatrix(0, AIRNAMEINDEX) = "Airing Vehicle"
    grdVehicle.Row = 0
    For llCol = NAMEINDEX To AIRNAMEINDEX Step 1
        grdVehicle.Col = llCol
        grdVehicle.CellBackColor = LIGHTBLUE
    Next llCol

End Sub

Private Sub mSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String
    Dim slDate As String
    Dim slTime As String
    Dim slDays As String
    Dim slHours As String
    Dim slMinutes As String
    Dim ilChar As Integer


    For llRow = grdVehicle.FixedRows To grdVehicle.Rows - 1 Step 1
        slStr = Trim$(grdVehicle.TextMatrix(llRow, NAMEINDEX))
        If slStr <> "" Then
            slSort = UCase$(Trim$(grdVehicle.TextMatrix(llRow, ilCol)))
            slStr = grdVehicle.TextMatrix(llRow, SORTINDEX)
            ilPos = InStr(1, slStr, "|", vbTextCompare)
            If ilPos > 1 Then
                slStr = Left$(slStr, ilPos - 1)
            End If
            If (ilCol <> imLastColSorted) Or ((ilCol = imLastColSorted) And (imLastSort = flexSortStringNoCaseDescending)) Then
                'Ascending
                If (ilCol = AIRNAMEINDEX) And (grdVehicle.TextMatrix(llRow, AIRNAMEINDEX) = "") Then
                    slSort = "zzzzzzzzzzzz"
                End If
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdVehicle.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            Else
                'Descending
                If (ilCol = AIRNAMEINDEX) And (grdVehicle.TextMatrix(llRow, AIRNAMEINDEX) = "") Then
                    slSort = "!!!!!!!!!!!!"
                End If
                slRow = Trim$(Str$(llRow))
                Do While Len(slRow) < 4
                    slRow = "0" & slRow
                Loop
                grdVehicle.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
            End If
        End If
    Next llRow
    If ilCol = imLastColSorted Then
        imLastColSorted = SORTINDEX
    Else
        imLastColSorted = -1
        imLastSort = -1
    End If
    gGrid_SortByCol grdVehicle, NAMEINDEX, SORTINDEX, imLastColSorted, imLastSort
    imLastColSorted = ilCol
End Sub

Private Sub mVehToGrid(ilVef As Integer, llRow As Long)

    If llRow >= grdVehicle.Rows Then
        grdVehicle.AddItem ""
    End If
    grdVehicle.Row = llRow
    grdVehicle.TextMatrix(llRow, NAMEINDEX) = Trim$(tgMVef(ilVef).sName)
    grdVehicle.Col = NAMEINDEX
    grdVehicle.CellBackColor = LIGHTYELLOW
    If tgMVef(ilVef).sState = "D" Then
        grdVehicle.TextMatrix(llRow, STATUSINDEX) = "Dormant"
    Else
        grdVehicle.TextMatrix(llRow, STATUSINDEX) = "Active"
    End If
    Select Case Trim$(tgMVef(ilVef).sType)
        Case "C"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Conventional"
        Case "S"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Selling"
        Case "A"
        Case "L"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Log"
        Case "V"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Virtual"
        Case "T"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Simul-cast"
        Case "P"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Package"
        Case "R"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Rep"
        Case "N"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "NTR"
        Case "G"
            grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Sport"
    End Select
    If Trim$(tgMVef(ilVef).sType) <> "S" Then
        grdVehicle.Col = TYPEINDEX
        grdVehicle.CellBackColor = LIGHTYELLOW
    End If
    grdVehicle.TextMatrix(llRow, ORIGTYPEINDEX) = Trim$(tgMVef(ilVef).sType)
    grdVehicle.TextMatrix(llRow, AIRNAMEINDEX) = ""
    grdVehicle.Col = AIRNAMEINDEX
    grdVehicle.CellBackColor = LIGHTYELLOW
    grdVehicle.TextMatrix(llRow, SORTINDEX) = ""
    grdVehicle.TextMatrix(llRow, VEFCODEINDEX) = tgMVef(ilVef).iCode
    llRow = llRow + 1

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mEnableBox                      *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Enable specified control       *
'*                                                     *
'*******************************************************
Private Sub mEnableBox()
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim slStr As String
    If (grdVehicle.Row < grdVehicle.FixedRows) Or (grdVehicle.Row >= grdVehicle.Rows) Or (grdVehicle.Col < grdVehicle.FixedCols) Or (grdVehicle.Col >= grdVehicle.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdVehicle.Row
    lmEnableCol = grdVehicle.Col
    bmCtrlVisible = True
    Select Case grdVehicle.Col
        Case NAMEINDEX
            'ignore
        Case STATUSINDEX
        Case TYPEINDEX
        Case AIRNAMEINDEX
            mPopAirVehicle
            edcDropDown.MaxLength = 40
            slStr = grdVehicle.Text
            If slStr = "Missing" Then
                slStr = ""
            End If
            If slStr = "[None]" Then
                lbcAirVehicle.ListIndex = 0
            Else
                edcDropDown.Text = slStr
            End If
            mSetListBoxHeight lbcAirVehicle, CLng(grdVehicle.Height / 2)
            imComboBoxIndex = lbcAirVehicle.ListIndex
    End Select
    mSetFocus
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer

    If (grdVehicle.Row < grdVehicle.FixedRows) Or (grdVehicle.Row >= grdVehicle.Rows) Or (grdVehicle.Col < grdVehicle.FixedCols) Or (grdVehicle.Col >= grdVehicle.Cols - 1) Then
        Exit Sub
    End If
    bmCtrlVisible = True
    llColPos = 0
    For ilCol = 0 To grdVehicle.Col - 1 Step 1
        llColPos = llColPos + grdVehicle.ColWidth(ilCol)
    Next ilCol
    Select Case grdVehicle.Col
        Case NAMEINDEX
        Case STATUSINDEX
            pbcToggle.Move grdVehicle.Left + llColPos + 30, grdVehicle.Top + grdVehicle.RowPos(grdVehicle.Row) + 30, grdVehicle.ColWidth(grdVehicle.Col) - 30, grdVehicle.RowHeight(grdVehicle.Row) - 15
            pbcToggle.Visible = True
            pbcToggle.SetFocus
        Case TYPEINDEX
            pbcToggle.Move grdVehicle.Left + llColPos + 30, grdVehicle.Top + grdVehicle.RowPos(grdVehicle.Row) + 30, grdVehicle.ColWidth(grdVehicle.Col) - 30, grdVehicle.RowHeight(grdVehicle.Row) - 15
            pbcToggle.Visible = True
            pbcToggle.SetFocus
        Case AIRNAMEINDEX
            edcDropDown.Move grdVehicle.Left + llColPos + 30, grdVehicle.Top + grdVehicle.RowPos(grdVehicle.Row) + 30, grdVehicle.ColWidth(grdVehicle.Col) - 30, grdVehicle.RowHeight(grdVehicle.Row) - 15
            edcDropDown.Visible = True
            edcDropDown.SetFocus
            
            edcDropDown.Move grdVehicle.Left + llColPos + 30, grdVehicle.Top + grdVehicle.RowPos(grdVehicle.Row) + 15, grdVehicle.ColWidth(grdVehicle.Col) - 30, grdVehicle.RowHeight(grdVehicle.Row) - 15
            'edcDropDown.Width = gSetCtrlWidth("Title1", lmCharacterWidth, edcDropDown.Width, 6)
            edcDropDown.Left = grdVehicle.Left + llColPos + grdVehicle.ColWidth(grdVehicle.Col) - edcDropDown.Width - cmcDropDown.Width
            edcDropDown.MaxLength = 40
            cmcDropDown.Move edcDropDown.Left + edcDropDown.Width, edcDropDown.Top, cmcDropDown.Width, edcDropDown.Height
            lbcAirVehicle.Move edcDropDown.Left, edcDropDown.Top + edcDropDown.Height, edcDropDown.Width + cmcDropDown.Width
            'If lbcCTE_1.Top + lbcCTE_1.Height > cmcCancel.Top Then
            '    lbcCTE_1.Top = edcDropDown.Top - lbcCTE_1.Height
            'End If
            edcDropDown.Visible = True
            cmcDropDown.Visible = True
            lbcAirVehicle.Visible = True
            edcDropDown.SetFocus

    End Select
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow()

'
'   mSetShow ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control whose value should be saved
'
    Dim slStr As String
    Dim llSvRow As Long
    Dim llSvCol As Long

    If (lmEnableRow >= grdVehicle.FixedRows) And (lmEnableRow < grdVehicle.Rows) Then
        'Set any field that that should only be set after user leaves the cell
        Select Case lmEnableCol
            Case NAMEINDEX
            Case STATUSINDEX
            Case TYPEINDEX
                llSvRow = grdVehicle.Row
                llSvCol = grdVehicle.Col
                grdVehicle.Row = lmEnableRow
                grdVehicle.Col = AIRNAMEINDEX
                If grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = "Conventional" Then
                    grdVehicle.CellBackColor = vbWhite
                Else
                    grdVehicle.CellBackColor = LIGHTYELLOW
                    grdVehicle.Text = ""
                End If
                grdVehicle.Row = llSvRow
                grdVehicle.Col = llSvCol
            Case AIRNAMEINDEX
                slStr = edcDropDown.Text
                grdVehicle.TextMatrix(lmEnableRow, lmEnableCol) = slStr
        End Select
    End If
    pbcToggle.Visible = False
    edcDropDown.Visible = False
    cmcDropDown.Visible = False
    lbcAirVehicle.Visible = False
    lmEnableRow = -1
    lmEnableCol = -1
    bmCtrlVisible = False
    mSetCommands
End Sub

Private Function mColOk() As Integer
    mColOk = True
    If grdVehicle.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
End Function


Private Sub mPopAirVehicle()
    Dim ilVefCode As Integer
    Dim ilVpf As Integer
    Dim ilVef As Integer
    Dim ilSAGroupNo As Integer
    
    lbcAirVehicle.Clear
    ilVefCode = Val(grdVehicle.TextMatrix(lmEnableRow, VEFCODEINDEX))
    ilVpf = gBinarySearchVpf(ilVefCode)
    If ilVpf <= 0 Then
        lbcAirVehicle.AddItem "[None]", 0
        lbcAirVehicle.ItemData(lbcAirVehicle.NewIndex) = 0
        Exit Sub
    End If
    ilSAGroupNo = tgVpf(ilVpf).iSAGroupNo
    For ilVpf = LBound(tgVpf) To UBound(tgVpf) - 1 Step 1
        If (tgVpf(ilVpf).iVefKCode <> ilVefCode) And (tgVpf(ilVpf).iSAGroupNo = ilSAGroupNo) Then
            ilVef = gBinarySearchVef(tgVpf(ilVpf).iVefKCode)
            If ilVef <> -1 Then
                If tgMVef(ilVef).sType = "A" Then
                    lbcAirVehicle.AddItem Trim$(tgMVef(ilVef).sName)
                    lbcAirVehicle.ItemData(lbcAirVehicle.NewIndex) = tgVpf(ilVpf).iVefKCode
                End If
            End If
        End If
    Next ilVpf
    lbcAirVehicle.AddItem "[None]", 0
    lbcAirVehicle.ItemData(lbcAirVehicle.NewIndex) = 0
End Sub

Sub mSetListBoxHeight(lbcCtrl As ListBox, llMaxHeight As Long)
'
'  flHeight = gListBoxHeight (ilNoRows, ilMaxRows)
'   Where:
'       ilNoRows (I) - current number of items within the list box
'       ilMaxRows (I) - max number of list box items to be displayed
'       flHeight (O) - height of list box in twips
'
    '+30 because of line above and below
    Dim llRowHeight As Long
    Dim slStr As String
    Dim ilMaxRow As Integer
    
    If lbcCtrl.ListCount > 0 Then
        'Determine standard height, set to small number so that only one row
        'height will be set (15 + 15 + RowHeight; 15 for size of boundaries)
        'Typical result is 300 (15 + 15 + 270)
'        lbcCtrl.Height = 10
        llRowHeight = 15 * SendMessageByString(lbcCtrl.hwnd, LB_GETITEMHEIGHT, 0, slStr)
        ilMaxRow = llMaxHeight / llRowHeight
        If lbcCtrl.ListCount <= ilMaxRow Then
'            lbcCtrl.Height = (lbcCtrl.Height - 30) * lbcCtrl.ListCount + 30 '375 + 255 * (ilNoRows - 1)
            lbcCtrl.Height = (llRowHeight) * lbcCtrl.ListCount + 30 '375 + 255 * (ilNoRows - 1)
        Else
'            lbcCtrl.Height = (lbcCtrl.Height - 30) * ilMaxRow + 30 '375 + 255 * (ilMaxRow - 1)
            lbcCtrl.Height = (llRowHeight) * ilMaxRow + 30 '375 + 255 * (ilMaxRow - 1)
        End If
    End If
End Sub

Sub mMatchLookAhead(edcTextBox As TextBox, lbcListBox As ListBox, ilBSMode As Integer, ilErrHighLightIndex As Integer)
'
'   gMatchLookAhead edcText, lbcCtrl, ilBSMode, ilHighlightIndex
'   Where:
'       edcText (I)- Text box control (containing input to be validated)
'       lbcCtrl (I)- List box control containing values to be matched
'       ilBSMode (I/O)- Backspace flag(True = backspace key was pressed, False =                        '       backspace key was not pressed)
'       ilHighlightIndex (I)- Selection to be highlighted if input is invalid
'

    Dim ilLen As Integer    'Length of current enter text
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim ilSelStart As Integer
    Dim ilBracket As Integer
    Dim ilSearch As Integer
    Dim ilSvLastFound As Integer
    Dim ilLoop As Integer
    Dim ilPos As Integer

    ilSelStart = edcTextBox.SelStart
    slStr = LTrim$(edcTextBox.Text)    'Remove leading blanks only
    ilLen = Len(edcTextBox.Text)
    ilIndex = lbcListBox.ListIndex
    If slStr = "" Then  'If space bar selected, text will be blank- ListIndex will contain a value
        lbcListBox.ListIndex = -1
        Exit Sub
        If lbcListBox.ListIndex >= 0 Then
            slStr = lbcListBox.List(lbcListBox.ListIndex)
            ilLen = 0
            ilIndex = -1    'Force dispaly of selected item by space bar
        Else
            Beep
            If ilErrHighLightIndex >= 0 Then
                lbcListBox.ListIndex = ilErrHighLightIndex
            End If
            Exit Sub
        End If
    End If
    If ilBSMode Then    'If backspace mode- reduce string by one character
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        ilBSMode = False
        If ilSelStart > 0 Then
            ilSelStart = ilSelStart - 1
        End If
    End If
    If Left$(slStr, 1) = "[" Then   'Search does not work when starting with [
        ilSvLastFound = -1
        For ilLoop = 0 To lbcListBox.ListCount - 1 Step 1
            ilPos = InStr(1, lbcListBox.List(ilLoop), slStr, 1)
            If ilPos = 1 Then
                ilSvLastFound = ilLoop
                Exit For
            Else
                If Left$(lbcListBox.List(ilLoop), 1) <> "[" Then
                    Exit For
                End If
            End If
        Next ilLoop
    Else
        'Test if matching string found in the combo box- if so display it (look ahead typing)
        'lbcListBox.ListIndex = 0
        gFndFirst lbcListBox, slStr
        ilBracket = False
        Do
            If gLastFound(lbcListBox) >= 0 Then
                If (Left$(lbcListBox.List(gLastFound(lbcListBox)), 1) = "[") And (Left$(slStr, 1) <> "[") Then
                    gFndNext lbcListBox, slStr
                    ilBracket = True
                Else
                    ilBracket = False
                End If
            Else
                ilBracket = False
            End If
        Loop While ilBracket
        'Test if another name matches encase names are not in sorted order
        ilSearch = True
        ilSvLastFound = gLastFound(lbcListBox)
        Do
            If gLastFound(lbcListBox) >= 0 Then
                If StrComp(slStr, lbcListBox.List(gLastFound(lbcListBox)), 1) = 0 Then
                    ilSvLastFound = gLastFound(lbcListBox)   'lbcListBox.LastFound
                    Exit Do
                End If
                gFndNext lbcListBox, slStr
            Else
                Exit Do
            End If
        Loop While ilSearch
    End If
    If ilSvLastFound >= 0 Then
        'If item found not same as current selected- change current
        If (ilIndex <> ilSvLastFound) Or ((ilIndex = ilSvLastFound) And Not ilBSMode) Then
            lbcListBox.ListIndex = ilSvLastFound 'This will cause a change event (reason for imChgMode)
        End If
'        If (ilIndex <> lbcListBox.LastFound) Or ((ilIndex = lbcListBox.LastFound) And Not ilBSMode) Then 'If indices not equal- highlight look ahead text
'            lbcListBox.SelStart = ilLen
'            lbcListBox.SelLength = Len(lbcListBox.Text)
'        End If
        ilErrHighLightIndex = ilSvLastFound
    Else
        Beep
        If (ilErrHighLightIndex >= 0) And (ilErrHighLightIndex < lbcListBox.ListCount) Then
            lbcListBox.ListIndex = ilErrHighLightIndex
            ilSelStart = 0
        End If
    End If
    If lbcListBox.ListIndex >= 0 Then
        edcTextBox.Text = lbcListBox.List(lbcListBox.ListIndex)
    Else
        edcTextBox.Text = lbcListBox.Text
    End If
    If ilSelStart <= Len(edcTextBox.Text) Then
        edcTextBox.SelStart = ilSelStart
        edcTextBox.SelLength = Len(edcTextBox.Text)
    Else
        edcTextBox.SelStart = 0
        edcTextBox.SelLength = Len(edcTextBox.Text)
    End If
End Sub

Sub mProcessArrowKey(ilShift As Integer, ilKeyCode As Integer, lbcCtrl As Control, ilLbcArrowSetting As Integer)
'
'   gProcessArrowKey Shift, KeyCode, lbcCtrl, imLbcArrowSetting
'   Where:
'       Shift (I)- Shift key state
'       KeyCode (I)- Key code
'       lbcCtrl (I)- list box control
'       ilLbcArrowSetting (I/O) - list box arrow setting flag
'                               True= make list box invisible (user click on item)
'                               False= retain list box visible state
'

    Dim ilLoop As Integer

    If (ilShift And ALTMASK) > 0 Then
        lbcCtrl.Visible = Not lbcCtrl.Visible
    ElseIf (ilShift And SHIFTMASK) > 0 Then
    Else
        ilLbcArrowSetting = True
        If ilKeyCode = KEYUP Then    'Up arrow
            If lbcCtrl.ListIndex > 0 Then
                lbcCtrl.ListIndex = lbcCtrl.ListIndex - 1
                If lbcCtrl.MultiSelect Then
                    For ilLoop = 0 To lbcCtrl.ListCount - 1 Step 1
                        If ilLoop = lbcCtrl.ListIndex Then
                            lbcCtrl.Selected(lbcCtrl.ListIndex) = True
                        Else
                            lbcCtrl.Selected(ilLoop) = False
                        End If
                    Next ilLoop
                End If
            End If
        Else
            If lbcCtrl.ListIndex < lbcCtrl.ListCount - 1 Then
                lbcCtrl.ListIndex = lbcCtrl.ListIndex + 1
                If lbcCtrl.MultiSelect Then
                    For ilLoop = 0 To lbcCtrl.ListCount - 1 Step 1
                        If ilLoop = lbcCtrl.ListIndex Then
                            lbcCtrl.Selected(lbcCtrl.ListIndex) = True
                        Else
                            lbcCtrl.Selected(ilLoop) = False
                        End If
                    Next ilLoop
                End If
            End If
        End If
        ilLbcArrowSetting = False
    End If
End Sub
Sub mProcessLbcClick(lbcCtrl As Control, edcCtrl As Control, ilChgMode As Integer, ilLbcArrowSetting As Integer)
'
'   gProcessLbcClick lbcSPerson, edcDropDown, imChgMode,                                                imLbcArrowSetting
'   Where:
'       lbcSPerson (I)- list box control
'       edcDropDown (I)- User edit control
'       imChgMode (I/O)- Change mode flag
'       imLbcArrowSetting (I/O) - list box arrow setting flag
'                               True= make list box invisible (user click on item)
'                               False= retain list box visible state
'

    Dim ilLbcArrowSet As Integer

    If ilChgMode = False Then
        ilChgMode = True
        ilLbcArrowSet = ilLbcArrowSetting
        edcCtrl.Text = lbcCtrl.List(lbcCtrl.ListIndex)
        If (edcCtrl.Visible) And (edcCtrl.Enabled) Then
            edcCtrl.SetFocus
        End If
        If Not ilLbcArrowSet Then
            lbcCtrl.Visible = False
        End If
        ilChgMode = False
    End If
    ilLbcArrowSetting = False
End Sub


Private Function mConvert() As Integer
    Dim llRow As Long
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVpfSell As Integer
    Dim ilVpfAir As Integer
    Dim llSvTotalRecords As Long
    Dim llSvRecordNo As Long
    
    
    ReDim tmRecreateSort(0 To 0) As RECREATESORT
    For llRow = grdVehicle.FixedRows To grdVehicle.Rows - 1 Step 1
        If grdVehicle.TextMatrix(llRow, NAMEINDEX) <> "" Then
            If grdVehicle.TextMatrix(llRow, STATUSINDEX) = "Active" Then
                slStr = grdVehicle.TextMatrix(llRow, VEFCODEINDEX)
                Do While Len(slStr) < 5
                    slStr = "0" & slStr
                Loop
                tmRecreateSort(UBound(tmRecreateSort)).sKey = slStr
                tmRecreateSort(UBound(tmRecreateSort)).iVefCode = Val(slStr)
                tmRecreateSort(UBound(tmRecreateSort)).bSellToConv = False
                tmRecreateSort(UBound(tmRecreateSort)).iAirAgreeVefCode = 0
                If (grdVehicle.TextMatrix(llRow, ORIGTYPEINDEX) = "S") And (grdVehicle.TextMatrix(llRow, TYPEINDEX) <> "Selling") Then
                    tmRecreateSort(UBound(tmRecreateSort)).bSellToConv = True
                    slStr = UCase(Trim$(grdVehicle.TextMatrix(llRow, AIRNAMEINDEX)))
                    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        If (tgMVef(ilLoop).sType = "A") Then
                            If slStr = UCase(Trim$(tgMVef(ilLoop).sName)) Then
                                tmRecreateSort(UBound(tmRecreateSort)).iAirAgreeVefCode = tgMVef(ilLoop).iCode
                                Exit For
                            End If
                        End If
                    Next ilLoop
                End If
                ReDim Preserve tmRecreateSort(0 To UBound(tmRecreateSort) + 1) As RECREATESORT
                If grdVehicle.TextMatrix(llRow, TYPEINDEX) = "Selling" Then
                    'Add Airing
                    ilVpfSell = gBinarySearchVpf(Val(grdVehicle.TextMatrix(llRow, VEFCODEINDEX)))
                    If ilVpfSell <> -1 Then
                        For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            If (tgMVef(ilLoop).sType = "A") And (tgMVef(ilLoop).sState <> "D") Then
                                ilVpfAir = gBinarySearchVpf(tgMVef(ilLoop).iCode)
                                If ilVpfAir <> -1 Then
                                    If tgVpf(ilVpfSell).iSAGroupNo = tgVpf(ilVpfAir).iSAGroupNo Then
                                        slStr = Trim$(Str$(tgVpf(ilVpfAir).iVefKCode))
                                        Do While Len(slStr) < 5
                                            slStr = "0" & slStr
                                        Loop
                                        tmRecreateSort(UBound(tmRecreateSort)).sKey = slStr
                                        tmRecreateSort(UBound(tmRecreateSort)).iVefCode = Val(slStr)
                                        tmRecreateSort(UBound(tmRecreateSort)).bSellToConv = False
                                        tmRecreateSort(UBound(tmRecreateSort)).iAirAgreeVefCode = 0
                                        ReDim Preserve tmRecreateSort(0 To UBound(tmRecreateSort) + 1) As RECREATESORT
                                    End If
                                End If
                            End If
                        Next ilLoop
                    End If
                End If
            End If
        End If
    Next llRow
    If UBound(tmRecreateSort) - 1 > 0 Then
        ArraySortTyp fnAV(tmRecreateSort(), 0), UBound(tmRecreateSort), 0, LenB(tmRecreateSort(0)), 0, LenB(tmRecreateSort(0).sKey), 0
    End If
    imFileNo = 0
    imTotalFiles = 73
    ReDim imDnfCode(0 To 0) As Integer
    ReDim lmCefCode(0 To 0) As Long
    ReDim imArfCode(0 To 0) As Integer
    ReDim lmPvfCode(0 To 0) As Long
    ReDim tmLcfDateInfo(0 To 0) As LCFDATEINFO
    ReDim lmLvfCode(0 To 0) As Long
    ReDim tmLibTitleCount(0 To 0) As LIBTITLECOUNT
    ReDim imLtfCode(0 To 0) As Integer
    ReDim imIhfCode(0 To 0) As Integer
    ReDim imItfCode(0 To 0) As Integer
    ReDim imIifCode(0 To 0) As Integer
    ReDim lmGhfCode(0 To 0) As Long
    ReDim imEtfCode(0 To 0) As Integer
    ReDim imCheckDnfCode(0 To 0) As Integer
    'ReDim lmDrfCode(0 To 0) As Long
    ReDim imRdfCode(0 To 0) As Integer
    ReDim imAdfCode(0 To 0) As Integer
    ReDim imSlfCode(0 To 0) As Integer
    ReDim imAgfCode(0 To 0) As Integer
    ReDim imArfCode(0 To 0) As Integer
    ReDim imTrfCode(0 To 0) As Integer
    ReDim imPnfCode(0 To 0) As Integer
    ReDim imSofCode(0 To 0) As Integer
    ReDim lmCxfCode(0 To 0) As Long
    ReDim imUrfCode(0 To 0) As Integer
    ReDim imSnfCode(0 To 0) As Integer
    ReDim lmAttCode(0 To 0) As Long
    ReDim imCstCode(0 To 0) As Integer
    ReDim imUstCode(0 To 0) As Integer
    ReDim lmRhtCode(0 To 0) As Long
    imPassNo = 1
    lmAdjCount = 1
    
    hmPvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPvf, "", smFromDBPath & "PVF.BTR", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imPvfRecLen = Len(tmPvf)
    ilRet = mCopyRecords("VEF.BTR", INDEXKEY0)
    ilRet = btrClose(hmPvf)
    btrDestroy hmPvf
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("VAF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("VBF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("VFF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("VOF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("VPF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("PIF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("PVF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    imPassNo = 1
    lmAdjCount = 2
    ilRet = mCopyRecords("LCF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    'Convert date to Monday if prior to lmPAStartdate
    For llRow = 0 To UBound(tmLcfDateInfo) - 1 Step 1
        If tmLcfDateInfo(llRow).lLatestDate < lmPAStartDate Then
            Do While gWeekDayLong(tmLcfDateInfo(llRow).lLatestDate) <> 0
                tmLcfDateInfo(llRow).lLatestDate = tmLcfDateInfo(llRow).lLatestDate - 1
            Loop
        End If
    Next llRow
    imPassNo = 2
    ilRet = mCopyRecords("LCF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    imPassNo = 1
    lmAdjCount = 1
    ilRet = mCopyRecords("PAF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("PFF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("VCF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("VLF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("IHF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("IIF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("ITF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("GHF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("GSF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    'Move here as both LCF and GSF add to lmLvfCode
    imPassNo = 1
    lmAdjCount = 2
    ilRet = mCopyRecords("LVF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    lmAdjCount = 1
    imPassNo = 2
    ilRet = mCopyRecords("LVF.BTR", INDEXKEY1)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    imPassNo = 1
    ilRet = mCopyRecords("LTF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("LEF.BTR", INDEXKEY1)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mCopyRecords("ENF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("ANF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("ETF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    imPassNo = 1
    lmAdjCount = 2
    ilRet = mCopyRecords("DNF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    llSvTotalRecords = lmTotalRecords
    llSvRecordNo = lmRecordNo
    lmAdjCount = 1
    If Not mOpenFile(hmFromDpf, hmToDpf, "DPF.BTR") Then
        mConvert = False
        Exit Function
    End If
    imDpfRecLen = Len(tmDpf)
    ilRet = mCopyRecords("DRF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCloseFile(hmFromDpf, hmToDpf)
    lmTotalRecords = llSvTotalRecords
    lmRecordNo = llSvRecordNo
    imPassNo = 2
    ilRet = mCopyRecords("DNF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    imPassNo = 1
    ilRet = mCopyRecords("DEF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    'ilRet = mCopyRecords("DPF.BTR", INDEXKEY0)
    'If imTerminate Then
    '    mConvert = False
    '    Exit Function
    'End If
    ilRet = mCopyRecords("RDF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    
    ilRet = mCopyRecords("SPF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("SAF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mCopyRecords("ADF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("AGF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("MNF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("PRF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("RNF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("RTF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("ARF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("TRF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("PNF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("SOF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("CMF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    'ilRet = mCopyRecords("URF.BTR", INDEXKEY0)
    ilRet = mUrfCopy()
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("SLF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("SNF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("SRF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mCopyRecords("ATT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("DAT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("EPT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("PFT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mCopyRecords("SHTT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("ARTT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("CLT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("TNT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mCopyRecords("CMT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("DFT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("FHT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("FIT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("FLT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("FMT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("MAT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("MET.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("MKT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("MNT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("SITE.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("SNT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("TZT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mCopyRecords("CCT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("CST.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mCopyRecords("RHT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("RET.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mCopyRecords("UST.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("DNT.MKD", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    
    ilRet = mCopyRecords("CXF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    ilRet = mCopyRecords("CEF.BTR", INDEXKEY0)
    If imTerminate Then
        mConvert = False
        Exit Function
    End If
    lacInfo(1).Caption = "ReCreator Completed"
    lacInfo(0).Caption = ""
    mConvert = True
End Function

Private Function mCheckSetting() As Integer
    Dim slStr As String
    Dim ilRet As Integer
    mCheckSetting = True
    smFromDBPath = Trim$(edcFrom.Text)
    If smFromDBPath = "" Then
        mCheckSetting = False
        MsgBox "Please defined the From Database location", vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    smFromDBPath = gSetPathEndSlash(smFromDBPath, False)
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFrom, "", smFromDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCheckSetting = False
        MsgBox "Check From Database path as Pervasive Error " & ilRet, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    imVefRecLen = Len(tmVef)
    ilRet = btrGetFirst(hmFrom, tmVef, imVefRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        mCheckSetting = False
        MsgBox "Check From Database path as Pervasive Error " & ilRet, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    ilRet = btrClose(hmFrom)
    btrDestroy hmFrom
    
    smToDBPath = Trim$(edcTo.Text)
    If smToDBPath = "" Then
        mCheckSetting = False
        MsgBox "Please defined the To Database location", vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    smToDBPath = gSetPathEndSlash(smToDBPath, False)
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTo, "", smToDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCheckSetting = False
        MsgBox "Check To Database path as Pervasive Error " & ilRet, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    imVefRecLen = Len(tmVef)
    ilRet = btrGetFirst(hmTo, tmVef, imVefRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_END_OF_FILE) Then
        mCheckSetting = False
        MsgBox "Check To Database path as Pervasive Error " & ilRet, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    ilRet = btrClose(hmTo)
    btrDestroy hmTo
    
    slStr = Trim$(cccPA.Text)
    If slStr = "" Then
        mCheckSetting = False
        MsgBox "Please defined the Program and Agreements Date", vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    If Not gIsDate(slStr) Then
        mCheckSetting = False
        MsgBox "Please check the Program and Agreements Date", vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    lmPAStartDate = gDateValue(slStr)
    slStr = Trim$(cccResearch.Text)
    If slStr = "" Then
        mCheckSetting = False
        MsgBox "Please defined the Research Date", vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    If Not gIsDate(slStr) Then
        mCheckSetting = False
        MsgBox "Please check the Research Date", vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    lmResearchStartDate = gDateValue(slStr)
End Function

Public Function mBinarySearchRecreateSort(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tmRecreateSort)
    ilMax = UBound(tmRecreateSort) - 1
    Do While ilMin <= ilMax
        ilMiddle = (CLng(ilMin) + ilMax) \ 2
        If ilCode = tmRecreateSort(ilMiddle).iVefCode Then
            'found the match
            mBinarySearchRecreateSort = ilMiddle
            Exit Function
        ElseIf ilCode < tmRecreateSort(ilMiddle).iVefCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchRecreateSort = -1
End Function

Private Function mCopyRecords(slFileName As String, ilAutoCodeKeyNo As Integer) As Integer
    Dim ilRecLen As Integer
    Dim llRecPos As Long
    Dim ilRet As Integer
    Dim llNoRec As Long
    Dim ilPos As Integer
    Dim slName As String
    Dim ilInsertRecord As Integer
    
    If imPassNo = 1 Then
        imFileNo = imFileNo + 1
        lacInfo(0).Caption = "File: " & imFileNo & " of " & imTotalFiles
    End If
    lacInfo(1).Caption = slFileName
    DoEvents
    ilPos = InStr(1, slFileName, ".", vbTextCompare)
    If ilPos <= 0 Then
        mCopyRecords = False
        MsgBox "File Name to convert in Error, Missing extension " & slFileName, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    slName = Left(slFileName, ilPos - 1)
    hmFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFrom, "", smFromDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
            ilRet = csiHandleValue(0, 7)
        End If
        mCopyRecords = False
        MsgBox "Check From Database path as Pervasive Error " & ilRet, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    hmTo = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTo, "", smToDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
            ilRet = csiHandleValue(0, 7)
        End If
        mCopyRecords = False
        MsgBox "Check To Database path as Pervasive Error " & ilRet, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    
    ilRecLen = btrRecordLength(hmFrom)  'Get and save record length
    If ilRecLen > Len(tmRecBuff) Then
        'Error message
        mCopyRecords = False
        MsgBox "Increase General Buffer size to " & ilRecLen, vbInformation + vbOKOnly, "Counterpoint"
        Exit Function
    End If
    If imPassNo = 1 Then
        lmTotalRecords = lmAdjCount * btrRecords(hmFrom)
        lmRecordNo = 0
    End If
    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlFile) 'Obtain number of records
    btrExtClear hmFrom   'Clear any previous extend operation
    ilRet = btrGetFirst(hmFrom, tmRecBuff, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hmFrom)
        btrDestroy hmFrom
        ilRet = btrClose(hmTo)
        btrDestroy hmTo
        Exit Function
    ElseIf ilRet <> BTRV_ERR_NONE Then
        If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
            ilRet = csiHandleValue(0, 7)
        End If
        mCopyRecords = False
        MsgBox "Get First Failed for " & slFileName & " with Error " & ilRet, vbInformation + vbOKOnly, "Counterpoint"
        ilRet = btrClose(hmFrom)
        btrDestroy hmFrom
        ilRet = btrClose(hmTo)
        btrDestroy hmTo
        Exit Function
    End If

    Call btrExtSetBounds(hmFrom, llNoRec, -1, "UC", slName, "") 'Set extract limits (all records)

    ilRet = btrExtAddField(hmFrom, 0, ilRecLen)  'Extract iCode field
    ''ilRet = btrExtGetNextExt(hlFile)  'Extract records
    'ilRet = btrExtGetNext(hmFrom, tmRecBuff, ilRecLen, llRecPos)    'Extract recor
    ilRet = mExtGetNext(slFileName, ilRecLen)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        ilRecLen = btrRecordLength(hmFrom)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlFile, tlCodeNameBuff, ilLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            'ilRet = btrExtGetNext(hmFrom, tmRecBuff, ilRecLen, llRecPos)
            ilRet = mExtGetNext(slFileName, ilRecLen)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If imTerminate Then
                mCopyRecords = False
                Exit Function
            End If
            lmRecordNo = lmRecordNo + 1
            lacInfo(1).Caption = slFileName & ": " & lmRecordNo & " of " & lmTotalRecords
            DoEvents
            ilInsertRecord = mTestRecord(slFileName)
            If ilInsertRecord Then
                'ilRet = btrInsert(hmTo, tmRecBuff, ilRecLen, ilAutoCodeKeyNo)
                ilRet = mInsert(slFileName, ilRecLen, ilAutoCodeKeyNo)
                If ilRet <> BTRV_ERR_NONE Then
                    If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mCopyRecords = False
                    MsgBox "Insert Failed for " & slFileName & " with Error " & ilRet, vbInformation + vbOKOnly, "Counterpoint"
                    ilRet = btrClose(hmFrom)
                    btrDestroy hmFrom
                    ilRet = btrClose(hmTo)
                    btrDestroy hmTo
                    Exit Function
                End If
            End If
            'ilRet = btrExtGetNext(hmFrom, tmRecBuff, ilRecLen, llRecPos)
            ilRet = mExtGetNext(slFileName, ilRecLen)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                'ilRet = btrExtGetNext(hmFrom, tmRecBuff, ilRecLen, llRecPos)
                ilRet = mExtGetNext(slFileName, ilRecLen)
            Loop
        Loop
    End If
    ilRet = btrClose(hmFrom)
    btrDestroy hmFrom
    ilRet = btrClose(hmTo)
    btrDestroy hmTo
    mCopyRecords = True
End Function

Private Function mTestRecord(slFileName As String) As Integer
    Dim ilVef As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llLogDate As Long
    Dim ilLoop As Integer
    Dim llTest As Long
    Dim blFound As Boolean
    Dim llIndex As Long
    Dim ilRet As Integer
    Dim llOffAirDate As Long
    Dim llDropDate As Long
    Dim llPvfCode As Long
    Dim llRow As Long
    
    mTestRecord = True
    If UCase(slFileName) = "VEF.BTR" Then
        'tmVef = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmVef.iCode)
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            If tmVef.sType = "S" Then
                If tmRecreateSort(ilVef).bSellToConv = True Then
                    tmVef.sType = "C"
                End If
            End If
            tmVef.sState = "A"
            mAddInteger tmVef.iDnfCode, imDnfCode()
            mAddInteger tmVef.iReallDnfCode, imDnfCode()
            If (tmVef.sType = "P") And (tmVef.lPvfCode > 0) Then
                llPvfCode = tmVef.lPvfCode
                mAddLong llPvfCode, lmPvfCode()
                Do While llPvfCode > 0
                    tmPvfSrchKey0.lCode = llPvfCode
                    ilRet = btrGetEqual(hmPvf, tmPvf, imPvfRecLen, tmPvfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    mAddLong tmPvf.lCode, lmPvfCode()
                    llPvfCode = tmPvf.lLkPvfCode
                Loop
            End If
        End If
        Exit Function
    End If
    If UCase(slFileName) = "VBF.BTR" Then
        'tmVbf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmVbf.iVefCode)
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            gUnpackDateLong tmVbf.iStartDate(0), tmVbf.iStartDate(1), llStartDate
            gUnpackDateLong tmVbf.iEndDate(0), tmVbf.iEndDate(1), llEndDate
            If (llEndDate = 0) Or (llEndDate >= lmPAStartDate) Then
                mAddLong tmVbf.lInsertionCefCode, lmCefCode()
            Else
                mTestRecord = False
            End If
        End If
        Exit Function
    End If
    If UCase(slFileName) = "VFF.BTR" Then
        'tmVff = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmVff.iVefCode)
        If ilVef <= -1 Then
            mTestRecord = False
        End If
        Exit Function
    End If
    If UCase(slFileName) = "VOF.BTR" Then
        'tmVof = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmVof.iVefCode)
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            mAddLong tmVof.lHd1CefCode, lmCefCode()
            mAddLong tmVof.lFt1CefCode, lmCefCode()
            mAddLong tmVof.lFt2CefCode, lmCefCode()
        End If
        Exit Function
    End If
    If UCase(slFileName) = "VPF.BTR" Then
        'tmVpf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmVpf.iVefKCode)
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            mAddLong tmVpf.lLgHd1CefCode, lmCefCode()
            mAddLong tmVpf.lLgNmCefCode, lmCefCode()
            mAddLong tmVpf.lEMailCefCode, lmCefCode()
            mAddLong tmVpf.lLgFt1CefCode, lmCefCode()
            mAddLong tmVpf.lLgFt2CefCode, lmCefCode()
            mAddInteger tmVpf.iFTPArfCode, imArfCode()
            mAddInteger tmVpf.iProducerArfCode, imArfCode()
            mAddInteger tmVpf.iProgProvArfCode, imArfCode()
            mAddInteger tmVpf.iCommProvArfCode, imArfCode()
            mAddInteger tmVpf.iAutoExptArfCode, imArfCode()
            mAddInteger tmVpf.iAutoImptArfCode, imArfCode()
        End If
        Exit Function
    End If
    If UCase(slFileName) = "VAF.BTR" Then
        'tmVaf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmVaf.iVefCode)
        If ilVef <= -1 Then
            mTestRecord = False
        End If
        Exit Function
    End If
    If UCase(slFileName) = "PIF.BTR" Then
        'tmPif = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmPif.iVefCode)
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            gUnpackDateLong tmPif.iStartDate(0), tmPif.iStartDate(1), llStartDate
            gUnpackDateLong tmPif.iEndDate(0), tmPif.iEndDate(1), llEndDate
            If (llEndDate <> 0) And (llEndDate < lmPAStartDate) Then
                mTestRecord = False
            End If
        End If
        Exit Function
    End If
    If UCase(slFileName) = "PVF.BTR" Then
        'tmLvf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(lmPvfCode) - 1 Step 1
            If lmPvfCode(llTest) = tmPvf.lCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If blFound Then
            For llRow = LBound(tmPvf.iVefCode) To UBound(tmPvf.iVefCode) Step 1
                ilVef = mBinarySearchRecreateSort(tmPvf.iVefCode(llRow))
                If ilVef <= -1 Then
                    tmPvf.iVefCode(llRow) = 0
                    tmPvf.iPctRate(llRow) = 0
                    tmPvf.iRdfCode(llRow) = 0
                    tmPvf.iNoSpot(llRow) = 0
                End If
            Next llRow
        Else
            mTestRecord = False
        End If
        Exit Function
    End If
    If UCase(slFileName) = "LCF.BTR" Then
        'tmLcf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmLcf.iVefCode)
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            ilVef = gBinarySearchVef(tmLcf.iVefCode)
            If ilVef <> -1 Then
                gUnpackDateLong tmLcf.iLogDate(0), tmLcf.iLogDate(1), llLogDate
                If imPassNo = 1 Then
                    blFound = False
                    For llTest = 0 To UBound(tmLcfDateInfo) - 1 Step 1
                        If tmLcfDateInfo(llTest).iVefCode = tmLcf.iVefCode Then
                            blFound = True
                            If (llLogDate >= lmPAStartDate) Then
                                tmLcfDateInfo(llTest).lLatestDate = lmPAStartDate
                            Else
                                If llLogDate > tmLcfDateInfo(llTest).lLatestDate Then
                                    tmLcfDateInfo(llTest).lLatestDate = llLogDate
                                End If
                            End If
                            Exit For
                        End If
                    Next llTest
                    If Not blFound Then
                        tmLcfDateInfo(UBound(tmLcfDateInfo)).iVefCode = tmLcf.iVefCode
                        If llLogDate >= lmPAStartDate Then
                            tmLcfDateInfo(UBound(tmLcfDateInfo)).lLatestDate = lmPAStartDate
                        Else
                            tmLcfDateInfo(UBound(tmLcfDateInfo)).lLatestDate = llLogDate
                        End If
                        ReDim Preserve tmLcfDateInfo(0 To UBound(tmLcfDateInfo) + 1) As LCFDATEINFO
                    End If
                    mTestRecord = False
                Else
                    blFound = False
                    For llTest = 0 To UBound(tmLcfDateInfo) - 1 Step 1
                        If tmLcfDateInfo(llTest).iVefCode = tmLcf.iVefCode Then
                            If (llLogDate >= tmLcfDateInfo(llTest).lLatestDate) Then
                                blFound = True
                            End If
                            Exit For
                        End If
                    Next llTest
                    'Include all games images
                    If (llLogDate = 0) Or (blFound) Or (tmLcf.iType > 0) Then
                        For ilLoop = LBound(tmLcf.lLvfCode) To UBound(tmLcf.lLvfCode) Step 1
                            mAddLong tmLcf.lLvfCode(ilLoop), lmLvfCode()
                        Next ilLoop
                    Else
                        mTestRecord = False
                    End If
                End If
            Else
                mTestRecord = False
            End If
        End If
        Exit Function
    End If
    If UCase(slFileName) = "LVF.BTR" Then
        'tmLvf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(lmLvfCode) - 1 Step 1
            If lmLvfCode(llTest) = tmLvf.lCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If blFound Then
            llIndex = -1
            For llTest = 0 To UBound(tmLibTitleCount) - 1 Step 1
                'If (tmLibTitleCount(llTest).lLvfCode = tmLvf.lCode) And (tmLibTitleCount(llTest).iLtfCode = tmLvf.iLtfCode) Then
                If (tmLibTitleCount(llTest).iLtfCode = tmLvf.iLtfCode) Then
                    llIndex = llTest
                    Exit For
                End If
            Next llTest
            If imPassNo = 1 Then
                If llIndex = -1 Then
                    tmLibTitleCount(UBound(tmLibTitleCount)).iLtfCode = tmLvf.iLtfCode
                    tmLibTitleCount(UBound(tmLibTitleCount)).iMinVersion = tmLvf.iVersion
                    ReDim Preserve tmLibTitleCount(0 To UBound(tmLibTitleCount) + 1) As LIBTITLECOUNT
                Else
                    If tmLvf.iVersion < tmLibTitleCount(llIndex).iMinVersion Then
                        tmLibTitleCount(llIndex).iMinVersion = tmLvf.iVersion
                    End If
                End If
                mTestRecord = False
            Else
                If llIndex <> -1 Then
                    tmLvf.iVersion = tmLvf.iVersion - tmLibTitleCount(llIndex).iMinVersion + 1
                    'tmRecBuff = tmLvf
                Else
                    mTestRecord = False
                End If
            End If
        Else
            mTestRecord = False
        End If
        Exit Function
    End If
    If UCase(slFileName) = "LTF.BTR" Then
        'tmLtf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(tmLibTitleCount) - 1 Step 1
            If (tmLibTitleCount(llTest).iLtfCode = tmLtf.iCode) Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        End If
        Exit Function
    End If
    If UCase(slFileName) = "LEF.BTR" Then
        'tmLef = tmRecBuff
        blFound = False
        'For llTest = 0 To UBound(tmLibTitleCount) - 1 Step 1
        '    If (tmLibTitleCount(llTest).lLvfCode = tmLef.lLvfCode) Then
        '        blFound = True
        '        Exit For
        '    End If
        'Next llTest
        For llTest = 0 To UBound(lmLvfCode) - 1 Step 1
            If lmLvfCode(llTest) = tmLef.lLvfCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        End If
        Exit Function
    End If
    If UCase(slFileName) = "PAF.BTR" Then
        'tmPaf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmPaf.iVefCode)
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            gUnpackDateLong tmPaf.iStartDate(0), tmPaf.iStartDate(1), llStartDate
            gUnpackDateLong tmPaf.iEndDate(0), tmPaf.iEndDate(1), llEndDate
            If (llEndDate <> 0) And (llEndDate < lmPAStartDate) Then
                mTestRecord = False
            End If
        End If
        Exit Function
    End If
    If UCase(slFileName) = "PFF.BTR" Then
        'tmPff = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmPff.iVefCode)
        If ilVef <= -1 Then
            mTestRecord = False
        End If
        Exit Function
    End If
    If UCase(slFileName) = "VCF.BTR" Then
        'tmVcf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmVcf.iSellCode)
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            If tmRecreateSort(ilVef).bSellToConv = False Then
                gUnpackDateLong tmVcf.iEffDate(0), tmVcf.iEffDate(1), llStartDate
                gUnpackDateLong tmVcf.iTermDate(0), tmVcf.iTermDate(1), llEndDate
                If (llEndDate <> 0) And (llEndDate < lmPAStartDate) Then
                    mTestRecord = False
                End If
            Else
                mTestRecord = False
            End If
        End If
        Exit Function
    End If
    If UCase(slFileName) = "VLF.BTR" Then
        'tmVlf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmVlf.iAirCode)
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            gUnpackDateLong tmVlf.iEffDate(0), tmVlf.iEffDate(1), llStartDate
            gUnpackDateLong tmVlf.iTermDate(0), tmVlf.iTermDate(1), llEndDate
            If (llEndDate <> 0) And (llEndDate < lmPAStartDate) Then
                mTestRecord = False
            End If
        End If
        Exit Function
    End If
    If UCase(slFileName) = "IHF.BTR" Then
        'tmIhf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmIhf.iVefCode)
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            mAddInteger tmIhf.iCode, imIhfCode()
            mAddInteger tmIhf.iItfCode, imItfCode()
            mAddInteger tmIhf.iIifCode, imIifCode()
        End If
        Exit Function
    End If
    If UCase(slFileName) = "IIF.BTR" Then
        'tmIif = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imIifCode) - 1 Step 1
            If imIifCode(llTest) = tmIif.iCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        Else
            mAddInteger tmIif.iItfCode, imItfCode()
        End If
        Exit Function
    End If
    If UCase(slFileName) = "ITF.BTR" Then
        'tmItf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imItfCode) - 1 Step 1
            If imItfCode(llTest) = tmItf.iCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        End If
        Exit Function
    End If
    If UCase(slFileName) = "GHF.BTR" Then
        'tmGhf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmGhf.iVefCode)
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            mAddLong tmGhf.lCode, lmGhfCode()
        End If
        Exit Function
    End If
    If UCase(slFileName) = "GSF.BTR" Then
        'tmGsf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(lmGhfCode) - 1 Step 1
            If lmGhfCode(llTest) = tmGsf.lGhfCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        Else
            mAddLong tmGsf.lLvfCode, lmLvfCode()
        End If
        Exit Function
    End If
    If UCase(slFileName) = "ENF.BTR" Then
        'tmEnf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmEnf.iVefCode)
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            mAddLong tmEnf.lCefCode, lmCefCode()
            mAddInteger tmEnf.iEtfCode, imEtfCode()
        End If
        Exit Function
    End If
    If UCase(slFileName) = "ETF.BTR" Then
        'tmEtf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imEtfCode) - 1 Step 1
            If imEtfCode(llTest) = tmEtf.iCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        End If
        Exit Function
    End If
    If UCase(slFileName) = "ANF.BTR" Then
        'All Records
        Exit Function
    End If

    If UCase(slFileName) = "DNF.BTR" Then
        'tmDnf = tmRecBuff
        If imPassNo = 1 Then
            gUnpackDateLong tmDnf.iBookDate(0), tmDnf.iBookDate(1), llStartDate
            If (llStartDate = 0) Or (llStartDate >= lmResearchStartDate) Then
                mAddInteger tmDnf.iCode, imCheckDnfCode()
            End If
            mTestRecord = False
        Else
            blFound = False
            For llTest = 0 To UBound(imDnfCode) - 1 Step 1
                If imDnfCode(llTest) = tmDnf.iCode Then
                    blFound = True
                    Exit For
                End If
            Next llTest
            If Not blFound Then
                mTestRecord = False
            End If
        End If
        Exit Function
    End If
    If UCase(slFileName) = "DRF.BTR" Then
        'tmDrf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imDnfCode) - 1 Step 1
            If imDnfCode(llTest) = tmDrf.iDnfCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            For llTest = 0 To UBound(imCheckDnfCode) - 1 Step 1
                If imCheckDnfCode(llTest) = tmDrf.iDnfCode Then
                    blFound = True
                    Exit For
                End If
            Next llTest
        End If
        If blFound Then
            ilVef = mBinarySearchRecreateSort(tmDrf.iVefCode)
            If ilVef <= -1 Then
                mTestRecord = False
            Else
                mAddInteger tmDrf.iRdfCode, imRdfCode()
                mAddInteger tmDrf.iDnfCode, imDnfCode()
                'mAddLong tmDrf.lCode, lmDrfCode()
                ilRet = mCopyDpf(tmDrf.lCode)
            End If
        Else
            mTestRecord = False
        End If
        Exit Function
    End If

    If UCase(slFileName) = "DEF.BTR" Then
        'tmDef = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imDnfCode) - 1 Step 1
            If imDnfCode(llTest) = tmDef.iDnfCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        End If
        Exit Function
    End If
    'If UCase(slFileName) = "DPF.BTR" Then
    '    tmDpf = tmRecBuff
    '    blFound = False
    '    For llTest = 0 To UBound(imDnfCode) - 1 Step 1
    '        If imDnfCode(llTest) = tmDpf.iDnfCode Then
    '            blFound = True
    '            Exit For
    '        End If
    '    Next llTest
    '    If blFound Then
    '        blFound = False
    '        For llTest = 0 To UBound(lmDrfCode) - 1 Step 1
    '            If lmDrfCode(llTest) = tmDpf.lDrfCode Then
    '                blFound = True
    '                Exit For
    '            End If
    '        Next llTest
    '    End If
    '    If Not blFound Then
    '        mTestRecord = False
    '    End If
    '    Exit Function
    'End If
    If UCase(slFileName) = "RDF.BTR" Then
        'tmRdf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imRdfCode) - 1 Step 1
            If imRdfCode(llTest) = tmRdf.iCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        End If
        Exit Function
    End If


    If UCase(slFileName) = "SPF.BTR" Then
        'All Records
        Exit Function
    End If
    
    If UCase(slFileName) = "SAF.BTR" Then
        'tmSaf = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmSaf.iVefCode)
        If ilVef <= -1 Then
            mTestRecord = False
        End If
        Exit Function
    End If

    If UCase(slFileName) = "ADF.BTR" Then
        'tmAdf = tmRecBuff
        If tmAdf.sState = "D" Then
            mTestRecord = False
        Else
            mAddInteger tmAdf.iCode, imAdfCode()
            mAddInteger tmAdf.iSlfCode, imSlfCode()
            mAddInteger tmAdf.iAgfCode, imAgfCode()
            mAddInteger tmAdf.iArfLkCode, imArfCode()
            mAddInteger tmAdf.iArfCntrCode, imArfCode()
            mAddInteger tmAdf.iArfInvCode, imArfCode()
            mAddInteger tmAdf.iTrfCode, imTrfCode()
            mAddInteger tmAdf.iPnfBuyer, imPnfCode()
            mAddInteger tmAdf.iPnfPay, imPnfCode()
            mAddInteger tmAdf.iUrfCode, imUrfCode()
        End If
        Exit Function
    End If
    If UCase(slFileName) = "AGF.BTR" Then
        'tmAgf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imAgfCode) - 1 Step 1
            If imAgfCode(llTest) = tmAgf.iCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        
        If (tmAgf.sState = "D") And (Not blFound) Then
            mTestRecord = False
        Else
            mAddInteger tmAgf.iSlfCode, imSlfCode()
            mAddInteger tmAgf.iArfLkCode, imArfCode()
            mAddInteger tmAgf.iArfCntrCode, imArfCode()
            mAddInteger tmAgf.iArfInvCode, imArfCode()
            mAddInteger tmAgf.iTrfCode, imTrfCode()
            mAddInteger tmAgf.iPnfBuyer, imPnfCode()
            mAddInteger tmAgf.iPnfPay, imPnfCode()
            mAddInteger tmAgf.iUrfCode, imUrfCode()
        End If
        Exit Function
    End If
    If UCase(slFileName) = "SLF.BTR" Then
        'tmSlf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imSlfCode) - 1 Step 1
            If imSlfCode(llTest) = tmSlf.iCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        
        If (tmSlf.sState = "D") And (Not blFound) Then
            mTestRecord = False
        Else
            mAddInteger tmSlf.iSofCode, imSofCode()
            'mAddInteger tmSlf.iUrfCode, imUrfCode()
        End If
        Exit Function
    End If

    If UCase(slFileName) = "MNF.BTR" Then
        'All Records
        Exit Function
    End If

    If UCase(slFileName) = "PRF.BTR" Then
        'tmPrf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imAdfCode) - 1 Step 1
            If imAdfCode(llTest) = tmPrf.iAdfCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        
        If (tmPrf.sState = "D") And (Not blFound) Then
            mTestRecord = False
        Else
            mAddInteger tmPrf.iPnfBuyer, imPnfCode()
            mAddInteger tmPrf.iUrfCode, imUrfCode()
        End If
        Exit Function
    End If

    If UCase(slFileName) = "RNF.BTR" Then
        'All Records
        Exit Function
    End If

    If UCase(slFileName) = "RTF.BTR" Then
        'All Records
        Exit Function
    End If

    If UCase(slFileName) = "ARF.BTR" Then
        'tmArf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imArfCode) - 1 Step 1
            If imArfCode(llTest) = tmArf.iCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        
        If (Not blFound) Then
            mTestRecord = False
        End If
        Exit Function
    End If

    If UCase(slFileName) = "TRF.BTR" Then
        'tmTrf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imTrfCode) - 1 Step 1
            If imTrfCode(llTest) = tmTrf.iCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        
        If (Not blFound) Then
            mTestRecord = False
        Else
            mAddInteger tmTrf.iUrfCode, imUrfCode()
        End If
        Exit Function
    End If

    If UCase(slFileName) = "PNF.BTR" Then
        'tmPnf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imPnfCode) - 1 Step 1
            If imPnfCode(llTest) = tmPnf.iCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        
        If (tmPnf.sState = "D") And (Not blFound) Then
            mTestRecord = False
        Else
            mAddLong tmPnf.lCxfCode, lmCxfCode()
            mAddLong tmPnf.lEMailCefCode, lmCefCode()
        End If
        Exit Function
    End If

    If UCase(slFileName) = "CXF.BTR" Then
        'tmCxf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(lmCxfCode) - 1 Step 1
            If lmCxfCode(llTest) = tmCxf.lCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        End If
        Exit Function
    End If
    
    If UCase(slFileName) = "SOF.BTR" Then
        'tmSof = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imSofCode) - 1 Step 1
            If imSofCode(llTest) = tmSof.iCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        
        If (tmSof.sState = "D") And (Not blFound) Then
            mTestRecord = False
        End If
        Exit Function
    End If

    If UCase(slFileName) = "CMF.BTR" Then
        'All Records
    End If

    'If UCase(slFileName) = "URF.BTR" Then
    '    tmUrf = tmRecBuff
    '    blFound = False
    '    For llTest = 0 To UBound(imUrfCode) - 1 Step 1
    '        If imUrfCode(llTest) = tmUrf.iCode Then
    '            blFound = True
    '            Exit For
    '        End If
    '    Next llTest
    '    blFound = True
    '    If (tmUrf.sDelete = "Y") And (Not blFound) Then
    '        mTestRecord = False
    '    Else
    '        mAddInteger tmUrf.iSnfCode, imSnfCode()
    '        mAddInteger tmUrf.iSlfCode, imSlfCode()
    '        ilVef = mBinarySearchRecreateSort(tmUrf.iVefCode)
    '        If ilVef <= -1 Then
    '            tmUrf.iVefCode = 0
    '        End If
    '        ilVef = mBinarySearchRecreateSort(tmUrf.iDefVeh)
    '        If ilVef <= -1 Then
    '            tmUrf.iDefVeh = 0
    '        End If
    '        tmRecBuff = tmUrf
    '    End If
    'End If

    If UCase(slFileName) = "SNF.BTR" Then
        'tmSnf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imSnfCode) - 1 Step 1
            If imSnfCode(llTest) = tmSnf.iCode Then
                blFound = True
                Exit For
            End If
        Next llTest

        If (Not blFound) Then
            mTestRecord = False
        End If
        Exit Function
    End If

    If UCase(slFileName) = "SRF.BTR" Then
        'tmSrf = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imSnfCode) - 1 Step 1
            If imSnfCode(llTest) = tmSrf.iSnfCode Then
                blFound = True
                Exit For
            End If
        Next llTest

        If (Not blFound) Then
            mTestRecord = False
        End If
        Exit Function
    End If

    If UCase(slFileName) = "ATT.MKD" Then
        'tmAtt = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmAtt.iVefCode)
        If ilVef <= -1 Then
            For llTest = 0 To UBound(tmRecreateSort) - 1 Step 1
                If tmRecreateSort(llTest).iAirAgreeVefCode = tmAtt.iVefCode Then
                    tmAtt.iVefCode = tmRecreateSort(llTest).iVefCode
                    ilVef = llTest
                    Exit For
                End If
            Next llTest
        End If
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            gUnpackDateLong tmAtt.iOffAir(0), tmAtt.iOffAir(1), llOffAirDate
            gUnpackDateLong tmAtt.iDropDate(0), tmAtt.iDropDate(1), llDropDate
            If ((llOffAirDate = 0) And (llDropDate = 0)) Or ((llOffAirDate >= lmPAStartDate) And (llDropDate >= lmPAStartDate)) Then
                mAddLong tmAtt.lCode, lmAttCode()
                gUnpackDateLong tmAtt.iOnAir(0), tmAtt.iOnAir(1), llStartDate
                If llStartDate < lmPAStartDate Then
                    gPackDateLong lmPAStartDate, tmAtt.iOnAir(0), tmAtt.iOnAir(1)
                End If
                gUnpackDateLong tmAtt.iAgreeStart(0), tmAtt.iAgreeStart(1), llStartDate
                If llStartDate < lmPAStartDate Then
                    gPackDateLong lmPAStartDate, tmAtt.iAgreeStart(0), tmAtt.iAgreeStart(1)
                End If
                'tmRecBuff = tmAtt
            Else
                mTestRecord = False
            End If
        End If
        Exit Function
    End If
    If UCase(slFileName) = "DAT.MKD" Then
        'tmDat = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(lmAttCode) - 1 Step 1
            If lmAttCode(llTest) = tmDat.lAtfCode Then
                blFound = True
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        Else
            ilVef = mBinarySearchRecreateSort(tmDat.iVefCode)
            If ilVef <= -1 Then
                For llTest = 0 To UBound(tmRecreateSort) - 1 Step 1
                    If tmRecreateSort(llTest).iAirAgreeVefCode = tmDat.iVefCode Then
                        tmDat.iVefCode = tmRecreateSort(llTest).iVefCode
                        Exit For
                    End If
                Next llTest
            End If
        End If
        Exit Function
    End If
    If UCase(slFileName) = "EPT.MKD" Then
        'tmEpt = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(lmAttCode) - 1 Step 1
            If lmAttCode(llTest) = tmEpt.lAttCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        Else
            ilVef = mBinarySearchRecreateSort(tmEpt.iVefCode)
            If ilVef <= -1 Then
                For llTest = 0 To UBound(tmRecreateSort) - 1 Step 1
                    If tmRecreateSort(llTest).iAirAgreeVefCode = tmEpt.iVefCode Then
                        tmEpt.iVefCode = tmRecreateSort(llTest).iVefCode
                        Exit For
                    End If
                Next llTest
            End If
        End If
        Exit Function
    End If
    If UCase(slFileName) = "PFT.MKD" Then
        'tmPft = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(lmAttCode) - 1 Step 1
            If lmAttCode(llTest) = tmPft.lAttCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        Else
            ilVef = mBinarySearchRecreateSort(tmPft.iVefCode)
            If ilVef <= -1 Then
                For llTest = 0 To UBound(tmRecreateSort) - 1 Step 1
                    If tmRecreateSort(llTest).iAirAgreeVefCode = tmPft.iVefCode Then
                        tmPft.iVefCode = tmRecreateSort(llTest).iVefCode
                        Exit For
                    End If
                Next llTest
            End If
            mAddLong tmGsf.lLvfCode, lmLvfCode()
        End If
        Exit Function
    End If
    
    If UCase(slFileName) = "ARTT.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "CLT.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "CMT.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "DFT.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "FLT.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "FMT.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "MAT.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "MET.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "MKT.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "MNT.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "SHTT.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "SITE.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "SNT.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "TNT.MKD" Then
        'All Records
        Exit Function
    End If
    If UCase(slFileName) = "TZT.MKD" Then
        'All Records
        Exit Function
    End If

    If UCase(slFileName) = "CCT.MKD" Then
        'tmCct = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmCct.iVefCode)
        If (ilVef <= -1) And (tmCct.iVefCode > 0) Then
            For llTest = 0 To UBound(tmRecreateSort) - 1 Step 1
                If tmRecreateSort(llTest).iAirAgreeVefCode = tmCct.iVefCode Then
                    tmCct.iVefCode = tmRecreateSort(llTest).iVefCode
                    ilVef = llTest
                    Exit For
                End If
            Next llTest
        End If
        If ilVef <= -1 Then
            mTestRecord = False
        End If
        Exit Function
    End If
    
    If UCase(slFileName) = "CST.MKD" Then
        'tmCst = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(imCstCode) - 1 Step 1
            If imCstCode(llTest) = tmCst.iCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        End If
        Exit Function
    End If
    
    If UCase(slFileName) = "RHT.MKD" Then
        'tmRht = tmRecBuff
        ilVef = mBinarySearchRecreateSort(tmRht.iVefCode)
        If (ilVef <= -1) And (tmRht.iVefCode > 0) Then
            For llTest = 0 To UBound(tmRecreateSort) - 1 Step 1
                If tmRecreateSort(llTest).iAirAgreeVefCode = tmRht.iVefCode Then
                    tmRht.iVefCode = tmRecreateSort(llTest).iVefCode
                    ilVef = llTest
                    Exit For
                End If
            Next llTest
        End If
        If ilVef <= -1 Then
            mTestRecord = False
        Else
            mAddLong tmRht.lCode, lmRhtCode()
        End If
        Exit Function
    End If
    
    If UCase(slFileName) = "RET.MKD" Then
        'tmRet = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(lmRhtCode) - 1 Step 1
            If lmRhtCode(llTest) = tmRet.lRhtCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        End If
        Exit Function
    End If

    If UCase(slFileName) = "UST.MKD" Then
        'tmUst = tmRecBuff
        blFound = False
        'For llTest = 0 To UBound(imUstCode) - 1 Step 1
        '    If imUstCode(llTest) = tmUst.iCode Then
        '        blFound = True
        '        Exit For
        '    End If
        'Next llTest
        blFound = True
        If (tmUst.iState = 1) And (Not blFound) Then
            mTestRecord = False
        Else
            mAddLong tmUst.lEMailCefCode, lmCefCode()
        End If
        Exit Function
    End If
    If UCase(slFileName) = "DNT.MKD" Then
        'All Records
        Exit Function
    End If

    If UCase(slFileName) = "CEF.BTR" Then
        'tmCef = tmRecBuff
        blFound = False
        For llTest = 0 To UBound(lmCefCode) - 1 Step 1
            If lmCefCode(llTest) = tmCef.lCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            mTestRecord = False
        End If
        Exit Function
    End If

End Function


Private Sub mAddLong(llCode As Long, llArray() As Long)
    Dim blFound As Boolean
    Dim llTest As Long
    
    If llCode > 0 Then
        blFound = False
        For llTest = 0 To UBound(llArray) - 1 Step 1
            If llArray(llTest) = llCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            llArray(UBound(llArray)) = llCode
            ReDim Preserve llArray(0 To UBound(llArray) + 1) As Long
        End If
    End If
End Sub

Private Sub mAddInteger(ilCode As Integer, ilArray() As Integer)
    Dim blFound As Boolean
    Dim llTest As Long
    
    If ilCode > 0 Then
        blFound = False
        For llTest = 0 To UBound(ilArray) - 1 Step 1
            If ilArray(llTest) = ilCode Then
                blFound = True
                Exit For
            End If
        Next llTest
        If Not blFound Then
            ilArray(UBound(ilArray)) = ilCode
            ReDim Preserve ilArray(0 To UBound(ilArray) + 1) As Integer
        End If
    End If
End Sub


Private Function mUrfCopy() As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    
    imFileNo = imFileNo + 1
    lacInfo(0).Caption = "File: " & imFileNo & " of " & imTotalFiles
    lacInfo(1).Caption = "URF.BTR"
    lmRecordNo = 0
    'hmFromUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    'ilRet = btrOpen(hmFromUrf, "", smFromDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet <> BTRV_ERR_NONE Then
    '    mUrfCopy = False
    '    Exit Function
    'End If
    'hmToUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    'ilRet = btrOpen(hmToUrf, "", smToDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet <> BTRV_ERR_NONE Then
    '    mUrfCopy = False
    '    Exit Function
    'End If
    If Not mOpenFile(hmFromUrf, hmToUrf, "URF.BTR") Then
        mUrfCopy = False
        Exit Function
    End If
    lmTotalRecords = btrRecords(hmFromUrf)
    imUrfRecLen = Len(tmUrf)
    ilRet = btrGetFirst(hmFromUrf, tmUrf, imUrfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If imTerminate Then
            mUrfCopy = False
            Exit Function
        End If
        lmRecordNo = lmRecordNo + 1
        lacInfo(1).Caption = "URF.BTR" & ": " & lmRecordNo & " of " & lmTotalRecords
        DoEvents
        gUrfDecrypt tmUrf
        mAddInteger tmUrf.iSnfCode, imSnfCode()
        mAddInteger tmUrf.iSlfCode, imSlfCode()
        ilVef = mBinarySearchRecreateSort(tmUrf.iVefCode)
        If ilVef <= -1 Then
            tmUrf.iVefCode = 0
        End If
        ilVef = mBinarySearchRecreateSort(tmUrf.iDefVeh)
        If ilVef <= -1 Then
            tmUrf.iDefVeh = 0
        End If
        gUrfEncrypt tmUrf
        ilRet = btrInsert(hmToUrf, tmUrf, imUrfRecLen, INDEXKEY0)
        ilRet = btrGetNext(hmFromUrf, tmUrf, imUrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    'ilRet = btrClose(hmFromUrf)
    'ilRet = btrClose(hmToUrf)
    'btrDestroy hmFromUrf
    'btrDestroy hmToUrf
    ilRet = mCloseFile(hmFromUrf, hmToUrf)
    mUrfCopy = True
End Function

Private Function mOpenFile(hlFrom As Integer, hlTo As Integer, slFileName As String) As Integer
    Dim ilRet As Integer
    
    hlFrom = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlFrom, "", smFromDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenFile = False
        Exit Function
    End If
    hlTo = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlTo, "", smToDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenFile = False
        Exit Function
    End If
    mOpenFile = True
End Function

Private Function mCloseFile(hlFrom As Integer, hlTo As Integer) As Integer
    Dim ilRet As Integer
    
    ilRet = btrClose(hlFrom)
    ilRet = btrClose(hlTo)
    btrDestroy hlFrom
    btrDestroy hlTo
    mCloseFile = True
End Function

Private Function mCopyDpf(llDrfCode As Long) As Integer
    Dim ilRet As Integer
    
    tmDpfSrchKey1.lDrfCode = llDrfCode
    tmDpfSrchKey1.iMnfDemo = 0
    ilRet = btrGetGreaterOrEqual(hmFromDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmDpf.lDrfCode = llDrfCode)
        If imTerminate Then
            mCopyDpf = False
            Exit Function
        End If
        DoEvents
        ilRet = btrInsert(hmToDpf, tmDpf, imDpfRecLen, INDEXKEY0)
        ilRet = btrGetNext(hmFromDpf, tmDpf, imDpfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mCopyDpf = True
End Function

Private Function mExtGetNext(slFileName As String, ilRecLen As Integer) As Integer
    Dim llRecPos As Long
    Dim ilRet As Integer
    
    Select Case UCase(slFileName)
    
        Case "VEF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmVef, ilRecLen, llRecPos)    'Extract record
        Case "VBF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmVbf, ilRecLen, llRecPos)    'Extract record
        Case "VFF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmVff, ilRecLen, llRecPos)    'Extract record
        Case "VOF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmVof, ilRecLen, llRecPos)    'Extract record
        Case "VPF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmVpf, ilRecLen, llRecPos)    'Extract record
        Case "VAF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmVaf, ilRecLen, llRecPos)    'Extract record
        Case "PVF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmPvf, ilRecLen, llRecPos)    'Extract record
        Case "PIF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmPif, ilRecLen, llRecPos)    'Extract record
        Case "LCF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmLcf, ilRecLen, llRecPos)    'Extract record
        Case "LVF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmLvf, ilRecLen, llRecPos)    'Extract record
        Case "LTF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmLtf, ilRecLen, llRecPos)    'Extract record
        Case "LEF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmLef, ilRecLen, llRecPos)    'Extract record
        Case "PAF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmPaf, ilRecLen, llRecPos)    'Extract record
        Case "PFF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmPff, ilRecLen, llRecPos)    'Extract record
        Case "VCF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmVcf, ilRecLen, llRecPos)    'Extract record
        Case "VLF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmVlf, ilRecLen, llRecPos)    'Extract record
        Case "IHF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmIhf, ilRecLen, llRecPos)    'Extract record
        Case "IIF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmIif, ilRecLen, llRecPos)    'Extract record
        Case "ITF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmItf, ilRecLen, llRecPos)    'Extract record
        Case "GHF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmGhf, ilRecLen, llRecPos)    'Extract record
        Case "GSF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmGsf, ilRecLen, llRecPos)    'Extract record
        Case "ENF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmEnf, ilRecLen, llRecPos)    'Extract record
        Case "ETF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmEtf, ilRecLen, llRecPos)    'Extract record
        Case "DNF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmDnf, ilRecLen, llRecPos)    'Extract record
        Case "DRF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmDrf, ilRecLen, llRecPos)    'Extract record
        Case "DEF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmDef, ilRecLen, llRecPos)    'Extract record
        Case "RDF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmRdf, ilRecLen, llRecPos)    'Extract record
        Case "SAF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmSaf, ilRecLen, llRecPos)    'Extract record
        Case "ADF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmAdf, ilRecLen, llRecPos)    'Extract record
        Case "AGF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmAgf, ilRecLen, llRecPos)    'Extract record
        Case "SLF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmSlf, ilRecLen, llRecPos)    'Extract record
        Case "PRF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmPrf, ilRecLen, llRecPos)    'Extract record
        Case "ARF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmArf, ilRecLen, llRecPos)    'Extract record
        Case "TRF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmTrf, ilRecLen, llRecPos)    'Extract record
        Case "PNF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmPnf, ilRecLen, llRecPos)    'Extract record
        Case "CXF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmCxf, ilRecLen, llRecPos)    'Extract record
        Case "SOF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmSof, ilRecLen, llRecPos)    'Extract record
        Case "SNF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmSnf, ilRecLen, llRecPos)    'Extract record
        Case "SRF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmSrf, ilRecLen, llRecPos)    'Extract record
        Case "ATT.MKD"
            ilRet = btrExtGetNext(hmFrom, tmAtt, ilRecLen, llRecPos)    'Extract record
        Case "DAT.MKD"
            ilRet = btrExtGetNext(hmFrom, tmDat, ilRecLen, llRecPos)    'Extract record
        Case "EPT.MKD"
            ilRet = btrExtGetNext(hmFrom, tmEpt, ilRecLen, llRecPos)    'Extract record
        Case "PfT.MKD"
            ilRet = btrExtGetNext(hmFrom, tmPft, ilRecLen, llRecPos)    'Extract record
        Case "CCT.MKD"
            ilRet = btrExtGetNext(hmFrom, tmCct, ilRecLen, llRecPos)    'Extract record
        Case "CST.MKD"
            ilRet = btrExtGetNext(hmFrom, tmCst, ilRecLen, llRecPos)    'Extract record
        Case "RHT.MKD"
            ilRet = btrExtGetNext(hmFrom, tmRht, ilRecLen, llRecPos)    'Extract record
        Case "RET.MKD"
            ilRet = btrExtGetNext(hmFrom, tmRet, ilRecLen, llRecPos)    'Extract record
        Case "UST.MKD"
            ilRet = btrExtGetNext(hmFrom, tmUst, ilRecLen, llRecPos)    'Extract record
        Case "CEF.BTR"
            ilRet = btrExtGetNext(hmFrom, tmCef, ilRecLen, llRecPos)    'Extract record
        Case Else
            ilRet = btrExtGetNext(hmFrom, tmRecBuff, ilRecLen, llRecPos)    'Extract record
    End Select
    mExtGetNext = ilRet
End Function

Private Function mInsert(slFileName As String, ilRecLen As Integer, ilAutoCodeKeyNo As Integer) As Integer
    Dim ilRet As Integer
    
    Select Case UCase(slFileName)
    
        Case "VEF.BTR"
            ilRet = btrInsert(hmTo, tmVef, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "VBF.BTR"
            ilRet = btrInsert(hmTo, tmVbf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "VFF.BTR"
            ilRet = btrInsert(hmTo, tmVff, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "VOF.BTR"
            ilRet = btrInsert(hmTo, tmVof, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "VPF.BTR"
            ilRet = btrInsert(hmTo, tmVpf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "VAF.BTR"
            ilRet = btrInsert(hmTo, tmVaf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "PIF.BTR"
            ilRet = btrInsert(hmTo, tmPif, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "PVF.BTR"
            ilRet = btrInsert(hmTo, tmPvf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "LCF.BTR"
            ilRet = btrInsert(hmTo, tmLcf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "LVF.BTR"
            ilRet = btrInsert(hmTo, tmLvf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "LTF.BTR"
            ilRet = btrInsert(hmTo, tmLtf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "LEF.BTR"
            ilRet = btrInsert(hmTo, tmLef, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "PAF.BTR"
            ilRet = btrInsert(hmTo, tmPaf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "PFF.BTR"
            ilRet = btrInsert(hmTo, tmPff, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "VCF.BTR"
            ilRet = btrInsert(hmTo, tmVcf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "VLF.BTR"
            ilRet = btrInsert(hmTo, tmVlf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "IHF.BTR"
            ilRet = btrInsert(hmTo, tmIhf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "IIF.BTR"
            ilRet = btrInsert(hmTo, tmIif, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "ITF.BTR"
            ilRet = btrInsert(hmTo, tmItf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "GHF.BTR"
            ilRet = btrInsert(hmTo, tmGhf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "GSF.BTR"
            ilRet = btrInsert(hmTo, tmGsf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "ENF.BTR"
            ilRet = btrInsert(hmTo, tmEnf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "ETF.BTR"
            ilRet = btrInsert(hmTo, tmEtf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "DNF.BTR"
            ilRet = btrInsert(hmTo, tmDnf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "DRF.BTR"
            ilRet = btrInsert(hmTo, tmDrf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "DEF.BTR"
            ilRet = btrInsert(hmTo, tmDef, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "RDF.BTR"
            ilRet = btrInsert(hmTo, tmRdf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "SAF.BTR"
            ilRet = btrInsert(hmTo, tmSaf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "ADF.BTR"
            ilRet = btrInsert(hmTo, tmAdf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "AGF.BTR"
            ilRet = btrInsert(hmTo, tmAgf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "SLF.BTR"
            ilRet = btrInsert(hmTo, tmSlf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "PRF.BTR"
            ilRet = btrInsert(hmTo, tmPrf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "ARF.BTR"
            ilRet = btrInsert(hmTo, tmArf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "TRF.BTR"
            ilRet = btrInsert(hmTo, tmTrf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "PNF.BTR"
            ilRet = btrInsert(hmTo, tmPnf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "CXF.BTR"
            ilRet = btrInsert(hmTo, tmCxf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "SOF.BTR"
            ilRet = btrInsert(hmTo, tmSof, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "SNF.BTR"
            ilRet = btrInsert(hmTo, tmSnf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "SRF.BTR"
            ilRet = btrInsert(hmTo, tmSrf, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "ATT.MKD"
            ilRet = btrInsert(hmTo, tmAtt, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "DAT.MKD"
            ilRet = btrInsert(hmTo, tmDat, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "EPT.MKD"
            ilRet = btrInsert(hmTo, tmEpt, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "PfT.MKD"
            ilRet = btrInsert(hmTo, tmPft, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "CCT.MKD"
            ilRet = btrInsert(hmTo, tmCct, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "CST.MKD"
            ilRet = btrInsert(hmTo, tmCst, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "RHT.MKD"
            ilRet = btrInsert(hmTo, tmRht, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "RET.MKD"
            ilRet = btrInsert(hmTo, tmRet, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "UST.MKD"
            ilRet = btrInsert(hmTo, tmUst, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case "CEF.BTR"
            ilRet = btrInsert(hmTo, tmCef, ilRecLen, ilAutoCodeKeyNo)    'Extract record
        Case Else
            ilRet = btrInsert(hmTo, tmRecBuff, ilRecLen, ilAutoCodeKeyNo)    'Extract record
    End Select
    mInsert = ilRet
End Function

