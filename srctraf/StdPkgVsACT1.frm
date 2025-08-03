VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form StdPkgVsACT1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   9435
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "StdPkgVsACT1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lbcDemo 
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
      Height          =   240
      ItemData        =   "StdPkgVsACT1.frx":08CA
      Left            =   7740
      List            =   "StdPkgVsACT1.frx":08CC
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4890
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmcExport 
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6030
      TabIndex        =   27
      Top             =   4770
      Width           =   1335
   End
   Begin VB.Frame frcSelection 
      Caption         =   "Package vs ACT1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   120
      TabIndex        =   0
      Top             =   15
      Width           =   9180
      Begin VB.ComboBox cbcBook 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7320
         TabIndex        =   17
         Top             =   900
         Width           =   1700
      End
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Audience"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   7950
         TabIndex        =   8
         Top             =   210
         Width           =   1845
      End
      Begin VB.ComboBox cbcVer 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2910
         TabIndex        =   21
         Top             =   1275
         Width           =   1875
      End
      Begin VB.ComboBox cbcLine 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5325
         TabIndex        =   23
         Top             =   1275
         Width           =   3675
      End
      Begin VB.ComboBox cbcStandard 
         Height          =   315
         Left            =   4890
         TabIndex        =   15
         Top             =   900
         Width           =   1700
      End
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Daypart"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6870
         TabIndex        =   7
         Top             =   210
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox ckcInclude 
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6060
         TabIndex        =   6
         Top             =   210
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox ckcCompare 
         Caption         =   "Contract"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   4
         Top             =   210
         Width           =   1140
      End
      Begin VB.CheckBox ckcCompare 
         Caption         =   "Package"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2655
         TabIndex        =   3
         Top             =   210
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox ckcCompare 
         Caption         =   "ACT1 Line-up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1125
         TabIndex        =   2
         Top             =   210
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.TextBox edcCntrNo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1050
         TabIndex        =   19
         Top             =   1275
         Width           =   1230
      End
      Begin VB.ComboBox cbcACT1PkgName 
         Height          =   315
         Left            =   2085
         TabIndex        =   13
         Top             =   900
         Width           =   1700
      End
      Begin VB.CommandButton cmcBrowser 
         Appearance      =   0  'Flat
         Caption         =   "&Browse..."
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
         Left            =   7485
         TabIndex        =   11
         Top             =   525
         Width           =   1485
      End
      Begin VB.TextBox edcACT1File 
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Top             =   525
         Width           =   5355
      End
      Begin VB.Label lacBook 
         Caption         =   "Book"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6780
         TabIndex        =   16
         Top             =   930
         Width           =   480
      End
      Begin VB.Label lacVer 
         Caption         =   "Ver #"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2355
         TabIndex        =   20
         Top             =   1290
         Width           =   570
      End
      Begin VB.Label lacStandard 
         Caption         =   "Standard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3945
         TabIndex        =   14
         Top             =   930
         Width           =   915
      End
      Begin VB.Label lacLine 
         Caption         =   "Line"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4830
         TabIndex        =   22
         Top             =   1305
         Width           =   570
      End
      Begin VB.Label lacInclude 
         Caption         =   "Include"
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
         Left            =   5325
         TabIndex        =   5
         Top             =   225
         Width           =   750
      End
      Begin VB.Label lacCompare 
         Caption         =   "Compare"
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
         Left            =   165
         TabIndex        =   1
         Top             =   225
         Width           =   885
      End
      Begin VB.Label lacCntrNo 
         Caption         =   "Contract #"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   18
         Top             =   1305
         Width           =   915
      End
      Begin VB.Label lacPkgName 
         Caption         =   "Package Name: ACT1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   930
         Width           =   1785
      End
      Begin VB.Label lacACT1File 
         Caption         =   "ACT1 Line-up PDF File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   9
         Top             =   540
         Width           =   1830
      End
   End
   Begin VB.CommandButton cmcCompare 
      Caption         =   "Compare"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4110
      TabIndex        =   26
      Top             =   4770
      Width           =   1335
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4725
      Width           =   45
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   25
      Top             =   4770
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCompare 
      Height          =   2610
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1875
      Visible         =   0   'False
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   4604
      _Version        =   393216
      Rows            =   3
      Cols            =   15
      FixedRows       =   2
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
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
      _Band(0).Cols   =   15
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComDlg.CommonDialog cdcSetup 
      Left            =   8745
      Top             =   4620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".Txt"
      Filter          =   "*.Txt|*.Txt|*.Doc|*.Doc|*.Asc|*.Asc"
      FilterIndex     =   1
      FontSize        =   0
      MaxFileSize     =   256
   End
End
Attribute VB_Name = "StdPkgVsACT1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  StdPkgVsACT1 - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private imCtrlKey As Integer
Private imShiftKey As Integer
Private imTerminate As Integer
Private lmLastClickedRow As Long
Private lmScrollTop As Long
Private lmEnableRow As Long
Private lmEnableCol As Long
Private smNowDate As String
Private lmNowDate As Long

Private imCompareColSorted As Integer
Private imCompareSort As Integer

Dim tmPkgVehicle() As SORTCODE
Dim smPkgVehicleTag As String

Dim tmDemo(0 To 44) As String
Dim smDemo As String
Dim imMnfDemo As Integer

Private tmPvf() As PVF
Private tmRdf As RDF

Dim hmDrf As Integer
Dim hmMnf As Integer
Dim hmDpf As Integer        'Demo Plus handle
Dim hmDef As Integer
Dim hmRaf As Integer

Private rst_vef As ADODB.Recordset
Private rst_pvf As ADODB.Recordset
Private rst_Chf As ADODB.Recordset
Private rst_Clf As ADODB.Recordset
Private rst_Cff As ADODB.Recordset
Private rst_dnf As ADODB.Recordset

Private Type GRIDINFO
    sKey As String * 20
    sACT1CallLetters As String * 10
    sPkgCallLetters As String * 10
    sCntrCallLetters As String * 10
    iACT1Units As Integer
    iPkgUnits As Integer
    iCntrUnits As Integer
    sACT1AQH As String * 10
    sPkgAQH As String * 10
    sCntrAQHOV As String * 10
    sCntrAQHDP As String * 10
    sACT1DP As String * 20
    sPkgDP As String * 20
    sCntrDP As String * 20
End Type
Private tmGridInfo() As GRIDINFO

'Grid Name

Const CACT1INDEX = 0
Const CPACKAGEINDEX = 1
Const CCONTRACTINDEX = 2
Const UACT1INDEX = 3
Const UPACKAGEINDEX = 4
Const UCONTRACTINDEX = 5
Const DACT1INDEX = 6
Const DPACKAGEINDEX = 7
Const DCONTRACTINDEX = 8
Const AACT1INDEX = 9
Const APACKAGEINDEX = 10
Const ACONTRACTOVINDEX = 11
Const ACONTRACTDPINDEX = 12
Const GRIDINFOROWINDEX = 13
Const SORTINDEX = 14


Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long

    'Blank rows within grid
'    gGrid_Clear grdCompare, True
    'Set color within cells
    grdCompare.RowHeight(0) = fgBoxGridH + 15
    For llRow = grdCompare.FixedRows To grdCompare.Rows - 1 Step 1
        For llCol = 0 To SORTINDEX Step 1
            grdCompare.Row = llRow
            grdCompare.Col = llCol
            grdCompare.CellBackColor = WHITE
            grdCompare.CellForeColor = BLACK
            grdCompare.TextMatrix(llRow, llCol) = ""
        Next llCol
        grdCompare.RowHeight(llRow) = fgBoxGridH + 15
    Next llRow
End Sub

Private Sub cbcACT1PkgName_Change()
    mSetCommands
End Sub

Private Sub cbcACT1PkgName_Click()
    mSetCommands
End Sub

Private Sub cbcLine_Change()
    mSetCommands
End Sub

Private Sub cbcLine_Click()
    mSetCommands
End Sub

Private Sub cbcStandard_Change()
    mSetCommands
End Sub

Private Sub cbcStandard_Click()
    mSetCommands
End Sub

Private Sub cbcVer_Change()
    cbcLine.Clear
    mSetCommands
End Sub

Private Sub cbcVer_Click()
    cbcLine.Clear
    mSetCommands
End Sub

Private Sub cbcVer_LostFocus()
    mPopLine
    mSetCommands
End Sub

Private Sub ckcCompare_Click(Index As Integer)
    Select Case Index
        Case 0  'ACT1 Line-up
            If ckcCompare(Index).Value = vbChecked Then
                lacACT1File.Enabled = True
                edcACT1File.Enabled = True
                cmcBrowser.Enabled = True
                cbcACT1PkgName.Enabled = True
                'disallow audience if ACT1 not included, otherwise we would have to ask which Demo
                ckcInclude(2).Enabled = True
            Else
                lacACT1File.Enabled = False
                edcACT1File.Enabled = False
                edcACT1File.Text = ""
                cmcBrowser.Enabled = False
                cbcACT1PkgName.Enabled = False
                cbcACT1PkgName.Text = ""
                ckcInclude(2).Enabled = False
                ckcInclude(2).Value = vbUnchecked
            End If
        Case 1  'Package
            If ckcCompare(Index).Value = vbChecked Then
                lacStandard.Enabled = True
                cbcStandard.Enabled = True
            Else
                lacStandard.Enabled = False
                cbcStandard.Enabled = False
                cbcStandard.Text = ""
            End If
        Case 2  'Contract
            If ckcCompare(Index).Value = vbChecked Then
                lacCntrNo.Enabled = True
                edcCntrNo.Enabled = True
                lacVer.Enabled = True
                cbcVer.Enabled = True
                lacLine.Enabled = True
                cbcLine.Enabled = True
            Else
                lacCntrNo.Enabled = False
                edcCntrNo.Enabled = False
                edcCntrNo.Text = ""
                lacVer.Enabled = False
                cbcVer.Enabled = False
                cbcVer.Text = ""
                lacLine.Enabled = False
                cbcLine.Enabled = False
                cbcLine.Text = ""
            End If
    End Select
    mSetCommands
End Sub

Private Sub ckcInclude_Click(Index As Integer)
    If ckcInclude(2).Value = vbChecked Then
        lacBook.Enabled = True
        cbcBook.Enabled = True
    Else
        lacBook.Enabled = False
        cbcBook.Enabled = False
        cbcBook.ListIndex = -1
    End If
    mSetCommands
End Sub

Private Sub cmcBrowser_Click()
    'igBrowserType = 8  'PDF
    'sgBrowseMaskFile = ""
    'sgBrowserDrivePath = Left(sgImportPath, Len(sgImportPath) - 1)
    'sgBrowserTitle = "ACT1 Line-up PDF"
    'Browser.Show vbModal
    'If igBrowserReturn = 1 Then
    '    'Pop daypart names
    '    edcACT1File.Text = sgBrowserFile
    'End If
    cdcSetup.flags = cdlOFNFileMustExist + cdlOFNLongNames
    cdcSetup.fileName = edcACT1File.Text
    cdcSetup.InitDir = Left$(sgImportPath, Len(sgImportPath) - 1)
    cdcSetup.Filter = "PDF(*.pdf)|*.pdf"
    cdcSetup.Action = 1    'Open
    edcACT1File.Text = cdcSetup.fileName
    If gFileExist(edcACT1File.Text) = 0 Then
        gShellAndWait StdPkgVsACT1, sgExePath & "PDFToText.exe" & " -table -clip -eol DOS " & """" & edcACT1File.Text & """", vbMinimizedFocus, True    'vbTrue
        mPopPackage
    End If
    mSetCommands
End Sub



Private Sub cmcBrowser_LostFocus()
    gChDrDir
End Sub

Private Sub cmcCompare_Click()
    gSetMousePointer grdCompare, grdCompare, vbHourglass
    grdCompare.Redraw = False
    smDemo = ""
    imMnfDemo = 0
    mSetGridColumns
    mSetGridTitles
    gGrid_IntegralHeight grdCompare, fgBoxGridH + 15
    mClearGrid
    mGetACT1Lineup
    mGetPkgInfo
    mGetCntrInfo
    mSortGridInfo
    mPopulate
    grdCompare.Redraw = True
    mSetCommands
    gSetMousePointer grdCompare, grdCompare, vbDefault
End Sub

Private Sub cmcDone_Click()
    mTerminate
End Sub

Private Sub edcACT1File_Change()
    mSetCommands
End Sub

Private Sub edcCntrNo_Change()
    cbcVer.Clear
    cbcLine.Clear
    mSetCommands
End Sub

Private Sub edcCntrNo_LostFocus()
    mPopVersion
    mSetCommands
End Sub

Private Sub Form_Activate()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                                                                                 *
'******************************************************************************************


    If imFirstTime Then
        imFirstTime = False
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    'Me.Width = (CLng(45) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.Width = (CLng(60) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    Me.Height = (CLng(90) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    gCenterStdAlone StdPkgVsACT1
    DoEvents
    mSetControls
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    mInit
    Screen.MousePointer = vbDefault
    Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer

    On Error Resume Next
    ilRet = btrClose(hmDrf)
    btrDestroy hmDrf
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    ilRet = btrClose(hmDpf)
    btrDestroy hmDpf
    ilRet = btrClose(hmDef)
    btrDestroy hmDef
    ilRet = btrClose(hmRaf)
    btrDestroy hmRaf
    
    rst_vef.Close
    rst_pvf.Close
    rst_Chf.Close
    rst_Clf.Close
    rst_Cff.Close
    rst_dnf.Close
    
    Erase tmPkgVehicle
    Erase tmGridInfo
    Erase tmPvf
    'ChDir$ sgCurDir
    gChDrDir
    Set StdPkgVsACT1 = Nothing
End Sub


Private Sub mInit()
    Dim ilRet As Integer
    Dim ilAdf As Integer

    gSetMousePointer grdCompare, grdCompare, vbHourglass
    'frcSelection.Caption = "Compare Standard Package " & Trim$(sgStdPkgName) & " against ACT1 Line-up Definition"
    frcSelection.Caption = "Compare Standard Package against ACT1 Line-up Definition"
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    smNowDate = Format$(gNow(), "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    lmLastClickedRow = -1
    lmScrollTop = grdCompare.FixedRows
    imCompareColSorted = -1
    imCompareSort = -1
    mClearGrid
    
    hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "DRF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: DRF.Btr)", StdPkgVsACT1
    On Error GoTo 0
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "MNF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: MNF.Btr)", StdPkgVsACT1
    On Error GoTo 0
    hmDpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Dpf.Btr)", StdPkgVsACT1
    On Error GoTo 0
    ' setup global variable for Demo Plus file (to see if any exists)
    lgDpfNoRecs = btrRecords(hmDpf)
    If lgDpfNoRecs = 0 Then
        lgDpfNoRecs = -1
    End If
    hmDef = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Def.Btr)", StdPkgVsACT1
    On Error GoTo 0
    hmRaf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Raf.Btr)", StdPkgVsACT1
    On Error GoTo 0
    
    smDemo = ""
    mPopDemo
    mPopStandard
    mPopBook

    Screen.MousePointer = vbDefault
    gSetMousePointer grdCompare, grdCompare, vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    gSetMousePointer grdCompare, grdCompare, vbDefault
    Exit Sub

End Sub

Private Sub mPopulate()
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilCol As Integer
    Dim ilIndex As Integer

    On Error GoTo ErrHand:

    grdCompare.Row = 0
'    For llCol = GRIDNAMEINDEX To ENTEREDDATEINDEX Step 1
'        grdCompare.Col = llCol
'        grdCompare.CellBackColor = vbBlack
'        grdCompare.CellBackColor = LIGHTBLUE
'    Next llCol
    grdCompare.RowHeight(0) = fgBoxGridH + 15
    llRow = grdCompare.FixedRows
    For ilLoop = 0 To UBound(tmGridInfo) - 1 Step 1
        If llRow >= grdCompare.Rows Then
            grdCompare.AddItem ""
        End If
        grdCompare.RowHeight(llRow) = fgBoxGridH + 15
        grdCompare.TextMatrix(llRow, CACT1INDEX) = Trim$(tmGridInfo(ilLoop).sACT1CallLetters)
        grdCompare.TextMatrix(llRow, CPACKAGEINDEX) = Trim$(tmGridInfo(ilLoop).sPkgCallLetters)
        grdCompare.TextMatrix(llRow, CCONTRACTINDEX) = Trim$(tmGridInfo(ilLoop).sCntrCallLetters)
        If Trim$(tmGridInfo(ilLoop).sACT1CallLetters) <> "" Then
            grdCompare.TextMatrix(llRow, UACT1INDEX) = Trim$(tmGridInfo(ilLoop).iACT1Units)
            grdCompare.TextMatrix(llRow, DACT1INDEX) = Trim$(tmGridInfo(ilLoop).sACT1DP)
            grdCompare.TextMatrix(llRow, AACT1INDEX) = Trim$(tmGridInfo(ilLoop).sACT1AQH)
        End If
        If Trim$(tmGridInfo(ilLoop).sPkgCallLetters) <> "" Then
            grdCompare.TextMatrix(llRow, UPACKAGEINDEX) = Trim$(tmGridInfo(ilLoop).iPkgUnits)
            grdCompare.TextMatrix(llRow, DPACKAGEINDEX) = Trim$(tmGridInfo(ilLoop).sPkgDP)
            grdCompare.TextMatrix(llRow, APACKAGEINDEX) = Trim$(tmGridInfo(ilLoop).sPkgAQH)
        End If
        If (Trim$(tmGridInfo(ilLoop).sCntrCallLetters) <> "") And (edcCntrNo.Text <> "") Then
            grdCompare.TextMatrix(llRow, UCONTRACTINDEX) = Trim$(tmGridInfo(ilLoop).iCntrUnits)
            grdCompare.TextMatrix(llRow, DCONTRACTINDEX) = Trim$(tmGridInfo(ilLoop).sCntrDP)
            grdCompare.TextMatrix(llRow, ACONTRACTOVINDEX) = Trim$(tmGridInfo(ilLoop).sCntrAQHOV)
            grdCompare.TextMatrix(llRow, ACONTRACTDPINDEX) = Trim$(tmGridInfo(ilLoop).sCntrAQHDP)
        End If
        grdCompare.TextMatrix(llRow, GRIDINFOROWINDEX) = ilLoop
        llRow = llRow + 1
    Next ilLoop
    For llRow = grdCompare.FixedRows To grdCompare.Rows - 1 Step 1
        slStr = Trim$(grdCompare.TextMatrix(llRow, CACT1INDEX)) & Trim$(grdCompare.TextMatrix(llRow, CPACKAGEINDEX)) & Trim$(grdCompare.TextMatrix(llRow, CCONTRACTINDEX))
        If slStr <> "" Then
            ilIndex = grdCompare.TextMatrix(llRow, GRIDINFOROWINDEX)
            ''For ilCol = CACT1INDEX To DPACKAGEINDEX Step 1
            'For ilCol = CACT1INDEX To UCONTRACTINDEX Step 1
            For ilCol = CACT1INDEX To ACONTRACTDPINDEX Step 1
                grdCompare.Row = llRow
                grdCompare.Col = ilCol
                grdCompare.CellBackColor = LIGHTYELLOW
                If ilCol <= UCONTRACTINDEX Then
                    If (Trim$(tmGridInfo(ilIndex).sPkgCallLetters) = "") And (ckcCompare(1).Value = vbChecked) Then
                        grdCompare.CellForeColor = vbRed
                    ElseIf (Trim$(tmGridInfo(ilIndex).sACT1CallLetters) = "") And (ckcCompare(0).Value = vbChecked) Then
                        grdCompare.CellForeColor = vbRed
                    ElseIf (Trim$(tmGridInfo(ilIndex).sCntrCallLetters) = "") And (edcCntrNo.Text <> "") Then
                        grdCompare.CellForeColor = vbRed
                    Else
                        If (ckcCompare(0).Value = vbChecked) And (ckcCompare(1).Value = vbChecked) Then
                            If tmGridInfo(ilIndex).iACT1Units <> tmGridInfo(ilIndex).iPkgUnits Then
                                grdCompare.CellForeColor = vbRed
                            End If
                        End If
                    End If
                End If
                If (ckcInclude(1).Value = vbChecked) And (ilCol >= DACT1INDEX) And (ilCol <= DCONTRACTINDEX) Then
                    If (ckcCompare(0).Value = vbChecked) And (ckcCompare(1).Value = vbChecked) Then
                        If (Trim$(tmGridInfo(ilIndex).sACT1DP) <> "") And (Trim$(tmGridInfo(ilIndex).sPkgDP) <> "") Then
                            If Trim$(tmGridInfo(ilIndex).sACT1DP) <> Trim$(tmGridInfo(ilIndex).sPkgDP) Then
                                grdCompare.CellForeColor = MAGENTA  'ORANGE
                            End If
                        End If
                    End If
                    If (ckcCompare(0).Value = vbChecked) And (ckcCompare(2).Value = vbChecked) Then
                        If (Trim$(tmGridInfo(ilIndex).sACT1DP) <> "") And (Trim$(tmGridInfo(ilIndex).sCntrDP) <> "") Then
                            If Trim$(tmGridInfo(ilIndex).sACT1DP) <> Trim$(tmGridInfo(ilIndex).sCntrDP) Then
                                grdCompare.CellForeColor = MAGENTA  'ORANGE
                            End If
                        End If
                    End If
                    If (ckcCompare(1).Value = vbChecked) And (ckcCompare(2).Value = vbChecked) Then
                        If (Trim$(tmGridInfo(ilIndex).sPkgDP) <> "") And (Trim$(tmGridInfo(ilIndex).sCntrDP) <> "") Then
                            If Trim$(tmGridInfo(ilIndex).sPkgDP) <> Trim$(tmGridInfo(ilIndex).sCntrDP) Then
                                grdCompare.CellForeColor = MAGENTA  'ORANGE
                            End If
                        End If
                    End If
                End If
            Next ilCol
        End If
    Next llRow
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    Resume Next
    On Error GoTo 0

End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    Dim llWidth As Long
    Dim llMinWidth As Long
    Dim flAdj As Single
    Dim ilColCount As Integer
    Dim ilColInc As Integer
    Dim ilLoop As Integer
    Dim llAdj As Long
    
    grdCompare.ColWidth(SORTINDEX) = 0
    grdCompare.ColWidth(GRIDINFOROWINDEX) = 0
    ilColCount = 0
    If ckcCompare(0).Value = vbChecked Then ilColCount = ilColCount + 1
    If ckcCompare(1).Value = vbChecked Then ilColCount = ilColCount + 1
    If ckcCompare(2).Value = vbChecked Then ilColCount = ilColCount + 1
    If ckcInclude(0).Value = vbChecked Then 'Units
        If ckcCompare(0).Value = vbChecked Then ilColCount = ilColCount + 1
        If ckcCompare(1).Value = vbChecked Then ilColCount = ilColCount + 1
        If ckcCompare(2).Value = vbChecked Then ilColCount = ilColCount + 1
    End If
    If ckcInclude(1).Value = vbChecked Then 'Daypart
        If ckcCompare(0).Value = vbChecked Then ilColCount = ilColCount + 2 'Daypart twice the size
        If ckcCompare(1).Value = vbChecked Then ilColCount = ilColCount + 2
        If ckcCompare(2).Value = vbChecked Then ilColCount = ilColCount + 2
    End If
    If ckcInclude(2).Value = vbChecked Then 'Audience
        If ckcCompare(0).Value = vbChecked Then ilColCount = ilColCount + 1
        If ckcCompare(1).Value = vbChecked Then ilColCount = ilColCount + 1
        If ckcCompare(2).Value = vbChecked Then ilColCount = ilColCount + 2 'OV + DP columns
    End If
    flAdj = 1# / ilColCount
    llAdj = 100 * flAdj - 1
    flAdj = llAdj / 100
    If ckcCompare(0).Value = vbChecked Then 'ACT1
        grdCompare.ColWidth(CACT1INDEX) = grdCompare.Width * flAdj
        grdCompare.ColWidth(UACT1INDEX) = grdCompare.Width * flAdj
        grdCompare.ColWidth(DACT1INDEX) = 2 * grdCompare.Width * flAdj
        grdCompare.ColWidth(AACT1INDEX) = grdCompare.Width * flAdj
    Else
        grdCompare.ColWidth(CACT1INDEX) = 0
        grdCompare.ColWidth(UACT1INDEX) = 0
        grdCompare.ColWidth(DACT1INDEX) = 0
        grdCompare.ColWidth(AACT1INDEX) = 0
    End If
    If ckcCompare(1).Value = vbChecked Then 'Package
        grdCompare.ColWidth(CPACKAGEINDEX) = grdCompare.Width * flAdj
        grdCompare.ColWidth(UPACKAGEINDEX) = grdCompare.Width * flAdj
        grdCompare.ColWidth(DPACKAGEINDEX) = 2 * grdCompare.Width * flAdj
        grdCompare.ColWidth(APACKAGEINDEX) = grdCompare.Width * flAdj
    Else
        grdCompare.ColWidth(CPACKAGEINDEX) = 0
        grdCompare.ColWidth(UPACKAGEINDEX) = 0
        grdCompare.ColWidth(DPACKAGEINDEX) = 0
        grdCompare.ColWidth(APACKAGEINDEX) = 0
    End If
    If ckcCompare(2).Value = vbChecked Then 'Contract
        grdCompare.ColWidth(CCONTRACTINDEX) = grdCompare.Width * flAdj
        grdCompare.ColWidth(UCONTRACTINDEX) = grdCompare.Width * flAdj
        grdCompare.ColWidth(DCONTRACTINDEX) = 2 * grdCompare.Width * flAdj
        grdCompare.ColWidth(ACONTRACTOVINDEX) = grdCompare.Width * flAdj
        grdCompare.ColWidth(ACONTRACTDPINDEX) = grdCompare.Width * flAdj
    Else
        grdCompare.ColWidth(CCONTRACTINDEX) = 0
        grdCompare.ColWidth(UCONTRACTINDEX) = 0
        grdCompare.ColWidth(DCONTRACTINDEX) = 0
        grdCompare.ColWidth(ACONTRACTOVINDEX) = 0
        grdCompare.ColWidth(ACONTRACTDPINDEX) = 0
    End If
    If ckcInclude(0).Value = vbUnchecked Then
        grdCompare.ColWidth(UACT1INDEX) = 0
        grdCompare.ColWidth(UPACKAGEINDEX) = 0
        grdCompare.ColWidth(UCONTRACTINDEX) = 0
    End If
    If ckcInclude(1).Value = vbUnchecked Then
        grdCompare.ColWidth(DACT1INDEX) = 0
        grdCompare.ColWidth(DPACKAGEINDEX) = 0
        grdCompare.ColWidth(DCONTRACTINDEX) = 0
    End If
    If ckcInclude(2).Value = vbUnchecked Then
        grdCompare.ColWidth(AACT1INDEX) = 0
        grdCompare.ColWidth(APACKAGEINDEX) = 0
        grdCompare.ColWidth(ACONTRACTOVINDEX) = 0
        grdCompare.ColWidth(ACONTRACTDPINDEX) = 0
    End If
    
    'llWidth = 0
    'For ilCol = CACT1INDEX To DCONTRACTINDEX Step 1
    '    llWidth = llWidth + grdCompare.ColWidth(ilCol)
    'Next ilCol
    'If ckcCompare(0).Value = vbChecked Then
    '    grdCompare.ColWidth(CACT1INDEX) = grdCompare.Width - llWidth - GRIDSCROLLWIDTH - 15
    'Else
    '    grdCompare.ColWidth(UACT1INDEX) = grdCompare.Width - llWidth - GRIDSCROLLWIDTH - 15
    'End If
    llWidth = GRIDSCROLLWIDTH + 45
    llMinWidth = grdCompare.Width
    For ilCol = 0 To grdCompare.Cols - 1 Step 1
        llWidth = llWidth + grdCompare.ColWidth(ilCol)
        If (grdCompare.ColWidth(ilCol) > 15) And (grdCompare.ColWidth(ilCol) < llMinWidth) Then
            llMinWidth = grdCompare.ColWidth(ilCol)
        End If
    Next ilCol
    llWidth = grdCompare.Width - llWidth
    If llWidth >= 15 Then
        Do
            llMinWidth = grdCompare.Width
            For ilCol = 0 To grdCompare.Cols - 1 Step 1
                If (grdCompare.ColWidth(ilCol) > 15) And (grdCompare.ColWidth(ilCol) < llMinWidth) Then
                    llMinWidth = grdCompare.ColWidth(ilCol)
                End If
            Next ilCol
            For ilCol = grdCompare.FixedCols To grdCompare.Cols - 1 Step 1
                If grdCompare.ColWidth(ilCol) > 15 Then
                    ilColInc = grdCompare.ColWidth(ilCol) / llMinWidth
                    For ilLoop = 1 To ilColInc Step 1
                        grdCompare.ColWidth(ilCol) = grdCompare.ColWidth(ilCol) + 15
                        llWidth = llWidth - 15
                        If llWidth < 15 Then
                            Exit For
                        End If
                    Next ilLoop
                    If llWidth < 15 Then
                        Exit For
                    End If
                End If
            Next ilCol
        Loop While llWidth >= 15
    End If
    grdCompare.Visible = True
    'Align columns to left
    gGrid_AlignAllColsLeft grdCompare
End Sub

Private Sub mSetGridTitles()
    'Set column titles
    grdCompare.TextMatrix(0, CACT1INDEX) = "Vehicle"
    grdCompare.TextMatrix(0, CPACKAGEINDEX) = "Vehicle"
    grdCompare.TextMatrix(0, CCONTRACTINDEX) = "Vehicle"
    grdCompare.TextMatrix(1, CACT1INDEX) = "ACT1"
    grdCompare.TextMatrix(1, CPACKAGEINDEX) = "Package"
    grdCompare.TextMatrix(1, CCONTRACTINDEX) = "Contract"
    grdCompare.TextMatrix(0, UACT1INDEX) = "Units"
    grdCompare.TextMatrix(0, UPACKAGEINDEX) = "Units"
    grdCompare.TextMatrix(0, UCONTRACTINDEX) = "Units"
    grdCompare.TextMatrix(1, UACT1INDEX) = "ACT1"
    grdCompare.TextMatrix(1, UPACKAGEINDEX) = "Package"
    grdCompare.TextMatrix(1, UCONTRACTINDEX) = "Contract"
    grdCompare.TextMatrix(0, DACT1INDEX) = "Daypart"
    grdCompare.TextMatrix(0, DPACKAGEINDEX) = "Daypart"
    grdCompare.TextMatrix(0, DCONTRACTINDEX) = "Daypart"
    grdCompare.TextMatrix(1, DACT1INDEX) = "ACT1"
    grdCompare.TextMatrix(1, DPACKAGEINDEX) = "Package"
    grdCompare.TextMatrix(1, DCONTRACTINDEX) = "Contract"
    grdCompare.TextMatrix(0, AACT1INDEX) = "Audience"
    grdCompare.TextMatrix(0, APACKAGEINDEX) = "Audience"
    grdCompare.TextMatrix(0, ACONTRACTOVINDEX) = "Audience"
    grdCompare.TextMatrix(0, ACONTRACTDPINDEX) = "Audience"
    grdCompare.TextMatrix(1, AACT1INDEX) = "ACT1"
    grdCompare.TextMatrix(1, APACKAGEINDEX) = "Package"
    grdCompare.TextMatrix(1, ACONTRACTOVINDEX) = "Contr: OV"
    grdCompare.TextMatrix(1, ACONTRACTDPINDEX) = "Contr: DP"
    grdCompare.Row = 0
    grdCompare.MergeCells = 2    'flexMergeRestrictColumns
    grdCompare.MergeRow(0) = True

End Sub

Private Sub mCompareGridSortCol(ilCol As Integer)
    Dim llRow As Long
    Dim slStr As String
    Dim slSort As String
    Dim ilPos As Integer
    Dim slRow As String

'    For llRow = grdCompare.FixedRows To grdCompare.Rows - 1 Step 1
'        slStr = Trim$(grdCompare.TextMatrix(llRow, GRIDNAMEINDEX))
'        If slStr <> "" Then
'            If ilCol = ENTEREDDATEINDEX Then
'                slStr = grdCompare.TextMatrix(llRow, ENTEREDDATEINDEX)
'                If slStr <> "" Then
'                    slSort = Trim$(str$(gDateValue(slStr)))
'                Else
'                    slSort = "0"
'                End If
'                Do While Len(slSort) < 6
'                    slSort = "0" & slSort
'                Loop
'            Else
'                slSort = UCase$(Trim$(grdCompare.TextMatrix(llRow, ilCol)))
'            End If
'            slStr = grdCompare.TextMatrix(llRow, SORTINDEX)
'            ilPos = InStr(1, slStr, "|", vbTextCompare)
'            If ilPos > 1 Then
'                slStr = Left$(slStr, ilPos - 1)
'            End If
'            If (ilCol <> imCompareColSorted) Or ((ilCol = imCompareColSorted) And (imCompareSort = flexSortStringNoCaseDescending)) Then
'                slRow = Trim$(str$(llRow))
'                Do While Len(slRow) < 4
'                    slRow = "0" & slRow
'                Loop
'                grdCompare.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
'            Else
'                slRow = Trim$(str$(llRow))
'                Do While Len(slRow) < 4
'                    slRow = "0" & slRow
'                Loop
'                grdCompare.TextMatrix(llRow, SORTINDEX) = slSort & slStr & "|" & slRow
'            End If
'        End If
'    Next llRow
'    If ilCol = imCompareColSorted Then
'        imCompareColSorted = SORTINDEX
'    Else
'        imCompareColSorted = -1
'        imCompareSort = -1
'    End If
'    gGrid_SortByCol grdCompare, GRIDNAMEINDEX, SORTINDEX, imCompareColSorted, imCompareSort
'    imCompareColSorted = ilCol
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
    gSetMousePointer grdCompare, grdCompare, vbDefault
    igManUnload = YES
    Unload StdPkgVsACT1
    igManUnload = NO
End Sub


Private Sub mSetControls()

    Dim ilGap As Integer
    Dim flChg As Single
    Dim llChg As Long
    Dim llWidth As Long
    
    ilGap = frcSelection.Width - cmcBrowser.Left - cmcBrowser.Width
    flChg = (Me.Width - 2 * frcSelection.Left) / frcSelection.Width
    'llChg = 10000 * flChg
    'flChg = llChg / 10000
    flChg = flChg + 0.04
    frcSelection.Width = Me.Width - 3 * frcSelection.Left
    cmcBrowser.Left = frcSelection.Width - cmcBrowser.Width - ilGap
    edcACT1File.Width = cmcBrowser.Left - edcACT1File.Left - ilGap
    
    ilGap = lacVer.Left - (edcCntrNo.Left + edcCntrNo.Width)
    
    llWidth = (frcSelection.Width - lacACT1File.Left - lacACT1File.Width - lacStandard.Width - lacBook.Width - 6 * ilGap - 90) / 3
    cbcACT1PkgName.Width = llWidth  'flChg * cbcACT1PkgName.Width
    lacStandard.Left = cbcACT1PkgName.Left + cbcACT1PkgName.Width + ilGap
    cbcStandard.Width = llWidth 'flChg * cbcStandard.Width
    cbcStandard.Left = lacStandard.Left + lacStandard.Width + 60
    lacBook.Left = cbcStandard.Left + cbcStandard.Width + ilGap
    cbcBook.Width = llWidth 'flChg * cbcBook.Width
    cbcBook.Left = lacBook.Left + lacBook.Width + 60
    
    edcCntrNo.Width = flChg * edcCntrNo.Width
    lacVer.Left = edcCntrNo.Left + edcCntrNo.Width + ilGap
    cbcVer.Width = flChg * cbcVer.Width
    cbcVer.Left = lacVer.Left + lacVer.Width + 60
    lacLine.Left = cbcVer.Left + cbcVer.Width + ilGap
    cbcLine.Width = flChg * cbcLine.Width
    cbcLine.Left = lacLine.Left + lacLine.Width + 60
    
    ilGap = cmcCompare.Left - (cmcDone.Left + cmcDone.Width)
    
    cmcDone.Top = Me.Height - cmcDone.Height - 120
    cmcCompare.Top = cmcDone.Top
    cmcExport.Top = cmcDone.Top
    cmcCompare.Left = Me.Width / 2 - cmcCompare.Width / 2
    cmcDone.Left = cmcCompare.Left - cmcDone.Width - ilGap
    cmcExport.Left = cmcCompare.Left + cmcCompare.Width + ilGap
    grdCompare.Move frcSelection.Left, frcSelection.Top + frcSelection.Height + 90, Me.Width - 360, cmcDone.Top - (frcSelection.Top + frcSelection.Height + 90)

End Sub


Private Sub mPaintRowColor(llRow As Long)
    Dim llCol As Long

'    grdCompare.Row = llRow
'    For llCol = GRIDNAMEINDEX To ENTEREDDATEINDEX Step 1
'        grdCompare.Col = llCol
'        If grdCompare.TextMatrix(llRow, SELECTEDINDEX) <> "1" Then
'            grdCompare.CellBackColor = LIGHTYELLOW
'        Else
'            grdCompare.CellBackColor = GRAY    'vbBlue
'        End If
'    Next llCol

End Sub

Private Sub grdCompare_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    Dim llCurrentRow As Long
    Dim llTopRow As Long
    Dim llCol As Long
    Dim ilFound As Integer
    Dim slStr As String

'    If Y < grdCompare.RowHeight(0) Then
'        grdCompare.Col = grdCompare.MouseCol
'        mCompareGridSortCol grdCompare.Col
'        grdCompare.Row = 0
'        grdCompare.Col = GNFCODEINDEX
'        Exit Sub
'    End If
'    ilFound = gGrid_GetRowCol(grdCompare, X, Y, llCurrentRow, llCol)
'    If llCurrentRow < grdCompare.FixedRows Then
'        Exit Sub
'    End If
'    If llCurrentRow >= grdCompare.FixedRows Then
'        If grdCompare.TextMatrix(llCurrentRow, GRIDNAMEINDEX) <> "" Then
'            llTopRow = grdCompare.TopRow
'            For llRow = grdCompare.FixedRows To grdCompare.Rows - 1 Step 1
'                If grdCompare.TextMatrix(llRow, GRIDNAMEINDEX) <> "" Then
'                    If llRow = llCurrentRow Then
'                        grdCompare.TextMatrix(llRow, SELECTEDINDEX) = "1"
'                    Else
'                        grdCompare.TextMatrix(llRow, SELECTEDINDEX) = "0"
'                    End If
'                    mPaintRowColor llRow
'                End If
'            Next llRow
'            grdCompare.TopRow = llTopRow
'            grdCompare.Row = llCurrentRow
'        End If
'    End If

End Sub


Private Sub mPopPackage()
    Dim slTextFile As String
    Dim blHeaderFound As Boolean
    Dim ilCode As Integer
    Dim ilName As Integer
    Dim ilScheduleTime As Integer
    Dim ilUnreported As Integer
    Dim slStr As String
    Dim oMyFileObj As FileSystemObject
    Dim MyFile As TextStream
    Dim slLine As String
    Dim blFindDemo As Boolean
    
    smDemo = ""
    blFindDemo = False
    cbcACT1PkgName.Clear
    blHeaderFound = False
    slTextFile = edcACT1File.Text
    slTextFile = Replace(UCase(slTextFile), ".PDF", ".Txt")
    Set oMyFileObj = New FileSystemObject
    If oMyFileObj.FILEEXISTS(slTextFile) Then
        Set MyFile = oMyFileObj.OpenTextFile(slTextFile, ForReading, False)
        slLine = MyFile.ReadLine
        Do While Not MyFile.AtEndOfStream
            slLine = UCase(slLine)
            If slLine <> "" Then
                If InStr(1, slLine, "MARKET") > 0 Then
                    'continue searching for demo
                    Do While Not MyFile.AtEndOfStream
                        If MyFile.AtEndOfStream Then
                            Exit Do
                        End If
                        slLine = MyFile.ReadLine
                        slLine = UCase(slLine)
                        If slLine <> "" Then
                            mGetDemo slLine
                            If smDemo <> "" Then
                                Exit Do
                            End If
                        End If
                    Loop
                    Exit Do
                End If
                If blHeaderFound Then
                    slStr = Trim$(Mid(slLine, ilCode, ilName - ilCode))
                    If ilUnreported > 0 Then
                        slStr = slStr & ": " & Trim$(Mid(slLine, ilScheduleTime, ilUnreported - ilScheduleTime))
                    Else
                        slStr = slStr & ": " & Trim$(Mid(slLine, ilScheduleTime))
                    End If
                    cbcACT1PkgName.AddItem Trim$(slStr)
                ElseIf (InStr(1, slLine, "CODE") > 0) And (InStr(1, slLine, "SCHEDULE TIME") > 0) Then
                    blHeaderFound = True
                    ilCode = InStr(1, slLine, "CODE")
                    ilName = InStr(1, slLine, "NAME")
                    ilScheduleTime = InStr(1, slLine, "SCHEDULE TIME")
                    ilUnreported = InStr(1, slLine, "UNREPORTED")
                    If ilUnreported <= 0 Then
                
                    End If
                End If
            End If
            'Get next line
            If MyFile.AtEndOfStream Then
                Exit Do
            End If
            slLine = MyFile.ReadLine
        Loop
        MyFile.Close
        Set MyFile = Nothing
    End If
    If smDemo <> "" Then
        ckcInclude(2).Caption = "Audience (" & smDemo & ")"
    Else
        ckcInclude(2).Caption = "Audience"
    End If
End Sub

Private Sub mGetACT1Lineup()
    Dim slTextFile As String
    Dim blHeaderFound As Boolean
    Dim ilBook As Integer
    Dim ilLineup As Integer
    Dim ilAQH As Integer
    Dim ilRating As Integer
    Dim slStr As String
    Dim oMyFileObj As FileSystemObject
    Dim MyFile As TextStream
    Dim slLine As String
    Dim slPkgName As String
    Dim slDaypartDay As String
    Dim slDaypartTime As String
    Dim ilPos As Integer
    Dim llUpper As Long
    Dim slVehicle As String
    Dim llGridInfo As Long
    Dim blLineupWithSchedule As Boolean
    Dim ilDemo As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slAQH As String
    ReDim tmGridInfo(0 To 0) As GRIDINFO
    If ckcCompare(0).Value = vbUnchecked Then
        Exit Sub
    End If
    blHeaderFound = False
    slStr = cbcACT1PkgName.Text
    ilPos = InStr(1, slStr, ":")
    slPkgName = UCase(Mid(slStr, 1, ilPos - 1))
    slStr = Trim$(UCase(Mid(slStr, ilPos + 1)))
    ilPos = InStr(1, slStr, " ")
    slDaypartDay = Trim$(Left(slStr, ilPos - 1))
    slDaypartTime = Trim$(Mid(slStr, ilPos + 1))
    slTextFile = edcACT1File.Text
    slTextFile = Replace(UCase(slTextFile), ".PDF", ".Txt")
    Set oMyFileObj = New FileSystemObject
    If oMyFileObj.FILEEXISTS(slTextFile) Then
        Set MyFile = oMyFileObj.OpenTextFile(slTextFile, ForReading, False)
        slLine = MyFile.ReadLine
        Do While Not MyFile.AtEndOfStream
            slLine = UCase(slLine)
            If slLine <> "" Then
                If blHeaderFound Then
                    If (InStr(1, slLine, slPkgName) >= ilLineup) And (((InStr(1, slLine, slDaypartDay) > 0) And (InStr(1, slLine, slDaypartTime) > 0) And (blLineupWithSchedule = True)) Or (blLineupWithSchedule = False)) Then
                        llUpper = UBound(tmGridInfo)
                        ilPos = InStr(1, slLine, "-")
                        If ilPos > 0 And ilPos < ilBook Then
                            slVehicle = Trim(Left(slLine, ilPos + 3))
                            llGridInfo = mBinarySearch(slVehicle)
                            If llGridInfo = -1 Then
                                llUpper = UBound(tmGridInfo)
                                tmGridInfo(llUpper).sKey = slVehicle
                                tmGridInfo(llUpper).sACT1CallLetters = slVehicle
                                tmGridInfo(llUpper).iACT1Units = 1
                                ilPos = InStr(1, slLine, "X]")
                                If ilPos > 0 Then
                                    tmGridInfo(llUpper).iACT1Units = Mid(slLine, ilPos - 1, 1)
                                End If
                                slAQH = Trim(Mid(slLine, ilAQH, ilRating - ilAQH))
                                slAQH = Replace(slAQH, ",", "")
                                tmGridInfo(llUpper).sACT1AQH = slAQH
                                If blLineupWithSchedule Then
                                    slStr = Trim(Mid(slLine, ilBook - 20, 20))
                                    ilPos = InStr(slStr, "M")
                                    If ilPos <= 0 Then ilPos = 999
                                    If InStr(slStr, "T") < ilPos And InStr(slStr, "T") > 0 Then ilPos = InStr(slStr, "T")
                                    If InStr(slStr, "W") < ilPos And InStr(slStr, "W") > 0 Then ilPos = InStr(slStr, "W")
                                    If InStr(slStr, "T") < ilPos And InStr(slStr, "T") > 0 Then ilPos = InStr(slStr, "T")
                                    If InStr(slStr, "F") < ilPos And InStr(slStr, "F") > 0 Then ilPos = InStr(slStr, "F")
                                    If InStr(slStr, "S") < ilPos And InStr(slStr, "S") > 0 Then ilPos = InStr(slStr, "S")
                                    If InStr(slStr, "S") < ilPos And InStr(slStr, "S") > 0 Then ilPos = InStr(slStr, "S")
                                    If ilPos <> 999 Then
                                        tmGridInfo(llUpper).sACT1DP = Mid(slStr, ilPos, ilBook - ilPos)
                                    Else
                                        tmGridInfo(llUpper).sACT1DP = ""
                                    End If
                                Else
                                    tmGridInfo(llUpper).sACT1DP = slDaypartDay & " " & slDaypartTime
                                End If
                                tmGridInfo(llUpper).sPkgCallLetters = ""
                                tmGridInfo(llUpper).iPkgUnits = 0
                                tmGridInfo(llUpper).sPkgAQH = ""
                                tmGridInfo(llUpper).sPkgDP = ""
                                tmGridInfo(llUpper).sCntrCallLetters = ""
                                tmGridInfo(llUpper).iCntrUnits = 0
                                tmGridInfo(llUpper).sCntrAQHOV = ""
                                tmGridInfo(llUpper).sCntrAQHDP = ""
                                tmGridInfo(llUpper).sCntrDP = ""
                                ReDim Preserve tmGridInfo(0 To llUpper + 1) As GRIDINFO
                                If UBound(tmGridInfo) - 1 > 0 Then
                                    ArraySortTyp fnAV(tmGridInfo(), 0), UBound(tmGridInfo), 0, LenB(tmGridInfo(0)), 0, LenB(tmGridInfo(0).sKey), 0
                                End If
                            Else
                                slAQH = Trim(Mid(slLine, ilAQH, ilRating - ilAQH))
                                slAQH = Replace(slAQH, ",", "")
                                tmGridInfo(llGridInfo).sACT1AQH = gAddStr(tmGridInfo(llGridInfo).sACT1AQH, slAQH)
                            End If
                        End If
                    Else
                        If (InStr(1, slLine, "BOOK") > 0) And (InStr(1, slLine, "LINEUP") > 0) And (InStr(1, slLine, "AQH") > 0) Then
                            blHeaderFound = True
                            ilBook = InStr(1, slLine, "BOOK")
                            ilLineup = InStr(1, slLine, "LINEUP")
                            ilAQH = InStr(1, slLine, "AQH")
                            ilRating = InStr(1, slLine, "RTG")
                            blLineupWithSchedule = False
                            If (InStr(1, slLine, "/SCHEDULE") > 0) Then
                                blLineupWithSchedule = True
                            End If
                        End If
                    End If
                ElseIf (InStr(1, slLine, "BOOK") > 0) And (InStr(1, slLine, "LINEUP") > 0) And (InStr(1, slLine, "AQH") > 0) Then
                    blHeaderFound = True
                    ilBook = InStr(1, slLine, "BOOK")
                    ilLineup = InStr(1, slLine, "LINEUP")
                    ilAQH = InStr(1, slLine, "AQH")
                    ilRating = InStr(1, slLine, "RTG")
                    blLineupWithSchedule = False
                    If (InStr(1, slLine, "/SCHEDULE") > 0) Then
                        blLineupWithSchedule = True
                    End If
                ElseIf (smDemo = "") And (blHeaderFound = False) Then
                    mGetDemo slLine
                    If smDemo <> "" Then
                        For ilDemo = LBound(tgDemoCode) To UBound(tgDemoCode) - 1 Step 1
                            slNameCode = tgDemoCode(ilDemo).sKey
                            ilRet = gParseItem(slNameCode, 1, "\", slName)
                            If InStr(1, Trim$(UCase(slName)), smDemo) > 0 Then
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                imMnfDemo = Val(slCode)
                                Exit For
                            End If
                        Next ilDemo
                    End If
                End If
            End If
            'Get next line
            If MyFile.AtEndOfStream Then
                Exit Do
            End If
            slLine = MyFile.ReadLine
        Loop
        MyFile.Close
        Set MyFile = Nothing
    End If

End Sub
Private Sub mGetPkgInfo()
    Dim slVehicle As String
    Dim ilUnits As Integer
    Dim llRow As Long
    Dim ilCol As Integer
    Dim llGridInfo As Long
    Dim llUpper As Long
    Dim slSQLQuery As String
    Dim ilVefCode As Integer
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim slDayPart As String
    Dim llPvfCode As Long
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim slCode As String
    
    ReDim tmPvf(0 To 0) As PVF
    If cbcStandard.ListIndex < 0 Then
        Exit Sub
    End If
    slNameCode = tmPkgVehicle(cbcStandard.ListIndex).sKey 'lbcDPNameCode.List(ilLoop)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilVefCode = Val(slCode)
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef = -1 Then
        Exit Sub
    End If
    llPvfCode = tgMVef(ilVef).lPvfCode
    Do While llPvfCode > 0
        slSQLQuery = "Select * from pvf_Package_Vehicle where pvfCode = " & llPvfCode
        Set rst_pvf = gSQLSelectCall(slSQLQuery)
        If Not rst_pvf.EOF Then
            mCreateUDTForPVF
        End If
        llPvfCode = rst_pvf!pvfLkPvfCode
    Loop
    For llRow = 0 To UBound(tmPvf) - 1 Step 1
        slVehicle = Trim$(tmPvf(llRow).sName)
        For ilCol = LBound(tmPvf(llRow).iNoSpot) To UBound(tmPvf(llRow).iNoSpot) Step 1
            ilVefCode = tmPvf(llRow).iVefCode(ilCol)
            ilVef = gBinarySearchVef(ilVefCode)
            If ilVef <> -1 Then
                ilUnits = tmPvf(llRow).iNoSpot(ilCol)
                ilRdf = gBinarySearchRdf(tmPvf(llRow).iRdfCode(ilCol))
                If ilRdf <> -1 Then
                    slDayPart = Trim$(tgMRdf(ilRdf).sName)
                Else
                    slDayPart = ""
                End If
                slVehicle = Trim$(tgMVef(ilVef).sName)
                llGridInfo = mBinarySearch(slVehicle)
                If llGridInfo = -1 Then
                    llUpper = UBound(tmGridInfo)
                    tmGridInfo(llUpper).sKey = slVehicle
                    tmGridInfo(llUpper).sACT1CallLetters = ""
                    tmGridInfo(llUpper).iACT1Units = 0
                    tmGridInfo(llUpper).sACT1AQH = ""
                    tmGridInfo(llUpper).sACT1DP = ""
                    tmGridInfo(llUpper).sPkgCallLetters = slVehicle
                    tmGridInfo(llUpper).iPkgUnits = ilUnits
                    tmGridInfo(llUpper).sPkgAQH = mGetPkgAud(ilVefCode, tmPvf(llRow).iRdfCode(ilCol))
                    tmGridInfo(llUpper).sPkgDP = slDayPart
                    tmGridInfo(llUpper).sCntrCallLetters = ""
                    tmGridInfo(llUpper).iCntrUnits = 0
                    tmGridInfo(llUpper).sCntrAQHOV = ""
                    tmGridInfo(llUpper).sCntrAQHDP = ""
                    tmGridInfo(llUpper).sCntrDP = ""
                    ReDim Preserve tmGridInfo(0 To llUpper + 1) As GRIDINFO
                    If UBound(tmGridInfo) - 1 > 0 Then
                        ArraySortTyp fnAV(tmGridInfo(), 0), UBound(tmGridInfo), 0, LenB(tmGridInfo(0)), 0, LenB(tmGridInfo(0).sKey), 0
                    End If
                Else
                    'Merge
                    tmGridInfo(llGridInfo).sPkgCallLetters = slVehicle
                    tmGridInfo(llGridInfo).iPkgUnits = ilUnits
                    tmGridInfo(llGridInfo).sPkgDP = slDayPart
                    tmGridInfo(llGridInfo).sPkgAQH = mGetPkgAud(ilVefCode, tmPvf(llRow).iRdfCode(ilCol))
                End If
            End If
        Next ilCol
    Next llRow
End Sub
    

Private Function mBinarySearch(slName As String) As Long

    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim ilresult As Integer
    
    mBinarySearch = -1
    llMin = LBound(tmGridInfo)
    llMax = UBound(tmGridInfo) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        ilresult = StrComp(Trim(tmGridInfo(llMiddle).sKey), slName, vbBinaryCompare)
        Select Case ilresult
            Case 0:
                mBinarySearch = llMiddle
                Exit Function
            Case 1:
                llMax = llMiddle - 1
            Case -1:
                llMin = llMiddle + 1
        End Select
    Loop
    Exit Function
End Function

Private Sub mSortGridInfo()
    Dim llRow As Long
    For llRow = 0 To UBound(tmGridInfo) - 1 Step 1
        If (Trim(tmGridInfo(llRow).sACT1CallLetters) <> "") And (Trim$(tmGridInfo(llRow).sPkgCallLetters) = "") And (Trim$(tmGridInfo(llRow).sCntrCallLetters) = "") Then
            tmGridInfo(llRow).sKey = "A" & tmGridInfo(llRow).sACT1CallLetters
        ElseIf (Trim$(tmGridInfo(llRow).sPkgCallLetters) <> "") And (Trim$(tmGridInfo(llRow).sACT1CallLetters) = "") And (Trim$(tmGridInfo(llRow).sCntrCallLetters) = "") Then
            tmGridInfo(llRow).sKey = "B" & tmGridInfo(llRow).sPkgCallLetters
        ElseIf (Trim$(tmGridInfo(llRow).sCntrCallLetters) <> "") And (Trim$(tmGridInfo(llRow).sACT1CallLetters) = "") And (Trim$(tmGridInfo(llRow).sPkgCallLetters) = "") Then
            tmGridInfo(llRow).sKey = "C" & tmGridInfo(llRow).sCntrCallLetters
        ElseIf (ckcCompare(0).Value = vbChecked) And (ckcCompare(1).Value = vbChecked) And (ckcCompare(2).Value = vbUnchecked) And ((Trim(tmGridInfo(llRow).sACT1CallLetters) = "") Or (Trim(tmGridInfo(llRow).sPkgCallLetters) = "")) Then
            If Trim$(tmGridInfo(llRow).sACT1CallLetters) <> "" Then
                tmGridInfo(llRow).sKey = "D" & tmGridInfo(llRow).sACT1CallLetters
            Else
                tmGridInfo(llRow).sKey = "D" & tmGridInfo(llRow).sPkgCallLetters
            End If
        ElseIf (ckcCompare(0).Value = vbChecked) And (ckcCompare(1).Value = vbUnchecked) And (ckcCompare(2).Value = vbChecked) And ((Trim(tmGridInfo(llRow).sACT1CallLetters) = "") Or (Trim(tmGridInfo(llRow).sCntrCallLetters) = "")) Then
            If Trim$(tmGridInfo(llRow).sACT1CallLetters) <> "" Then
                tmGridInfo(llRow).sKey = "E" & tmGridInfo(llRow).sACT1CallLetters
            Else
                tmGridInfo(llRow).sKey = "E" & tmGridInfo(llRow).sPkgCallLetters
            End If
        ElseIf (ckcCompare(0).Value = vbUnchecked) And (ckcCompare(1).Value = vbChecked) And (ckcCompare(2).Value = vbChecked) And ((Trim(tmGridInfo(llRow).sPkgCallLetters) = "") Or (Trim(tmGridInfo(llRow).sCntrCallLetters) = "")) Then
            If Trim$(tmGridInfo(llRow).sACT1CallLetters) <> "" Then
                tmGridInfo(llRow).sKey = "F" & tmGridInfo(llRow).sACT1CallLetters
            Else
                tmGridInfo(llRow).sKey = "F" & tmGridInfo(llRow).sPkgCallLetters
            End If
        Else
            If Trim$(tmGridInfo(llRow).iACT1Units) <> tmGridInfo(llRow).iPkgUnits Then
                tmGridInfo(llRow).sKey = "H" & tmGridInfo(llRow).sACT1CallLetters
            Else
                tmGridInfo(llRow).sKey = "X" & tmGridInfo(llRow).sACT1CallLetters
            End If
        End If
    Next llRow
    If UBound(tmGridInfo) - 1 > 0 Then
        ArraySortTyp fnAV(tmGridInfo(), 0), UBound(tmGridInfo), 0, LenB(tmGridInfo(0)), 0, LenB(tmGridInfo(0).sKey), 0
    End If
End Sub

Private Sub cmcExport_Click()
    Dim hlExport As Integer
    Dim ilRet As Integer
    Dim slToFile As String
    Dim llRow As Long
    Dim llCol As Long
    Dim slStr As String
    Dim slStdPkgName As String
    Dim slOpenName As String
    Dim slDate As String
    Dim slTime As String
    Dim slFMonth As String
    Dim slFDay As String
    Dim slFYear As String
    
    slOpenName = ""
    
    If ckcCompare(0).Value = vbChecked Then
        slOpenName = "ACT1_Vs_"
    End If
    If ckcCompare(1).Value = vbChecked Then
        slOpenName = slOpenName & "Pkg_Vs_"
    End If
    If ckcCompare(2).Value = vbChecked Then
        slOpenName = slOpenName & "Cntr_Vs_"
    End If
    slOpenName = Left(slOpenName, Len(slOpenName) - 4)
    gCurrDateTime slDate, slTime, slFMonth, slFDay, slFYear
    slOpenName = slOpenName & "_" & slFMonth & slFDay & slFYear & "_" & Format(slTime, "HHMMSS")
    slToFile = sgExportPath & slOpenName & ".csv"
    ilRet = 0
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        Kill slToFile
    End If
    ilRet = 0
    ilRet = gFileOpen(slToFile, "Output", hlExport)
    If ilRet <> 0 Then
        MsgBox "Open " & slToFile & ", Error #" & Str(Err.Number), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
        Exit Sub
    End If
    If ckcCompare(0).Value = vbChecked Then
        slStr = "ACT1 Line-up File: " & Trim$(edcACT1File.Text)
        Print #hlExport, slStr
        slStr = "ACT1 Package: " & Trim$(cbcACT1PkgName.Text)
        Print #hlExport, slStr
    End If
    If ckcCompare(1).Value = vbChecked Then
        slStr = "Standard Package: " & Trim$(cbcStandard.Text)
        Print #hlExport, slStr
    End If
    If ckcCompare(2).Value = vbChecked Then
        slStr = "Contract: " & Trim$(edcCntrNo.Text)
        Print #hlExport, slStr
        slStr = "Version: " & Trim$(cbcVer.Text)
        Print #hlExport, slStr
        slStr = "Line: " & Trim$(cbcLine.Text)
        Print #hlExport, slStr
    End If
    If ckcInclude(2).Value = vbChecked Then
        slStr = "Demo: " & smDemo
        Print #hlExport, slStr
    End If
    Print #hlExport, ""
    slStr = "ACT1 Vehicle" & ","
    slStr = slStr & "Package Vehicle" & ","
    slStr = slStr & "Contract Vehicle" & ","
    slStr = slStr & "" '& ","
    slStr = slStr & "ACT1 Units" & ","
    slStr = slStr & "Package Units" & ","
    slStr = slStr & "Contract Units" & ","
    slStr = slStr & "" '& ","
    slStr = slStr & "ACT1 Daypart" & ","
    slStr = slStr & "Package Daypart" & ","
    slStr = slStr & "Contract Daypart" & ","
    slStr = slStr & "" '& ","
    slStr = slStr & "ACT1 Audience" & ","
    slStr = slStr & "Package Audience" & ","
    slStr = slStr & "Contract AQH Overrides" & ","
    slStr = slStr & "Contract AQH Daypart" & ","
    Print #hlExport, slStr
    For llRow = grdCompare.FixedRows To grdCompare.Rows - 1 Step 1
        slStr = Trim$(grdCompare.TextMatrix(llRow, CACT1INDEX)) & Trim$(grdCompare.TextMatrix(llRow, CPACKAGEINDEX)) & Trim$(grdCompare.TextMatrix(llRow, CCONTRACTINDEX))
        If slStr <> "" Then
            slStr = ""
            For llCol = CACT1INDEX To ACONTRACTDPINDEX Step 1
                slStr = slStr & Trim$(grdCompare.TextMatrix(llRow, llCol)) & ","
            Next llCol
            Print #hlExport, Left$(slStr, Len(slStr) - 1)
        End If
    Next llRow
    Close #hlExport
    MsgBox "Create file: " & slToFile, vbOKOnly + vbApplicationModal, "File Saved"
    Exit Sub
End Sub

Private Sub mSetCommands()
    Dim ilNoChk As Integer
    
    cmcCompare.Enabled = True
    ilNoChk = 0
    If ckcCompare(0).Value = vbChecked Then ilNoChk = ilNoChk + 1
    If ckcCompare(1).Value = vbChecked Then ilNoChk = ilNoChk + 1
    If ckcCompare(2).Value = vbChecked Then ilNoChk = ilNoChk + 1
    If ilNoChk < 2 Then
        cmcCompare.Enabled = False
    End If
    If ckcCompare(0).Value = vbChecked Then
        If edcACT1File.Text = "" Then
            cmcCompare.Enabled = False
        End If
        If cbcACT1PkgName.Text = "" Or cbcACT1PkgName.ListIndex < 0 Then
            cmcCompare.Enabled = False
        End If
    End If
    If ckcCompare(1).Value = vbChecked Then
        If cbcStandard.Text = "" Then
            cmcCompare.Enabled = False
        End If
    End If
    If ckcCompare(2).Value = vbChecked Then
        If edcCntrNo.Text = "" Then
            cmcCompare.Enabled = False
        End If
        If cbcVer.Text = "" Or cbcVer.ListIndex < 0 Then
            cmcCompare.Enabled = False
        End If
        If cbcLine.Text = "" Or cbcLine.ListIndex < 0 Then
            cmcCompare.Enabled = False
        End If
    End If
    If ckcInclude(2).Value = vbChecked Then
        If cbcBook.Text = "" Or cbcBook.ListIndex < 0 Then
            cmcCompare.Enabled = False
        End If
    End If
    If cmcCompare.Enabled Then
        If grdCompare.Visible Then
            cmcExport.Enabled = True
        Else
            cmcExport.Enabled = False
        End If
    Else
        cmcExport.Enabled = False
    End If
End Sub

Private Sub mPopStandard()
    Dim ilLoop As Integer
    Dim slName As String
    Dim slCode As String
    Dim ilRet As Integer
    
    cbcStandard.Clear
    ilRet = gPopUserVehicleBox(StdPkgVsACT1, VEHSTDPKG + ACTIVEVEH + DORMANTVEH, cbcStandard, tmPkgVehicle(), smPkgVehicleTag)
    'slName = StdPkg.cbcSelect.Text
    'gFindMatch slName, 0, cbcStandard
    'If gLastFound(cbcStandard) >= 0 Then
    '    cbcStandard.ListIndex = gLastFound(cbcStandard)
    'End If
    
End Sub


Private Sub mCreateUDTForPVF()
    Dim ilUpper As Integer
    ilUpper = UBound(tmPvf)
    tmPvf(ilUpper).lCode = rst_pvf!pvfCode
    tmPvf(ilUpper).sName = rst_pvf!pvfName
    tmPvf(ilUpper).lLkPvfCode = rst_pvf!pvfLkPvfCode
    tmPvf(ilUpper).iVefCode(0) = rst_pvf!pvfVefCode1
    tmPvf(ilUpper).iVefCode(1) = rst_pvf!pvfVefCode2
    tmPvf(ilUpper).iVefCode(2) = rst_pvf!pvfVefCode3
    tmPvf(ilUpper).iVefCode(3) = rst_pvf!pvfVefCode4
    tmPvf(ilUpper).iVefCode(4) = rst_pvf!pvfVefCode5
    tmPvf(ilUpper).iVefCode(5) = rst_pvf!pvfVefCode6
    tmPvf(ilUpper).iVefCode(6) = rst_pvf!pvfVefCode7
    tmPvf(ilUpper).iVefCode(7) = rst_pvf!pvfVefCode8
    tmPvf(ilUpper).iVefCode(8) = rst_pvf!pvfVefCode9
    tmPvf(ilUpper).iVefCode(9) = rst_pvf!pvfVefCode10
    tmPvf(ilUpper).iVefCode(10) = rst_pvf!pvfVefCode11
    tmPvf(ilUpper).iVefCode(11) = rst_pvf!pvfVefCode12
    tmPvf(ilUpper).iVefCode(12) = rst_pvf!pvfVefCode13
    tmPvf(ilUpper).iVefCode(13) = rst_pvf!pvfVefCode14
    tmPvf(ilUpper).iVefCode(14) = rst_pvf!pvfVefCode15
    tmPvf(ilUpper).iVefCode(15) = rst_pvf!pvfVefCode16
    tmPvf(ilUpper).iVefCode(16) = rst_pvf!pvfVefCode17
    tmPvf(ilUpper).iVefCode(17) = rst_pvf!pvfVefCode18
    tmPvf(ilUpper).iVefCode(18) = rst_pvf!pvfVefCode19
    tmPvf(ilUpper).iVefCode(19) = rst_pvf!pvfVefCode20
    tmPvf(ilUpper).iVefCode(20) = rst_pvf!pvfVefCode21
    tmPvf(ilUpper).iVefCode(21) = rst_pvf!pvfVefCode22
    tmPvf(ilUpper).iVefCode(22) = rst_pvf!pvfVefCode23
    tmPvf(ilUpper).iVefCode(23) = rst_pvf!pvfVefCode24
    tmPvf(ilUpper).iVefCode(24) = rst_pvf!pvfVefCode25
    tmPvf(ilUpper).iRdfCode(0) = rst_pvf!pvfRdfCode1
    tmPvf(ilUpper).iRdfCode(1) = rst_pvf!pvfRdfCode2
    tmPvf(ilUpper).iRdfCode(2) = rst_pvf!pvfRdfCode3
    tmPvf(ilUpper).iRdfCode(3) = rst_pvf!pvfRdfCode4
    tmPvf(ilUpper).iRdfCode(4) = rst_pvf!pvfRdfCode5
    tmPvf(ilUpper).iRdfCode(5) = rst_pvf!pvfRdfCode6
    tmPvf(ilUpper).iRdfCode(6) = rst_pvf!pvfRdfCode7
    tmPvf(ilUpper).iRdfCode(7) = rst_pvf!pvfRdfCode8
    tmPvf(ilUpper).iRdfCode(8) = rst_pvf!pvfRdfCode9
    tmPvf(ilUpper).iRdfCode(9) = rst_pvf!pvfRdfCode10
    tmPvf(ilUpper).iRdfCode(10) = rst_pvf!pvfRdfCode11
    tmPvf(ilUpper).iRdfCode(11) = rst_pvf!pvfRdfCode12
    tmPvf(ilUpper).iRdfCode(12) = rst_pvf!pvfRdfCode13
    tmPvf(ilUpper).iRdfCode(13) = rst_pvf!pvfRdfCode14
    tmPvf(ilUpper).iRdfCode(14) = rst_pvf!pvfRdfCode15
    tmPvf(ilUpper).iRdfCode(15) = rst_pvf!pvfRdfCode16
    tmPvf(ilUpper).iRdfCode(16) = rst_pvf!pvfRdfCode17
    tmPvf(ilUpper).iRdfCode(17) = rst_pvf!pvfRdfCode18
    tmPvf(ilUpper).iRdfCode(18) = rst_pvf!pvfRdfCode19
    tmPvf(ilUpper).iRdfCode(19) = rst_pvf!pvfRdfCode20
    tmPvf(ilUpper).iRdfCode(20) = rst_pvf!pvfRdfCode21
    tmPvf(ilUpper).iRdfCode(21) = rst_pvf!pvfRdfCode22
    tmPvf(ilUpper).iRdfCode(22) = rst_pvf!pvfRdfCode23
    tmPvf(ilUpper).iRdfCode(23) = rst_pvf!pvfRdfCode24
    tmPvf(ilUpper).iRdfCode(24) = rst_pvf!pvfRdfCode25
    tmPvf(ilUpper).iNoSpot(0) = rst_pvf!pvfNoSpot1
    tmPvf(ilUpper).iNoSpot(1) = rst_pvf!pvfNoSpot2
    tmPvf(ilUpper).iNoSpot(2) = rst_pvf!pvfNoSpot3
    tmPvf(ilUpper).iNoSpot(3) = rst_pvf!pvfNoSpot4
    tmPvf(ilUpper).iNoSpot(4) = rst_pvf!pvfNoSpot5
    tmPvf(ilUpper).iNoSpot(5) = rst_pvf!pvfNoSpot6
    tmPvf(ilUpper).iNoSpot(6) = rst_pvf!pvfNoSpot7
    tmPvf(ilUpper).iNoSpot(7) = rst_pvf!pvfNoSpot8
    tmPvf(ilUpper).iNoSpot(8) = rst_pvf!pvfNoSpot9
    tmPvf(ilUpper).iNoSpot(9) = rst_pvf!pvfNoSpot10
    tmPvf(ilUpper).iNoSpot(10) = rst_pvf!pvfNoSpot11
    tmPvf(ilUpper).iNoSpot(11) = rst_pvf!pvfNoSpot12
    tmPvf(ilUpper).iNoSpot(12) = rst_pvf!pvfNoSpot13
    tmPvf(ilUpper).iNoSpot(13) = rst_pvf!pvfNoSpot14
    tmPvf(ilUpper).iNoSpot(14) = rst_pvf!pvfNoSpot15
    tmPvf(ilUpper).iNoSpot(15) = rst_pvf!pvfNoSpot16
    tmPvf(ilUpper).iNoSpot(16) = rst_pvf!pvfNoSpot17
    tmPvf(ilUpper).iNoSpot(17) = rst_pvf!pvfNoSpot18
    tmPvf(ilUpper).iNoSpot(18) = rst_pvf!pvfNoSpot19
    tmPvf(ilUpper).iNoSpot(19) = rst_pvf!pvfNoSpot20
    tmPvf(ilUpper).iNoSpot(20) = rst_pvf!pvfNoSpot21
    tmPvf(ilUpper).iNoSpot(21) = rst_pvf!pvfNoSpot22
    tmPvf(ilUpper).iNoSpot(22) = rst_pvf!pvfNoSpot23
    tmPvf(ilUpper).iNoSpot(23) = rst_pvf!pvfNoSpot24
    tmPvf(ilUpper).iNoSpot(24) = rst_pvf!pvfNoSpot25
    tmPvf(ilUpper).iPctRate(0) = rst_pvf!pvfPctRate1
    tmPvf(ilUpper).iPctRate(1) = rst_pvf!pvfPctRate2
    tmPvf(ilUpper).iPctRate(2) = rst_pvf!pvfPctRate3
    tmPvf(ilUpper).iPctRate(3) = rst_pvf!pvfPctRate4
    tmPvf(ilUpper).iPctRate(4) = rst_pvf!pvfPctRate5
    tmPvf(ilUpper).iPctRate(5) = rst_pvf!pvfPctRate6
    tmPvf(ilUpper).iPctRate(6) = rst_pvf!pvfPctRate7
    tmPvf(ilUpper).iPctRate(7) = rst_pvf!pvfPctRate8
    tmPvf(ilUpper).iPctRate(8) = rst_pvf!pvfPctRate9
    tmPvf(ilUpper).iPctRate(9) = rst_pvf!pvfPctRate10
    tmPvf(ilUpper).iPctRate(10) = rst_pvf!pvfPctRate11
    tmPvf(ilUpper).iPctRate(11) = rst_pvf!pvfPctRate12
    tmPvf(ilUpper).iPctRate(12) = rst_pvf!pvfPctRate13
    tmPvf(ilUpper).iPctRate(13) = rst_pvf!pvfPctRate14
    tmPvf(ilUpper).iPctRate(14) = rst_pvf!pvfPctRate15
    tmPvf(ilUpper).iPctRate(15) = rst_pvf!pvfPctRate16
    tmPvf(ilUpper).iPctRate(16) = rst_pvf!pvfPctRate17
    tmPvf(ilUpper).iPctRate(17) = rst_pvf!pvfPctRate18
    tmPvf(ilUpper).iPctRate(18) = rst_pvf!pvfPctRate19
    tmPvf(ilUpper).iPctRate(19) = rst_pvf!pvfPctRate20
    tmPvf(ilUpper).iPctRate(20) = rst_pvf!pvfPctRate21
    tmPvf(ilUpper).iPctRate(21) = rst_pvf!pvfPctRate22
    tmPvf(ilUpper).iPctRate(22) = rst_pvf!pvfPctRate23
    tmPvf(ilUpper).iPctRate(23) = rst_pvf!pvfPctRate24
    tmPvf(ilUpper).iPctRate(24) = rst_pvf!pvfPctRate25
    ReDim Preserve tmPvf(0 To ilUpper + 1) As PVF
End Sub

Private Sub mPopLine()
    Dim slSQLQuery As String
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim llChfCode As Long
    
    cbcLine.Clear
    If edcCntrNo.Text = "" Then
        Exit Sub
    End If
    If cbcVer.ListIndex < 0 Then
        Exit Sub
    End If
    llChfCode = cbcVer.ItemData(cbcVer.ListIndex)
    slSQLQuery = "Select clfLine, clfRdfCode, clfChfCode, clfvefcode from clf_Contract_Line where clfchfcode = " & llChfCode
    slSQLQuery = slSQLQuery & " And clfType In('O', 'A', 'E')" & " Order By clfLine"
    Set rst_Clf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Clf.EOF
        ilVef = gBinarySearchVef(rst_Clf!clfVefCode)
        If ilVef <> -1 Then
            ilRdf = gBinarySearchRdf(rst_Clf!clfRdfCode)
            If ilRdf <> -1 Then
                cbcLine.AddItem rst_Clf!clfLine & " " & Trim$(tgMVef(ilVef).sName) & " " & Trim$(tgMRdf(ilRdf).sName)
                cbcLine.ItemData(cbcLine.NewIndex) = rst_Clf!clfChfCode
            End If
        End If
        rst_Clf.MoveNext
    Loop
    
End Sub

Private Sub mGetCntrInfo()
    Dim slVehicle As String
    Dim ilUnits As Integer
    Dim llRow As Long
    Dim ilCol As Integer
    Dim llGridInfo As Long
    Dim llUpper As Long
    Dim slSQLQuery As String
    Dim ilVefCode As Integer
    Dim ilVef As Integer
    Dim ilRdf As Integer
    Dim slDayPart As String
    Dim llChfCode As Long
    Dim slStr As String
    Dim ilPos As Integer
    Dim slPkLineNo As String
    Dim ilDay As Integer
    Dim slAQHOv As String
    Dim slAQHDp As String
    Dim ilDays(0 To 6) As Integer
    
    If cbcLine.ListIndex < 0 Then
        Exit Sub
    End If
    slStr = cbcLine.Text
    ilPos = InStr(1, slStr, " ")
    If ilPos <= 0 Then
        Exit Sub
    End If
    slPkLineNo = Left(slStr, ilPos - 1)
    llChfCode = cbcLine.ItemData(cbcLine.ListIndex)
    slSQLQuery = "Select * from clf_contract_line where clfChfCode = " & llChfCode
    slSQLQuery = slSQLQuery & " And clfPkLineNo = " & slPkLineNo
    Set rst_Clf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Clf.EOF
        ilVef = gBinarySearchVef(rst_Clf!clfVefCode)
        If ilVef <> -1 Then
            slVehicle = Trim$(tgMVef(ilVef).sName)
            ilRdf = gBinarySearchRdf(rst_Clf!clfRdfCode)
            If ilRdf <> -1 Then
                'Get units by week
                ilUnits = 0
                For ilDay = 0 To 6 Step 1
                    ilDays(ilDay) = False
                Next ilDay
                slSQLQuery = "Select sum(cffSpotsWk) as Units"
                slSQLQuery = slSQLQuery & " From cff_contract_Flight where cffChfCode = " & llChfCode
                slSQLQuery = slSQLQuery & " And cffClfLine = " & rst_Clf!clfLine
                slSQLQuery = slSQLQuery & " And cffDyWk = 'W'"
                Set rst_Cff = gSQLSelectCall(slSQLQuery)
                If Not rst_Cff.EOF Then
                    ilUnits = ilUnits + Val(rst_Cff!Units)
                End If
                'Get units by daily
                slSQLQuery = "Select sum(cffMo+cffTu+cffWe+cffTh+cffFr+cffSa+cffSu) as Units "
                slSQLQuery = slSQLQuery & " From cff_contract_Flight where cffChfCode = " & llChfCode
                slSQLQuery = slSQLQuery & " And cffClfLine = " & rst_Clf!clfLine
                slSQLQuery = slSQLQuery & " And cffDyWk = 'D'"
                Set rst_Cff = gSQLSelectCall(slSQLQuery)
                If Not rst_Cff.EOF Then
                    If rst_Cff!Units <> Null Then
                        ilUnits = ilUnits + Val(rst_Cff!Units)
                    End If
                End If
                slSQLQuery = "Select cffMo, cffTu, cffWe, cffTh, cffFr, cffSa, cffSu"
                slSQLQuery = slSQLQuery & " From cff_contract_Flight where cffChfCode = " & llChfCode
                slSQLQuery = slSQLQuery & " And cffClfLine = " & rst_Clf!clfLine
                Set rst_Cff = gSQLSelectCall(slSQLQuery)
                If Not rst_Cff.EOF Then
                    If rst_Cff!cffMo > 0 Then ilDays(0) = True
                    If rst_Cff!cffTu > 0 Then ilDays(1) = True
                    If rst_Cff!cffWe > 0 Then ilDays(2) = True
                    If rst_Cff!cffTh > 0 Then ilDays(3) = True
                    If rst_Cff!CffFr > 0 Then ilDays(4) = True
                    If rst_Cff!cffSa > 0 Then ilDays(5) = True
                    If rst_Cff!cffSu > 0 Then ilDays(6) = True
                End If
                
                slDayPart = Trim$(tgMRdf(ilRdf).sName)
                llGridInfo = mBinarySearch(slVehicle)
                If llGridInfo = -1 Then
                    llUpper = UBound(tmGridInfo)
                    tmGridInfo(llUpper).sKey = slVehicle
                    tmGridInfo(llUpper).sACT1CallLetters = ""
                    tmGridInfo(llUpper).iACT1Units = 0
                    tmGridInfo(llUpper).sACT1AQH = ""
                    tmGridInfo(llUpper).sACT1DP = ""
                    tmGridInfo(llUpper).sPkgCallLetters = ""
                    tmGridInfo(llUpper).iPkgUnits = 0
                    tmGridInfo(llUpper).sPkgAQH = ""
                    tmGridInfo(llUpper).sPkgDP = ""
                    tmGridInfo(llUpper).sCntrCallLetters = slVehicle
                    tmGridInfo(llUpper).iCntrUnits = ilUnits
                    mGetLnAud ilDays(), slAQHOv, slAQHDp
                    tmGridInfo(llUpper).sCntrAQHOV = slAQHOv
                    tmGridInfo(llUpper).sCntrAQHDP = slAQHDp
                    tmGridInfo(llUpper).sCntrDP = slDayPart
                    ReDim Preserve tmGridInfo(0 To llUpper + 1) As GRIDINFO
                    If UBound(tmGridInfo) - 1 > 0 Then
                        ArraySortTyp fnAV(tmGridInfo(), 0), UBound(tmGridInfo), 0, LenB(tmGridInfo(0)), 0, LenB(tmGridInfo(0).sKey), 0
                    End If
                Else
                    'Merge
                    tmGridInfo(llGridInfo).sCntrCallLetters = slVehicle
                    tmGridInfo(llGridInfo).iCntrUnits = ilUnits
                    tmGridInfo(llGridInfo).sCntrDP = slDayPart
                    mGetLnAud ilDays(), slAQHOv, slAQHDp
                    tmGridInfo(llGridInfo).sCntrAQHOV = slAQHOv
                    tmGridInfo(llGridInfo).sCntrAQHDP = slAQHDp
                End If
            End If
        End If
        rst_Clf.MoveNext
    Loop
End Sub

Private Sub mPopVersion()
    Dim slSQLQuery As String
    Dim slStr As String
    Dim slStatus As String
    Dim ilCntRevNo As Integer
    Dim ilVerNo As Integer
    Dim ilExtRevNo As Integer
    
    cbcVer.Clear
    If edcCntrNo.Text = "" Then
        Exit Sub
    End If
    slSQLQuery = "Select * from chf_Contract_Header where chfCntrNo = " & edcCntrNo.Text
    slSQLQuery = slSQLQuery & " Order By chfOHDDate Desc"
    Set rst_Chf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Chf.EOF
        slStatus = rst_Chf!chfStatus
        ilCntRevNo = rst_Chf!chfCntRevNo
        ilVerNo = rst_Chf!chfPropVer
        ilExtRevNo = rst_Chf!chfExtRevNo
        If (slStatus = "W") Or (slStatus = "C") Or (slStatus = "I") Or (slStatus = "D") Then
            If ilCntRevNo > 0 Then
                slStr = "R" & Trim$(Str$(ilCntRevNo)) & "-" & Trim$(Str$(ilExtRevNo))
            Else
                slStr = "V" & Trim$(Str$(ilVerNo))
            End If
        Else
            slStr = "R" & Trim$(Str$(ilCntRevNo)) & "-" & Trim$(Str$(ilExtRevNo))
        End If
    
        Select Case slStatus
            Case "W"
                If ilCntRevNo > 0 Then
                    slStr = slStr & " Rev Working"
                Else
                    slStr = slStr & " Working"
                End If
            Case "D"
                slStr = slStr & " Rejected"
            Case "C"
                If ilCntRevNo > 0 Then
                    slStr = slStr & " Rev Completed"
                Else
                    slStr = slStr & " Completed"
                End If
            Case "I"
                If ilCntRevNo > 0 Then
                    slStr = slStr & " Rev Unapproved"
                Else
                    slStr = slStr & " Unapproved"
                End If
            Case "G"
                slStr = slStr & " Approved Hold"
            Case "N"
                slStr = slStr & " Approved Order"
            Case "H"
                slStr = slStr & " Hold"
            Case "O"
                slStr = slStr & " Order"
        End Select
        slStr = slStr & " " & Format(rst_Chf!ChfStartDate, "m/d/yy") & "-" & Format(rst_Chf!ChfEndDate, "m/d/yy")
        cbcVer.AddItem slStr
        cbcVer.ItemData(cbcVer.NewIndex) = rst_Chf!chfCode
        rst_Chf.MoveNext
    Loop
End Sub

Private Sub mPopDemo()
    Dim ilRet As Integer
    
    tmDemo(0) = "12-17"
    tmDemo(1) = "12-20"
    tmDemo(2) = "12-24"
    tmDemo(3) = "12-34"
    tmDemo(4) = "12-44"
    tmDemo(5) = "12-49"
    tmDemo(6) = "12-54"
    tmDemo(7) = "12-64"
    tmDemo(8) = "12+"
    tmDemo(9) = "18-20"
    tmDemo(10) = "18-24"
    tmDemo(11) = "18-34"
    tmDemo(12) = "18-44"
    tmDemo(13) = "18-49"
    tmDemo(14) = "18-54"
    tmDemo(15) = "18-64"
    tmDemo(16) = "18+"
    tmDemo(17) = "21-24"
    tmDemo(18) = "21-34"
    tmDemo(19) = "21-44"
    tmDemo(20) = "21-49"
    tmDemo(21) = "21-54"
    tmDemo(22) = "21-64"
    tmDemo(23) = "21+"
    tmDemo(24) = "25-34"
    tmDemo(25) = "25-44"
    tmDemo(26) = "25-49"
    tmDemo(27) = "25-54"
    tmDemo(28) = "25-64"
    tmDemo(29) = "25+"
    tmDemo(30) = "35-44"
    tmDemo(31) = "35-49"
    tmDemo(32) = "35-54"
    tmDemo(33) = "35-64"
    tmDemo(34) = "35+"
    tmDemo(35) = "45-49"
    tmDemo(36) = "45-54"
    tmDemo(37) = "45-64"
    tmDemo(38) = "45+"
    tmDemo(39) = "50-54"
    tmDemo(40) = "50-64"
    tmDemo(41) = "50+"
    tmDemo(42) = "55-64"
    tmDemo(43) = "55+"
    tmDemo(44) = "65+"

    ilRet = gPopMnfPlusFieldsBox(StdPkgVsACT1, lbcDemo, tgDemoCode(), sgDemoCodeTag, "D")

End Sub

Private Sub mGetDemo(slLine As String)
    Dim ilDemo As Integer
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    
    If smDemo = "" Then
        For ilDemo = 0 To UBound(tmDemo) Step 1
            ilPos1 = InStr(1, slLine, tmDemo(ilDemo))
            If ilPos1 >= 0 Then
                ilPos2 = InStr(1, slLine, "BOY")
                If ilPos2 > 0 And ilPos2 < ilPos1 Then
                    smDemo = "M" & tmDemo(ilDemo)
                    Exit Sub
                End If
                ilPos2 = InStr(1, slLine, "GIRL")
                If ilPos2 > 0 And ilPos2 < ilPos1 Then
                    smDemo = "W" & tmDemo(ilDemo)
                    Exit Sub
                End If
                ilPos2 = InStr(1, slLine, "TEEN")
                If ilPos2 > 0 And ilPos2 < ilPos1 Then
                    smDemo = "A" & tmDemo(ilDemo)
                    Exit Sub
                End If
                ilPos2 = InStr(1, slLine, "FEMALE")
                If ilPos2 > 0 And ilPos2 < ilPos1 Then
                    smDemo = "W" & tmDemo(ilDemo)
                    Exit Sub
                End If
                ilPos2 = InStr(1, slLine, "MALE")
                If ilPos2 > 0 And ilPos2 < ilPos1 Then
                    smDemo = "M" & tmDemo(ilDemo)
                    Exit Sub
                End If
                ilPos2 = InStr(1, slLine, "WOMEN")
                If ilPos2 > 0 And ilPos2 < ilPos1 Then
                    smDemo = "W" & tmDemo(ilDemo)
                    Exit Sub
                End If
                ilPos2 = InStr(1, slLine, "MEN")
                If ilPos2 > 0 And ilPos2 < ilPos1 Then
                    smDemo = "M" & tmDemo(ilDemo)
                    Exit Sub
                End If
                ilPos2 = InStr(1, slLine, "ADULT")
                If ilPos2 > 0 And ilPos2 < ilPos1 Then
                    smDemo = "A" & tmDemo(ilDemo)
                    Exit Sub
                End If
                ilPos2 = InStr(1, slLine, "PERSON")
                If ilPos2 > 0 And ilPos2 < ilPos1 Then
                    smDemo = "A" & tmDemo(ilDemo)
                    Exit Sub
                End If
           End If
        Next ilDemo
    End If
End Sub

Private Function mGetPkgAud(ilVefCode As Integer, ilRdfCode As Integer) As String
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim ilDnfCode As Integer
    Dim ilVef As Integer
    Dim ilDay As Integer
    Dim ilRdf As Integer
    Dim llDate As Long
    Dim ilTime As Integer
    Dim llRafCode As Long
    Dim ilMnfSocEco As Integer
    Dim llAvgAud As Long
    Dim llPopEst As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim ilRet As Integer
    ReDim ilDays(0 To 6) As Integer
    
    mGetPkgAud = ""
    If ckcInclude(2).Value = vbUnchecked Then
        Exit Function
    End If
    llOvStartTime = 0
    llOvEndTime = 0
    ilDnfCode = -1
    If cbcBook.ListIndex <= 0 Then
        Exit Function
    End If
    If cbcBook.ListIndex = 1 Then
        ilVef = gBinarySearchVef(ilVefCode)
        If ilVef <> -1 Then
            ilDnfCode = tgMVef(ilVef).iDnfCode
        End If
    Else
        ilDnfCode = cbcBook.ItemData(cbcBook.ListIndex)
    End If
    'Build record into tmPBDP
    For ilDay = 0 To 6 Step 1
        ilDays(ilDay) = False
    Next ilDay
    For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
        If ilRdfCode = tgMRdf(ilRdf).iCode Then
            tmRdf = tgMRdf(ilRdf)
            For ilTime = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1  'Row
                If (tmRdf.iStartTime(0, ilTime) <> 1) Or (tmRdf.iStartTime(1, ilTime) <> 0) Then
                    For ilDay = 1 To 7 Step 1
                        If tmRdf.sWkDays(ilTime, ilDay - 1) = "Y" Then
                            ilDays(ilDay - 1) = True
                        End If
                    Next ilDay
                End If
            Next ilTime
            Exit For
        End If
    Next ilRdf
    If (ilDnfCode > 0) And (ilVefCode > 0) And (imMnfDemo > 0) Then
        llDate = 0
        llRafCode = 0
        ilMnfSocEco = 0
        ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, ilVefCode, ilMnfSocEco, imMnfDemo, llDate, llDate, ilRdfCode, llOvStartTime, llOvEndTime, ilDays(), "S", llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
    Else
        llAvgAud = 0
    End If
    If llAvgAud > 0 Then
        mGetPkgAud = Trim$(Str$(llAvgAud))
    End If

End Function

Private Sub mGetLnAud(ilDayLn() As Integer, slAQHOv As String, slAQHDp As String)
    Dim ilVefCode As Integer
    Dim ilRdfCode As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim ilDnfCode As Integer
    Dim ilVef As Integer
    Dim ilDay As Integer
    Dim ilRdf As Integer
    Dim llDate As Long
    Dim ilTime As Integer
    Dim llRafCode As Long
    Dim ilMnfSocEco As Integer
    Dim llAvgAud As Long
    Dim llPopEst As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim blOverrideDefined As Boolean
    Dim ilRet As Integer
    ReDim ilDaysDP(0 To 6) As Integer
    
    slAQHOv = ""
    slAQHDp = ""
    If ckcInclude(2).Value = vbUnchecked Then
        Exit Sub
    End If
    ilVefCode = rst_Clf!clfVefCode
    ilRdfCode = rst_Clf!clfRdfCode
    If (InStr(1, rst_Clf!clfStartTime, " ", vbBinaryCompare) = 0) And (InStr(1, rst_Clf!clfEndTime, " ", vbBinaryCompare) = 0) Then
        llOvStartTime = 0
        llOvEndTime = 0
        blOverrideDefined = False
    Else
        llOvStartTime = gTimeToLong(Format(rst_Clf!clfStartTime, "h:mm:ssa/p"), False)
        llOvEndTime = gTimeToLong(Format(rst_Clf!clfEndTime, "h:mm:ssa/p"), True)
        blOverrideDefined = True
    End If
    ilDnfCode = rst_Clf!clfDnfCode
    'Build record into tmPBDP
    For ilDay = 0 To 6 Step 1
        ilDaysDP(ilDay) = False
    Next ilDay
    For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
        If ilRdfCode = tgMRdf(ilRdf).iCode Then
            tmRdf = tgMRdf(ilRdf)
            For ilTime = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1  'Row
                If (tmRdf.iStartTime(0, ilTime) <> 1) Or (tmRdf.iStartTime(1, ilTime) <> 0) Then
                    For ilDay = 1 To 7 Step 1
                        If tmRdf.sWkDays(ilTime, ilDay - 1) = "Y" Then
                            ilDaysDP(ilDay - 1) = True
                        End If
                        If ilDayLn(ilDay - 1) <> ilDaysDP(ilDay - 1) Then blOverrideDefined = True
                    Next ilDay
                End If
            Next ilTime
            Exit For
        End If
    Next ilRdf
    If (ilDnfCode > 0) And (ilVefCode > 0) And (imMnfDemo > 0) Then
        llDate = 0
        llRafCode = 0
        ilMnfSocEco = 0
        If blOverrideDefined Then
            ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, ilVefCode, ilMnfSocEco, imMnfDemo, llDate, llDate, ilRdfCode, llOvStartTime, llOvEndTime, ilDayLn(), "S", llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
            If llAvgAud > 0 Then
                slAQHOv = llAvgAud
            End If
        End If
        llOvStartTime = 0
        llOvEndTime = 0
        ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, ilVefCode, ilMnfSocEco, imMnfDemo, llDate, llDate, ilRdfCode, llOvStartTime, llOvEndTime, ilDaysDP(), "S", llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
            If llAvgAud > 0 Then
                slAQHDp = llAvgAud
            End If
    End If

End Sub


Private Sub mPopBook()
    Dim slSQLQuery As String
    
    slSQLQuery = "Select dnfCode, dnfBookName, dnfBookDate from DNF_Demo_Rsrch_Names "
    slSQLQuery = slSQLQuery & " Order By dnfBookDate Desc, dnfBookName"
    Set rst_dnf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_dnf.EOF
        cbcBook.AddItem Trim$(rst_dnf!dnfBookName) & ":" & rst_dnf!dnfBookDate
        cbcBook.ItemData(cbcBook.NewIndex) = rst_dnf!dnfCode
        rst_dnf.MoveNext
    Loop
    cbcBook.AddItem "[Default Vehicle]", 0
    cbcBook.ItemData(cbcBook.NewIndex) = 0
    cbcBook.AddItem "[None]", 0
    cbcBook.ItemData(cbcBook.NewIndex) = 0
    
End Sub
