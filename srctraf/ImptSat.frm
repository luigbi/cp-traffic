VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ImptSat 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5700
   ClientLeft      =   795
   ClientTop       =   1485
   ClientWidth     =   7755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5700
   ScaleWidth      =   7755
   Begin VB.CommandButton cmcBrowse 
      Caption         =   ".."
      Height          =   285
      Left            =   5520
      TabIndex        =   24
      Top             =   735
      Width           =   375
   End
   Begin VB.PictureBox plcPop 
      Height          =   375
      Left            =   1455
      ScaleHeight     =   315
      ScaleWidth      =   3900
      TabIndex        =   22
      Top             =   1185
      Width           =   3960
      Begin VB.TextBox edcPop 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   30
         TabIndex        =   23
         Top             =   30
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmcPop 
      Appearance      =   0  'Flat
      Caption         =   "Import Browser.."
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
      Left            =   6030
      TabIndex        =   20
      Top             =   1200
      Width           =   1605
   End
   Begin VB.ListBox lbcDemo 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "ImptSat.frx":0000
      Left            =   6975
      List            =   "ImptSat.frx":0002
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2265
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   210
      Left            =   2475
      TabIndex        =   18
      Top             =   3075
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox plcDefault 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      ScaleHeight     =   255
      ScaleWidth      =   6000
      TabIndex        =   9
      Top             =   2655
      Width           =   6000
      Begin VB.CheckBox ckcDefault 
         Caption         =   "Rating Book Name"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2070
         TabIndex        =   10
         Top             =   15
         Value           =   1  'Checked
         Width           =   1935
      End
   End
   Begin VB.TextBox edcBookDate 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1455
      MaxLength       =   10
      TabIndex        =   8
      Top             =   2145
      Width           =   1275
   End
   Begin VB.TextBox edcBookName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1455
      MaxLength       =   30
      TabIndex        =   6
      Top             =   1695
      Width           =   2670
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7230
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4905
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7155
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4170
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7110
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4605
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.ListBox lbcErrors 
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
      Height          =   1710
      Left            =   1020
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3450
      Visible         =   0   'False
      Width           =   5340
   End
   Begin VB.CommandButton cmcAud 
      Appearance      =   0  'Flat
      Caption         =   "Import Browser.."
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
      Left            =   6030
      TabIndex        =   4
      Top             =   735
      Width           =   1605
   End
   Begin VB.PictureBox plcAud 
      Height          =   375
      Left            =   1455
      ScaleHeight     =   315
      ScaleWidth      =   3900
      TabIndex        =   2
      Top             =   705
      Width           =   3960
      Begin VB.TextBox edcAud 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmcFileConv 
      Appearance      =   0  'Flat
      Caption         =   "Convert &Files"
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
      Left            =   2100
      TabIndex        =   12
      Top             =   5310
      Width           =   1830
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
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
      Left            =   4245
      TabIndex        =   11
      Top             =   5310
      Width           =   1050
   End
   Begin VB.Label lacPop 
      Appearance      =   0  'Flat
      Caption         =   "Population File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      TabIndex        =   21
      Top             =   1305
      Width           =   1275
   End
   Begin VB.Label lacBookDate 
      Appearance      =   0  'Flat
      Caption         =   "Book Date"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      TabIndex        =   7
      Top             =   2190
      Width           =   1095
   End
   Begin VB.Label lacBookName 
      Appearance      =   0  'Flat
      Caption         =   "Book Name"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      TabIndex        =   5
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   105
      Top             =   5220
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacFileType 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   60
      TabIndex        =   13
      Top             =   3045
      Width           =   2190
   End
   Begin VB.Label lbcAud 
      Appearance      =   0  'Flat
      Caption         =   "Audience File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   90
      TabIndex        =   1
      Top             =   810
      Width           =   1320
   End
End
Attribute VB_Name = "ImptSat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''******************************************************************************************
''***** VB Compress Pro 6.11.32 generated this copy of ImptSat.frm on Wed 6/17/09 @ 12:56 PM
''***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
''******************************************************************************************
'
'' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
'' Proprietary Software, Do not copy
''
'' File Name: ImptSat.Frm
''
'' Release: 1.0
''
'' Description:
''   This file contains the import contract conversion input screen code
'Option Explicit
'Option Compare Text
'Dim imFirstActivate As Integer
'Dim lmTotalNoBytes As Long
'Dim lmProcessedNoBytes As Long
'Dim imTestAddStdDemo As Integer
'Dim lmPercent As Long
'Dim imShowMsg As Integer
'Dim lmLen As Long
'Dim imUpdateMode As Integer 'Update mopde: True or False.  If True and ListenerOrUSA is USA, don't import population
'Dim hmFrom As Integer   'From file hanle
'Dim hmTo As Integer   'From file hanle
'Dim hmDnf As Integer    'file handle
'Dim tmDnf As DNF
'Dim imDnfRecLen As Integer  'Record length
'Dim tmDnfSrchKey As INTKEY0
'Dim hmDrf As Integer    'file handle
'Dim tmDrf As DRF
'Dim imDrfRecLen As Integer  'Record length
'Dim tmDrfSrchKey As DRFKEY0
'Dim tmDrfSrchKey2 As LONGKEY0
'Dim tmPrevDrf() As DRF      'Used in Update mode
'
'Dim hmDpf As Integer    'file handle
'Dim tmDpf As DPF
'Dim imDpfRecLen As Integer  'Record length
'Dim tmDpfSrchKey As LONGKEY0
'Dim tmDpfSrchKey1 As DPFKEY1
'Dim tmPrevDpf() As DPF      'Used in Update mode
'
'Dim hmRdf As Integer    'file handle
'Dim tmRdf As RDF
'Dim imRdfRecLen As Integer  'Record length
'Dim hmVef As Integer    'file handle
'Dim tmVef As VEF
'Dim imVefRecLen As Integer  'Record length
'Dim tmVefSrchKey As INTKEY0
'Dim imVefCodeInDnf() As Integer   'Array if vehicles code which are vehicles with data in dnf
'Dim smVehNotFound() As String
'Dim imRejectedVefCode() As Integer  'Vehicles rejected by the user
'Dim hmMnf As Integer    'file handle
'Dim tmMnf As MNF        'Record structure
'Dim imMnfRecLen As Integer  'Record length
''Dim hmDsf As Integer 'Delete Stamp file handle
''Dim tmDsf As DSF        'DSF record image
''Dim tmDsfSrchKey As LONGKEY0    'DSF key record image
''Dim imDsfRecLen As Integer        'DSF record length
''Dim tmRec As LPOPREC
'Dim imTerminate As Integer
'Dim imConverting As Integer
'Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
'Dim smNowDate As String
'Dim lmNowDate As Long
'Dim smSyncDate As String
'Dim smSyncTime As String
'Dim smDataForm As String
'
'Dim smFieldValues(1 To 60) As String    '25 fields generated in a record
'Dim smSvFields(1 To 60) As String    '25 fields generated in a record
'
'Dim tmDPInfo() As SATDPINFO
'Dim tmSatDemo() As SATDEMO
'
'Dim tmSatExtraPop() As SATEXTRAPOP
'
'Dim tmNameCode() As SORTCODE
'Dim smNameCodeTag As String
'Dim lmMaxWidth As Long
'
'
'
''*******************************************************
''*                                                     *
''*      Procedure Name:mObtainDemo                     *
''*                                                     *
''*             Created:6/13/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Populate tgMnfCDemo             *
''*                                                     *
''*******************************************************
'Private Function mObtainDemo() As Integer
''
''   ilRet = mObtainDemo ()
''   Where:
''       tgMnfCDemo() (I)- MNF record structure
''       ilRet (O)- True = populated; False = error
''
'    Dim ilRecLen As Integer     'Record length
'    Dim llNoRec As Long         'Number of records in Mnf
'    Dim llRecPos As Long        'Record location
'    Dim ilRet As Integer
'    Dim ilOffset As Integer
'    Dim ilUpperBound As Integer
'    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
'
'    ReDim tgMnfSDemo(1 To 1) As MNF
'    ilUpperBound = 1
'    ilRecLen = Len(tgMnfSDemo(1)) 'btrRecordLength(hmMnf)  'Get and save record length
'    'llNoRec = btrRecords(hmMnf) 'Obtain number of records
'    llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlFile) 'Obtain number of records
'    btrExtClear hmMnf   'Clear any previous extend operation
'    ilRet = btrGetFirst(hmMnf, tmMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    If ilRet = BTRV_ERR_END_OF_FILE Then
'        mObtainDemo = True
'        Exit Function
'    Else
'        If ilRet <> BTRV_ERR_NONE Then
'            mObtainDemo = False
'            Exit Function
'        End If
'    End If
'    Call btrExtSetBounds(hmMnf, llNoRec, 0, "UC", "MNF", "") 'Set extract limits (all records)
'    tlCharTypeBuff.sType = "D"
'    ilOffset = 2 'gFieldOffset("Mnf", "MnfType")
'    ilRet = btrExtAddLogicConst(hmMnf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
'    ilOffset = 0 'gFieldOffset("Mnf", "MnfCode")
'    ilRet = btrExtAddField(hmMnf, ilOffset, ilRecLen)  'Extract iCode field
'    If ilRet <> BTRV_ERR_NONE Then
'        mObtainDemo = False
'        Exit Function
'    End If
'    ilRet = btrExtGetNext(hmMnf, tmMnf, ilRecLen, llRecPos)
'    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
'        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
'            mObtainDemo = False
'            Exit Function
'        End If
'        ilRecLen = Len(tgMnfSDemo(1))  'Extract operation record size
'        Do While ilRet = BTRV_ERR_REJECT_COUNT
'            ilRet = btrExtGetNext(hmMnf, tmMnf, ilRecLen, llRecPos)
'        Loop
'        Do While ilRet = BTRV_ERR_NONE
'            If tmMnf.iGroupNo <= 0 Then
'                ilUpperBound = UBound(tgMnfSDemo)
'                tgMnfSDemo(ilUpperBound) = tmMnf
'                ilUpperBound = ilUpperBound + 1
'                ReDim Preserve tgMnfSDemo(1 To ilUpperBound) As MNF
'            End If
'            ilRet = btrExtGetNext(hmMnf, tmMnf, ilRecLen, llRecPos)
'            Do While ilRet = BTRV_ERR_REJECT_COUNT
'                ilRet = btrExtGetNext(hmMnf, tmMnf, ilRecLen, llRecPos)
'            Loop
'        Loop
'    End If
'    mObtainDemo = True
'    Exit Function
'End Function
'
'
''*******************************************************
''*                                                     *
''*      Procedure Name:mAddStdDemo                     *
''*                                                     *
''*             Created:6/13/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Add Standard Demos              *
''*                                                     *
''*******************************************************
'Private Function mAddStdDemo() As Integer
''
''   ilRet = mAddStdDemo ()
''   Where:
''       ilRet (O)- True = populated; False = error
''
'    Dim ilRet As Integer
'    Dim ilLoop As Integer
'    Dim ilFound As Integer
'    Dim ilIndex As Integer
'    Dim slSyncDate As String
'    Dim slSyncTime As String
'    Dim ilAddMissingOnly As Integer
'
'    If Not imTestAddStdDemo Then
'        mAddStdDemo = True
'        Exit Function
'    End If
'    imTestAddStdDemo = False
'    ReDim ilfilter(0 To 1) As Integer
'    ReDim slFilter(0 To 1) As String
'    ReDim ilOffset(0 To 1) As Integer
'    ilfilter(0) = CHARFILTER
'    slFilter(0) = "D"
'    ilOffset(0) = gFieldOffset("Mnf", "MnfType") '2
'    ilfilter(1) = INTEGERFILTER
'    slFilter(1) = "0"
'    ilOffset(1) = gFieldOffset("Mnf", "MnfGroupNo") '2
'    lbcDemo.Clear
'    ilRet = gIMoveListBox(Research, lbcDemo, tmNameCode(), smNameCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilfilter(), slFilter(), ilOffset())
'    smNameCodeTag = ""
'    If lbcDemo.ListCount > 0 Then
'        'Test if 20 exist
'        For ilLoop = 1 To lbcDemo.ListCount - 1 Step 1
'            If InStr(1, lbcDemo.List(ilLoop), "20", vbTextCompare) > 0 Then
'                mAddStdDemo = True
'                Exit Function
'            End If
'        Next ilLoop
'        'Add in missing demos
'        ilAddMissingOnly = True
'    Else
'        ilAddMissingOnly = False
'    End If
'    lbcDemo.Clear
'    gDemoPop lbcDemo   'Get demo names
'    gGetSyncDateTime slSyncDate, slSyncTime
'    For ilLoop = 1 To lbcDemo.ListCount - 1 Step 1
'        ilFound = False
'        If ilAddMissingOnly Then
'            For ilIndex = LBound(tmNameCode) To UBound(tmNameCode) - 1 Step 1
'                If InStr(1, Trim$(tmNameCode(ilIndex).sKey), Trim$(lbcDemo.List(ilLoop)), vbTextCompare) > 0 Then
'                    ilFound = True
'                    Exit For
'                End If
'            Next ilIndex
'        End If
'        If Not ilFound Then
'            tmMnf.iCode = 0
'            tmMnf.sType = "D"
'            tmMnf.sName = lbcDemo.List(ilLoop)
'            tmMnf.sRPU = ""
'            tmMnf.sUnitType = ""
'            tmMnf.iMerge = 0
'            tmMnf.iGroupNo = 0
'            tmMnf.sCodeStn = ""
'            tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
'            tmMnf.iAutoCode = tmMnf.iCode
'            ilRet = btrInsert(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0)
'            Do
'                'tmMnfSrchKey.iCode = tmMnf.iCode
'                'ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'                tmMnf.iRemoteID = tgUrf(0).iRemoteUserID
'                tmMnf.iAutoCode = tmMnf.iCode
'                gPackDate slSyncDate, tmMnf.iSyncDate(0), tmMnf.iSyncDate(1)
'                gPackTime slSyncTime, tmMnf.iSyncTime(0), tmMnf.iSyncTime(1)
'                ilRet = btrUpdate(hmMnf, tmMnf, imMnfRecLen)
'            Loop While ilRet = BTRV_ERR_CONFLICT
'        End If
'    Next ilLoop
'    mAddStdDemo = True
'    Exit Function
'End Function
'
'
'' MsgBox parameters
''Const vbOkOnly = 0                 ' OK button only
''Const vbCritical = 16          ' Critical message
''Const vbApplicationModal = 0
''Const INDEXKEY0 = 0
'Private Sub cmcCancel_Click()
'    If imConverting Then
'        imTerminate = True
'        Exit Sub
'    End If
'    mTerminate
'End Sub
'Private Sub cmcFileConv_Click()
'    Dim slAudName As String
'    Dim slPopName As String
'    Dim slBookName As String
'    Dim slBookDate As String
'    Dim ilRet As Integer
'    Dim ilLoop As Integer
'    Dim ilPrevDnfCode As Integer
'    Dim slToFileName As String
'    Dim slDateTime As String
'    Dim slFileDate As String
'    Dim tlDnf As DNF
'
'    lacFileType.Caption = ""
'    lbcErrors.Clear
'    lbcErrors.Visible = True
'    imShowMsg = True
'    lmLen = 0
'    lmMaxWidth = 0
'    '7/31/06:  Retain Population for USA update
'    imUpdateMode = False
'    slAudName = Trim$(edcAud.Text)
'    If slAudName = "" Then
'        MsgBox "Audience File Name Must be Defined", vbExclamation, "Name Error"
'        edcAud.SetFocus
'        Exit Sub
'    End If
'    slPopName = Trim$(edcPop.Text)
'    If slPopName = "" Then
'        MsgBox "Population File Name Must be Defined", vbExclamation, "Name Error"
'        edcAud.SetFocus
'        Exit Sub
'    End If
'    'Test if Book Name Exist
'    slBookName = Trim$(edcBookName.Text)
'    If slBookName = "" Then
'        MsgBox "Book Name Must be Defined", vbExclamation, "Name Error"
'        edcBookName.SetFocus
'        Exit Sub
'    End If
'    slBookDate = Trim$(edcBookDate.Text)
'    If slBookDate = "" Then
'        MsgBox "Book Date Must be Defined", vbExclamation, "Name Error"
'        edcBookDate.SetFocus
'        Exit Sub
'    End If
'    If Not gValidDate(slBookDate) Then
'        MsgBox "Invalid Date", vbExclamation, "Date Error"
'        edcBookName.SetFocus
'        Exit Sub
'    End If
'    ReDim tmPrevDrf(0 To 0) As DRF
'    ReDim tmPrevDpf(0 To 0) As DPF
'    ilRet = mBookNameUsed(slBookName, slBookDate, ilPrevDnfCode)
'    If ilRet = 1 Then
'        MsgBox "Book Name Previously Used but Dates Don't Match", vbExclamation, "Name Error"
'        edcBookName.SetFocus
'        Exit Sub
'    End If
'    If ilRet = 2 Then
'        Screen.MousePointer = vbDefault
'        'ilRet = MsgBox(slBookName & " previously imported, override", vbYesNo + vbQuestion + vbApplicationModal, "Override")
'        'If ilRet = vbNo Then
'        '    cmcCancel.SetFocus
'        '    Exit Sub
'        'End If
'        sgGenMsg = "Book Name and Date Previously Imported: Replace all Research, Update and Add Research or Cancel this operation"
'        sgCMCTitle(0) = "Replace"
'        sgCMCTitle(1) = "Update"
'        sgCMCTitle(2) = "Cancel"
'        sgCMCTitle(3) = ""
'        igDefCMC = 0
'        igEditBox = 0
'        GenMsg.Show vbModal
'        DoEvents
'
'        If igAnsCMC = 0 Then
'            'Remove records associated with book previously imported
'            Screen.MousePointer = vbHourglass
'            gGetSyncDateTime smSyncDate, smSyncTime
'            ilRet = mRemovePrevDnf(ilPrevDnfCode)
'            If Not ilRet Then
'                cmcCancel.SetFocus
'                Exit Sub
'            End If
'        ElseIf igAnsCMC = 1 Then
'            'Get previously defined records so that it can be determined if import is update or insert mode
'            Screen.MousePointer = vbHourglass
'            mGetPrevDrfDpf ilPrevDnfCode
'            mSetUpdateMode ilPrevDnfCode
'        Else
'            cmcCancel.SetFocus
'            Exit Sub
'        End If
'    End If
'    'Check file names
'    If (InStr(slAudName, ":") = 0) And (Left$(slAudName, 2) <> "\\") Then
'        slAudName = sgImportPath & slAudName
'    End If
'    If (InStr(slPopName, ":") = 0) And (Left$(slPopName, 2) <> "\\") Then
'        slPopName = sgImportPath & slPopName
'    End If
'    lmProcessedNoBytes = 0
'    ilRet = 0
'    ReDim smVehNotFound(0 To 0) As String
'    On Error GoTo cmcFileConvErr:
'    hmFrom = FreeFile
'    Open slAudName For Input Access Read As hmFrom
'    If ilRet <> 0 Then
'        Screen.MousePointer = vbDefault
'        Close hmFrom
'        MsgBox "Unable to find " & slAudName, vbExclamation, "Name Error"
'        edcAud.SetFocus
'        Exit Sub
'    End If
'    '2 required becuase of two passes
'    lmTotalNoBytes = 2 * LOF(hmFrom) 'The Loc returns current position \128
'    Close hmFrom
'    On Error GoTo cmcFileConvErr:
'    hmFrom = FreeFile
'    Open slPopName For Input Access Read As hmFrom
'    If ilRet <> 0 Then
'        Screen.MousePointer = vbDefault
'        Close hmFrom
'        MsgBox "Unable to find " & slPopName, vbExclamation, "Name Error"
'        edcPop.SetFocus
'        Exit Sub
'    End If
'    lmTotalNoBytes = lmTotalNoBytes + LOF(hmFrom) 'The Loc returns current position \128
'    Close hmFrom
'    ilRet = 0
'    hmTo = FreeFile
'    slToFileName = sgDBPath & "Messages\" & "ImptSat.Txt"
'    slDateTime = FileDateTime(slToFileName)
'    If ilRet = 0 Then
'        slFileDate = Format$(slDateTime, "m/d/yy")
'        If gDateValue(slFileDate) = lmNowDate Then  'Append
'            Open slToFileName For Append As hmTo
'        Else
'            Open slToFileName For Output As hmTo
'        End If
'    Else
'        ilRet = 0
'        Open slToFileName For Output As hmTo
'    End If
'    If ilRet <> 0 Then
'        Screen.MousePointer = vbDefault
'        MsgBox "Open " & slToFileName & " Error #" & str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'        cmcCancel.SetFocus
'        Exit Sub
'    End If
'    Screen.MousePointer = vbHourglass
'    imConverting = True
'    smDataForm = "8"
'    gGetSyncDateTime smSyncDate, smSyncTime
'    plcGauge.Value = 0
'    lmPercent = 0
'    Print #hmTo, "Import Satellite Research on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
'    Print #hmTo, ""
'    ReDim imVefCodeInDnf(1 To 1) As Integer
'    ReDim imRejectedVefCode(1 To 1) As Integer
'    tmDnf.iCode = 0
'    tmDnf.sBookName = slBookName
'    gPackDate slBookDate, tmDnf.iBookDate(0), tmDnf.iBookDate(1)
'    gPackDate smNowDate, tmDnf.iEnteredDate(0), tmDnf.iEnteredDate(1)
'    tmDnf.iUrfCode = tgUrf(0).iCode
'    tmDnf.sType = "I"
'    tmDnf.sForm = smDataForm
'    tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
'    tmDnf.iAutoCode = tmDnf.iCode
'    If UBound(tmPrevDrf) <= LBound(tmPrevDrf) Then
'        '7/31/06:  Retain Population for USA update
'        tmDnf.sExactTime = "N"
'        tmDnf.sSource = "S"
'        tmDnf.sEstListenerOrUSA = "L"
'        ilRet = btrInsert(hmDnf, tmDnf, imDnfRecLen, INDEXKEY0)
'    Else
'        Do
'            tmDnf.iCode = tmPrevDrf(LBound(tmPrevDrf)).iDnfCode
'            tmDnfSrchKey.iCode = tmDnf.iCode
'            ilRet = btrGetEqual(hmDnf, tlDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'            '7/31/06:  Retain Population for USA update
'            tmDnf.sEstListenerOrUSA = tlDnf.sEstListenerOrUSA
'            ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
'        Loop While ilRet = BTRV_ERR_CONFLICT
'    End If
'    If ilRet <> BTRV_ERR_NONE Then
'        Print #hmTo, "Error when Adding Demo Name File (DNF)" & str$(ilRet)
'        Close hmTo
'        If gOkAddStrToListBox("Error Adding DNF", lmLen, imShowMsg) Then
'            lbcErrors.AddItem "Error Adding DNF"
'        Else
'            imShowMsg = False
'        End If
'        imConverting = False
'        mTerminate
'        Exit Sub
'    End If
'    'If tgSpf.sRemoteUsers = "Y" Then
'        Do
'            tmDnfSrchKey.iCode = tmDnf.iCode
'            ilRet = btrGetEqual(hmDnf, tlDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'            tmDnf.iRemoteID = tgUrf(0).iRemoteUserID
'            tmDnf.iAutoCode = tmDnf.iCode
'            gPackDate smSyncDate, tmDnf.iSyncDate(0), tmDnf.iSyncDate(1)
'            gPackTime smSyncTime, tmDnf.iSyncTime(0), tmDnf.iSyncTime(1)
'            ilRet = btrUpdate(hmDnf, tmDnf, imDnfRecLen)
'        Loop While ilRet = BTRV_ERR_CONFLICT
'    'End If
'    'lmCount = 0
'    'Process population file
'    lacFileType.Caption = "Processing Population"
'    Print #hmTo, "** Processing Population: " & slPopName & " **"
'    If Not mConvPop(slPopName) Then
'        Print #hmTo, "Import Satellite Research terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
'        Close hmTo
'        imConverting = False
'        mTerminate
'        Exit Sub
'    End If
'    lacFileType.Caption = "Processing Audience"
'    Print #hmTo, "** Processing Audience: " & slAudName & " **"
'    If Not mConvAud(slAudName) Then
'        Print #hmTo, "Import Satellite Research terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
'        Close hmTo
'        imConverting = False
'        mTerminate
'        Exit Sub
'    End If
'    lacFileType.Caption = "Processing Pre-Defined Daypart Audience"
'    Print #hmTo, "** Processing Pre-defined Daypart Audience: " & slAudName & " **"
'    If Not mConvDpfAud(slAudName) Then
'        Print #hmTo, "Import Satellite Research terminated on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
'        Close hmTo
'        imConverting = False
'        mTerminate
'        Exit Sub
'    End If
'    If ckcDefault.Value = vbChecked Then
'        For ilLoop = LBound(imVefCodeInDnf) To UBound(imVefCodeInDnf) - 1 Step 1
'            Do
'                tmVefSrchKey.iCode = imVefCodeInDnf(ilLoop)
'                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'                If ilRet <> BTRV_ERR_NONE Then
'                    Exit Do
'                End If
'                tmVef.iDnfCode = tmDnf.iCode
'                ilRet = btrUpdate(hmVef, tmVef, imVefRecLen)
'            Loop While ilRet = BTRV_ERR_CONFLICT
'            ilRet = gBinarySearchVef(tmVef.iCode)
'            If ilRet <> -1 Then
'                tgMVef(ilRet) = tmVef
'            End If
'        Next ilLoop
'    End If
'    Print #hmTo, "Import Satellite Research successfully completed on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
'    Close hmTo
'    ilRet = mObtainBookName()
'    lacFileType.Caption = "Done"
'    plcGauge.Value = 100
'    bmResearchSaved = true
'    cmcCancel.Caption = "&Done"
'    cmcCancel.SetFocus
'    imConverting = False
'    Screen.MousePointer = vbDefault
'    Exit Sub
'cmcFileConvErr:
'    ilRet = Err.Number
'    Resume Next
'End Sub
'Private Sub cmcAud_Click()
'    lacFileType.Caption = ""
'    igBrowserType = 7   'Text
'    sgBrowseMaskFile = "*"
'    Browser.Show vbModal
'    If igBrowserReturn = 1 Then
'        edcAud.Text = sgBrowserFile
'    End If
'    DoEvents
'    edcAud.SetFocus
'    If InStr(1, sgCurDir, ":") > 0 Then
'        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
'        ChDir sgCurDir
'    End If
'End Sub
'
'Private Sub cmcPop_Click()
'    lacFileType.Caption = ""
'    igBrowserType = 7   'Text
'    sgBrowseMaskFile = "*"
'    Browser.Show vbModal
'    If igBrowserReturn = 1 Then
'        edcPop.Text = sgBrowserFile
'    End If
'    DoEvents
'    edcPop.SetFocus
'    If InStr(1, sgCurDir, ":") > 0 Then
'        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
'        ChDir sgCurDir
'    End If
'End Sub
'
'Private Sub cmcReport_Click()
'    mReport
'End Sub
'Private Sub cmcReport_GotFocus()
'    lacFileType.Caption = ""
'End Sub
'Private Sub edcBookDate_GotFocus()
'    lacFileType.Caption = ""
'    gCtrlGotFocus ActiveControl
'End Sub
'Private Sub edcBookName_GotFocus()
'    lacFileType.Caption = ""
'    gCtrlGotFocus ActiveControl
'End Sub
'
'Private Sub edcAud_Change()
'    edcBookName.Text = ""
'End Sub
'
'Private Sub edcAud_GotFocus()
'    lacFileType.Caption = ""
'    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
'        imFirstFocus = False
'        'Show branner
'    End If
'    gCtrlGotFocus ActiveControl
'End Sub
'Private Sub edcLinkDestHelpMsg_Change()
'    igParentRestarted = True
'End Sub
'Private Sub Form_Activate()
'    If Not imFirstActivate Then
'        DoEvents    'Process events so pending keys are not sent to this
'                    'form when keypreview turn on
'        Me.KeyPreview = True
'        Exit Sub
'    End If
'    imFirstActivate = False
'    Me.KeyPreview = True
'    Me.Refresh
'End Sub
'
'Private Sub Form_Deactivate()
'    Me.KeyPreview = False
'End Sub
'
'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
'        gFunctionKeyBranch KeyCode
'        plcDefault.Visible = False
'        plcDefault.Visible = True
'        plcAud.Visible = False
'        plcAud.Visible = True
'    End If
'
'End Sub
'
'Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
'    sgDoneMsg = CmdStr
'    igChildDone = True
'    Cancel = 0
'End Sub
'Private Sub Form_Load()
'    mInit
'    If imTerminate Then
'        cmcCancel_Click
'    End If
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
''Rm**    ilRet = btrReset(hgHlf)
''Rm**    btrDestroy hgHlf
'    'btrStopAppl
'    'End
'End Sub
'Private Sub imcHelp_Click()
'    'Traffic!cdcSetup.HelpFile = sgHelpPath & "traffic.hlp"
'    'Traffic!cdcSetup.HelpCommand = cdlHelpIndex
'    'Traffic!cdcSetup.Action = 6
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mBookNameUsed                   *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Test if book name used before  *
''*                                                     *
''*******************************************************
'Private Function mBookNameUsed(slBookName As String, slBookDate As String, ilPrevDnfCode As Integer) As Integer
'    'Dim llNoRec As Long         'Number of records in Sof
'    'Dim slName As String
'    Dim llDate As Long
'    'Dim ilExtLen As Integer
'    'Dim llRecPos As Long        'Record location
'    'Dim ilRet As Integer
'    'Dim ilOffset As Integer
'    Dim llTestDate As Long
'    'Dim tlDnf As DNF
'    Dim ilLoop As Integer
'
'    llTestDate = gDateValue(slBookDate)
'    'ilExtLen = Len(tlDnf)  'Extract operation record size
'    'llNoRec = gExtNoRec(ilExtLen)'btrRecords(hmDnf) 'Obtain number of records
'    'btrExtClear hmDnf   'Clear any previous extend operation
'    'ilRet = btrGetFirst(hmDnf, tlDnf, imDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    'If ilRet = BTRV_ERR_END_OF_FILE Then
'    '    mBookNameUsed = False
'    '    Exit Function
'    'End If
'    'Call btrExtSetBounds(hmDnf, llNoRec, -1, "UC") 'Set extract limits (all records including first)
'    'ilOffset = 0
'    'ilRet = btrExtAddField(hmDnf, ilOffset, imDnfRecLen)  'Extract iCode field
'    'If ilRet <> BTRV_ERR_NONE Then
'    '    mBookNameUsed = False
'    '    Exit Function
'    'End If
'    ''ilRet = btrExtGetNextExt(hmDnf)    'Extract record
'    'ilRet = btrExtGetNext(hmDnf, tlDnf, ilExtLen, llRecPos)
'    'If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
'    '    If ilRet <> BTRV_ERR_NONE Then
'    '        mBookNameUsed = False
'    '        Exit Function
'    '    End If
'    '    ilExtLen = Len(tlDnf)  'Extract operation record size
'    '    'ilRet = btrExtGetFirst(hmDnf, tlDnfExt, ilExtLen, llRecPos)
'    '    Do While ilRet = BTRV_ERR_REJECT_COUNT
'    '        ilRet = btrExtGetNext(hmDnf, tlDnf, ilExtLen, llRecPos)
'    '    Loop
'    '    Do While ilRet = BTRV_ERR_NONE
'    '        gUnpackDateLong tlDnf.iBookDate(0), tlDnf.iBookDate(1), llDate
'    '        If (StrComp(Trim$(tlDnf.sBookName), Trim$(slBookName), 1) = 0) And (llDate = llTestDate) Then
'    '            mBookNameUsed = True
'    '            ilPrevDnfCode = tlDnf.iCode
'    '            Exit Function
'    '        End If
'    '        ilRet = btrExtGetNext(hmDnf, tlDnf, ilExtLen, llRecPos)
'    '        Do While ilRet = BTRV_ERR_REJECT_COUNT
'    '            ilRet = btrExtGetNext(hmDnf, tlDnf, ilExtLen, llRecPos)
'    '        Loop
'    '    Loop
'    'End If
'    'mBookNameUsed = False
'    mBookNameUsed = 0   'No
'    ilPrevDnfCode = 0
'    For ilLoop = LBound(tgDnfBook) To UBound(tgDnfBook) - 1 Step 1
'        If StrComp(Trim$(slBookName), Trim$(tgDnfBook(ilLoop).sBookName), 1) = 0 Then
'            mBookNameUsed = 1
'            gUnpackDateLong tgDnfBook(ilLoop).iBookDate(0), tgDnfBook(ilLoop).iBookDate(1), llDate
'            If (llDate = llTestDate) Then
'                mBookNameUsed = 2
'                ilPrevDnfCode = tgDnfBook(ilLoop).iCode
'            End If
'            Exit Function
'        End If
'    Next ilLoop
'    Exit Function
'End Function
'
''*******************************************************
''*                                                     *
''*      Procedure Name:mConvDpfAud                     *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Convert DPF records            *
''*                                                     *
''*******************************************************
'Private Function mConvDpfAud(slFromFile As String) As Integer
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  slStr                         ilPos1                        ilPos2                    *
''*                                                                                        *
''******************************************************************************************
'
'    Dim ilRet As Integer
'    Dim ilHeaderFd As Integer
'    Dim slLine As String
'    Dim ilLoop As Integer
'    Dim ilIndex As Integer
'    Dim ilPos As Integer
'    Dim llPercent As Long
'    Dim ilAge As Integer
'    Dim ilAdj As Integer
'    Dim ilAddFlag As Integer
'    Dim ilEof As Integer
'    Dim slSexChar As String
'    Dim slDemoAge As String
'    Dim slChar As String
'    Dim ilCol As Integer
'    Dim slChannelCode As String
'    Dim slVehicleName As String
'    Dim slARBCode As String
'    Dim ilVef As Integer
'    Dim ilFound As Integer
'    Dim ilVefCode As Integer
'    Dim ilSatIndex As Integer
'    Dim ilPop As Integer
'    Dim slDemoName As String
'    Dim ilPRet As Integer
'
'    ilRet = 0
'    On Error GoTo mConvDpfAudErr:
'    hmFrom = FreeFile
'    Open slFromFile For Input Access Read As hmFrom
'    If ilRet <> 0 Then
'        Close hmFrom
'        MsgBox "Open " & slFromFile & ", Error #" & str$(ilRet), vbExclamation, "Open Error"
'        edcAud.SetFocus
'        mConvDpfAud = False
'        Exit Function
'    End If
'    DoEvents
'    If imTerminate Then
'        Close hmFrom
'        mTerminate
'        mConvDpfAud = False
'        Exit Function
'    End If
'    lmLen = 0
'    imShowMsg = True
'    ilHeaderFd = False
'    ilAddFlag = False
'    slLine = ""
'    ilRet = 0
'    On Error GoTo mConvDpfAudErr:
'    Do
'        ilRet = 0
'        On Error GoTo mConvDpfAudErr:
''        If ilRet <> 0 Then
''            Close hmFrom
''            MsgBox "Input Error #" & Str$(ilRet) & " when reading Daypart File", vbExclamation, "Read Error"
''            mTerminate
''            mConvDpfAud = False
''            Exit Function
''        End If
'        Line Input #hmFrom, slLine
'        On Error GoTo 0
'        If ilRet = 62 Then
'            Exit Do
'        End If
'        If Len(slLine) > 0 Then
'            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
'                ilEof = True
'            Else
'                DoEvents
'                If imTerminate Then
'                    Close hmFrom
'                    mTerminate
'                    mConvDpfAud = False
'                    Exit Function
'                End If
'                gParseCDFields slLine, False, smFieldValues()
'                'Determine field Type
'                'Header Record
'                '   ,,Mon-Sun 6am-Mid,.....
'                '   ,,Total Persons, Total Persons,....
'                'Demo Record
'                '   Males   12+, Vehicle Name, 12000
'                '   Females 18+, Vehicle Name, 1345
'                '
'                If Not ilHeaderFd Then
'                    If (InStr(1, Trim$(slLine), "Total Persons", 1) > 0) Or (InStr(1, Trim$(slLine), "Primary Persons", 1) > 0) Then
'                        ilHeaderFd = True
'                    End If
'                Else
'                    'Obtain Vehicle
''                    slStationCode = ""
''                    slVehicleName = ""
''                    ilPos1 = InStr(1, smFieldValues(2), "Ch", vbTextCompare)
''                    ilPos2 = InStr(1, smFieldValues(2), "-", vbTextCompare)
''                    If (ilPos1 > 0) And (ilPos2 > 0) Then
''                        ''slStationCode = Left$(smFieldValues(2), ilPos2 - 1)
''                        'slStationCode = Left$(smFieldValues(2), ilPos1 + 1) & Mid$(smFieldValues(2), ilPos1 + 3, ilPos2 - (ilPos1 + 3))
''                        slStationCode = Mid$(smFieldValues(2), ilPos1 + 2, ilPos2 - (ilPos1 + 2))
''                        slVehicleName = Trim$(Mid$(smFieldValues(2), ilPos2 + 1))
''                    Else
''                        slVehicleName = Trim$(smFieldValues(2))
''                    End If
''                    ilVefCode = -1
''                    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
''                        If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Or (tgMVef(ilLoop).sType = "V") Then
''                            'If ((slStationCode <> "") And (StrComp(slStationCode, Trim$(tgMVef(ilLoop).sCodeStn), 1) = 0)) Or (StrComp(slVehicleName, Trim$(tgMVef(ilLoop).sName), 1) = 0) Then
''                            slStr = tgMVef(ilLoop).sCodeStn
''                            If InStr(1, slStr, "Ch", vbTextCompare) > 0 Then
''                                slStr = Mid$(slStr, 3)  'Remove Ch
''                            End If
''                            'Jim 1/18/06:  Replace Or with And operator
''                            'If ((slStationCode <> "") And (Val(slStr)) = Val(slStationCode)) Or (StrComp(slVehicleName, Trim$(tgMVef(ilLoop).sName), 1) = 0) Then
''                            If ((slStationCode <> "") And (Val(slStr)) = Val(slStationCode)) And (StrComp(slVehicleName, Trim$(tgMVef(ilLoop).sName), 1) = 0) Then
''                                ilVefCode = tgMVef(ilLoop).iCode
''                            End If
''                        End If
''                    Next ilLoop
'                    ilVefCode = mFindChName(slChannelCode, slVehicleName, slARBCode)
'                    If ilVefCode > 0 Then
'                        ilFound = False
'                        For ilLoop = LBound(imVefCodeInDnf) To UBound(imVefCodeInDnf) - 1 Step 1
'                            If imVefCodeInDnf(ilLoop) = ilVefCode Then
'                                ilFound = True
'                                Exit For
'                            End If
'                        Next ilLoop
'                        If Not ilFound Then
'                            imVefCodeInDnf(UBound(imVefCodeInDnf)) = ilVefCode
'                            ReDim Preserve imVefCodeInDnf(1 To UBound(imVefCodeInDnf) + 1) As Integer
'                        End If
'                        ilAge = 0
'                        ilIndex = 0
'                        If ((InStr(1, smFieldValues(1), "Males", vbTextCompare) > 0) Or ((InStr(1, smFieldValues(1), "Females", vbTextCompare) > 0))) And (InStr(1, smFieldValues(1), "+", vbTextCompare) > 0) Then
'                            slSexChar = UCase$(Left$(smFieldValues(1), 1))
'                            ilPos = 0
'                            If (slSexChar = "M") Or (slSexChar = "F") Then
'                                'Scan for xx-yy
'                                ilIndex = 2
'                                Do While ilIndex < Len(smFieldValues(1))
'                                    slChar = Mid$(smFieldValues(1), ilIndex, 1)
'                                    If (slChar >= "0") And (slChar <= "9") Then
'                                        ilPos = ilIndex
'                                        Exit Do
'                                    End If
'                                    ilIndex = ilIndex + 1
'                                Loop
'                            End If
'                            If ilPos > 0 Then
'                                If (slSexChar = "M") Then
'                                    ilAdj = 0
'                                    slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
'                                Else
'                                    ilAdj = 9
'                                    slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
'                                End If
'                                Select Case slDemoAge
'                                    Case "12+"
'                                        ilIndex = 1 + ilAdj
'                                    Case "18+"
'                                        ilIndex = 2 + ilAdj
'                                    Case "21+"
'                                        ilIndex = 3 + ilAdj
'                                    Case "25+"
'                                        ilIndex = 4 + ilAdj
'                                    Case "35+"
'                                        ilIndex = 5 + ilAdj
'                                    Case "45+"
'                                        ilIndex = 6 + ilAdj
'                                    Case "50+"
'                                        ilIndex = 7 + ilAdj
'                                    Case "55+"
'                                        ilIndex = 8 + ilAdj
'                                    Case "65+"
'                                        ilIndex = 9 + ilAdj
'                                    Case Else
'                                        ilIndex = 0
'                                End Select
'                            Else
'                                ilIndex = 0
'                            End If
'                        Else
'                            slSexChar = UCase$(Left$(smFieldValues(1), 1))
'                            If (slSexChar = "M") Or (slSexChar = "F") Then
'                                'Scan for xx-yy
'                                ilIndex = 2
'                                Do While ilIndex < Len(smFieldValues(1))
'                                    slChar = Mid$(smFieldValues(1), ilIndex, 1)
'                                    If (slChar >= "0") And (slChar <= "9") Then
'                                        ilPos = ilIndex
'                                        Exit Do
'                                    End If
'                                    ilIndex = ilIndex + 1
'                                Loop
'                            End If
'                            If ilPos > 0 Then
'                                If (slSexChar = "M") Then
'                                    ilAdj = 0
'                                    slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
'                                ElseIf (slSexChar = "F") Then
'                                    ilAdj = 9
'                                    slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
'                                End If
'                                Select Case slDemoAge
'                                    Case "12-17"
'                                        ilIndex = 1 + ilAdj
'                                    Case "18-20"
'                                        ilIndex = 2 + ilAdj
'                                    Case "21-24"
'                                        ilIndex = 3 + ilAdj
'                                    Case "25-34"
'                                        ilIndex = 4 + ilAdj
'                                    Case "35-44"
'                                        ilIndex = 5 + ilAdj
'                                    Case "45-49"
'                                        ilIndex = 6 + ilAdj
'                                    Case "50-54"
'                                        ilIndex = 7 + ilAdj
'                                    Case "55-64"
'                                        ilIndex = 8 + ilAdj
'                                    Case "65+"
'                                        ilIndex = 9 + ilAdj
'                                    Case Else
'                                        ilIndex = 0
'                                End Select
'                                If ilIndex <> 0 Then
'                                    ilAddFlag = True
'                                    For ilCol = 0 To UBound(tmDPInfo) - 1 Step 1
'                                        ilFound = False
'                                        For ilVef = LBound(tmSatDemo) To UBound(tmSatDemo) - 1 Step 1
'                                            If (ilVefCode = tmSatDemo(ilVef).iVefCode) And (ilCol = tmSatDemo(ilVef).iDPIndex) Then
'                                                ilFound = True
'                                                ilSatIndex = ilVef
'                                                Exit For
'                                            End If
'                                        Next ilVef
'                                        If Not ilFound Then
'                                            ilSatIndex = UBound(tmSatDemo)
'                                            tmSatDemo(ilSatIndex).iDPIndex = ilCol
'                                            tmSatDemo(ilSatIndex).iVefCode = ilVefCode
'                                            For ilLoop = 1 To 18 Step 1
'                                                tmSatDemo(ilSatIndex).lDemo(ilLoop) = -1
'                                                tmSatDemo(ilSatIndex).lPlus(ilLoop) = 0
'                                            Next ilLoop
'                                            ReDim Preserve tmSatDemo(0 To ilSatIndex + 1) As SATDEMO
'                                        End If
'                                        tmSatDemo(ilSatIndex).lDemo(ilIndex) = Val(smFieldValues(tmDPInfo(ilCol).iStartCol))
'                                    Next ilCol
'                                End If
'                            Else
'                                ilIndex = 0
'                            End If
'                        End If
'                        If ilIndex = 0 Then
'                            slSexChar = UCase$(Left$(smFieldValues(1), 1))
'                            ilPos = 0
'                            If (slSexChar = "M") Or (slSexChar = "F") Or (slSexChar = "P") Then
'                                'Scan for xx-yy
'                                ilIndex = 2
'                                Do While ilIndex < Len(smFieldValues(1))
'                                    slChar = Mid$(smFieldValues(1), ilIndex, 1)
'                                    If (slChar >= "0") And (slChar <= "9") Then
'                                        ilPos = ilIndex
'                                        Exit Do
'                                    End If
'                                    ilIndex = ilIndex + 1
'                                Loop
'                            End If
'                            If ilPos > 0 Then
'                                slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
'                                'Save Population to be used with DPF
'                                If (slSexChar = "M") Or (slSexChar = "B") Then
'                                    slDemoName = "M" & slDemoAge
'                                ElseIf (slSexChar = "W") Or (slSexChar = "G") Or (slSexChar = "F") Then
'                                    slDemoName = "W" & slDemoAge
'                                ElseIf (slSexChar = "T") Or (slSexChar = "P") Then
'                                    If InStr(1, slDemoAge, "12", 1) > 0 Then
'                                        slDemoName = "P" & slDemoAge
'                                    Else
'                                        slDemoName = "A" & slDemoAge
'                                    End If
'                                End If
'                                slDemoName = Trim$(slDemoName)
'                                For ilLoop = LBound(tgMnfSDemo) To UBound(tgMnfSDemo) Step 1
'                                    If StrComp(Trim$(tgMnfSDemo(ilLoop).sName), slDemoName, 1) = 0 Then
'                                        'find drf and pop
'                                        For ilCol = 0 To UBound(tmDPInfo) - 1 Step 1
'                                            ilFound = False
'                                            For ilVef = LBound(tmSatDemo) To UBound(tmSatDemo) - 1 Step 1
'                                                If (ilVefCode = tmSatDemo(ilVef).iVefCode) And (ilCol = tmSatDemo(ilVef).iDPIndex) Then
'                                                    ilFound = True
'                                                    'Find Pop
'                                                    For ilPop = LBound(tmSatExtraPop) To UBound(tmSatExtraPop) - 1 Step 1
'                                                        If tmSatExtraPop(ilPop).iMnfDemo = tgMnfSDemo(ilLoop).iCode Then
'                                                            tmDpf.lCode = 0
'                                                            tmDpf.iDnfCode = tmDnf.iCode
'                                                            tmDpf.lDrfCode = tmSatDemo(ilVef).lDrfCode
'                                                            tmDpf.iMnfDemo = tmSatExtraPop(ilPop).iMnfDemo
'                                                            tmDpf.lPop = tmSatExtraPop(ilPop).lPop
''                                                            If tgSpf.sSAudData = "H" Then
''                                                                tmDpf.lDemo = (Val(smFieldValues(tmDPInfo(ilCol).iStartCol)) + 50) \ 100
''                                                            Else
''                                                                tmDpf.lDemo = (Val(smFieldValues(tmDPInfo(ilCol).iStartCol)) + 500) \ 1000
''                                                            End If
'                                                            If tgSpf.sSAudData = "H" Then
'                                                                tmDpf.lDemo = (Val(smFieldValues(tmDPInfo(ilCol).iStartCol)) + 50) \ 100
'                                                            ElseIf tgSpf.sSAudData = "N" Then
'                                                                tmDpf.lDemo = (Val(smFieldValues(tmDPInfo(ilCol).iStartCol)) + 5) \ 10
'                                                            ElseIf tgSpf.sSAudData = "U" Then
'                                                                tmDpf.lDemo = Val(smFieldValues(tmDPInfo(ilCol).iStartCol))
'                                                            Else
'                                                                tmDpf.lDemo = (Val(smFieldValues(tmDPInfo(ilCol).iStartCol)) + 500) \ 1000
'                                                            End If
'                                                            If tmDpf.lDemo > 0 Then
'                                                                'ilPRet = btrInsert(hmDpf, tmDpf, imDpfRecLen, INDEXKEY0)
'                                                                ilPRet = mAddDpf()
'                                                                If ilPRet <> BTRV_ERR_NONE Then
'                                                                    If (ilPRet = 30000) Or (ilPRet = 30001) Or (ilPRet = 30002) Or (ilPRet = 30003) Then
'                                                                        ilPRet = csiHandleValue(0, 7)
'                                                                    End If
'                                                                    Print #hmTo, "Warning: Error when Adding Demo Plus Data File (DPF)" & str$(ilPRet) & " for " & "Ch " & slChannelCode & " Vehicle " & slVehicleName & " ARB Code " & slARBCode
'                                                                    lbcErrors.AddItem "Error Adding DPF" & " for " & "Ch " & slChannelCode & " Vehicle " & slVehicleName & " ARB Code " & slARBCode
'                                                                End If
'                                                            End If
'                                                            Exit For
'                                                        End If
'                                                    Next ilPop
'                                                    Exit For
'                                                End If
'                                            Next ilVef
'                                            If Not ilFound Then
'                                                'Add a blank DRF, then add dpf
'                                            End If
'                                        Next ilCol
'                                        Exit For
'                                    End If
'                                Next ilLoop
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
'            llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
'            If llPercent >= 100 Then
'                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
'                    llPercent = 99
'                Else
'                    llPercent = 100
'                End If
'            End If
'            If plcGauge.Value <> llPercent Then
'                plcGauge.Value = llPercent
'            End If
'        End If
'    Loop Until ilEof
'
'    Close hmFrom
'    plcGauge.Value = 100
'    mConvDpfAud = True
'    Exit Function
'mConvDpfAudErr:
'    ilRet = Err.Number
'    Resume Next
'End Function
'
''*******************************************************
''*                                                     *
''*      Procedure Name:mConvAud                    *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Convert CHF                    *
''*                                                     *
''*******************************************************
'Private Function mConvAud(slFromFile As String) As Integer
'    Dim ilRet As Integer
'    Dim ilHeaderFd As Integer
'    Dim slLine As String
'    Dim ilLoop As Integer
'    Dim ilIndex As Integer
'    Dim slStr As String
'    Dim ilPos As Integer
'    Dim ilPos1 As Integer
'    Dim ilPos2 As Integer
'    Dim llPercent As Long
'    Dim ilAge As Integer
'    Dim ilAdj As Integer
'    Dim ilAddFlag As Integer
'    Dim ilDay As Integer
'    Dim ilEof As Integer
'    Dim slSexChar As String
'    Dim slDemoAge As String
'    Dim slChar As String
'    Dim ilCol As Integer
'    Dim ilSY As Integer
'    Dim ilEY As Integer
'    Dim ilTIndex As Integer
'    Dim slDay As String
'    Dim slTime As String
'    Dim slChannelCode As String
'    Dim slVehicleName As String
'    Dim slARBCode As String
'    Dim ilVef As Integer
'    Dim ilFound As Integer
'    Dim ilVefCode As Integer
'    Dim ilSatIndex As Integer
'    Dim slSvLine As String
'    Dim llRif As Long
'    Dim ilRdf As Integer
'    Dim ilTimeCount As Integer
'    Dim ilNoTimes As Integer
'    Dim ilOk As Integer
'    Dim ilDayIndex As Integer
'
'    ilRet = 0
'    On Error GoTo mConvAudErr:
'    hmFrom = FreeFile
'    Open slFromFile For Input Access Read As hmFrom
'    If ilRet <> 0 Then
'        Close hmFrom
'        MsgBox "Open " & slFromFile & ", Error #" & str$(ilRet), vbExclamation, "Open Error"
'        edcAud.SetFocus
'        mConvAud = False
'        Exit Function
'    End If
'    DoEvents
'    If imTerminate Then
'        Close hmFrom
'        mTerminate
'        mConvAud = False
'        Exit Function
'    End If
'    ilNoTimes = 1
'    lmLen = 0
'    imShowMsg = True
'    ilHeaderFd = False
'    ilAddFlag = False
'    slLine = ""
'    ilRet = 0
'    On Error GoTo mConvAudErr:
'    Do
'        ilRet = 0
'        On Error GoTo mConvAudErr:
''        If ilRet <> 0 Then
''            Close hmFrom
''            MsgBox "Input Error #" & Str$(ilRet) & " when reading Daypart File", vbExclamation, "Read Error"
''            mTerminate
''            mConvAud = False
''            Exit Function
''        End If
'        Line Input #hmFrom, slLine
'        On Error GoTo 0
'        If ilRet = 62 Then
'            Exit Do
'        End If
'        If Len(slLine) > 0 Then
'            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
'                ilEof = True
'            Else
'                DoEvents
'                If imTerminate Then
'                    Close hmFrom
'                    mTerminate
'                    mConvAud = False
'                    Exit Function
'                End If
'                gParseCDFields slLine, False, smFieldValues()
'                'Determine field Type
'                'Header Record
'                '   ,,Mon-Sun 6am-Mid,.....
'                '   ,,Total Persons, Total Persons,....
'                'Demo Record
'                '   Males   12+, Vehicle Name, 12000
'                '   Females 18+, Vehicle Name, 1345
'                '
'                If Not ilHeaderFd Then
'                    If (InStr(1, Trim$(slLine), "Total Persons", 1) > 0) Or (InStr(1, Trim$(slLine), "Primary Persons", 1) > 0) Then
'                        ilHeaderFd = True
'                        'Determine number of Demo columns
'                        ilPos = 3
'                        ReDim tmDPInfo(0 To 0) As SATDPINFO
'                        ReDim tmSatDemo(0 To 0) As SATDEMO
'                        Do While ilPos <= UBound(smFieldValues)
'                            If (InStr(1, Trim$(smFieldValues(ilPos)), "Total Persons", 1) > 0) Or (InStr(1, Trim$(smFieldValues(ilPos)), "Primary Persons", 1) > 0) Then
'                                tmDPInfo(UBound(tmDPInfo)).iStartCol = ilPos
'                                tmDPInfo(UBound(tmDPInfo)).iType = 0  'Vehicle
'                                tmDPInfo(UBound(tmDPInfo)).iBkNm = 0  'Daypart Book Name
'                                tmDPInfo(UBound(tmDPInfo)).iSY = -1
'                                tmDPInfo(UBound(tmDPInfo)).iEY = -1
'                                ReDim Preserve tmDPInfo(0 To UBound(tmDPInfo) + 1) As SATDPINFO
'                            End If
'                            ilPos = ilPos + 1
'                        Loop
'                        gParseCDFields slSvLine, False, smSvFields()
'                        For ilCol = 0 To UBound(tmDPInfo) - 1 Step 1
'                            slStr = UCase$(Trim$(smSvFields(tmDPInfo(ilCol).iStartCol)))
'                            If tmDPInfo(ilCol).iSY = -1 Then
'                                ilSY = -1
'                                ilTIndex = InStr(1, slStr, " ", 1) + 1
'                                slDay = Trim$(Mid$(slStr, 1, ilTIndex - 2))
'                                Select Case slDay
'                                    Case "MO", "MON"
'                                        ilSY = 0
'                                        ilEY = 0
'                                    Case "TU", "TUE"
'                                        ilSY = 1
'                                        ilEY = 1
'                                    Case "WE", "WED"
'                                        ilSY = 2
'                                        ilEY = 2
'                                    Case "TH", "THU"
'                                        ilSY = 3
'                                        ilEY = 3
'                                    Case "FR", "FRI"
'                                        ilSY = 4
'                                        ilEY = 4
'                                    Case "SA", "SAT"
'                                        ilSY = 5
'                                        ilEY = 5
'                                        If Mid$(slStr, 3, 2) = "SU" Then
'                                            ilEY = 6
'                                        End If
'                                    Case "SU", "SUN"
'                                        ilSY = 6
'                                        ilEY = 6
'                                    Case "MF", "MON-FRI", "MO-FR"
'                                        ilSY = 0
'                                        ilEY = 4
'                                    Case "MSA", "MON-SAT", "MO-SA"
'                                        ilSY = 0
'                                        ilEY = 5
'                                    Case "MSU", "MON-SUN", "MO-SU"
'                                        ilSY = 0
'                                        ilEY = 6
'                                    Case "MS", "MON-SUN", "MO-SU"
'                                        ilSY = 0
'                                        ilEY = 6
'                                    Case "SS", "SAT-SUN", "SA-SU", "WEEKEND"
'                                        ilSY = 5
'                                        ilEY = 6
'                                    Case "FSU", "FRI-SUN", "FR-SU"
'                                        ilSY = 4
'                                        ilEY = 6
'                                    Case "FSA", "FRI-SAT", "FR-SA"
'                                        ilSY = 4
'                                        ilEY = 5
'                                End Select
'                                If ilSY <> -1 Then
''                                    slChar = Mid$(slStr, ilTIndex, 1)
''                                    If (slChar >= "0") And (slChar <= "9") Then
''                                        slStr = Mid$(slStr, ilTIndex)
''                                    Else
''                                        slStr = Mid$(slStr, ilTIndex + 1)
''                                    End If
''                                    slTime = Mid$(slStr, 1, 1)
''                                    slStr = Mid$(slStr, 2)
''                                    slChar = Mid$(slStr, 1, 1)
''                                    If (slChar >= "0") And (slChar <= "9") Then
''                                        slTime = slTime & Mid$(slStr, 1, 2)
''                                        slStr = Mid$(slStr, 3)
''                                    Else
''                                        If Mid$(slStr, 1, 1) = "-" Then
''                                            slTime = slTime & right$(slStr, 1)
''                                        Else
''                                            slTime = slTime & Mid$(slStr, 1, 1)
''                                            slStr = Mid$(slStr, 2)
''                                        End If
''                                    End If
'                                    ilPos = InStr(ilTIndex, slStr, "-", vbTextCompare)
'                                    If ilPos > 0 Then
'                                        slTime = Mid$(slStr, ilTIndex, ilPos - ilTIndex)
'                                        If StrComp(slTime, "Mid", vbTextCompare) = 0 Then
'                                            slTime = "12AM"
'                                        End If
'                                        If gValidTime(slTime) Then
'                                            tmDPInfo(ilCol).lStartTime = gTimeToCurrency(slTime, False)
''                                            slChar = Mid$(slStr, 1, 1)
''                                            If slChar = "-" Then
''                                                slStr = Trim$(Mid$(slStr, 2))
''                                            End If
''                                            slTime = Mid$(slStr, 1, 1)
''                                            slStr = Mid$(slStr, 2)
''                                            slChar = Mid$(slStr, 1, 1)
''                                            If (slChar >= "0") And (slChar <= "9") Then
''                                                slTime = slTime & Mid$(slStr, 1, 2)
''                                                slStr = Mid$(slStr, 3)
''                                            Else
''                                                slTime = slTime & Mid$(slStr, 1, 1)
''                                                slStr = Mid$(slStr, 2)
''                                            End If
'                                            slTime = Mid$(slStr, ilPos + 1)
'                                            If StrComp(slTime, "Mid", vbTextCompare) = 0 Then
'                                                slTime = "12AM"
'                                            End If
'                                            If gValidTime(slTime) Then
'                                                tmDPInfo(ilCol).iSY = ilSY
'                                                tmDPInfo(ilCol).iEY = ilEY
'                                                tmDPInfo(ilCol).lEndTime = gTimeToCurrency(slTime, False)
'                                                For ilDay = 0 To 6 Step 1
'                                                    tmDPInfo(ilCol).sDays(ilDay) = "N"
'                                                Next ilDay
'                                                For ilDay = ilSY To ilEY Step 1
'                                                    tmDPInfo(ilCol).sDays(ilDay) = "Y"
'                                                Next ilDay
'                                                tmDPInfo(ilCol).iType = 1  'Daypart
'                                            End If
'                                        End If
'                                    End If
'                                End If
'                            End If
'                        Next ilCol
'                    Else
'                        slSvLine = slLine
'                    End If
'                Else
''                    'Obtain Vehicle
''                    slStationCode = ""
''                    slVehicleName = ""
''                    ilPos1 = InStr(1, smFieldValues(2), "Ch", vbTextCompare)
''                    ilPos2 = InStr(1, smFieldValues(2), "-", vbTextCompare)
''                    If (ilPos1 > 0) And (ilPos2 > 0) Then
''                        'slStationCode = Left$(smFieldValues(2), ilPos1 + 1) & Mid$(smFieldValues(2), ilPos1 + 3, ilPos2 - (ilPos1 + 3))
''                        slStationCode = Mid$(smFieldValues(2), ilPos1 + 2, ilPos2 - (ilPos1 + 2))
''                        slVehicleName = Trim$(Mid$(smFieldValues(2), ilPos2 + 1))
''                    Else
''                        slVehicleName = Trim$(smFieldValues(2))
''                    End If
''                    ilVefCode = -1
''                    For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
''                        If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Or (tgMVef(ilLoop).sType = "V") Then
''                            'If ((slStationCode <> "") And (StrComp(slStationCode, Trim$(tgMVef(ilLoop).sCodeStn), 1) = 0)) Or (StrComp(slVehicleName, Trim$(tgMVef(ilLoop).sName), 1) = 0) Then
''                            slStr = tgMVef(ilLoop).sCodeStn
''                            If InStr(1, slStr, "Ch", vbTextCompare) > 0 Then
''                                slStr = Mid$(slStr, 3)  'Remove Ch
''                            End If
''                            'Jim Request 1/18/06:  Replace Or with And
''                            'If ((slStationCode <> "") And (Val(slStr)) = Val(slStationCode)) Or (StrComp(slVehicleName, Trim$(tgMVef(ilLoop).sName), 1) = 0) Then
''                            If ((slStationCode <> "") And (Val(slStr)) = Val(slStationCode)) And (StrComp(slVehicleName, Trim$(tgMVef(ilLoop).sName), 1) = 0) Then
''                                ilVefCode = tgMVef(ilLoop).iCode
''                            End If
''                        End If
''                    Next ilLoop
'                    ilVefCode = mFindChName(slChannelCode, slVehicleName, slARBCode)
'                    If ilVefCode > 0 Then
'                        ilFound = False
'                        For ilLoop = LBound(imVefCodeInDnf) To UBound(imVefCodeInDnf) - 1 Step 1
'                            If imVefCodeInDnf(ilLoop) = ilVefCode Then
'                                ilFound = True
'                                Exit For
'                            End If
'                        Next ilLoop
'                        If Not ilFound Then
'                            imVefCodeInDnf(UBound(imVefCodeInDnf)) = ilVefCode
'                            ReDim Preserve imVefCodeInDnf(1 To UBound(imVefCodeInDnf) + 1) As Integer
'                        End If
'                        ilAge = 0
'                        If ((InStr(1, smFieldValues(1), "Males", vbTextCompare) > 0) Or ((InStr(1, smFieldValues(1), "Females", vbTextCompare) > 0))) And (InStr(1, smFieldValues(1), "+", vbTextCompare) > 0) Then
'                            ilPos1 = InStr(1, smFieldValues(1), "Males", vbTextCompare) + 5
'                            ilPos2 = InStr(1, smFieldValues(1), "+", vbTextCompare)
'                            slStr = Trim$(Mid$(smFieldValues(1), ilPos1, ilPos2 - ilPos1))
'                            ilAge = Val(slStr)
'                            ilAdj = 0
'                            If (InStr(1, smFieldValues(1), "Females", vbTextCompare) > 0) And (InStr(1, smFieldValues(1), "+", vbTextCompare) > 0) Then
'                                ilAdj = 9
'                            End If
'                            ilIndex = 0
'                            Select Case ilAge
'                                Case 12
'                                    ilIndex = 1 + ilAdj
'                                Case 18
'                                    ilIndex = 2 + ilAdj
'                                Case 21
'                                    ilIndex = 3 + ilAdj
'                                Case 25
'                                    ilIndex = 4 + ilAdj
'                                Case 35
'                                    ilIndex = 5 + ilAdj
'                                Case 45
'                                    ilIndex = 6 + ilAdj
'                                Case 50
'                                    ilIndex = 7 + ilAdj
'                                Case 55
'                                    ilIndex = 8 + ilAdj
'                                Case 65
'                                    ilIndex = 9 + ilAdj
'                                Case Else
'                                    ilIndex = 0
'                            End Select
'                            If ilIndex <> 0 Then
'                                ilAddFlag = True
'                                For ilCol = 0 To UBound(tmDPInfo) - 1 Step 1
'                                    ilFound = False
'                                    For ilVef = LBound(tmSatDemo) To UBound(tmSatDemo) - 1 Step 1
'                                        If (ilVefCode = tmSatDemo(ilVef).iVefCode) And (ilCol = tmSatDemo(ilVef).iDPIndex) Then
'                                            ilFound = True
'                                            ilSatIndex = ilVef
'                                            Exit For
'                                        End If
'                                    Next ilVef
'                                    If Not ilFound Then
'                                        ilSatIndex = UBound(tmSatDemo)
'                                        tmSatDemo(ilSatIndex).iDPIndex = ilCol
'                                        tmSatDemo(ilSatIndex).iVefCode = ilVefCode
'                                        For ilLoop = 1 To 18 Step 1
'                                            tmSatDemo(ilSatIndex).lDemo(ilLoop) = -1
'                                            tmSatDemo(ilSatIndex).lPlus(ilLoop) = 0
'                                        Next ilLoop
'                                        ReDim Preserve tmSatDemo(0 To ilSatIndex + 1) As SATDEMO
'                                    End If
'                                    tmSatDemo(ilSatIndex).lPlus(ilIndex) = Val(smFieldValues(tmDPInfo(ilCol).iStartCol))
'                                Next ilCol
'                            End If
'                        Else
'                            slSexChar = UCase$(Left$(smFieldValues(1), 1))
'                            If (slSexChar = "M") Or (slSexChar = "F") Then
'                                'Scan for xx-yy
'                                ilIndex = 2
'                                Do While ilIndex < Len(smFieldValues(1))
'                                    slChar = Mid$(smFieldValues(1), ilIndex, 1)
'                                    If (slChar >= "0") And (slChar <= "9") Then
'                                        ilPos = ilIndex
'                                        Exit Do
'                                    End If
'                                    ilIndex = ilIndex + 1
'                                Loop
'                            End If
'                            If ilPos > 0 Then
'                                If (slSexChar = "M") Then
'                                    ilAdj = 0
'                                    slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
'                                ElseIf (slSexChar = "F") Then
'                                    ilAdj = 9
'                                    slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
'                                End If
'                                Select Case slDemoAge
'                                    Case "12-17"
'                                        ilIndex = 1 + ilAdj
'                                    Case "18-20"
'                                        ilIndex = 2 + ilAdj
'                                    Case "21-24"
'                                        ilIndex = 3 + ilAdj
'                                    Case "25-34"
'                                        ilIndex = 4 + ilAdj
'                                    Case "35-44"
'                                        ilIndex = 5 + ilAdj
'                                    Case "45-49"
'                                        ilIndex = 6 + ilAdj
'                                    Case "50-54"
'                                        ilIndex = 7 + ilAdj
'                                    Case "55-64"
'                                        ilIndex = 8 + ilAdj
'                                    Case "65+"
'                                        ilIndex = 9 + ilAdj
'                                    Case Else
'                                        ilIndex = 0
'                                End Select
'                                If ilIndex <> 0 Then
'                                    ilAddFlag = True
'                                    For ilCol = 0 To UBound(tmDPInfo) - 1 Step 1
'                                        ilFound = False
'                                        For ilVef = LBound(tmSatDemo) To UBound(tmSatDemo) - 1 Step 1
'                                            If (ilVefCode = tmSatDemo(ilVef).iVefCode) And (ilCol = tmSatDemo(ilVef).iDPIndex) Then
'                                                ilFound = True
'                                                ilSatIndex = ilVef
'                                                Exit For
'                                            End If
'                                        Next ilVef
'                                        If Not ilFound Then
'                                            ilSatIndex = UBound(tmSatDemo)
'                                            tmSatDemo(ilSatIndex).iDPIndex = ilCol
'                                            tmSatDemo(ilSatIndex).iVefCode = ilVefCode
'                                            For ilLoop = 1 To 18 Step 1
'                                                tmSatDemo(ilSatIndex).lDemo(ilLoop) = -1
'                                                tmSatDemo(ilSatIndex).lPlus(ilLoop) = 0
'                                            Next ilLoop
'                                            ReDim Preserve tmSatDemo(0 To ilSatIndex + 1) As SATDEMO
'                                        End If
'                                        tmSatDemo(ilSatIndex).lDemo(ilIndex) = Val(smFieldValues(tmDPInfo(ilCol).iStartCol))
'                                    Next ilCol
'                                End If
'                            End If
'                        End If
'                    Else
''                        '1/18/05 Jim:  Match Station Code and Vehicle name required
''                        'slStr = smFieldValues(2)
''                        ilFound = False
''                        For ilLoop = LBound(smVehNotFound) To UBound(smVehNotFound) - 1 Step 1
''                            If StrComp(slStr, smVehNotFound(ilLoop), 1) = 0 Then
''                                ilFound = True
''                            End If
''                        Next ilLoop
''                        If Not ilFound Then
''                            smVehNotFound(UBound(smVehNotFound)) = slStr
''                            ReDim Preserve smVehNotFound(0 To UBound(smVehNotFound) + 1) As String
''                            Print #hmTo, "Unable to Find Vehicle " & slStr & " record not added"
''                            If gOkAddStrToListBox("Unable to Find Vehicle " & slStr, lmLen, imShowMsg) Then
''                                lbcErrors.AddItem "Unable to Find Vehicle " & slStr
''                            Else
''                                imShowMsg = False
''                            End If
''                        End If
'                    End If
'                End If
'            End If
'            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
'            llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
'            If llPercent >= 100 Then
'                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
'                    llPercent = 99
'                Else
'                    llPercent = 100
'                End If
'            End If
'            If plcGauge.Value <> llPercent Then
'                plcGauge.Value = llPercent
'            End If
'        End If
'    Loop Until ilEof
'    If ilHeaderFd And ilAddFlag Then
'        For ilCol = 0 To UBound(tmDPInfo) - 1 Step 1
'            tmDrf.lCode = 0
'            tmDrf.iDnfCode = tmDnf.iCode
'            tmDrf.sDemoDataType = "D"
'            tmDrf.iMnfSocEco = 0
'            tmDrf.sInfoType = "D"
'            tmDrf.iRdfcode = 0
'            gPackTimeLong tmDPInfo(ilCol).lStartTime, tmDrf.iStartTime(0), tmDrf.iStartTime(1)
'            gPackTimeLong tmDPInfo(ilCol).lEndTime, tmDrf.iEndTime(0), tmDrf.iEndTime(1)
'            For ilDay = 0 To 6 Step 1
'                tmDrf.sDay(ilDay) = tmDPInfo(ilCol).sDays(ilDay)
'            Next ilDay
'            tmDrf.sProgCode = ""
'            tmDrf.iStartTime2(0) = 1
'            tmDrf.iStartTime2(1) = 0
'            tmDrf.iEndTime2(0) = 1
'            tmDrf.iEndTime2(1) = 0
'            tmDrf.iQHIndex = 0
'            tmDrf.iCount = 0
'            tmDrf.sExStdDP = "N"
'            tmDrf.sExRpt = "N"
'            tmDrf.sDataType = "A"
'            For ilVef = LBound(tmSatDemo) To UBound(tmSatDemo) - 1 Step 1
'                If ilCol = tmSatDemo(ilVef).iDPIndex Then
'                    For ilLoop = 1 To 18 Step 1
'                        tmDrf.lDemo(ilLoop) = 0
'                    Next ilLoop
'                    tmDrf.iVefCode = tmSatDemo(ilVef).iVefCode
'                    For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
'                        If (tgMRif(llRif).iVefCode = tmDrf.iVefCode) Then
'                            ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfcode)
'                            If ilRdf <> -1 Then
'                                'Ignore Dormant Dayparts- Jim request 6/23/04 because of demo to XM
'                                If tgMRdf(ilRdf).sState <> "D" Then
'                                    If (tgMRdf(ilRdf).iLtfCode(0) = 0) And (tgMRdf(ilRdf).iLtfCode(1) = 0) And (tgMRdf(ilRdf).iLtfCode(2) = 0) Then
'                                        ilTimeCount = 0
'                                        For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
'                                            If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
'                                                ilTimeCount = ilTimeCount + 1
'                                            End If
'                                        Next ilIndex
'                                        If ilTimeCount = ilNoTimes Then
'                                            ilTimeCount = 0
'                                            For ilIndex = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
'                                                If (tgMRdf(ilRdf).iStartTime(0, ilIndex) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilIndex) <> 0) Then
'                                                    ilOk = False
'                                                    If ((tgMRdf(ilRdf).iStartTime(0, ilIndex) = tmDrf.iStartTime(0)) And (tgMRdf(ilRdf).iStartTime(1, ilIndex) = tmDrf.iStartTime(1))) Then
'                                                        If (tgMRdf(ilRdf).iEndTime(0, ilIndex) = tmDrf.iEndTime(0)) And (tgMRdf(ilRdf).iEndTime(1, ilIndex) = tmDrf.iEndTime(1)) Then
'                                                            ilOk = True
'                                                        End If
'                                                    End If
'                                                    If ((tgMRdf(ilRdf).iStartTime(0, ilIndex) = tmDrf.iStartTime2(0)) And (tgMRdf(ilRdf).iStartTime(1, ilIndex) = tmDrf.iStartTime2(1))) Then
'                                                        If (tgMRdf(ilRdf).iEndTime(0, ilIndex) = tmDrf.iEndTime2(0)) And (tgMRdf(ilRdf).iEndTime(1, ilIndex) = tmDrf.iEndTime2(1)) Then
'                                                            ilOk = True
'                                                        End If
'                                                    End If
'                                                    If ilOk Then
'                                                        'Exact time match- check days
'                                                        ilOk = True
'                                                        For ilDayIndex = 0 To 6 Step 1
'                                                            If (tmDrf.sDay(ilDayIndex) = "Y") And (tgMRdf(ilRdf).sWkDays(ilIndex, ilDayIndex + 1) <> "Y") Then
'                                                                ilOk = False
'                                                                Exit For
'                                                            ElseIf (tmDrf.sDay(ilDayIndex) = "N") And (tgMRdf(ilRdf).sWkDays(ilIndex, ilDayIndex + 1) <> "N") Then
'                                                                ilOk = False
'                                                                Exit For
'                                                            End If
'                                                        Next ilDayIndex
'                                                        If ilOk Then
'                                                            ilTimeCount = ilTimeCount + 1
'                                                            If ilTimeCount = ilNoTimes Then
'                                                                tmDrf.iRdfcode = tgMRdf(ilRdf).iCode
'                                                                Exit For
'                                                            End If
'                                                        End If
'                                                    End If
'                                                End If
'                                            Next ilIndex
'                                        End If
'                                    End If
'                                End If
'                            End If
'                            If tmDrf.iRdfcode > 0 Then
'                                Exit For
'                            End If
'                        End If
'                    Next llRif
'                    'Add Record
'                    tmDrf.sForm = smDataForm
'                    tmDrf.lCode = 0
'                    tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
'                    tmDrf.lAutoCode = tmDrf.lCode
'                    gPackDate smSyncDate, tmDrf.iSyncDate(0), tmDrf.iSyncDate(1)
'                    gPackTime smSyncTime, tmDrf.iSyncTime(0), tmDrf.iSyncTime(1)
'                    For ilLoop = 1 To 8 Step 1
'                        If tmSatDemo(ilVef).lDemo(ilLoop) >= 0 Then
'                            tmDrf.lDemo(ilLoop) = tmSatDemo(ilVef).lDemo(ilLoop)
'                        Else
'                            tmDrf.lDemo(ilLoop) = tmSatDemo(ilVef).lPlus(ilLoop) - tmSatDemo(ilVef).lPlus(ilLoop + 1)
'                            If tmDrf.lDemo(ilLoop) < 0 Then
'                                tmDrf.lDemo(ilLoop) = 0
'                            End If
'                        End If
'                    Next ilLoop
'                    tmDrf.lDemo(9) = tmSatDemo(ilVef).lPlus(9)
'                    For ilLoop = 10 To 17 Step 1
'                        If tmSatDemo(ilVef).lDemo(ilLoop) >= 0 Then
'                            tmDrf.lDemo(ilLoop) = tmSatDemo(ilVef).lDemo(ilLoop)
'                        Else
'                            tmDrf.lDemo(ilLoop) = tmSatDemo(ilVef).lPlus(ilLoop) - tmSatDemo(ilVef).lPlus(ilLoop + 1)
'                            If tmDrf.lDemo(ilLoop) < 0 Then
'                                tmDrf.lDemo(ilLoop) = 0
'                            End If
'                        End If
'                    Next ilLoop
'                    tmDrf.lDemo(18) = tmSatDemo(ilVef).lPlus(18)
'                    For ilLoop = 1 To 18 Step 1
''                        If tgSpf.sSAudData = "H" Then
''                            tmDrf.lDemo(ilLoop) = (tmDrf.lDemo(ilLoop) + 50) \ 100
''                        Else
''                            tmDrf.lDemo(ilLoop) = (tmDrf.lDemo(ilLoop) + 500) \ 1000
''                        End If
'                        If tgSpf.sSAudData = "H" Then
'                            tmDrf.lDemo(ilLoop) = (tmDrf.lDemo(ilLoop) + 50) \ 100
'                        ElseIf tgSpf.sSAudData = "N" Then
'                            tmDrf.lDemo(ilLoop) = (tmDrf.lDemo(ilLoop) + 5) \ 10
'                        ElseIf tgSpf.sSAudData = "U" Then
'                            tmDrf.lDemo(ilLoop) = tmDrf.lDemo(ilLoop)
'                        Else
'                            tmDrf.lDemo(ilLoop) = (tmDrf.lDemo(ilLoop) + 500) \ 1000
'                        End If
'                    Next ilLoop
'                    ilRet = mAddDrf()
'                    tmSatDemo(ilVef).lDrfCode = tmDrf.lCode
'                End If
'            Next ilVef
'        Next ilCol
'    End If
'
'    Close hmFrom
'    'plcGauge.Value = 100
'    mConvAud = True
'    Exit Function
'mConvAudErr:
'    ilRet = Err.Number
'    Resume Next
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:mConvPop                        *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Convert CHF                    *
''*                                                     *
''*******************************************************
'Private Function mConvPop(slFromFile As String) As Integer
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilUpdate                      llDrfCode                     llDrf                     *
''*  tlDrf                                                                                 *
''******************************************************************************************
'
'    Dim ilRet As Integer
'    Dim ilHeaderFd As Integer
'    Dim slLine As String
'    Dim ilLoop As Integer
'    Dim ilIndex As Integer
'    Dim ilPos As Integer
'    Dim llPercent As Long
'    Dim ilAge As Integer
'    Dim ilAdj As Integer
'    Dim ilAddFlag As Integer
'    Dim ilDay As Integer
'    Dim ilEof As Integer
'    Dim slSexChar As String
'    Dim slDemoAge As String
'    Dim slChar As String
'    Dim slDemoName As String
'    ReDim llDemo(1 To 18) As Long
'    ReDim llPlus(1 To 18) As Long
'
'    ReDim tmSatExtraPop(1 To 1) As SATEXTRAPOP
'
'    ilRet = 0
'    On Error GoTo mConvPopErr:
'    hmFrom = FreeFile
'    Open slFromFile For Input Access Read As hmFrom
'    If ilRet <> 0 Then
'        Close hmFrom
'        MsgBox "Open " & slFromFile & ", Error #" & str$(ilRet), vbExclamation, "Open Error"
'        edcAud.SetFocus
'        mConvPop = False
'        Exit Function
'    End If
'    DoEvents
'    If imTerminate Then
'        Close hmFrom
'        mTerminate
'        mConvPop = False
'        Exit Function
'    End If
'    ilHeaderFd = False
'    ilAddFlag = False
'    slLine = ""
'    tmDrf.lCode = 0
'    tmDrf.iDnfCode = tmDnf.iCode
'    tmDrf.sDemoDataType = "P"
'    tmDrf.iMnfSocEco = 0
'    tmDrf.iVefCode = 0
'    tmDrf.sInfoType = ""
'    tmDrf.iRdfcode = 0
'    tmDrf.sProgCode = ""
'    tmDrf.iStartTime(0) = 1
'    tmDrf.iStartTime(1) = 0
'    tmDrf.iEndTime(0) = 1
'    tmDrf.iEndTime(1) = 0
'    tmDrf.iStartTime2(0) = 1
'    tmDrf.iStartTime2(1) = 0
'    tmDrf.iEndTime2(0) = 1
'    tmDrf.iEndTime2(1) = 0
'    For ilDay = 0 To 6 Step 1
'        tmDrf.sDay(ilDay) = "Y"
'    Next ilDay
'    tmDrf.iQHIndex = 0
'    tmDrf.iCount = 0
'    tmDrf.sExStdDP = "N"
'    tmDrf.sExRpt = "N"
'    tmDrf.sDataType = "A"
'    For ilLoop = 1 To 18 Step 1
'        tmDrf.lDemo(ilLoop) = 0
'        llDemo(ilLoop) = -1
'        llPlus(ilLoop) = 0
'    Next ilLoop
'    On Error GoTo mConvPopErr:
'    Do
'        ilRet = 0
'        On Error GoTo mConvPopErr:
'        Line Input #hmFrom, slLine
'        On Error GoTo 0
'        If ilRet = 62 Then
'            Exit Do
'        End If
'        If Len(slLine) > 0 Then
'            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
'                ilEof = True
'            Else
'                DoEvents
'                If imTerminate Then
'                    Close hmFrom
'                    mTerminate
'                    mConvPop = False
'                    Exit Function
'                End If
'                gParseCDFields slLine, False, smFieldValues()
'                'Determine field Type
'                'Header Record
'                '   Demographics,Total Estimate Number of Primary and........
'                'Demo Record
'                '   Males   12+, 12000
'                '   Females 18+,
'                '
'                If Not ilHeaderFd Then
'                    If InStr(1, Trim$(slLine), "Demographic", 1) > 0 Then
'                        If (InStr(1, RTrim$(slLine), "Total", 1) > 0) Or (InStr(1, RTrim$(slLine), "Number of Primary", 1) > 0) Then
'                            ilHeaderFd = True
'                        End If
'                    End If
'                Else
'                    ilAge = 0
'                    ilIndex = 0
'                    If ((InStr(1, smFieldValues(1), "Males", vbTextCompare) > 0) Or ((InStr(1, smFieldValues(1), "Females", vbTextCompare) > 0))) And (InStr(1, smFieldValues(1), "+", vbTextCompare) > 0) Then
'                        slSexChar = UCase$(Left$(smFieldValues(1), 1))
'                        ilPos = 0
'                        If (slSexChar = "M") Or (slSexChar = "F") Then
'                            'Scan for xx-yy
'                            ilIndex = 2
'                            Do While ilIndex < Len(smFieldValues(1))
'                                slChar = Mid$(smFieldValues(1), ilIndex, 1)
'                                If (slChar >= "0") And (slChar <= "9") Then
'                                    ilPos = ilIndex
'                                    Exit Do
'                                End If
'                                ilIndex = ilIndex + 1
'                            Loop
'                        End If
'                        If ilPos > 0 Then
'                            If (slSexChar = "M") Then
'                                ilAdj = 0
'                                slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
'                            Else
'                                ilAdj = 9
'                                slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
'                            End If
'                            Select Case slDemoAge
'                                Case "12+"
'                                    ilIndex = 1 + ilAdj
'                                Case "18+"
'                                    ilIndex = 2 + ilAdj
'                                Case "21+"
'                                    ilIndex = 3 + ilAdj
'                                Case "25+"
'                                    ilIndex = 4 + ilAdj
'                                Case "35+"
'                                    ilIndex = 5 + ilAdj
'                                Case "45+"
'                                    ilIndex = 6 + ilAdj
'                                Case "50+"
'                                    ilIndex = 7 + ilAdj
'                                Case "55+"
'                                    ilIndex = 8 + ilAdj
'                                Case "65+"
'                                    ilIndex = 9 + ilAdj
'                                Case Else
'                                    ilIndex = 0
'                            End Select
'                            If ilIndex <> 0 Then
'                                ilAddFlag = True
'                                llPlus(ilIndex) = Val(smFieldValues(2))
'                            End If
'                        Else
'                            ilIndex = 0
'                        End If
'                    Else
'                        slSexChar = UCase$(Left$(smFieldValues(1), 1))
'                        ilPos = 0
'                        If (slSexChar = "M") Or (slSexChar = "F") Then
'                            'Scan for xx-yy
'                            ilIndex = 2
'                            Do While ilIndex < Len(smFieldValues(1))
'                                slChar = Mid$(smFieldValues(1), ilIndex, 1)
'                                If (slChar >= "0") And (slChar <= "9") Then
'                                    ilPos = ilIndex
'                                    Exit Do
'                                End If
'                                ilIndex = ilIndex + 1
'                            Loop
'                        End If
'                        If ilPos > 0 Then
'                            If (slSexChar = "M") Then
'                                ilAdj = 0
'                                slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
'                            Else
'                                ilAdj = 9
'                                slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
'                            End If
'                            Select Case slDemoAge
'                                Case "12-17"
'                                    ilIndex = 1 + ilAdj
'                                Case "18-20"
'                                    ilIndex = 2 + ilAdj
'                                Case "21-24"
'                                    ilIndex = 3 + ilAdj
'                                Case "25-34"
'                                    ilIndex = 4 + ilAdj
'                                Case "35-44"
'                                    ilIndex = 5 + ilAdj
'                                Case "45-49"
'                                    ilIndex = 6 + ilAdj
'                                Case "50-54"
'                                    ilIndex = 7 + ilAdj
'                                Case "55-64"
'                                    ilIndex = 8 + ilAdj
'                                Case "65+"
'                                    ilIndex = 9 + ilAdj
'                                Case Else
'                                    ilIndex = 0
'                            End Select
'                            If ilIndex <> 0 Then
'                                ilAddFlag = True
'                                llDemo(ilIndex) = Val(smFieldValues(2))
'                            End If
'                        Else
'                            ilIndex = 0
'                        End If
'                    End If
'                    If ilIndex = 0 Then
'                        slSexChar = UCase$(Left$(smFieldValues(1), 1))
'                        ilPos = 0
'                        If (slSexChar = "M") Or (slSexChar = "F") Or (slSexChar = "P") Then
'                            'Scan for xx-yy
'                            ilIndex = 2
'                            Do While ilIndex < Len(smFieldValues(1))
'                                slChar = Mid$(smFieldValues(1), ilIndex, 1)
'                                If (slChar >= "0") And (slChar <= "9") Then
'                                    ilPos = ilIndex
'                                    Exit Do
'                                End If
'                                ilIndex = ilIndex + 1
'                            Loop
'                        End If
'                        If ilPos > 0 Then
'                            slDemoAge = Trim$(Mid$(smFieldValues(1), ilPos))
'                            'Save Population to be used with DPF
'                            If (slSexChar = "M") Or (slSexChar = "B") Then
'                                slDemoName = "M" & slDemoAge
'                            ElseIf (slSexChar = "W") Or (slSexChar = "G") Then
'                                slDemoName = "W" & slDemoAge
'                            ElseIf (slSexChar = "T") Or (slSexChar = "P") Then
'                                If InStr(1, slDemoAge, "12", 1) > 0 Then
'                                    slDemoName = "P" & slDemoAge
'                                Else
'                                    slDemoName = "A" & slDemoAge
'                                End If
'                            End If
'                            slDemoName = Trim$(slDemoName)
'                            For ilLoop = LBound(tgMnfSDemo) To UBound(tgMnfSDemo) Step 1
'                                If StrComp(Trim$(tgMnfSDemo(ilLoop).sName), slDemoName, 1) = 0 Then
'                                    tmSatExtraPop(UBound(tmSatExtraPop)).iMnfDemo = tgMnfSDemo(ilLoop).iCode
''                                    If tgSpf.sSAudData = "H" Then
''                                        tmSatExtraPop(UBound(tmSatExtraPop)).lPop = (Val(smFieldValues(2)) + 50) \ 100
''                                    Else
''                                        tmSatExtraPop(UBound(tmSatExtraPop)).lPop = (Val(smFieldValues(2)) + 500) \ 1000
''                                    End If
'                                    If tgSpf.sSAudData = "H" Then
'                                        tmSatExtraPop(UBound(tmSatExtraPop)).lPop = (Val(smFieldValues(2)) + 50) \ 100
'                                    ElseIf tgSpf.sSAudData = "N" Then
'                                        tmSatExtraPop(UBound(tmSatExtraPop)).lPop = (Val(smFieldValues(2)) + 5) \ 10
'                                    ElseIf tgSpf.sSAudData = "U" Then
'                                        tmSatExtraPop(UBound(tmSatExtraPop)).lPop = Val(smFieldValues(2))
'                                    Else
'                                        tmSatExtraPop(UBound(tmSatExtraPop)).lPop = (Val(smFieldValues(2)) + 500) \ 1000
'                                    End If
'                                    ReDim Preserve tmSatExtraPop(1 To UBound(tmSatExtraPop) + 1) As SATEXTRAPOP
'                                    Exit For
'                                End If
'                            Next ilLoop
'                        End If
'                    End If
'                End If
'            End If
'            lmProcessedNoBytes = lmProcessedNoBytes + Len(slLine) + 2 'Loc(hmFrom)
'            llPercent = (lmProcessedNoBytes * CSng(100)) / lmTotalNoBytes
'            If llPercent >= 100 Then
'                If lmProcessedNoBytes + 3 < lmTotalNoBytes Then
'                    llPercent = 99
'                Else
'                    llPercent = 100
'                End If
'            End If
'            If plcGauge.Value <> llPercent Then
'                plcGauge.Value = llPercent
'            End If
'        End If
'    Loop Until ilEof
'    '7/31/06:  Retain Population for USA update
'    If ilHeaderFd And ilAddFlag And (imUpdateMode = False) Then
'        tmDrf.sForm = smDataForm
'        tmDrf.lCode = 0
'        tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
'        tmDrf.lAutoCode = tmDrf.lCode
'        gPackDate smSyncDate, tmDrf.iSyncDate(0), tmDrf.iSyncDate(1)
'        gPackTime smSyncTime, tmDrf.iSyncTime(0), tmDrf.iSyncTime(1)
'        For ilLoop = 1 To 8 Step 1
'            If llDemo(ilLoop) >= 0 Then
'                tmDrf.lDemo(ilLoop) = llDemo(ilLoop)
'            Else
'                tmDrf.lDemo(ilLoop) = llPlus(ilLoop) - llPlus(ilLoop + 1)
'                If tmDrf.lDemo(ilLoop) < 0 Then
'                    tmDrf.lDemo(ilLoop) = 0
'                End If
'            End If
'        Next ilLoop
'        tmDrf.lDemo(9) = llPlus(9)
'        For ilLoop = 10 To 17 Step 1
'            If llDemo(ilLoop) >= 0 Then
'                tmDrf.lDemo(ilLoop) = llDemo(ilLoop)
'            Else
'                tmDrf.lDemo(ilLoop) = llPlus(ilLoop) - llPlus(ilLoop + 1)
'                If tmDrf.lDemo(ilLoop) < 0 Then
'                    tmDrf.lDemo(ilLoop) = 0
'                End If
'            End If
'        Next ilLoop
'        tmDrf.lDemo(18) = llPlus(18)
'        For ilLoop = 1 To 18 Step 1
''            If tgSpf.sSAudData = "H" Then
''                tmDrf.lDemo(ilLoop) = (tmDrf.lDemo(ilLoop) + 50) \ 100
''            Else
''                tmDrf.lDemo(ilLoop) = (tmDrf.lDemo(ilLoop) + 500) \ 1000
''            End If
'            If tgSpf.sSAudData = "H" Then
'                tmDrf.lDemo(ilLoop) = (tmDrf.lDemo(ilLoop) + 50) \ 100
'            ElseIf tgSpf.sSAudData = "N" Then
'                tmDrf.lDemo(ilLoop) = (tmDrf.lDemo(ilLoop) + 5) \ 10
'            ElseIf tgSpf.sSAudData = "U" Then
'                tmDrf.lDemo(ilLoop) = tmDrf.lDemo(ilLoop)
'            Else
'                tmDrf.lDemo(ilLoop) = (tmDrf.lDemo(ilLoop) + 500) \ 1000
'            End If
'        Next ilLoop
'        ilRet = mAddDrf()
'
'    End If
'    Close hmFrom
'    'plcGauge.Value = 100
'    mConvPop = True
'    Exit Function
'mConvPopErr:
'    ilRet = Err.Number
'    Resume Next
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:gGetRecLength                   *
''*                                                     *
''*             Created:10/09/93      By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Obtain the record length from   *
''*                     the database                    *
''*                                                     *
''*******************************************************
'Private Function mGetRecLength(slFileName As String) As Integer
''
''   ilRecLen = mGetRecLength(slName)
''   Where:
''       slName (I)- Name of the file
''       ilRecLen (O)- record length within the file
''
'    Dim hlFile As Integer
'    Dim ilRet As Integer
'    hlFile = CBtrvTable(ONEHANDLE) 'CBtrvObj
'    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        mGetRecLength = -ilRet
'        ilRet = btrClose(hlFile)
'        btrDestroy hlFile
'        Exit Function
'    End If
'    mGetRecLength = btrRecordLength(hlFile)  'Get and save record length
'    ilRet = btrClose(hlFile)
'    btrDestroy hlFile
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:mInit                           *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Initialize modular             *
''*                                                     *
''*******************************************************
'Private Sub mInit()
''
''   mInit
''   Where:
''
'    Dim ilRet As Integer
'    imTerminate = False
'    imFirstActivate = True
'    'mParseCmmdLine
'    Screen.MousePointer = vbHourglass
'    bmResearchSaved = False
'    imTestAddStdDemo = True
'    imConverting = False
'    imFirstFocus = True
'    lmTotalNoBytes = 0
'    lmProcessedNoBytes = 0
'    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
'    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptSat
'    On Error GoTo 0
'    imRdfRecLen = Len(tmRdf)
'    hmMnf = CBtrvTable(TWOHANDLES) 'CBtrvObj
'    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptSat
'    On Error GoTo 0
'    imMnfRecLen = Len(tmMnf)
'    hmVef = CBtrvTable(TWOHANDLES) 'CBtrvObj
'    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptSat
'    On Error GoTo 0
'    imVefRecLen = Len(tmVef)
'    hmDrf = CBtrvTable(TWOHANDLES) 'CBtrvObj
'    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptSat
'    On Error GoTo 0
'    imDrfRecLen = Len(tmDrf)
'    hmDpf = CBtrvTable(TWOHANDLES) 'CBtrvObj
'    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptSat
'    On Error GoTo 0
'    imDpfRecLen = Len(tmDpf)
'    hmDnf = CBtrvTable(TWOHANDLES) 'CBtrvObj
'    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo mInitErr
'    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptSat
'    On Error GoTo 0
'    imDnfRecLen = Len(tmDnf)
''    hmDsf = CBtrvTable(TWOHANDLES) 'CBtrvObj
''    ilRet = btrOpen(hmDsf, "", sgDBPath & "Dsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
''    On Error GoTo mInitErr
''    gBtrvErrorMsg ilRet, "mInit (btrOpen)", ImptSat
''    On Error GoTo 0
''    imDsfRecLen = Len(tmDsf)
'    'Populate arrays to determine if records exist
'    ilRet = mAddStdDemo()
'    ilRet = mObtainBookName()
'    If ilRet = False Then
'        Screen.MousePointer = vbDefault
'        MsgBox "Obtain Book Name Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
'        imTerminate = True
'        Exit Sub
'    End If
'    ilRet = gObtainRcfRifRdf()
'    'ilRet = mObtainDaypart()
'    If ilRet = False Then
'        Screen.MousePointer = vbDefault
'        MsgBox "Obtain Daypart Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
'        imTerminate = True
'        Exit Sub
'    End If
'    ilRet = mObtainDemo()
'    If ilRet = False Then
'        Screen.MousePointer = vbDefault
'        MsgBox "Obtain Demo Error", vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
'        imTerminate = True
'        Exit Sub
'    End If
'    smNowDate = Format$(gNow(), "m/d/yy")
'    lmNowDate = gDateValue(smNowDate)
'
'    'smRptTime = Format$(Now, "h:m:s AM/PM")
'    'gPackTime smRptTime, tmIcf.iTime(0), tmIcf.iTime(1)
'    gCenterStdAlone ImptSat
'    If mTestRecLengths() Then
'        Screen.MousePointer = vbDefault
'        imTerminate = True
'        Exit Sub
'    End If
'    Screen.MousePointer = vbDefault
'    'imcHelp.Picture = Traffic!imcHelp.Picture
'    Exit Sub
'mInitErr:
'    On Error GoTo 0
'    imTerminate = True
'    Exit Sub
'
'    ilRet = Err.Number
'    Resume Next
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mObtainBookName                 *
''*                                                     *
''*             Created:6/13/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Populate tgCompMnf for          *
''*                     collection                      *
''*                                                     *
''*******************************************************
'Private Function mObtainBookName() As Integer
''
''   ilRet = mObtainBookName ()
''   Where:
''       tgCompMnf() (I)- MNFCOMPEXT record structure to be created
''       ilRet (O)- True = populated; False = error
''
'    Dim llNoRec As Long         'Number of records in Mnf
'    Dim ilExtLen As Integer
'    Dim llRecPos As Long        'Record location
'    Dim ilRet As Integer
'    Dim ilOffset As Integer
'    Dim ilUpperBound As Integer
'
'    ReDim tgDnfBook(0 To 0) As DNF
'    ilUpperBound = UBound(tgDnfBook)
'    'ilRet = btrGetFirst(hmDnf, tgDnfBook(ilUpperBound), imDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    'Do While ilRet = BTRV_ERR_NONE
'    '    ilUpperBound = ilUpperBound + 1
'    '    ReDim Preserve tgDnfBook(1 To ilUpperBound) As DNF
'    '    ilRet = btrGetNext(hmDnf, tgDnfBook(ilUpperBound), imDnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    'Loop
'    ilExtLen = Len(tgDnfBook(0))  'Extract operation record size
'    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hmDnf) 'Obtain number of records
'    btrExtClear hmDnf   'Clear any previous extend operation
'    ilRet = btrGetFirst(hmDnf, tgDnfBook(0), imDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    If ilRet = BTRV_ERR_END_OF_FILE Then
'        mObtainBookName = True
'        Exit Function
'    End If
'    Call btrExtSetBounds(hmDnf, llNoRec, -1, "UC", "DNF", "") 'Set extract limits (all records including first)
'    ilOffset = 0
'    ilRet = btrExtAddField(hmDnf, ilOffset, imDnfRecLen)  'Extract iCode field
'    If ilRet <> BTRV_ERR_NONE Then
'        mObtainBookName = False
'        Exit Function
'    End If
'    'ilRet = btrExtGetNextExt(hmDnf)    'Extract record
'    ilRet = btrExtGetNext(hmDnf, tgDnfBook(ilUpperBound), ilExtLen, llRecPos)
'    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
'        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
'            mObtainBookName = False
'            Exit Function
'        End If
'        ilExtLen = Len(tgDnfBook(0))  'Extract operation record size
'        Do While ilRet = BTRV_ERR_REJECT_COUNT
'            ilRet = btrExtGetNext(hmDnf, tgDnfBook(ilUpperBound), ilExtLen, llRecPos)
'        Loop
'        Do While ilRet = BTRV_ERR_NONE
'            ilUpperBound = ilUpperBound + 1
'            ReDim Preserve tgDnfBook(1 To ilUpperBound) As DNF
'            ilRet = btrExtGetNext(hmDnf, tgDnfBook(ilUpperBound), ilExtLen, llRecPos)
'            Do While ilRet = BTRV_ERR_REJECT_COUNT
'                ilRet = btrExtGetNext(hmDnf, tgDnfBook(ilUpperBound), ilExtLen, llRecPos)
'            Loop
'        Loop
'    End If
'    mObtainBookName = True
'    Exit Function
'End Function
''*******************************************************
''*                                                     *
''*      Procedure Name:mRemovePrevDnf                  *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: Remove record of previouly     *
''*                      imported book                  *
''*                                                     *
''*******************************************************
'Private Function mRemovePrevDnf(ilPrevDnfCode As Integer) As Integer
'    Dim ilRet As Integer
'    Dim ilPRet As Integer
'    Dim tlDrf As DRF
'    Dim tlDnf As DNF
'    Do
'        tmDrfSrchKey.iDnfCode = ilPrevDnfCode
'        tmDrfSrchKey.sDemoDataType = ""
'        tmDrfSrchKey.iMnfSocEco = 0
'        tmDrfSrchKey.iVefCode = 0
'        tmDrfSrchKey.sInfoType = ""
'        tmDrfSrchKey.iRdfcode = 0
'        ilRet = btrGetGreaterOrEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'        If ilRet <> BTRV_ERR_NONE Then
'            Exit Do
'        End If
'        If tlDrf.iDnfCode <> ilPrevDnfCode Then
'            Exit Do
'        End If
'        'tmRec = tlDrf
'        'ilRet = gGetByKeyForUpdate("DRF", hmDrf, tmRec)
'        'tlDrf = tmRec
'        'If ilRet <> BTRV_ERR_NONE Then
'        '    mRemovePrevDnf = False
'        '    ilRet = MsgBox("Remove Not Completed, Try Later", vbOkOnly + vbExclamation, "Remove")
'        '    Exit Function
'        'End If
'        tmDrfSrchKey2.lCode = tlDrf.lCode
'        ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'        If ilRet <> BTRV_ERR_NONE Then
'            Exit Do
'        End If
'        ilRet = btrDelete(hmDrf)
'        If ilRet = BTRV_ERR_NONE Then
'            Do
'                tmDpfSrchKey1.lDrfCode = tlDrf.lCode
'                tmDpfSrchKey1.iMnfDemo = 0
'                ilPRet = btrGetGreaterOrEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'                If ilPRet = BTRV_ERR_NONE Then
'                    If tmDpf.lDrfCode <> tlDrf.lCode Then
'                        Exit Do
'                    End If
'                    tmDpfSrchKey1.lDrfCode = tlDrf.lCode
'                    tmDpfSrchKey1.iMnfDemo = tmDpf.iMnfDemo
'                    ilPRet = btrGetEqual(hmDpf, tmDpf, imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'                    If ilPRet = BTRV_ERR_NONE Then
'                        ilPRet = btrDelete(hmDpf)
'                    End If
'                Else
'                    Exit Do
'                End If
'            Loop While ilPRet = BTRV_ERR_NONE
'        End If
'    Loop
'    ilRet = BTRV_ERR_NONE
'    Do
'        tmDnfSrchKey.iCode = ilPrevDnfCode
'        ilRet = btrGetEqual(hmDnf, tlDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
'        If ilRet <> BTRV_ERR_NONE Then
'            mRemovePrevDnf = False
'            ilRet = MsgBox("Remove Not Completed, Try Later", vbOkOnly + vbExclamation, "Remove")
'            Exit Function
'        End If
'        ilRet = btrDelete(hmDnf)
'    Loop While ilRet = BTRV_ERR_CONFLICT
'    If ilRet <> BTRV_ERR_NONE Then
'        mRemovePrevDnf = False
'        ilRet = MsgBox("Remove Not Completed, Try Later", vbOkOnly + vbExclamation, "Remove")
'        Exit Function
'    End If
''    If tgSpf.sRemoteUsers = "Y" Then
''        tmDsf.lCode = 0
''        tmDsf.sFileName = "DNF"
''        gPackDate smSyncDate, tmDsf.iSyncDate(0), tmDsf.iSyncDate(1)
''        gPackTime smSyncTime, tmDsf.iSyncTime(0), tmDsf.iSyncTime(1)
''        tmDsf.iRemoteID = tlDnf.iRemoteID
''        tmDsf.lAutoCode = tlDnf.iAutoCode
''        tmDsf.iSourceID = tgUrf(0).iRemoteUserID
''        tmDsf.lCntrNo = 0
''        ilRet = btrInsert(hmDsf, tmDsf, imDsfRecLen, INDEXKEY0)
''    End If
'    mRemovePrevDnf = True
'    Exit Function
'End Function
'Private Sub mReport()
'    Dim slStr As String
'    If igStdAloneMode Then
'        Exit Sub
'    End If
'    'If Not gWinRoom(igNoExeWinRes(RPTSELEXE)) Then
'    '    Exit Sub
'    'End If
'    igRptCallType = CHFCONVMENU
'    'Screen.MousePointer = vbHourGlass  'Wait
'    'igChildDone = False
'    'edcLinkSrceDoneMsg.Text = ""
'    'If (Not igStdAloneMode) And (imShowHelpMsg) Then
'        If igTestSystem Then
'            slStr = "ImptSat^Test\" & sgUserName & "\" & Trim$(str$(igRptCallType))
'        Else
'            slStr = "ImptSat^Prod\" & sgUserName & "\" & Trim$(str$(igRptCallType))
'        End If
'    'Else
'    '    If igTestSystem Then
'    '        slStr = "ImptSat^Test^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
'    '    Else
'    '        slStr = "ImptSat^Prod^NOHELP\" & sgUserName & "\" & Trim$(Str$(igRptCallType))
'    '    End If
'    'End If
'    ''lgShellRet = Shell(sgExePath & "RptSel.Exe " & slStr, 1)
'    'lgShellRet = Shell(sgExePath & "RptList.Exe " & slStr, 1)
'    'ImptSat.Enabled = False
'    'Do While Not igChildDone
'    '    DoEvents
'    'Loop
'    'ImptSat.Enabled = True
'    sgCommandStr = slStr
'    RptList.Show vbModal
'    slStr = sgDoneMsg
'    'edcLinkSrceDoneMsg.Text = "Ok"  'Tell child received message-unload
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mTerminate                      *
''*                                                     *
''*             Created:5/17/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments: terminate form                 *
''*                                                     *
''*******************************************************
'Private Sub mTerminate()
''
''   mTerminate
''   Where:
''
'    Dim ilRet As Integer
'    Erase imRejectedVefCode
'    Erase tmNameCode
'    Erase tgDnfBook
'    Erase tmDPInfo
'    Erase tmSatDemo
'    Erase tmSatExtraPop
'    'Erase tgRdf
'    Erase imVefCodeInDnf
'    Erase smVehNotFound
'    Erase tmPrevDrf
'    Erase tmPrevDpf
'    ilRet = btrClose(hmRdf)
'    btrDestroy hmRdf
'    ilRet = btrClose(hmMnf)
'    btrDestroy hmMnf
'    ilRet = btrClose(hmVef)
'    btrDestroy hmVef
'    ilRet = btrClose(hmDrf)
'    btrDestroy hmDrf
'    ilRet = btrClose(hmDpf)
'    btrDestroy hmDpf
'    ilRet = btrClose(hmDnf)
'    btrDestroy hmDnf
''    ilRet = btrClose(hmDsf)
''    btrDestroy hmDsf
'    Screen.MousePointer = vbDefault
'    'igParentRestarted = False
'    'If Not igStdAloneMode Then
'    '    If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
'    '        edcLinkDestHelpMsg.LinkExecute "@" & "Done"
'    '    Else
'    '        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
'    '        edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
'    '        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
'    '        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
'    '        edcLinkDestHelpMsg.LinkExecute "Done"
'    '    End If
'    '    Do While Not igParentRestarted
'    '        DoEvents
'    '    Loop
'    'End If
'    Screen.MousePointer = vbDefault
'    If bmResearchSaved Then
'        If (Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES Then
'            ilRet = MsgBox("Please update the Vehicle default ACT1 Lineup codes if required", vbOKOnly + vbInformation, "Warning")
'        End If
'    End If
'
'    igManUnload = YES
'    Unload ImptSat
'    Set ImptSat = Nothing   'Remove data segment
'    igManUnload = NO
'End Sub
''*******************************************************
''*                                                     *
''*      Procedure Name:mTestRecLengths                 *
''*                                                     *
''*             Created:4/12/93       By:D. LeVine      *
''*            Modified:              By:               *
''*                                                     *
''*            Comments:Test if record lengths match    *
''*                                                     *
''*******************************************************
'Private Function mTestRecLengths() As Integer
'    Dim ilSizeError As Integer
'    Dim ilSize As Integer
'    ilSizeError = False
'    ilSize = mGetRecLength("Rdf.Btr")
'    If ilSize <> Len(tmRdf) Then
'        If ilSize > 0 Then
'            MsgBox "Rdf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmRdf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
'            ilSizeError = True
'        Else
'            MsgBox "Rdf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
'            ilSizeError = True
'        End If
'    End If
'    ilSize = mGetRecLength("Mnf.Btr")
'    If ilSize <> Len(tmMnf) Then
'        If ilSize > 0 Then
'            MsgBox "Mnf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmMnf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
'            ilSizeError = True
'        Else
'            MsgBox "Mnf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
'            ilSizeError = True
'        End If
'    End If
'    ilSize = mGetRecLength("Vef.Btr")
'    If ilSize <> Len(tmVef) Then
'        If ilSize > 0 Then
'            MsgBox "Vef size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmVef)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
'            ilSizeError = True
'        Else
'            MsgBox "Vef error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
'            ilSizeError = True
'        End If
'    End If
'    ilSize = mGetRecLength("Drf.Btr")
'    If ilSize <> Len(tmDrf) Then
'        If ilSize > 0 Then
'            MsgBox "Drf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmDrf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
'            ilSizeError = True
'        Else
'            MsgBox "Drf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
'            ilSizeError = True
'        End If
'    End If
'    ilSize = mGetRecLength("Dnf.Btr")
'    If ilSize <> Len(tmDnf) Then
'        If ilSize > 0 Then
'            MsgBox "Dnf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmDnf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
'            ilSizeError = True
'        Else
'            MsgBox "Dnf error: " & str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "Initialize Error"
'            ilSizeError = True
'        End If
'    End If
'    mTestRecLengths = ilSizeError
'End Function
'Private Sub plcDefault_Paint()
'    plcDefault.CurrentX = 0
'    plcDefault.CurrentY = 0
'    plcDefault.Print "Set as Vehicle Default"
'End Sub
'Private Sub plcScreen_Paint()
'    plcScreen.CurrentX = 0
'    plcScreen.CurrentY = 0
'    plcScreen.Print "Import Satellite Data"
'End Sub
'
'Private Sub mGetPrevDrfDpf(ilPrevDnfCode As Integer)
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilDrf                                                                                 *
''******************************************************************************************
'
'    Dim ilRet As Integer
'    Dim llDrfUpper As Long
'    Dim llDpfUpper As Long
'    Dim llDrf As Long
'
'    llDrfUpper = UBound(tmPrevDrf)
'    llDpfUpper = UBound(tmPrevDpf)
'    tmDrfSrchKey.iDnfCode = ilPrevDnfCode
'    tmDrfSrchKey.sDemoDataType = ""
'    tmDrfSrchKey.iMnfSocEco = 0
'    tmDrfSrchKey.iVefCode = 0
'    tmDrfSrchKey.sInfoType = ""
'    tmDrfSrchKey.iRdfcode = 0
'    ilRet = btrGetGreaterOrEqual(hmDrf, tmPrevDrf(llDrfUpper), imDrfRecLen, tmDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'    Do While (ilRet = BTRV_ERR_NONE) And (tmPrevDrf(llDrfUpper).iDnfCode = ilPrevDnfCode)
'        llDrfUpper = llDrfUpper + 1
'        ReDim Preserve tmPrevDrf(0 To llDrfUpper) As DRF
'        ilRet = btrGetNext(hmDrf, tmPrevDrf(llDrfUpper), imDrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'    Loop
'    For llDrf = 0 To llDrfUpper - 1 Step 1
'        tmDpfSrchKey1.lDrfCode = tmPrevDrf(llDrf).lCode
'        tmDpfSrchKey1.iMnfDemo = 0
'        ilRet = btrGetGreaterOrEqual(hmDpf, tmPrevDpf(llDpfUpper), imDpfRecLen, tmDpfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'        Do While (ilRet = BTRV_ERR_NONE) And (tmPrevDpf(llDpfUpper).lDrfCode = tmPrevDrf(llDrf).lCode)
'            llDpfUpper = llDpfUpper + 1
'            ReDim Preserve tmPrevDpf(0 To llDpfUpper) As DPF
'            ilRet = btrGetNext(hmDpf, tmPrevDpf(llDpfUpper), imDpfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'    Next llDrf
'End Sub
'
'Private Function mAddDrf() As Integer
'    Dim ilRet As Integer
'    Dim ilUpdate As Integer
'    Dim llDrfCode As Long
'    Dim llDrf As Long
'    Dim ilDay As Integer
'    Dim tlDrf As DRF
'
'    ilUpdate = False
'    llDrfCode = -1
'    If UBound(tmPrevDrf) > LBound(tmPrevDrf) Then
'        For llDrf = LBound(tmPrevDrf) To UBound(tmPrevDrf) - 1 Step 1
'            If (tmDrf.iVefCode = tmPrevDrf(llDrf).iVefCode) And (tmDrf.iRdfcode = tmPrevDrf(llDrf).iRdfcode) Then
'                If (tmDrf.sDemoDataType = tmPrevDrf(llDrf).sDemoDataType) And (tmDrf.iMnfSocEco = tmPrevDrf(llDrf).iMnfSocEco) And (tmDrf.sInfoType = tmPrevDrf(llDrf).sInfoType) Then
'                    If (tmDrf.iStartTime(0) = tmPrevDrf(llDrf).iStartTime(0)) And (tmDrf.iStartTime(1) = tmPrevDrf(llDrf).iStartTime(1)) Then
'                        If (tmDrf.iEndTime(0) = tmPrevDrf(llDrf).iEndTime(0)) And (tmDrf.iEndTime(1) = tmPrevDrf(llDrf).iEndTime(1)) Then
'                            llDrfCode = tmPrevDrf(llDrf).lCode
'                            For ilDay = 0 To 6 Step 1
'                                If tmDrf.sDay(ilDay) <> tmPrevDrf(llDrf).sDay(ilDay) Then
'                                    llDrfCode = -1
'                                    Exit For
'                                End If
'                            Next ilDay
'                            If llDrfCode <> -1 Then
'                                ilUpdate = True
'                                Exit For
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        Next llDrf
'    End If
'    If Not ilUpdate Then
'        ilRet = btrInsert(hmDrf, tmDrf, imDrfRecLen, INDEXKEY2)
'    Else
'        Do
'            tmDrfSrchKey2.lCode = llDrfCode
'            ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
'            If ilRet = BTRV_ERR_NONE Then
'                tmDrf.lCode = llDrfCode
'                ilRet = btrUpdate(hmDrf, tmDrf, imDrfRecLen)
'            Else
'                Exit Do
'            End If
'        Loop While ilRet = BTRV_ERR_CONFLICT
'    End If
'
'    'If tgSpf.sRemoteUsers = "Y" Then
'        Do
'            tmDrfSrchKey2.lCode = tmDrf.lCode
'            ilRet = btrGetEqual(hmDrf, tlDrf, imDrfRecLen, tmDrfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
'            tmDrf.iRemoteID = tgUrf(0).iRemoteUserID
'            tmDrf.lAutoCode = tmDrf.lCode
'            gPackDate smSyncDate, tmDrf.iSyncDate(0), tmDrf.iSyncDate(1)
'            gPackTime smSyncTime, tmDrf.iSyncTime(0), tmDrf.iSyncTime(1)
'            ilRet = btrUpdate(hmDrf, tmDrf, imDrfRecLen)
'        Loop While ilRet = BTRV_ERR_CONFLICT
'    'End If
'    mAddDrf = ilRet
'End Function
'
'Private Function mAddDpf() As Integer
'    Dim ilRet As Integer
'    Dim ilUpdate As Integer
'    Dim llDpfCode As Long
'    Dim llDpf As Long
'    Dim tlDpf As DPF
'
'    ilUpdate = False
'    llDpfCode = -1
'    If UBound(tmPrevDpf) > LBound(tmPrevDpf) Then
'        For llDpf = LBound(tmPrevDpf) To UBound(tmPrevDpf) - 1 Step 1
'            If (tmDpf.lDrfCode = tmPrevDpf(llDpf).lDrfCode) And (tmDpf.iMnfDemo = tmPrevDpf(llDpf).iMnfDemo) Then
'                llDpfCode = tmPrevDpf(llDpf).lCode
'                ilUpdate = True
'                Exit For
'            End If
'        Next llDpf
'    End If
'    If Not ilUpdate Then
'        ilRet = btrInsert(hmDpf, tmDpf, imDpfRecLen, INDEXKEY0)
'    Else
'        Do
'            tmDpfSrchKey.lCode = llDpfCode
'            ilRet = btrGetEqual(hmDpf, tlDpf, imDpfRecLen, tmDpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'            If ilRet = BTRV_ERR_NONE Then
'                tmDpf.lCode = llDpfCode
'                '7/31/06:  Retain Population for USA update
'                If imUpdateMode Then
'                    tmDpf.lPop = tlDpf.lPop
'                End If
'                ilRet = btrUpdate(hmDpf, tmDpf, imDpfRecLen)
'            Else
'                Exit Do
'            End If
'        Loop While ilRet = BTRV_ERR_CONFLICT
'    End If
'    mAddDpf = ilRet
'End Function
'
'Private Function mFindChName(slChannelCode As String, slVehicleName As String, slARBCode As String) As Integer
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *
''*  ilAsk                         ilRet                         ilRejected                *
''*                                                                                        *
''******************************************************************************************
'
'    Dim ilPos1 As Integer
'    Dim ilPos2 As Integer
'    Dim ilVefCode As Integer
'    Dim ilVef As Integer
'    Dim ilVpf As Integer
'    Dim ilLoop As Integer
'    Dim slMsg As String
'    Dim ilFound As Integer
'    Dim slStr As String
'    Dim llRet As Long
'    Dim llValue As Long
'    Dim llRg As Long
'
''    ilRejected = False
'    slChannelCode = ""
'    slVehicleName = ""
'    ilPos1 = InStr(1, smFieldValues(2), "Ch", vbTextCompare)
'    ilPos2 = InStr(1, smFieldValues(2), "-", vbTextCompare)
'    If (ilPos1 > 0) And (ilPos2 > 0) Then
'        'slStationCode = Left$(smFieldValues(2), ilPos1 + 1) & Mid$(smFieldValues(2), ilPos1 + 3, ilPos2 - (ilPos1 + 3))
'        slChannelCode = Mid$(smFieldValues(2), ilPos1 + 2, ilPos2 - (ilPos1 + 2))
'        slVehicleName = Trim$(Mid$(smFieldValues(2), ilPos2 + 1))
'    Else
'        slVehicleName = Trim$(smFieldValues(2))
'    End If
'    slARBCode = Trim$(smFieldValues(3))
''    ilVefCode = -1
''    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
''        If tgMVef(ilVef).sState <> "D" Then
''            If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "S") Or (tgMVef(ilVef).sType = "V") Then
''                'If ((slStationCode <> "") And (StrComp(slStationCode, Trim$(tgMVef(ilVef).sCodeStn), 1) = 0)) Or (StrComp(slVehicleName, Trim$(tgMVef(ilVef).sName), 1) = 0) Then
''                slStr = tgMVef(ilVef).sCodeStn
''                If InStr(1, slStr, "Ch", vbTextCompare) > 0 Then
''                    slStr = Mid$(slStr, 3)  'Remove Ch
''                End If
''                'Jim Request 1/18/06:  Replace Or with And
''                'If ((slStationCode <> "") And (Val(slStr)) = Val(slStationCode)) Or (StrComp(slVehicleName, Trim$(tgMVef(ilVef).sName), 1) = 0) Then
''                If ((slStationCode <> "") And (Val(slStr)) = Val(slStationCode)) And (StrComp(slVehicleName, Trim$(tgMVef(ilVef).sName), 1) = 0) Then
''                    ilVefCode = tgMVef(ilVef).iCode
''                    mFindChName = ilVefCode
''                    Exit Function
''                End If
''            End If
''        End If
''    Next ilVef
''    'Match only Channel
''    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
''        If tgMVef(ilVef).sState <> "D" Then
''            If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "S") Or (tgMVef(ilVef).sType = "V") Then
''                slStr = tgMVef(ilVef).sCodeStn
''                If InStr(1, slStr, "Ch", vbTextCompare) > 0 Then
''                    slStr = Mid$(slStr, 3)  'Remove Ch
''                End If
''                'Jim Request 1/18/06:  Replace Or with And
''                'If ((slStationCode <> "") And (Val(slStr)) = Val(slStationCode)) Or (StrComp(slVehicleName, Trim$(tgMVef(ilVef).sName), 1) = 0) Then
''                If ((slStationCode <> "") And (Val(slStr)) = Val(slStationCode)) Then
''                    For ilLoop = LBound(imVefCodeInDnf) To UBound(imVefCodeInDnf) - 1 Step 1
''                        If imVefCodeInDnf(ilLoop) = tgMVef(ilVef).iCode Then
''                            ilVefCode = tgMVef(ilVef).iCode
''                            mFindChName = ilVefCode
''                            Exit Function
''                        End If
''                    Next ilLoop
''                End If
''            End If
''        End If
''    Next ilVef
''    'Match only Channel
''    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
''        If tgMVef(ilVef).sState <> "D" Then
''            If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "S") Or (tgMVef(ilVef).sType = "V") Then
''                slStr = tgMVef(ilVef).sCodeStn
''                If InStr(1, slStr, "Ch", vbTextCompare) > 0 Then
''                    slStr = Mid$(slStr, 3)  'Remove Ch
''                End If
''                'Jim Request 1/18/06:  Replace Or with And
''                'If ((slStationCode <> "") And (Val(slStr)) = Val(slStationCode)) Or (StrComp(slVehicleName, Trim$(tgMVef(ilVef).sName), 1) = 0) Then
''                ilAsk = False
''                If ((slStationCode <> "") And (Val(slStr)) = Val(slStationCode)) Then
''                    ilAsk = True
''                    For ilLoop = LBound(imRejectedVefCode) To UBound(imRejectedVefCode) - 1 Step 1
''                        If imRejectedVefCode(ilLoop) = tgMVef(ilVef).iCode Then
''                            ilAsk = False
''                            ilRejected = True
''                            Exit For
''                        End If
''                    Next ilLoop
''                End If
''                If ilAsk Then
''                    ilRet = MsgBox("Channel Numbers " & slStationCode & " Match but Channel Names Do Not. Arb =" & slVehicleName & " Counterpoint Name =" & Trim$(tgMVef(ilVef).sName) & ", Import Research Anyway", vbInformation + vbYesNo, "Research")
''                    If ilRet = vbYes Then
''                        slMsg = "Imported: " & "Channel Numbers " & slStationCode & " Match but Channel Names Do Not. Arb =" & slVehicleName & " Counterpoint Name =" & Trim$(tgMVef(ilVef).sName)
''                        Print #hmTo, slMsg
''                        If gOkAddStrToListBox(slMsg, lmLen, imShowMsg) Then
''                            lbcErrors.AddItem slMsg
''                            If (Traffic.pbcArial.TextWidth(slMsg)) > lmMaxWidth Then
''                                lmMaxWidth = Traffic.pbcArial.TextWidth(slMsg)
''                                If lmMaxWidth > lbcErrors.Width Then
''                                    llValue = lmMaxWidth / 15 + 120
''                                    llRg = 0
''                                    llRet = SendMessageByNum(lbcErrors.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
''                                End If
''                            End If
''                        Else
''                            imShowMsg = False
''                        End If
''                        ilVefCode = tgMVef(ilVef).iCode
''                        mFindChName = ilVefCode
''                        Exit Function
''                    End If
''                    ilRejected = True
''                    slMsg = "Rejected: " & "Channel Numbers " & slStationCode & " Match but Channel Names Do Not. Arb =" & slVehicleName & " Counterpoint Name =" & Trim$(tgMVef(ilVef).sName)
''                    Print #hmTo, slMsg
''                    If gOkAddStrToListBox(slMsg, lmLen, imShowMsg) Then
''                        lbcErrors.AddItem slMsg
''                        If (Traffic.pbcArial.TextWidth(slMsg)) > lmMaxWidth Then
''                            lmMaxWidth = Traffic.pbcArial.TextWidth(slMsg)
''                            If lmMaxWidth > lbcErrors.Width Then
''                                llValue = lmMaxWidth / 15 + 120
''                                llRg = 0
''                                llRet = SendMessageByNum(lbcErrors.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
''                            End If
''                        End If
''                    Else
''                        imShowMsg = False
''                    End If
''                    imRejectedVefCode(UBound(imRejectedVefCode)) = tgMVef(ilVef).iCode
''                    ReDim Preserve imRejectedVefCode(1 To UBound(imRejectedVefCode) + 1) As Integer
''                End If
''            End If
''        End If
''    Next ilVef
''    'Match only Name
''    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
''        If tgMVef(ilVef).sState <> "D" Then
''            If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "S") Or (tgMVef(ilVef).sType = "V") Then
''                slStr = tgMVef(ilVef).sCodeStn
''                If InStr(1, slStr, "Ch", vbTextCompare) > 0 Then
''                    slStr = Mid$(slStr, 3)  'Remove Ch
''                End If
''                'Jim Request 1/18/06:  Replace Or with And
''                'If ((slStationCode <> "") And (Val(slStr)) = Val(slStationCode)) Or (StrComp(slVehicleName, Trim$(tgMVef(ilVef).sName), 1) = 0) Then
''                If (StrComp(slVehicleName, Trim$(tgMVef(ilVef).sName), 1) = 0) Then
''                    For ilLoop = LBound(imVefCodeInDnf) To UBound(imVefCodeInDnf) - 1 Step 1
''                        If imVefCodeInDnf(ilLoop) = tgMVef(ilVef).iCode Then
''                            ilVefCode = tgMVef(ilVef).iCode
''                            mFindChName = ilVefCode
''                            Exit Function
''                        End If
''                    Next ilLoop
''                End If
''            End If
''        End If
''    Next ilVef
''    'Match only Name
''    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
''        If tgMVef(ilVef).sState <> "D" Then
''            If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "S") Or (tgMVef(ilVef).sType = "V") Then
''                slStr = tgMVef(ilVef).sCodeStn
''                If InStr(1, slStr, "Ch", vbTextCompare) > 0 Then
''                    slStr = Mid$(slStr, 3)  'Remove Ch
''                End If
''                'Jim Request 1/18/06:  Replace Or with And
''                'If ((slStationCode <> "") And (Val(slStr)) = Val(slStationCode)) Or (StrComp(slVehicleName, Trim$(tgMVef(ilVef).sName), 1) = 0) Then
''                ilAsk = False
''                If (StrComp(slVehicleName, Trim$(tgMVef(ilVef).sName), 1) = 0) Then
''                    ilAsk = True
''                    For ilLoop = LBound(imRejectedVefCode) To UBound(imRejectedVefCode) - 1 Step 1
''                        If imRejectedVefCode(ilLoop) = tgMVef(ilVef).iCode Then
''                            ilAsk = False
''                            ilRejected = True
''                            Exit For
''                        End If
''                    Next ilLoop
''                End If
''                If ilAsk Then
''                    ilRet = MsgBox("Channel Names " & slVehicleName & " Match but Channel Numbers Do Not. Arb =" & slStationCode & " Counterpoint Name =" & slStr & ", Import Research Anyway", vbInformation + vbYesNo, "Research")
''                    If ilRet = vbYes Then
''                        slMsg = "Imported: " & "Channel Names " & slVehicleName & " Match but Channel Numbers Do Not. Arb =" & slStationCode & " Counterpoint Name =" & slStr
''                        Print #hmTo, slMsg
''                        If gOkAddStrToListBox(slMsg, lmLen, imShowMsg) Then
''                            lbcErrors.AddItem slMsg
''                            If (Traffic.pbcArial.TextWidth(slMsg)) > lmMaxWidth Then
''                                lmMaxWidth = Traffic.pbcArial.TextWidth(slMsg)
''                                If lmMaxWidth > lbcErrors.Width Then
''                                    llValue = lmMaxWidth / 15 + 120
''                                    llRg = 0
''                                    llRet = SendMessageByNum(lbcErrors.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
''                                End If
''                            End If
''                        Else
''                            imShowMsg = False
''                        End If
''                        ilVefCode = tgMVef(ilVef).iCode
''                        mFindChName = ilVefCode
''                        Exit Function
''                    End If
''                    ilRejected = True
''                    slMsg = "Rejected: " & "Channel Names " & slVehicleName & " Match but Channel Numbers Do Not. Arb =" & slStationCode & " Counterpoint Name =" & slStr
''                    Print #hmTo, slMsg
''                    If gOkAddStrToListBox(slMsg, lmLen, imShowMsg) Then
''                        lbcErrors.AddItem slMsg
''                        If (Traffic.pbcArial.TextWidth(slMsg)) > lmMaxWidth Then
''                            lmMaxWidth = Traffic.pbcArial.TextWidth(slMsg)
''                            If lmMaxWidth > lbcErrors.Width Then
''                                llValue = lmMaxWidth / 15 + 120
''                                llRg = 0
''                                llRet = SendMessageByNum(lbcErrors.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
''                            End If
''                        End If
''                    Else
''                        imShowMsg = False
''                    End If
''                    imRejectedVefCode(UBound(imRejectedVefCode)) = tgMVef(ilVef).iCode
''                    ReDim Preserve imRejectedVefCode(1 To UBound(imRejectedVefCode) + 1) As Integer
''                End If
''            End If
''        End If
''    Next ilVef
''    If Not ilRejected Then
''        slStr = smFieldValues(2)
''        ilFound = False
''        For ilLoop = LBound(smVehNotFound) To UBound(smVehNotFound) - 1 Step 1
''            If StrComp(slStr, smVehNotFound(ilLoop), 1) = 0 Then
''                ilFound = True
''            End If
''        Next ilLoop
''        If Not ilFound Then
''            smVehNotFound(UBound(smVehNotFound)) = slStr
''            ReDim Preserve smVehNotFound(0 To UBound(smVehNotFound) + 1) As String
''            slMsg = "Unable to Find Vehicle " & slStr & " record not added"
''            Print #hmTo, slMsg
''            If gOkAddStrToListBox(slMsg, lmLen, imShowMsg) Then
''                lbcErrors.AddItem slMsg
''                If (Traffic.pbcArial.TextWidth(slMsg)) > lmMaxWidth Then
''                    lmMaxWidth = Traffic.pbcArial.TextWidth(slMsg)
''                    If lmMaxWidth > lbcErrors.Width Then
''                        llValue = lmMaxWidth / 15 + 120
''                        llRg = 0
''                        llRet = SendMessageByNum(lbcErrors.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
''                    End If
''                End If
''            Else
''                imShowMsg = False
''            End If
''        End If
''    End If
'    '6/28/06:  Replace testing for Match of channel and vehicle name with only testing for matching Arbitron Code
'    For ilVpf = LBound(tgVpf) To UBound(tgVpf) - 1 Step 1
'        If StrComp(slARBCode, Trim$(tgVpf(ilVpf).sARBCode), vbTextCompare) = 0 Then
'            ilVef = gBinarySearchVef(tgVpf(ilVpf).iVefKCode)
'            If ilVef <> -1 Then
'                ilVefCode = tgMVef(ilVef).iCode
'                mFindChName = ilVefCode
'                Exit Function
'            End If
'        End If
'    Next ilVpf
'    ilFound = False
'    slStr = smFieldValues(2) & " ARB Code " & smFieldValues(3)
'    For ilLoop = LBound(smVehNotFound) To UBound(smVehNotFound) - 1 Step 1
'        If StrComp(slStr, smVehNotFound(ilLoop), 1) = 0 Then
'            ilFound = True
'        End If
'    Next ilLoop
'    If Not ilFound Then
'        smVehNotFound(UBound(smVehNotFound)) = slStr
'        ReDim Preserve smVehNotFound(0 To UBound(smVehNotFound) + 1) As String
'        slMsg = "Unable to Find Vehicle " & slStr & " record not added"
'        Print #hmTo, slMsg
'        If gOkAddStrToListBox(slMsg, lmLen, imShowMsg) Then
'            lbcErrors.AddItem slMsg
'            If (Traffic.pbcArial.TextWidth(slMsg)) > lmMaxWidth Then
'                lmMaxWidth = Traffic.pbcArial.TextWidth(slMsg)
'                If lmMaxWidth > lbcErrors.Width Then
'                    llValue = lmMaxWidth / 15 + 120
'                    llRg = 0
'                    llRet = SendMessageByNum(lbcErrors.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
'                End If
'            End If
'        Else
'            imShowMsg = False
'        End If
'    End If
'    mFindChName = 0
'    Exit Function
'End Function
'
'Private Sub mSetUpdateMode(ilPrevDnfCode As Integer)
''7/31/06:  Retain Population for USA update
'    Dim ilRet As Integer
'
'    tmDnfSrchKey.iCode = ilPrevDnfCode
'    ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
'    If ilRet = BTRV_ERR_NONE Then
'        If Trim$(tmDnf.sEstListenerOrUSA) = "U" Then
'            imUpdateMode = True
'        End If
'    End If
'End Sub
Private Sub cmcAud_Click()

End Sub

