VERSION 5.00
Begin VB.Form DDFReorg 
   Appearance      =   0  'Flat
   Caption         =   "DDF Reorganize"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   1260
   ClientWidth     =   11370
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
   LinkTopic       =   "DoneMsg"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5970
   ScaleWidth      =   11370
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmcCompare 
      Appearance      =   0  'Flat
      Caption         =   "&Compare DDF's"
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
      Left            =   6375
      TabIndex        =   20
      Top             =   5505
      Width           =   1455
   End
   Begin VB.CommandButton cmcDone 
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
      Left            =   3765
      TabIndex        =   9
      Top             =   5505
      Width           =   1050
   End
   Begin VB.CommandButton cmcProcess 
      Appearance      =   0  'Flat
      Caption         =   "&Process"
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
      Height          =   285
      Left            =   5070
      TabIndex        =   10
      Top             =   5505
      Width           =   1050
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   540
      Top             =   5430
   End
   Begin VB.PictureBox plcDDF 
      FillStyle       =   0  'Solid
      Height          =   5040
      Left            =   150
      ScaleHeight     =   4980
      ScaleWidth      =   11055
      TabIndex        =   0
      Top             =   105
      Width           =   11115
      Begin VB.PictureBox pbcGauge 
         Height          =   210
         Left            =   7185
         ScaleHeight     =   150
         ScaleWidth      =   3660
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4620
         Width           =   3720
         Begin VB.Line lncGauge 
            BorderColor     =   &H80000002&
            BorderStyle     =   3  'Dot
            X1              =   60
            X2              =   3615
            Y1              =   75
            Y2              =   75
         End
      End
      Begin VB.PictureBox pbcNoFiles 
         Height          =   210
         Left            =   7185
         ScaleHeight     =   150
         ScaleWidth      =   3660
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3870
         Width           =   3720
         Begin VB.Line lncNoFiles 
            BorderColor     =   &H80000002&
            BorderStyle     =   3  'Dot
            X1              =   60
            X2              =   3615
            Y1              =   75
            Y2              =   75
         End
      End
      Begin VB.PictureBox plcFile 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3675
         ScaleHeight     =   255
         ScaleWidth      =   3255
         TabIndex        =   17
         Top             =   300
         Width           =   3255
         Begin VB.OptionButton rbcFile 
            Caption         =   "Copy (xxx.new)"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1245
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   1605
         End
         Begin VB.OptionButton rbcFile 
            Caption         =   "Create"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   375
            TabIndex        =   18
            Top             =   0
            Value           =   -1  'True
            Width           =   885
         End
      End
      Begin VB.DirListBox lbcDBPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   7380
         TabIndex        =   12
         Top             =   990
         Width           =   3300
      End
      Begin VB.DriveListBox cbcDBDrive 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7380
         TabIndex        =   11
         Top             =   600
         Width           =   3300
      End
      Begin VB.ListBox lbcNewFile 
         Appearance      =   0  'Flat
         Height          =   2175
         Left            =   3660
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   2730
         Width           =   3300
      End
      Begin VB.ListBox lbcOldFile 
         Appearance      =   0  'Flat
         Height          =   2175
         Left            =   165
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   2730
         Width           =   3300
      End
      Begin VB.DriveListBox cbcNewDrive 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3660
         TabIndex        =   6
         Top             =   600
         Width           =   3300
      End
      Begin VB.DirListBox lbcNewPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   3660
         TabIndex        =   7
         Top             =   990
         Width           =   3300
      End
      Begin VB.DirListBox lbcOldPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1440
         Left            =   165
         TabIndex        =   3
         Top             =   975
         Width           =   3300
      End
      Begin VB.DriveListBox cbcOldDrive 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   165
         TabIndex        =   2
         Top             =   600
         Width           =   3300
      End
      Begin VB.Label lacFileName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7080
         TabIndex        =   16
         Top             =   4170
         Width           =   3915
      End
      Begin VB.Label lacNoFiles 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7080
         TabIndex        =   14
         Top             =   3600
         Width           =   3915
      End
      Begin VB.Label lacTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7080
         TabIndex        =   15
         Top             =   4335
         Width           =   3915
      End
      Begin VB.Label lacDB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "DATA Path"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7380
         TabIndex        =   13
         Top             =   60
         Width           =   3300
      End
      Begin VB.Label lacOldDDF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "DDF Path: Old"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   210
         TabIndex        =   1
         Top             =   60
         Width           =   3300
      End
      Begin VB.Label lacNewDDF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "DDF Path: New"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3705
         TabIndex        =   5
         Top             =   75
         Width           =   3300
      End
   End
   Begin VB.Label lacRename 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   23
      Top             =   5175
      Width           =   11100
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Spec"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Spec..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Spec"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save Spec &As..."
      End
      Begin VB.Menu mnuFileSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Ex&it"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help On ..."
      End
      Begin VB.Menu mnuHelpSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About ..."
      End
   End
End
Attribute VB_Name = "DDFReorg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: DDFReorg.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the DDFReorg screen code
'
'Script file layout (file name is FieldSet.New):
'Form                                          Example
'[Table Name]                                  [ADF Advertisers]
'PageSize=4096                                 PageSize=4096
'BlankTruncation=Yes
'DataCompression=Yes
'IndexBalancing=Yes
'KeyOnly=Yes
'FreeSpace=10   or 20   or 30
'Init:NewFieldName=init value for field name    Init:adfPriceType=N
'                  or name of field to get value from Init:adfPriceType=adfXXX
'                  or name of field plus adjusted Init:chfiSlspCommPct1 = chfComm1/100
'Rename:OldFieldName=NewFieldName               Rename:adfOldXYZ=adfNewXYZ
'Adjust:FieldName=Oper Value                    Adjust:adfTarget=*10
'                    (allowed Oper = + - * /)
'
'
'
'Command Line parameters
'  /N:New DDF path
'  /O:Old DDF path
'  /F:Database location
'  /CreateOnly:List of files to Create Only (Don't copy records)
'              pipe character as separator with no spaces
'              /CreateOnly:ast.mkd|cptt.mkd
'  /Run    run compare and process
'  /Exit   Exit after process
'  /No_btr   ignore btr files
'  /No_Mkd   ignore mkd files
'  /RemoteUser  only compare files that exists in the database
'
Option Explicit
Option Compare Text
Dim smDBPath As String
Dim smCreateDBPath As String
Dim smOldDDFPath As String
Dim smNewDDFPath As String
Dim smCreateOnlyArray() As String
Dim hmDB As Integer
Dim lmTotalNoRec As Long
Dim lmProcessedNoRec As Long
Dim hmTo As Integer   'To file hanle
Dim hmFrom As Integer   'From file hanle
Dim smCmmdFileName As String
Dim hmFile As Integer
Dim hmOldFile As Integer
Dim hmNewFile As Integer
Dim imFileID As Integer
Dim lmOldActualRecLen As Long
Dim lmOldBaseRecLen As Long
Dim lmNewActualRecLen As Long
Dim lmNewBaseRecLen As Long
Dim imVarLenNote As Integer
Dim tmOldRec As BASEREC
Dim smFieldValue As String
Dim imPageSize(0 To 3) As Integer
Dim tmOldDDFFile() As DDFFILE    'DDF File- File Names
Dim tmNewDDFFile() As DDFFILE    'DDF File- File Names
Dim tmOldDDFField() As DDFFIELD    'DDF File- Field Names
Dim tmNewDDFField() As DDFFIELD    'DDF File- Field Names
Dim tmFieldSrchKey1 As DDFFIELD1
Dim tmOldDDFIndex() As DDFINDEX    'DDF File- Index Names
Dim tmNewDDFIndex() As DDFINDEX    'DDF File- Index Names
Dim tmIndexSrchKey0 As DDFINDEX0
Dim tmInitValues() As INITVALUES
Dim tmAdjustValues() As ADJUSTVALUES
Dim smHeading As String
Dim imExporting As Integer
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imReadingFile As Integer
Dim imKeyDownRow As Integer
Dim imNewVarLenRec As Integer
Dim lmOldVarLenRec As Long
Dim imOldVarType As Integer    '12 or 13
Dim imNewVarType As Integer    '12 or 13
Dim lmMaxNoRec As Long
Dim imVBCAdjust As Integer
Dim imEnableOnCancel As Integer 'Enable cmcDone and cmcProcess on Cancel
Dim imProcURF As Integer    'True=Processing URF
Dim imExit As Integer
Dim imRun As Integer
Dim imRU As Integer
Dim imSelectBTR As Integer
Dim imSelectMKD As Integer
Dim imSelectENG As Integer
Dim lmActualRecLen As Long
Dim imBaseRecLen As Long
Dim tmGetDDFIndex As DDFINDEX
Dim tmGetDDFFile As DDFFILE
Dim tmGetDDFFIELD As DDFFIELD
Dim bErrorFound As Boolean

Public Sub mGetIndexFields()
    Dim tlDDFField As DDFFIELD
    tlDDFField.iOffset = 0
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFIndex.iFileID = Val(smFieldValue)
    tlDDFField.iOffset = 2
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFIndex.iFieldID = Val(smFieldValue)
    tlDDFField.iOffset = 4
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFIndex.iNumber = Val(smFieldValue)
    tlDDFField.iOffset = 6
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFIndex.iPart = Val(smFieldValue)
    tlDDFField.iOffset = 8
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFIndex.iFlag = Val(smFieldValue)
End Sub

Public Sub mGetFileFields()
    Dim tlDDFField As DDFFIELD
    tlDDFField.iOffset = 0
    tlDDFField.iSize = 2
    tlDDFField.sDataType = Chr(1)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFFile.iFileID = Val(smFieldValue)
    tlDDFField.iOffset = 2
    tlDDFField.iSize = 20
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFFile.sName = smFieldValue
    tlDDFField.iOffset = 22
    tlDDFField.iSize = 64
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFFile.sLocation = smFieldValue
    tlDDFField.iOffset = 86
    tlDDFField.iSize = 1
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFFile.sFlags = smFieldValue
    tlDDFField.iOffset = 87
    tlDDFField.iSize = 10
    tlDDFField.sDataType = Chr(0)
    tlDDFField.sDec = ""
    mGetField tlDDFField
    tmGetDDFFile.sReserved = smFieldValue
End Sub

Public Sub mGetFieldFields()
    Dim tlGDDFField As DDFFIELD
    tlGDDFField.iOffset = 0
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.iFieldID = Val(smFieldValue)
    tlGDDFField.iOffset = 2
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.iFileID = Val(smFieldValue)
    tlGDDFField.iOffset = 4
    tlGDDFField.iSize = 20
    tlGDDFField.sDataType = Chr(0)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.sName = smFieldValue
    tmGetDDFFIELD.sDataType = Chr(tmOldRec.sChar(25))
    tlGDDFField.iOffset = 25
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.iOffset = Val(smFieldValue)
    tlGDDFField.iOffset = 27
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.iSize = Val(smFieldValue)
    tlGDDFField.iOffset = 29
    tlGDDFField.iSize = 1
    tlGDDFField.sDataType = Chr(0)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.sDec = smFieldValue
    tlGDDFField.iOffset = 30
    tlGDDFField.iSize = 2
    tlGDDFField.sDataType = Chr(1)
    tlGDDFField.sDec = ""
    mGetField tlGDDFField
    tmGetDDFFIELD.iFlags = Val(smFieldValue)
End Sub

Private Sub cbcDBDrive_Change()
    lbcDBPath.Path = cbcDBDrive.Drive
    mSetCommands
End Sub

Private Sub cbcNewDrive_Change()
    lbcNewPath.Path = cbcNewDrive.Drive
    lbcNewFile.Clear
    mSetCommands
End Sub
Private Sub cbcOldDrive_Change()
    lbcOldPath.Path = cbcOldDrive.Drive
    lbcOldFile.Clear
    cbcNewDrive.Drive = cbcOldDrive.Drive
    cbcDBDrive.Drive = cbcOldDrive.Drive
    mSetCommands
End Sub

Private Sub cmcCompare_Click()
    lacRename.Caption = "Checking..."
    Screen.MousePointer = vbHourglass
    mSetFiles
    Screen.MousePointer = vbDefault
    If cmcProcess.Enabled = False Then
        lacRename.Caption = "Nothing to do, DDFs up to date."
    End If
End Sub

Private Sub cmcDone_Click()
    mTerminate
End Sub

Private Sub cmcProcess_Click()
    bErrorFound = False
    mProcessRecords
    If imExit Then
        mTerminate
    End If
    cmcProcess.Enabled = False
    cmcDone.Caption = "&Done"
    If bErrorFound = True Then
        lacRename.Caption = "Errors were detected during processing.  Check Log"
    End If
End Sub

Private Sub Form_Load()
    sgExePath = App.Path
    mInit
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
    DDFReorg.Show
    DoEvents
    If imRun Then
        cmcCompare_Click
        If cmcProcess.Enabled Then
            cmcProcess_Click
        Else
            'No changes in DDF's
            mMoveDDFs
            mTerminate
        End If
    ElseIf imRU Then
        cmcCompare_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    Erase tmOldDDFFile
    Erase tmNewDDFFile
    Erase tmOldDDFField
    Erase tmNewDDFField
    Erase tmOldDDFIndex
    Erase tmNewDDFIndex
    'Reset used instead of Close to cause # Clients on network to be decrement
    Set DDFReorg = Nothing   'Remove data segment
    btrStopAppl
    End
End Sub

Private Sub lbcDBPath_Change()
    smDBPath = lbcDBPath.Path & "\"
    sgDBPath = smDBPath
End Sub

Private Sub lbcNewFile_Click()
    mSetCommands
End Sub

Private Sub lbcNewPath_Change()
    smNewDDFPath = lbcNewPath.Path & "\"
    Screen.MousePointer = vbHourglass
    mFileBuild smNewDDFPath, ".DDF", tmNewDDFFile()
    mFilePop tmNewDDFFile(), lbcNewFile
    cmcProcess.Enabled = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub lbcOldPath_Change()
    Dim ilPos As Integer
    Dim ilPos1 As Integer
    Dim slNewDDFPath As String
    
    smOldDDFPath = lbcOldPath.Path & "\"
    Screen.MousePointer = vbHourglass
    mFileBuild smOldDDFPath, ".DDF", tmOldDDFFile()
    mFilePop tmOldDDFFile(), lbcOldFile
    cmcProcess.Enabled = False
    slNewDDFPath = ""
    ilPos = InStr(1, smOldDDFPath, "\V", vbTextCompare)
    If ilPos > 0 Then
        ilPos1 = InStr(ilPos + 1, smOldDDFPath, "\", vbTextCompare)
        If ilPos1 > 0 Then
            slNewDDFPath = Left$(smOldDDFPath, ilPos1) & "NewDDF"
        Else
            slNewDDFPath = smOldDDFPath & "\NewDDF"
        End If
    Else
        ilPos = InStr(1, smOldDDFPath, "\Prod", vbTextCompare)
        If ilPos > 0 Then
            ilPos1 = InStr(ilPos + 1, smOldDDFPath, "\", vbTextCompare)
            If ilPos1 > 0 Then
                slNewDDFPath = Left$(smOldDDFPath, ilPos1) & "NewDDF"
            Else
                slNewDDFPath = smOldDDFPath & "\NewDDF"
            End If
        Else
            ilPos = InStr(1, smOldDDFPath, "\Test", vbTextCompare)
            If ilPos > 0 Then
                ilPos1 = InStr(ilPos + 1, smOldDDFPath, "\", vbTextCompare)
                If ilPos1 > 0 Then
                    slNewDDFPath = Left$(smOldDDFPath, ilPos1) & "NewDDF"
                Else
                    slNewDDFPath = smOldDDFPath & "\NewDDF"
                End If
            End If
        End If
    End If
    If slNewDDFPath <> "" Then
        lbcNewPath.Path = slNewDDFPath
    End If
    lbcDBPath.Path = lbcOldPath.Path
    Screen.MousePointer = vbDefault
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gBtrvErrorMsg                   *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Show error btrieve message      *
'*                                                     *
'*******************************************************
Private Sub mBtrvErrorMsg(ilErrCode As Integer, slMess As String)
'
'   slMess = "Open Arf"
'   gBtrvErrorMsg ilRet, slMess
'   Where:
'       ilErrCode (I)- Error number from btrieve
'       slMess (I) - Message to be added to the btrieve error message
'       MainForm (I)- Name of Form to unload if error exists
'
    Dim slErrMsg As String
    If ilErrCode = BTRV_ERR_NONE Then
        Exit Sub
    End If
    slErrMsg = "Btrieve: "
    Select Case ilErrCode
        Case BTRV_ERR_INVALID_OP    '1
            slErrMsg = slErrMsg & "Invalid Operation (1)"
        Case BTRV_ERR_IO_ERR        '2
            slErrMsg = slErrMsg & "I/O Error (2)"
        Case BTRV_ERR_NOT_OPEN      '3
            slErrMsg = slErrMsg & "File not Opened (3)"
        Case BTRV_ERR_KEY_NOT_FOUND '4
            slErrMsg = slErrMsg & "Key Value not Found (4)"
        Case BTRV_ERR_DUPLICATE_KEY '5
            slErrMsg = slErrMsg & "Duplicate Key Value (5)"
        Case BTRV_ERR_INVALID_KEY   '6
            slErrMsg = slErrMsg & "Invalid Key Number (6)"
        Case BTRV_ERR_DIFF_KEY      '7
            slErrMsg = slErrMsg & "Different Key Number (7)"
        Case BTRV_ERR_INVALID_POS '8
            slErrMsg = slErrMsg & "Invalid Positioning (8)"
        Case BTRV_ERR_END_OF_FILE '9
            slErrMsg = slErrMsg & "End of File (9)"
        Case BTRV_ERR_MOD_KEY_VALUE '10
            slErrMsg = slErrMsg & "Modifiable Key Value Error (10)"
        Case BTRV_ERR_INVALID_FNAME '11
            slErrMsg = slErrMsg & "Invalid File Name (11)"
        Case BTRV_ERR_FILE_NOT_FOUND    '12
            slErrMsg = slErrMsg & "File Not Found (12)"
        Case BTRV_ERR_EXT_FILE      '13
            slErrMsg = slErrMsg & "Extended File Error (13)"
        Case BTRV_ERR_PREIMAGE_OPEN '14
            slErrMsg = slErrMsg & "Per-image Open Error (14)"
        Case BTRV_ERR_PREIMAGE_IO   '15
            slErrMsg = slErrMsg & "Pre-image I/O Error (15)"
        Case BTRV_ERR_EXPANSION '16
            slErrMsg = slErrMsg & "Expansion Error (16)"
        Case BTRV_ERR_CLOSE '17
            slErrMsg = slErrMsg & "Close Error (17)"
        Case BTRV_ERR_DISK_FULL '18
            slErrMsg = slErrMsg & "Disk Full (18)"
        Case BTRV_ERR_UNRECOVERABLE '19
            slErrMsg = slErrMsg & "Unrecoverable Error (19)"
        Case BTRV_ERR_MGR_INACTIVE  '20
            slErrMsg = slErrMsg & "Record Manager Inactive (20)"
        Case BTRV_ERR_KEYBUF_LENGTH '21
            slErrMsg = slErrMsg & "Key Buffer Too Short (21)"
        Case BTRV_ERR_DATABUF_LENGTH    '22
            slErrMsg = slErrMsg & "Data Buffer Length (22)"
        Case BTRV_ERR_POSBLK_LENGTH '23
            slErrMsg = slErrMsg & "Position Block Length (23)"
        Case BTRV_ERR_PAGESIZE  '24
            slErrMsg = slErrMsg & "Page Size Error (24)"
        Case BTRV_ERR_CREATE_IO '25
            slErrMsg = slErrMsg & "Create I/O Error"
        Case BTRV_ERR_NUMKEYS   '26
            slErrMsg = slErrMsg & "Number of Keys (26)"
        Case BTRV_ERR_INVALID_KEYPOS    '27
            slErrMsg = slErrMsg & "Invalid Key Position (27)"
        Case BTRV_ERR_REC_LENGTH    '28
            slErrMsg = slErrMsg & "Invalid Record Length (28)"
        Case BTRV_ERR_KEY_LENGTH    '29
            slErrMsg = slErrMsg & "Invalid Key Length (29)"
        Case BTRV_ERR_NOT_BTRV_FILE '30
            slErrMsg = slErrMsg & "Not a Btrieve File (30)"
        Case BTRV_ERR_ALREADY_EXTD  '31
            slErrMsg = slErrMsg & "File Already Extended (31)"
        Case BTRV_ERR_EXTD_IO   '32
            slErrMsg = slErrMsg & "Extend I/O Error (32)"
        Case BTRV_ERR_INVALID_EXT_NAME  '34
            slErrMsg = slErrMsg & "Invalid Extension Name (34)"
        Case BTRV_ERR_DIRECTORY '35
            slErrMsg = slErrMsg & "Directory Error (35)"
        Case BTRV_ERR_TRANSACTION   '36
            slErrMsg = slErrMsg & "Transaction Error"
        Case BTRV_ERR_TRANS_ACTIVE  '37
            slErrMsg = slErrMsg & "Transaction in Active (37)"
        Case BTRV_ERR_TRANS_FILE_IO '38
            slErrMsg = slErrMsg & "Transaction Control File I/O Error (38)"
        Case BTRV_ERR_END_ABORT_TRANS   '39
            slErrMsg = slErrMsg & "End/Abort Transaction Error (39)"
        Case BTRV_ERR_TRANS_MAX_FILES   '40
            slErrMsg = slErrMsg & "Transaction Max Files (40)"
        Case BTRV_ERR_OP_NOT_ALLOWED    '41
            slErrMsg = slErrMsg & "Operation Not allowed (41)"
        Case BTRV_ERR_ACCEL_ACCESS  '42
            slErrMsg = slErrMsg & "Incomplete Accelerated Access (42)"
        Case BTRV_ERR_INVALID_REC_ADDR  '43
            slErrMsg = slErrMsg & "Invalid Record Address (43)"
        Case BTRV_ERR_NULL_KEY_PATH '44
            slErrMsg = slErrMsg & "Null Key Path (44)"
        Case BTRV_ERR_INCON_KEY_FLAGS   '45
            slErrMsg = slErrMsg & "Inconsistent Key Flags (45)"
        Case BTRV_ERR_ACCESS_DENIED '46
            slErrMsg = slErrMsg & "Access to File Denied (46)"
        Case BTRV_ERR_MAX_OPEN_FILES    '47
            slErrMsg = slErrMsg & "Maximum Open Files (47)"
        Case BTRV_ERR_INVALID_ALT_SEQ   '48
            slErrMsg = slErrMsg & "Invalid Alternate Sequence Definition (48)"
        Case BTRV_ERR_KEY_TYPE  '49
            slErrMsg = slErrMsg & "Key Type Error"
        Case BTRV_ERR_OWNER_SET '50
            slErrMsg = slErrMsg & "Owner Already Set (50)"
        Case BTRV_ERR_INVALID_OWNER '51
            slErrMsg = slErrMsg & "Invalid Owner (51)"
        Case BTRV_ERR_WRITE_CACHE   '52
            slErrMsg = slErrMsg & "Error Writing Cache (52)"
        Case BTRV_ERR_INVALID_INTF  '53
            slErrMsg = slErrMsg & "Invalid Interface (53)"
        Case BTRV_ERR_VARIABLE_PAGE '54
            slErrMsg = slErrMsg & "Variable Page Error (54)"
        Case 55
            slErrMsg = slErrMsg & "Autoincrement Error (55)"
        Case BTRV_ERR_INCOMPLT_INDEX    '56
            slErrMsg = slErrMsg & "Incomplete Index"
        Case BTRV_ERR_EXPAND_MEM    '57
            slErrMsg = slErrMsg & "Expanded Memory Error (57)"
        Case BTRV_ERR_COMPBUF_SIZE  '58
            slErrMsg = slErrMsg & "Compression Buffer Too Short (58)"
        Case BTRV_ERR_FILE_EXISTS   '59
            slErrMsg = slErrMsg & "File Already Exists (59)"
        Case BTRV_ERR_REJECT_COUNT  '60
            slErrMsg = slErrMsg & "Reject Count Reached (60)"
        Case BTRV_ERR_WORK_SPACE_SIZE   '61
            slErrMsg = slErrMsg & "Work Space Too Small (61)"
        Case BTRV_ERR_INCORRECT_DESCP   '62
            slErrMsg = slErrMsg & "Incorrect Description (62)"
        Case BTRV_ERR_INVALID_EXTDBUF   '63
            slErrMsg = slErrMsg & "Invalid Extended Insert Buffer (63)"
        Case BTRV_ERR_FILTER_LIMIT  '64
            slErrMsg = slErrMsg & "Filter Limit Reached (64)"
        Case BTRV_ERR_INCOR_FLD_OFFSET  '65
            slErrMsg = slErrMsg & "Incorrect Field Offset (65)"
        Case BTRV_ERR_AUTO_TRANS_ABORT  '74
            slErrMsg = slErrMsg & "Automatic Transaction Abort (74)"
        Case BTRV_ERR_DEADLOCK_DETECT   '78
            slErrMsg = slErrMsg & "Deadlock Detected (78)"
        Case BTRV_ERR_CONFLICT  '80
            slErrMsg = slErrMsg & "Conflict Error (80)"
        Case BTRV_ERR_LOCK  '81
            slErrMsg = slErrMsg & "Lock Error (81)"
        Case BTRV_ERR_LOST_POS  '82
            slErrMsg = slErrMsg & "Lost Position (82)"
        Case BTRV_ERR_READ_TRANS    '83
            slErrMsg = slErrMsg & "Read Outside Transaction (83)"
        Case BTRV_ERR_REC_LOCKED    '84
            slErrMsg = slErrMsg & "Record Locked (84)"
        Case BTRV_ERR_FILE_LOCKED   '85
            slErrMsg = slErrMsg & "File Locked (85)"
        Case BTRV_ERR_FILE_TBL_FULL '86
            slErrMsg = slErrMsg & "File Table Full (86)"
        Case BTRV_ERR_HNDL_TBL_FULL '87
            slErrMsg = slErrMsg & "Handle Table Full (87)"
        Case BTRV_ERR_INCOM_MODE    '88
            slErrMsg = slErrMsg & "Incompatible Mode Error (88)"
        Case BTRV_ERR_REDIR_DEV_FULL    '90
            slErrMsg = slErrMsg & "Redirection Device Table Full (90)"
        Case BTRV_ERR_SERVER    '91
            slErrMsg = slErrMsg & "Serve Error (91)"
        Case BTRV_ERR_TRANS_TBL_FULL    '92
            slErrMsg = slErrMsg & "Transaction Table Full (92)"
        Case BTRV_ERR_INCOM_LOCK    '93
            slErrMsg = slErrMsg & "Incompatible Lock Type (93)"
        Case BTRV_ERR_PERMISSION    '94
            slErrMsg = slErrMsg & "Permission Error (94)"
        Case BTRV_ERR_SESSION   '95
            slErrMsg = slErrMsg & "Session No Longer Valid (95)"
        Case BTRV_ERR_COMM_ENV  '96
            slErrMsg = slErrMsg & "Communications Environment Error (96)"
        Case BTRV_ERR_DATA_MSGSIZE  '97
            slErrMsg = slErrMsg & "Data Message too Small (97)"
        Case BTRV_ERR_INTERNAL_TRANS    '98
            slErrMsg = slErrMsg & "Internal Transaction Error (98)"
    End Select
    slErrMsg = slErrMsg & ": " & slMess
    MsgBox slErrMsg, vbOkOnly + vbCritical + vbApplicationModal, "Btrieve Error"
    bErrorFound = True
    mTerminate
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mCopyFields                     *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Copy fields from tmOldRec to    *
'*                     tgNewRec                        *
'*                                                     *
'*******************************************************
Private Sub mCopyFields()
'
'   slFieldSet(I)- Path plus name of file containing initial values ("FieldSet.New")
'   slName(I)- Table name as defined in file.ddf i.e. adf Advertiser
'
    Dim ilLoop As Integer
    Dim ilNewField As Integer
    Dim ilOldField As Integer
    Dim ilFound As Integer
    Dim ilRet As Integer
    Dim ilMatchField As Integer
    For ilNewField = 0 To UBound(tmNewDDFField) - 1 Step 1
        ilFound = False
        ilMatchField = False
        smFieldValue = ""
        For ilOldField = 0 To UBound(tmOldDDFField) - 1 Step 1
            If StrComp(Trim$(tmNewDDFField(ilNewField).sName), Trim$(tmOldDDFField(ilOldField).sName), 1) = 0 Then
                ilFound = True
                If (tmNewDDFField(ilNewField).sDataType = tmOldDDFField(ilOldField).sDataType) And (tmNewDDFField(ilNewField).iSize = tmOldDDFField(ilOldField).iSize) Then
                    If (Asc(tmNewDDFField(ilNewField).sDataType) = 5) Or (Asc(tmOldDDFField(ilOldField).sDataType) = 6) Then
                        If (Asc(tmNewDDFField(ilNewField).sDec) = Asc(tmOldDDFField(ilOldField).sDec)) Then
                            ilMatchField = True
                        Else
                            ilMatchField = False
                        End If
                    Else
                        ilMatchField = True
                    End If
                Else
                    ilMatchField = False
                End If
                For ilLoop = LBound(tmAdjustValues) To UBound(tmAdjustValues) - 1 Step 1
                    If StrComp(Trim$(tmNewDDFField(ilNewField).sName), Trim$(tmAdjustValues(ilLoop).sName), 1) = 0 Then
                        ilMatchField = False
                        Exit For
                    End If
                Next 'ilLoop
                If ilMatchField Then
                    mMoveField tmOldDDFField(ilOldField), tmNewDDFField(ilNewField)
                Else
                    mGetField tmOldDDFField(ilOldField)
                    For ilLoop = LBound(tmAdjustValues) To UBound(tmAdjustValues) - 1 Step 1
                        If StrComp(Trim$(tmNewDDFField(ilNewField).sName), Trim$(tmAdjustValues(ilLoop).sName), 1) = 0 Then
                            Select Case tmAdjustValues(ilLoop).iOper
                                Case 1  'Add
                                    smFieldValue = gAddStr(smFieldValue, tmAdjustValues(ilLoop).sValue)
                                Case 2  'Sub
                                    smFieldValue = gSubStr(smFieldValue, tmAdjustValues(ilLoop).sValue)
                                Case 3  'Mul
                                    smFieldValue = gMulStr(smFieldValue, tmAdjustValues(ilLoop).sValue)
                                Case 4  'Div
                                    smFieldValue = gDivStr(smFieldValue, tmAdjustValues(ilLoop).sValue)
                            End Select
                            Exit For
                        End If
                    Next 'ilLoop
                    mPutField tmNewDDFField(ilNewField)
                End If
                Exit For
            End If
        Next 'ilOldField
        If Not ilFound Then
            For ilLoop = LBound(tmInitValues) To UBound(tmInitValues) - 1 Step 1
                If StrComp(Trim$(tmNewDDFField(ilNewField).sName), Trim$(tmInitValues(ilLoop).sName), 1) = 0 Then
                    smFieldValue = Trim$(tmInitValues(ilLoop).sInitValue)
                    'Test if this is a field name
                    For ilOldField = 0 To UBound(tmOldDDFField) - 1 Step 1
                        If StrComp(Trim$(tmOldDDFField(ilOldField).sName), Trim$(smFieldValue), 1) = 0 Then
                            mGetField tmOldDDFField(ilOldField)
                            Select Case tmInitValues(ilLoop).iOper
                                Case 1  'Add
                                    smFieldValue = gAddStr(smFieldValue, tmInitValues(ilLoop).sValue)
                                Case 2  'Sub
                                    smFieldValue = gSubStr(smFieldValue, tmInitValues(ilLoop).sValue)
                                Case 3  'Mul
                                    smFieldValue = gMulStr(smFieldValue, tmInitValues(ilLoop).sValue)
                                Case 4  'Div
                                    smFieldValue = gDivStr(smFieldValue, tmInitValues(ilLoop).sValue)
                            End Select
                            Exit For
                        End If
                    Next ilOldField
                    Exit For
                End If
            Next 'ilLoop
            mPutField tmNewDDFField(ilNewField)
        End If
    Next 'ilNewField
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mFieldBuild                     *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list box with field    *
'*                     names                           *
'*                                                     *
'*******************************************************
Private Sub mFieldBuild(slDDFPath As String, slExt As String, ilFileID As Integer, tlDDFField() As DDFFIELD)
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilMax As Integer
    Dim ilOffset As Integer
    Dim ilUpper As Integer
    Dim ilRecLen As Integer
    ReDim tlDDFField(0 To 0) As DDFFIELD
    ReDim tlField(0 To 0) As DDFFIELD
    ilUpper = UBound(tlField)
    hmFile = CBtrvTable()    'CBtrvObj()
    ilRet = btrOpen(hmFile, "", slDDFPath & "Field" & slExt, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        bErrorFound = True
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        'mBtrvErrorMsg ilRet, "btrOpen:" & lbcDDF(0).List(lbcDDF(0).ListIndex)
        Exit Sub
    End If
    ilRecLen = Len(tlField(ilUpper)) 'Len(tmRec)
    lmActualRecLen = Len(tlField(ilUpper)) 'Len(tmRec)
    imBaseRecLen = lmActualRecLen
    tmFieldSrchKey1.iFileID = ilFileID
    ilRet = btrGetGreaterOrEqual(hmFile, tmOldRec, lmActualRecLen, tmFieldSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        bErrorFound = True
        mBtrvErrorMsg ilRet, "btrGetGreaterOrEqual: " & "FIELD.DDF"
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Exit Sub
    End If
    Do While (ilRet = BTRV_ERR_NONE)
        If imTerminate Then
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            mTerminate
            Exit Sub
        End If
        mGetFieldFields
        tlField(ilUpper) = tmGetDDFFIELD
        If (tlField(ilUpper).iFileID <> ilFileID) Then
            Exit Do
        End If
        If Asc(tlField(ilUpper).sDataType) <= 17 Then
            ilUpper = ilUpper + 1
            ReDim Preserve tlField(0 To ilUpper) As DDFFIELD
        End If
        lmActualRecLen = imBaseRecLen
        ilRet = btrGetNext(hmFile, tmOldRec, lmActualRecLen, BTRV_LOCK_NONE)
    Loop

    ilRet = btrClose(hmFile)
    btrDestroy hmFile
    
    'Sort by offset
    ilMax = 0
    For ilLoop = 0 To UBound(tlField) - 1 Step 1
        If tlField(ilLoop).iOffset > ilMax Then
            ilMax = tlField(ilLoop).iOffset
        End If
    Next ilLoop
    ilUpper = UBound(tlDDFField)
    ilOffset = 0
    Do
        For ilLoop = 0 To UBound(tlField) - 1 Step 1
            If tlField(ilLoop).iOffset = ilOffset Then
                tlDDFField(ilUpper) = tlField(ilLoop)
                ilUpper = ilUpper + 1
                ReDim Preserve tlDDFField(0 To ilUpper) As DDFFIELD
                Exit For
            End If
        Next ilLoop
        ilOffset = ilOffset + 1
    Loop While ilOffset <= ilMax
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mFileBuild                      *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build Table of file names       *
'*                                                     *
'*******************************************************
Private Sub mFileBuild(slDDFPath As String, slExt As String, tlDDFFile() As DDFFILE)
    '   slDDFPath(I)- DDF Path
    '   slExt(I)- Extension (.ddf or .new)
    '   tlDDFFile(O)- File names
    '
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim ilRecLen As Integer
    ReDim tlDDFFile(0 To 0) As DDFFILE
    ilUpper = UBound(tlDDFFile)
    hmFile = CBtrvTable()    'CBtrvObj()
    ilRet = btrOpen(hmFile, "", slDDFPath & "File" & slExt, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        'bErrorFound = True
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Exit Sub
    End If
    ilRecLen = Len(tlDDFFile(ilUpper)) 'Len(tmRec)
    lmActualRecLen = Len(tlDDFFile(ilUpper)) 'Len(tmRec)
    imBaseRecLen = lmActualRecLen
    ilRet = btrGetFirst(hmFile, tmOldRec, lmActualRecLen, INDEXKEY0, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        bErrorFound = True
        mBtrvErrorMsg ilRet, "btrGetFirst: " & "FILE.DDF"
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Exit Sub
    End If
    Do While ilRet = BTRV_ERR_NONE
        If imTerminate Then
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            mTerminate
            Exit Sub
        End If
        mGetFileFields
        tlDDFFile(ilUpper) = tmGetDDFFile
        If (InStr(1, Trim$(tlDDFFile(ilUpper).sName), "X$", 1) <> 1) And (StrComp(Left$(tlDDFFile(ilUpper).sName, 3), "TVI", vbTextCompare) <> 0) Then
            ilUpper = ilUpper + 1
            ReDim Preserve tlDDFFile(0 To ilUpper) As DDFFILE
        End If
        lmActualRecLen = imBaseRecLen
        ilRet = btrGetNext(hmFile, tmOldRec, lmActualRecLen, BTRV_LOCK_NONE)
    Loop

    ilRet = btrClose(hmFile)
    btrDestroy hmFile
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mFilePop                        *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list box with file     *
'*                     names                           *
'*                                                     *
'*******************************************************
Private Sub mFilePop(tlDDFFile() As DDFFILE, lbcFile As Control)
    Dim ilLoop As Integer
    Dim slStr As String
    Dim ilPos As Integer
    lbcFile.Clear
    For ilLoop = 0 To UBound(tlDDFFile) - 1 Step 1
        If imSelectBTR And imSelectMKD And imSelectENG Then
            lbcFile.AddItem Trim$(tlDDFFile(ilLoop).sName)
        ElseIf imSelectBTR Then
            slStr = Trim$(gRemoveZero(tlDDFFile(ilLoop).sLocation))
            ilPos = InStr(slStr, ".")
            slStr = Mid$(slStr, ilPos + 1)
            If StrComp(slStr, "btr", 1) = 0 Then
                lbcFile.AddItem Trim$(tlDDFFile(ilLoop).sName)
            End If
        ElseIf imSelectMKD Then
            slStr = Trim$(gRemoveZero(tlDDFFile(ilLoop).sLocation))
            ilPos = InStr(slStr, ".")
            slStr = Mid$(slStr, ilPos + 1)
            If StrComp(slStr, "Mkd", 1) = 0 Then
                lbcFile.AddItem Trim$(tlDDFFile(ilLoop).sName)
            End If
        ElseIf imSelectENG Then
            slStr = Trim$(gRemoveZero(tlDDFFile(ilLoop).sLocation))
            ilPos = InStr(slStr, ".")
            slStr = Mid$(slStr, ilPos + 1)
            If StrComp(slStr, "Eng", 1) = 0 Then
                lbcFile.AddItem Trim$(tlDDFFile(ilLoop).sName)
            End If
        End If
    Next ilLoop
    mSetCommands
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetField                       *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get field from record           *
'*                                                     *
'*******************************************************
Private Sub mGetField(tlDDFField As DDFFIELD)
'Taken from ShowFix vb6
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilSize As Integer
    Dim ilValue As Integer
    Dim llValue As Long
    Dim ilDyMn As Integer
    Dim ilYear As Integer
    Dim ilHsSec As Integer
    Dim ilMinHr As Integer
    Dim slMoney As String
    Dim slHex As String
    ilIndex = tlDDFField.iOffset + 1
    ilSize = tlDDFField.iSize
    smFieldValue = "Error"
    Select Case Asc(tlDDFField.sDataType)
        Case 0  'String
            smFieldValue = ""
            For ilLoop = 1 To ilSize Step 1
                If (tmOldRec.sChar(ilIndex)) = 0 Then
                    smFieldValue = smFieldValue & " "
                Else
                    smFieldValue = smFieldValue & Chr(tmOldRec.sChar(ilIndex))
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 1  'Integer
            If ilSize = 1 Then
                ilValue = tmOldRec.sChar(ilIndex) 'Asc(tmOldRec.sChar(ilIndex))
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 2 Then
                HMemCpy ilValue, tmOldRec.sChar(ilIndex), 2
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 4 Then
                HMemCpy llValue, tmOldRec.sChar(ilIndex), 4
                smFieldValue = Trim$(Str$(llValue))
            End If
        Case 2  'IEEE Float
        Case 3  'Btrieve Date
            ilDyMn = 256 * (tmOldRec.sChar(ilIndex + 1)) + (tmOldRec.sChar(ilIndex))
            ilYear = 256 * (tmOldRec.sChar(ilIndex + 3)) + (tmOldRec.sChar(ilIndex + 2))
            gUnpackDate ilDyMn, ilYear, smFieldValue
        Case 4  'Btrieve Time
            ilHsSec = 256 * (tmOldRec.sChar(ilIndex + 1)) + (tmOldRec.sChar(ilIndex))
            ilMinHr = 256 * (tmOldRec.sChar(ilIndex + 3)) + (tmOldRec.sChar(ilIndex + 2))
            gUnpackTime ilHsSec, ilMinHr, "A", "1", smFieldValue
        Case 5  'COBOL Decimal COMP-3
            slMoney = ""
            For ilLoop = 1 To ilSize Step 1
                slMoney = slMoney & Chr(tmOldRec.sChar(ilIndex))
                ilIndex = ilIndex + 1
            Next ilLoop
            gPDNToStr slMoney, Asc(tlDDFField.sDec), smFieldValue
        Case 6  'COBOL Money
            slMoney = ""
            For ilLoop = 1 To ilSize Step 1
                slMoney = slMoney & Chr(tmOldRec.sChar(ilIndex))
                ilIndex = ilIndex + 1
            Next ilLoop
            gPDNToStr slMoney, Asc(tlDDFField.sDec), smFieldValue
        Case 7  'Logical
            If ilSize = 1 Then
                ilValue = (tmOldRec.sChar(ilIndex))
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 2 Then
                HMemCpy ilValue, tmOldRec.sChar(ilIndex), 2
                smFieldValue = Trim$(Str$(ilValue))
            End If
        Case 8  'COBOL Numeric
        Case 9  'BASIC Float
        Case 10 'Pascal LString
            ilValue = tmOldRec.sChar(ilIndex)
            smFieldValue = ""
            ilIndex = ilIndex + 1
            For ilLoop = 1 To ilValue Step 1
                If (tmOldRec.sChar(ilIndex)) = 0 Then
                    smFieldValue = smFieldValue & " "
                Else
                    smFieldValue = smFieldValue & Chr(tmOldRec.sChar(ilIndex))
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 11 'C ZString
            smFieldValue = ""
            Do While (tmOldRec.sChar(ilIndex)) <> 0
                smFieldValue = smFieldValue & Chr(tmOldRec.sChar(ilIndex))
                ilIndex = ilIndex + 1
            Loop
        Case 12 'Variable Length Note
        Case 13 'LVar (length, Characters, 0, 0); Length does not include zeros
            HMemCpy ilValue, tmOldRec.sChar(ilIndex), 2
            smFieldValue = ""
            ilIndex = ilIndex + 2
            For ilLoop = 1 To ilValue Step 1
                If (tmOldRec.sChar(ilIndex)) = 0 Then
                    smFieldValue = smFieldValue & " "
                Else
                    smFieldValue = smFieldValue & Chr(tmOldRec.sChar(ilIndex))
                End If
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 14 'Unsigned Binary
            HMemCpy ilValue, tmOldRec.sChar(ilIndex), 2
            llValue = gUnsignIntToLong(ilValue)
            smFieldValue = Trim$(Str$(llValue))
        Case 15 'AutoIncrement
            If ilSize = 2 Then
                HMemCpy ilValue, tmOldRec.sChar(ilIndex), 2
                smFieldValue = Trim$(Str$(ilValue))
            ElseIf ilSize = 4 Then
                HMemCpy llValue, tmOldRec.sChar(ilIndex), 4
                smFieldValue = Trim$(Str$(llValue))
            Else
                smFieldValue = "Error"
            End If
        Case 16 'Bit
        Case 17 'COBOL Numeric STS
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mIndexBuild                     *
'*                                                     *
'*             Created:7/04/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build Index table               *
'*                                                     *
'*******************************************************
Private Sub mIndexBuild(slDDFPath As String, slExt As String, ilFileID As Integer, tlDDFIndex() As DDFINDEX)
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim ilRecLen As Integer
    Dim ilLoop As Integer
    Dim ilMaxIndex As Integer
    Dim ilMaxSeg As Integer
    Dim ilIndex As Integer
    Dim ilSeg As Integer
    ReDim tlIndex(0 To 0) As DDFINDEX
    ReDim tlDDFIndex(0 To 0) As DDFINDEX
    ilUpper = UBound(tlIndex)
    hmFile = CBtrvTable()    'CBtrvObj()
    ilRet = btrOpen(hmFile, "", slDDFPath & "Index" & slExt, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        bErrorFound = True
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Exit Sub
    End If
    ilRecLen = Len(tlIndex(ilUpper)) 'Len(tmRec)
    lmActualRecLen = Len(tlIndex(ilUpper)) 'Len(tmRec)
    imBaseRecLen = lmActualRecLen
    tmIndexSrchKey0.iFileID = ilFileID
    ilRet = btrGetGreaterOrEqual(hmFile, tmOldRec, lmActualRecLen, tmIndexSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        bErrorFound = True
        mBtrvErrorMsg ilRet, "btrGetGreaterOrEqual: " & "Index.DDF"
        ilRet = btrClose(hmFile)
        btrDestroy hmFile
        Exit Sub
    End If
    Do While (ilRet = BTRV_ERR_NONE)
        If imTerminate Then
            ilRet = btrClose(hmFile)
            btrDestroy hmFile
            mTerminate
            Exit Sub
        End If
        mGetIndexFields
        tlIndex(ilUpper) = tmGetDDFIndex
        If (tlIndex(ilUpper).iFileID <> ilFileID) Then
            Exit Do
        End If
        ilUpper = ilUpper + 1
        ReDim Preserve tlIndex(0 To ilUpper) As DDFINDEX
        lmActualRecLen = imBaseRecLen
        ilRet = btrGetNext(hmFile, tmOldRec, lmActualRecLen, BTRV_LOCK_NONE)
    Loop
    ilRet = btrClose(hmFile)
    btrDestroy hmFile
    'Sort index into index, segment order
    ReDim tlDDFIndex(0 To ilUpper) As DDFINDEX
    ilMaxIndex = 0
    For ilLoop = 0 To UBound(tlIndex) - 1 Step 1
        If tlIndex(ilLoop).iNumber > ilMaxIndex Then
            ilMaxIndex = tlIndex(ilLoop).iNumber
        End If
    Next ilLoop
    ilUpper = LBound(tlDDFIndex)
    For ilIndex = 0 To ilMaxIndex Step 1
        ilMaxSeg = 0
        For ilLoop = 0 To UBound(tlIndex) - 1 Step 1
            If (tlIndex(ilLoop).iNumber = ilIndex) And (tlIndex(ilLoop).iPart > ilMaxSeg) Then
                ilMaxSeg = tlIndex(ilLoop).iPart
            End If
        Next ilLoop
        For ilSeg = 0 To ilMaxSeg Step 1
            For ilLoop = 0 To UBound(tlIndex) - 1 Step 1
                If (tlIndex(ilLoop).iNumber = ilIndex) And (tlIndex(ilLoop).iPart = ilSeg) Then
                    tlDDFIndex(ilUpper) = tlIndex(ilLoop)
                    ilUpper = ilUpper + 1
                    Exit For
                End If
            Next ilLoop
        Next ilSeg
    Next ilIndex
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
    Dim llNoRec As Long 'Number of records
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim slDrive As String
    Dim slFromFile As String
    Dim slCommand As String
    Dim ilSpace As Integer
    Dim slStr As String
    Dim slPath As String
    ReDim tmKeyFile(0 To 0) As KEYFILE
    ReDim tmOldDDFLinkField(0 To 0) As DDFFIELD
    ReDim tmOldDDFField(0 To 0) As DDFFIELD
    imPageSize(0) = 512
    imPageSize(1) = 1024
    imPageSize(2) = 2048
    imPageSize(3) = 4096
    imExit = False
    imRun = False
    imRU = False
    imSelectBTR = True
    imSelectMKD = True
    imSelectENG = True
    smNewDDFPath = ""
    smOldDDFPath = ""
    smDBPath = ""
    lncNoFiles.Visible = False
    lncGauge.Visible = False
    If Command$ <> "Debug" Then
        slCommand = Command$
        ilPos = InStr(1, slCommand, "/C:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, slCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(slCommand, ilPos + 3))
                If ilPos - 1 > 0 Then
                    slCommand = Trim$(Left$(slCommand, ilPos - 1))
                Else
                    slCommand = ""
                End If
            Else
                slStr = Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3)
                If ilPos - 1 > 0 Then
                    slCommand = Left$(slCommand, ilPos - 1) & Mid$(slCommand, ilSpace)
                Else
                    slCommand = Mid$(slCommand, ilSpace)
                End If
            End If
            smCmmdFileName = slStr   'Selection File
        End If
        ilPos = InStr(1, slCommand, "/O:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, slCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(slCommand, ilPos + 3))
                If ilPos - 1 > 0 Then
                    slCommand = Trim$(Left$(slCommand, ilPos - 1))
                Else
                    slCommand = ""
                End If
            Else
                slStr = Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3)
                If ilPos - 1 > 0 Then
                    slCommand = Left$(slCommand, ilPos - 1) & Mid$(slCommand, ilSpace)
                Else
                    slCommand = Mid$(slCommand, ilSpace)
                End If
            End If
            smOldDDFPath = slStr   'Selection File
        End If
        ilPos = InStr(1, slCommand, "/N:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, slCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(slCommand, ilPos + 3))
                If ilPos - 1 > 0 Then
                    slCommand = Trim$(Left$(slCommand, ilPos - 1))
                Else
                    slCommand = ""
                End If
            Else
                slStr = Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3)
                If ilPos - 1 > 0 Then
                    slCommand = Left$(slCommand, ilPos - 1) & Mid$(slCommand, ilSpace)
                Else
                    slCommand = Mid$(slCommand, ilSpace)
                End If
            End If
            smNewDDFPath = slStr   'Selection File
        End If
        ilPos = InStr(1, slCommand, "/F:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, slCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(slCommand, ilPos + 3))
                If ilPos - 1 > 0 Then
                    slCommand = Trim$(Left$(slCommand, ilPos - 1))
                Else
                    slCommand = ""
                End If
            Else
                slStr = Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3)
                If ilPos - 1 > 0 Then
                    slCommand = Left$(slCommand, ilPos - 1) & Mid$(slCommand, ilSpace)
                Else
                    slCommand = Mid$(slCommand, ilSpace)
                End If
            End If
            smDBPath = slStr   'Selection File
        End If
        If (smOldDDFPath <> "") And (smDBPath = "") Then
            smDBPath = smOldDDFPath
        ElseIf (smDBPath <> "") And (smOldDDFPath = "") Then
            smOldDDFPath = smDBPath
        End If
        If (smOldDDFPath <> "") And (smNewDDFPath = "") Then
            If StrComp(CurDir$, smOldDDFPath, 1) <> 0 Then
                smNewDDFPath = CurDir$
            End If
        End If
        ilPos = InStr(1, slCommand, "/Exit", 1)
        If ilPos > 0 Then
            imExit = True
        End If
        ilPos = InStr(1, slCommand, "/Run", 1)
        If (ilPos > 0) And (smOldDDFPath <> "") And (smNewDDFPath <> "") And (smDBPath <> "") Then
            imRun = True
        End If
        ilPos = InStr(1, slCommand, "/RemoteUser", 1)
        If ilPos > 0 Then
            imRU = True
        End If
        ilPos = InStr(1, slCommand, "/No_BTR", 1)
        If ilPos > 0 Then
            imSelectBTR = False
        End If
        ilPos = InStr(1, slCommand, "/No_MKD", 1)
        If ilPos > 0 Then
            imSelectMKD = False
        End If
        ilPos = InStr(1, slCommand, "/No_ENG", 1)
        If ilPos > 0 Then
            imSelectENG = False
        End If
        ilPos = InStr(1, slCommand, "/CreateOnly:", 1)
        If ilPos > 0 Then
            slStr = ""
            ilSpace = InStr(ilPos, slCommand, " ")
            If ilSpace = 0 Then
                slStr = Trim$(Mid$(slCommand, ilPos + 12))
                If ilPos - 1 > 0 Then
                    slCommand = Trim$(Left$(slCommand, ilPos - 1))
                Else
                    slCommand = ""
                End If
            Else
                slStr = Mid$(slCommand, ilPos + 12, ilSpace - ilPos - 12)
                If ilPos - 1 > 0 Then
                    slCommand = Left$(slCommand, ilPos - 1) & Mid$(slCommand, ilSpace)
                Else
                    slCommand = Mid$(slCommand, ilSpace)
                End If
            End If
        Else
            slStr = ""
        End If
        smCreateOnlyArray = Split(slStr, "|")
    Else
        smCmmdFileName = ""   'Selection File
        smCreateOnlyArray = Split("", "|")
    End If
    hmDB = CBtrvMngrInit(0) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    mInitVar
    mInitBox
    sgDBPath = smDBPath
    
    mnuFilePrint.Enabled = False
    ilPos = InStr(smOldDDFPath, ":")
    If ilPos > 0 Then
        slDrive = Left(smOldDDFPath, ilPos)
        slPath = Mid$(smOldDDFPath, ilPos + 1)
        If Right$(slPath, 1) = "\" Then
            slPath = Left$(slPath, Len(slPath) - 1)
        Else
            smOldDDFPath = smOldDDFPath & "\"
        End If
        On Error Resume Next
        cbcOldDrive.Drive = slDrive
        slStr = lbcOldPath.Path
        ilPos = InStr(slStr, ":")
        If ilPos > 0 Then
            slStr = Mid$(slStr, ilPos + 1)
        End If
        If StrComp(slStr, slPath, 1) = 0 Then
            lbcOldPath_Change
        Else
            lbcOldPath.Path = slPath
        End If
        On Error GoTo 0
    End If
    ilPos = InStr(smNewDDFPath, ":")
    If ilPos > 0 Then
        slDrive = Left(smNewDDFPath, ilPos)
        slPath = Mid$(smNewDDFPath, ilPos + 1)
        If Right$(slPath, 1) = "\" Then
            slPath = Left$(slPath, Len(slPath) - 1)
        Else
            smNewDDFPath = smNewDDFPath & "\"
        End If
        On Error Resume Next
        cbcNewDrive.Drive = slDrive
        slStr = lbcNewPath.Path
        ilPos = InStr(slStr, ":")
        If ilPos > 0 Then
            slStr = Mid$(slStr, ilPos + 1)
        End If
        If StrComp(slStr, slPath, 1) = 0 Then
            lbcNewPath_Change
        Else
            lbcNewPath.Path = slPath
        End If
        On Error GoTo 0
    End If
    ilPos = InStr(smDBPath, ":")
    If ilPos > 0 Then
        slDrive = Left(smDBPath, ilPos)
        slPath = Mid$(smDBPath, ilPos + 1)
        If Right$(slPath, 1) = "\" Then
            slPath = Left$(slPath, Len(slPath) - 1)
        Else
            smDBPath = smDBPath & "\"
        End If
        On Error Resume Next
        cbcDBDrive.Drive = slDrive
        slStr = lbcDBPath.Path
        ilPos = InStr(slStr, ":")
        If ilPos > 0 Then
            slStr = Mid$(slStr, ilPos + 1)
        End If
        If StrComp(slStr, slPath, 1) = 0 Then
            lbcDBPath_Change
        Else
            lbcDBPath.Path = slPath
        End If
        On Error GoTo 0
    End If
    DDFReorg.Height = 6915
    gCenterStdAlone DDFReorg
    imReadingFile = False
    imTerminate = False
    Exit Sub
mInitErr:
    ilRet = Err
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
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
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mInitValue                      *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain initial values for fields*
'*                     not defined                     *
'*                                                     *
'*******************************************************
Private Sub mInitValues(slFieldSet As String, slName As String)
'
'   slFieldSet(I)- Path plus name of file containing initial values and rename fields ("FieldSet.New")
'   slName(I)- Table name as defined in file.ddf i.e. adf Advertiser
'
    Dim ilLoop As Integer
    Dim ilNewField As Integer
    Dim ilOldField As Integer
    Dim ilFound As Integer
    Dim ilNewDataType As Integer
    Dim slReturn As String * 130
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim slLine As String
    Dim ilCompared As Integer
    Dim slStr As String
    Dim ilPos As Integer
    ReDim tmInitValues(0 To 0) As INITVALUES
    ReDim tmAdjustValues(0 To 0) As ADJUSTVALUES
    slLine = "    Rename Fields:"
    For ilOldField = 0 To UBound(tmOldDDFField) - 1 Step 1
        On Error Resume Next
        slStr = "Rename:" & Trim$(tmOldDDFField(ilOldField).sName)
        ilRet = GetPrivateProfileString(slName, slStr, "Not Found", slReturn, 128, slFieldSet)
        If Left$(slReturn, ilRet) <> "Not Found" Then
            slLine = slLine & " " & Trim$(tmOldDDFField(ilOldField).sName)
            tmOldDDFField(ilOldField).sName = Left$(slReturn, ilRet)
            slLine = slLine & " To " & Trim$(tmOldDDFField(ilOldField).sName)
            Print #hmTo, slLine
            slLine = "               "
        End If
    Next ilOldField
    ilUpper = UBound(tmInitValues)
    slLine = "    New Fields:"
    For ilNewField = 0 To UBound(tmNewDDFField) - 1 Step 1
        ilFound = False
        For ilOldField = 0 To UBound(tmOldDDFField) - 1 Step 1
            If StrComp(Trim$(tmNewDDFField(ilNewField).sName), Trim$(tmOldDDFField(ilOldField).sName), 1) = 0 Then
                ilFound = True
                Exit For
            End If
        Next ilOldField
        If Not ilFound Then
            slLine = slLine & " " & Trim$(tmNewDDFField(ilNewField).sName)
            On Error Resume Next
            slStr = "Init:" & Trim$(tmNewDDFField(ilNewField).sName)
            ilRet = GetPrivateProfileString(slName, slStr, "Not Found", slReturn, 128, slFieldSet)
            If Left$(slReturn, ilRet) <> "Not Found" Then
                tmInitValues(ilUpper).iOper = 0
                tmInitValues(ilUpper).sValue = ""
                slStr = Left$(slReturn, ilRet)
                ilPos = InStr(slStr, "+")
                If ilPos > 0 Then
                    tmInitValues(ilUpper).iOper = 1
                Else
                    ilPos = InStr(slStr, "-")
                    If ilPos > 0 Then
                        tmInitValues(ilUpper).iOper = 2
                    Else
                        ilPos = InStr(slStr, "*")
                        If ilPos > 0 Then
                            tmInitValues(ilUpper).iOper = 3
                        Else
                            ilPos = InStr(slStr, "/")
                            If ilPos > 0 Then
                                tmInitValues(ilUpper).iOper = 4
                            End If
                        End If
                    End If
                End If
                tmInitValues(ilUpper).sName = Trim$(tmNewDDFField(ilNewField).sName)
                If tmInitValues(ilUpper).iOper <= 0 Then
                    tmInitValues(ilUpper).sInitValue = slStr
                    tmInitValues(ilUpper).sValue = ""
                Else
                    tmInitValues(ilUpper).sInitValue = Left$(slStr, ilPos - 1)  'Left$(slReturn, ilRet)
                    tmInitValues(ilUpper).sValue = Mid$(slStr, ilPos + 1)
                End If
                slLine = slLine & "=" & Trim$(tmInitValues(ilUpper).sInitValue)
                ilUpper = ilUpper + 1
                ReDim Preserve tmInitValues(0 To ilUpper) As INITVALUES
            End If
            If Len(slLine) > 50 Then
                Print #hmTo, slLine
                slLine = "               "
            End If
            On Error GoTo 0
        End If
    Next ilNewField
    If Len(slLine) > 15 Then
        Print #hmTo, slLine
    End If
    ilUpper = UBound(tmAdjustValues)
    slLine = "    Adjust Fields:"
    For ilNewField = 0 To UBound(tmNewDDFField) - 1 Step 1
        On Error Resume Next
        slStr = "Adjust:" & Trim$(tmNewDDFField(ilNewField).sName)
        ilRet = GetPrivateProfileString(slName, slStr, "Not Found", slReturn, 128, slFieldSet)
        If Left$(slReturn, ilRet) <> "Not Found" Then
            tmAdjustValues(ilUpper).sName = Trim$(tmNewDDFField(ilNewField).sName)
            slStr = Left$(slReturn, ilRet)
            slLine = slLine & " " & Trim$(tmNewDDFField(ilNewField).sName)
            slLine = slLine & "=" & Trim$(slStr)
            ilPos = InStr(slStr, "+")
            If ilPos > 0 Then
                tmAdjustValues(ilUpper).iOper = 1
            Else
                ilPos = InStr(slStr, "-")
                If ilPos > 0 Then
                    tmAdjustValues(ilUpper).iOper = 2
                Else
                    ilPos = InStr(slStr, "*")
                    If ilPos > 0 Then
                        tmAdjustValues(ilUpper).iOper = 3
                    Else
                        ilPos = InStr(slStr, "/")
                        If ilPos > 0 Then
                            tmAdjustValues(ilUpper).iOper = 4
                        End If
                    End If
                End If
            End If
            If ilPos > 0 Then
                slStr = Mid$(slStr, ilPos + 1)
                tmAdjustValues(ilUpper).sValue = slStr
                ilUpper = ilUpper + 1
                ReDim Preserve tmAdjustValues(0 To ilUpper) As ADJUSTVALUES
            End If
        End If
        If Len(slLine) > 50 Then
            Print #hmTo, slLine
            slLine = "                  "
        End If
        On Error GoTo 0
    Next ilNewField
    If Len(slLine) > 18 Then
        Print #hmTo, slLine
    End If
    slLine = "    Changed Fields:"
    For ilNewField = 0 To UBound(tmNewDDFField) - 1 Step 1
        For ilOldField = 0 To UBound(tmOldDDFField) - 1 Step 1
            If StrComp(Trim$(tmNewDDFField(ilNewField).sName), Trim$(tmOldDDFField(ilOldField).sName), 1) = 0 Then
                ilCompared = True
                If StrComp(Trim$(tmOldDDFField(ilOldField).sDataType), Trim$(tmNewDDFField(ilNewField).sDataType), 1) <> 0 Then
                    ilCompared = False
                End If
                If tmOldDDFField(ilOldField).iSize <> tmNewDDFField(ilNewField).iSize Then
                    ilCompared = False
                End If
                If StrComp(Trim$(tmOldDDFField(ilOldField).sDec), Trim$(tmNewDDFField(ilNewField).sDec), 1) <> 0 Then
                    ilCompared = False
                End If
                If tmOldDDFField(ilOldField).iFlags <> tmNewDDFField(ilNewField).iFlags Then
                    ilCompared = False
                End If
                If Not ilCompared Then
                    slLine = slLine & " " & Trim$(tmNewDDFField(ilNewField).sName)
                    If Len(slLine) > 60 Then
                        Print #hmTo, slLine
                        slLine = "                   "
                    End If
                End If
                Exit For
            End If
        Next ilOldField
    Next ilNewField
    If Len(slLine) > 19 Then
        Print #hmTo, slLine
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gInitGlobalVar                  *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Initialize global variables     *
'*                                                     *
'*******************************************************
Private Sub mInitVar()
'
'   gInitGlobalVar
'
    'smDBPath = sgDBPath
    'smOldDDFPath = sgDDFPath
    lmTotalNoRec = 0
    lmProcessedNoRec = 0
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveField                      *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Move field into record          *
'*                                                     *
'*******************************************************
Private Sub mMoveField(tlOldDDFField As DDFFIELD, tlNewDDFField As DDFFIELD)
    Dim ilLoop As Integer
    Dim ilOldIndex As Integer
    Dim ilNewIndex As Integer
    Dim ilSize As Integer
    Dim ilValue As Integer
    Dim slInt2 As INT2
    ilOldIndex = tlOldDDFField.iOffset + 1
    ilNewIndex = tlNewDDFField.iOffset + 1
    ilSize = tlOldDDFField.iSize        'Old and new size match
    Select Case Asc(tlOldDDFField.sDataType)    'Old and new match
        Case 0  'String
            For ilLoop = 1 To ilSize Step 1
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
            Next 'ilLoop
        Case 1  'Integer
            For ilLoop = 1 To ilSize Step 1
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
            Next 'ilLoop
        Case 2  'IEEE Float
        Case 3  'Btrieve Date
            For ilLoop = 1 To 4 Step 1
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
            Next 'ilLoop
        Case 4  'Btrieve Time
            For ilLoop = 1 To 4 Step 1
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
            Next 'ilLoop
        Case 5  'COBOL Decimal COMP-3
            For ilLoop = 1 To ilSize Step 1
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
            Next 'ilLoop
        Case 6  'COBOL Money
            For ilLoop = 1 To ilSize Step 1
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
            Next 'ilLoop
        Case 7  'Logical
            For ilLoop = 1 To ilSize Step 1
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
            Next 'ilLoop
        Case 8  'COBOL Numeric
        Case 9  'BASIC Float
        Case 10 'Pascal LString
            ilValue = Asc(tmOldRec.sChar(ilOldIndex)) + 1
            For ilLoop = 1 To ilValue Step 1
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
            Next 'ilLoop
        Case 11 'C ZString
            Do While Asc(tmOldRec.sChar(ilOldIndex)) <> 48
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
                If ilOldIndex = ilSize Then
                    ilNewIndex = ilNewIndex - 1
                    Do While Asc(tgNewRec.sChar(ilNewIndex)) = 32
                        ilNewIndex = ilNewIndex - 1
                    Loop
                    tgNewRec.sChar(ilNewIndex + 1) = Chr(0)
                    ilOldIndex = ilNewIndex
                    Exit Do
                End If
            Loop
            tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
        Case 12 'Variable Length Note
            ilValue = lmOldActualRecLen - lmOldBaseRecLen
            imVarLenNote = ilValue
            For ilLoop = 1 To ilValue Step 1
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
            Next 'ilLoop
        Case 13 'LVar (length, Characters, 0, 0); Length does not include zeros
            ''slInt2.sInt2 = tmOldRec.sChar(ilOldIndex) & tmOldRec.sChar(ilOldIndex + 1)
            ''HMemCpy ilValue, slInt2, 2
            'ilValue = 256 * (tmOldRec.sChar(ilOldIndex + 1)) + (tmOldRec.sChar(ilOldIndex))
            HMemCpy ilValue, tmOldRec.sChar(ilOldIndex), 2
            ilValue = ilValue + 2
            For ilLoop = 1 To ilValue Step 1
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
            Next 'ilLoop
        Case 14 'Unsigned Binary
            For ilLoop = 1 To ilSize Step 1
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
            Next 'ilLoop
        Case 15 'AutoIncrement
            For ilLoop = 1 To ilSize Step 1
                tgNewRec.sChar(ilNewIndex) = tmOldRec.sChar(ilOldIndex)
                ilOldIndex = ilOldIndex + 1
                ilNewIndex = ilNewIndex + 1
            Next 'ilLoop
        Case 16 'Bit
        Case 17 'COBOL Numeric STS
    End Select
End Sub

Private Sub mnuFileExit_Click()
    mTerminate
End Sub

Private Sub mnuFileOpen_Click()
    cmcDone.Enabled = False
    cmcProcess.Enabled = False
    'plcFOpen.Caption = "Open File"
    'plcFOpen.Visible = True
    'plcFOpen.ZOrder vbBringToFront
End Sub

Private Sub mnuFilePSetup_Click()
'    cdcSetup.Flags = cdlPDPrintSetup
'    cdcSetup.Action = 5    'DLG_PRINT
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mProcessRecords                 *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain records                  *
'*                                                     *
'*******************************************************
Private Sub mProcessRecords()
    Dim ilRet As Integer
    Dim ilNewLoop As Integer
    Dim ilOldLoop As Integer
    Dim slName As String
    Dim slNewFile As String
    Dim slOldFile As String
    Dim slExtName As String
    Dim ilPos As Integer
    Dim ilFile As Integer
    Dim llRecNo As Long
    Dim ilField As Integer
    Dim llNoProcessed As Long
    'Dim ilExtLen As Integer
    Dim llExtLen As Long
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilIndex As Integer
    Dim slInt2 As INT2
    Dim ilValue As Integer
    Dim ilFileFlags As Integer
    Dim ilNoKeys As Integer
    Dim ilNewField As Integer
    Dim ilLoop As Integer
    Dim slLine As String
    Dim slFieldSet As String
    Dim slReturn As String * 130
    Dim slReturnAlt As String * 258
    Dim ilPageSize As Integer
    Dim ilPages As Integer
    Dim llNoFiles As Long
    Dim llNoFilesProc As Long
    Dim ilCRet As Integer
    Dim ilAltCollateReq As Integer
    Dim slAltCollateFile As String  'Name of alternate collate file
    Dim slAltCollateSeq As String  'Name of alternate collate Sequence
    Dim slSeqName As String
    Dim slStamp As String
    Dim llPercent As Long
    Dim fs As New FileSystemObject
    Dim slFileName As String
    Dim ilExtNumber As Integer
    Dim slExtension As String
    Dim blCopyRecords As Boolean
    
    Screen.MousePointer = vbHourglass
    On Error GoTo gGetDateTimeErr:
    lncNoFiles.X2 = lncNoFiles.X1
    lncGauge.X2 = lncGauge.X1
    lacRename.Caption = ""

    smDBPath = lbcDBPath.Path & "\"
    sgDBPath = smDBPath
    smCreateDBPath = smDBPath '& "DDFReorg\"
    ilRet = 0
    slFileName = smDBPath & "Messages"
    If Not fs.FolderExists(slFileName) Then
        'MsgBox ("Messages = " & sgDBPath & "Messages\" & " path is invalid.  Counterpoint will attempt to create the directory")
        fs.CreateFolder (slFileName)
        If Not fs.FolderExists(slFileName) Then
            MsgBox "Counterpoint was unable to create the folder: " & slFileName & " Please have your IT manager add this folder.", vbOkOnly + vbApplicationModal, "DDFReorg"
            Exit Sub
        End If
    End If
    ilRet = 0
    hmTo = FreeFile
    Open smDBPath & "messages\" & "DDFReorg.Txt" For Output As hmTo
    If ilRet <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Open " & "Unable to create DDFReorg.Txt" & " in " & smDBPath & "messages\, Error " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
        bErrorFound = True
'        plcGauge.Visible = False
'        plcNoFiles.Visible = False
        lncNoFiles.Visible = False
        lncGauge.Visible = False
        Exit Sub
    End If
    llNoFilesProc = 0
    llNoFiles = 0
    For ilNewLoop = 0 To lbcNewFile.ListCount - 1 Step 1
        If lbcNewFile.Selected(ilNewLoop) Then
            llNoFiles = llNoFiles + 1
        End If
    Next ilNewLoop
    slLine = Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmTo, "DDFReorg: Creating files on " & slLine
    Print #hmTo, "  Database Path " & smDBPath
    Print #hmTo, "  Old DDF Path " & smOldDDFPath
    Print #hmTo, "  New DDF Path " & smNewDDFPath
    For ilNewLoop = 0 To lbcNewFile.ListCount - 1 Step 1
        If lbcNewFile.Selected(ilNewLoop) Then
            llNoFilesProc = llNoFilesProc + 1
            lacNoFiles.Caption = "Processing" & Str(llNoFilesProc) & " of" & Str(llNoFiles) & " files"
            'plcNoFiles.Value = 100 * llNoFilesProc / llNoFiles
            llPercent = 100 * (llNoFilesProc / llNoFiles)
            If llPercent >= 100 Then
                llPercent = 100
            End If
            lncNoFiles.X2 = ((pbcNoFiles.Width - 2 * lncNoFiles.X1) * llPercent) / 100
            If llPercent > 0 Then
                lncNoFiles.Visible = True
            End If
            DoEvents
            slName = Trim$(lbcNewFile.List(ilNewLoop))
            slLine = Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
            Print #hmTo, "  Processing file: " & slName & " on " & slLine
            For ilFile = 0 To UBound(tmNewDDFFile) - 1 Step 1
                If StrComp(Trim$(tmNewDDFFile(ilFile).sName), slName, 1) = 0 Then
                    'Extract File Name C:\xxx\xxx.btr or c:xxx.btr
                    slNewFile = Trim$(gRemoveZero(tmNewDDFFile(ilFile).sLocation))
                    ilPos = InStr(slNewFile, ":")
                    If ilPos > 0 Then
                        slNewFile = Mid$(slNewFile, ilPos + 1)
                    End If
                    ilPos = InStr(slNewFile, "\")
                    Do While ilPos > 0
                        slNewFile = Mid$(slNewFile, ilPos + 1)
                        ilPos = InStr(slNewFile, "\")
                    Loop
                    ilPos = InStr(slNewFile, ".")
                    slNewFile = Left(slNewFile, ilPos - 1) & "_.NEW" '".NEW"
                    ilFileFlags = Asc(tmNewDDFFile(ilFile).sFlags)  'Bit 4 indicates a dictionary file if set (1), and a
                                                                    'and a user-defined defined file cleared (0). The other
                                                                    'bits appear to be unused
                    mFieldBuild smNewDDFPath, ".DDF", tmNewDDFFile(ilFile).iFileID, tmNewDDFField()
                    mIndexBuild smNewDDFPath, ".DDF", tmNewDDFFile(ilFile).iFileID, tmNewDDFIndex()
                    Exit For
                End If
            Next ilFile
            imProcURF = False
            If InStr(1, slNewFile, "urf.", 1) > 0 Then
                imProcURF = True
            End If
            ReDim tmOldDDFField(0 To 0) As DDFFIELD
            slOldFile = ""
            For ilOldLoop = 0 To lbcOldFile.ListCount - 1 Step 1
                If StrComp(Trim$(lbcOldFile.List(ilOldLoop)), slName, 1) = 0 Then
                    For ilFile = 0 To UBound(tmOldDDFFile) - 1 Step 1
                        If StrComp(Trim$(tmOldDDFFile(ilFile).sName), slName, 1) = 0 Then
                            'Extract File Name C:\xxx\xxx.btr or c:xxx.btr
                            slOldFile = Trim$(gRemoveZero(tmOldDDFFile(ilFile).sLocation))
                            ilPos = InStr(slOldFile, ":")
                            If ilPos > 0 Then
                                slOldFile = Mid$(slOldFile, ilPos + 1)
                            End If
                            ilPos = InStr(slOldFile, "\")
                            Do While ilPos > 0
                                slOldFile = Mid$(slOldFile, ilPos + 1)
                                ilPos = InStr(slOldFile, "\")
                            Loop
                            mFieldBuild smOldDDFPath, ".DDF", tmOldDDFFile(ilFile).iFileID, tmOldDDFField()
                            Exit For
                        End If
                    Next ilFile
                    Exit For
                End If
            Next ilOldLoop
            llRecNo = 0
            'Create new file---
            ilRet = btrClose(hmNewFile)
            btrDestroy hmNewFile
            ilRet = btrClose(hmOldFile)
            btrDestroy hmOldFile
            lmNewBaseRecLen = 0
            imNewVarLenRec = False
            imNewVarType = 0
            For ilField = 0 To UBound(tmNewDDFField) - 1 Step 1
                If (Asc(tmNewDDFField(ilField).sDataType) = 12) Then
                    imNewVarLenRec = True
                    imNewVarType = 12
                ElseIf (Asc(tmNewDDFField(ilField).sDataType) = 13) Then
                    imNewVarLenRec = True
                    imNewVarType = 13
                Else
                    lmNewBaseRecLen = lmNewBaseRecLen + tmNewDDFField(ilField).iSize
                End If
            Next ilField
            'Flag bits
            '0x0001       File contains variable-length records
            '0x0002       Trailing blanks are truncated
            '0x0004       Pages were preallocated
            '0x0008       Data in file is compressed
            '0x0010       File contains keys only
            '0x0020       Index balancing is in effect
            '0x0040       Free space threshold is 10%
            '0x0080       Free space threshold is 20%
            '0x00C0       Free space threshold is 30%
            '0x0100       File has reserved space for duplicate pointers
            '0x0200       System Data (0x1000 must be 0)
            '0x0400       Specific key number was assigned
            '0x0800       File contains VATs (version 6.1+ only)
            '
            On Error GoTo gGetDateTimeErr:
            ilRet = 0
            slStamp = FileDateTime(smDBPath & "Replication.ini")
            If ilRet = 0 Then
                ilFileFlags = &H200
            Else
                ilFileFlags = 0
            End If
            slFieldSet = smNewDDFPath & "FieldSet.New" 'smDBPath & "FieldSet.New"
            If imNewVarLenRec Then
                ilFileFlags = ilFileFlags + 1
            Else
                ilFileFlags = ilFileFlags + 0
            End If
            On Error Resume Next
            ilRet = GetPrivateProfileString(slName, "BlankTruncation", "Not Found", slReturn, 128, slFieldSet)
            If Left$(slReturn, ilRet) = "Yes" Then
                ilFileFlags = ilFileFlags + &H2
            End If
            On Error Resume Next
            ilRet = GetPrivateProfileString(slName, "DataCompression", "Not Found", slReturn, 128, slFieldSet)
            If Left$(slReturn, ilRet) = "Yes" Then
                ilFileFlags = ilFileFlags + &H8
            End If
            On Error Resume Next
            ilRet = GetPrivateProfileString(slName, "KeyOnly", "Not Found", slReturn, 128, slFieldSet)
            If Left$(slReturn, ilRet) = "Yes" Then
                ilFileFlags = ilFileFlags + &H10
            End If
            On Error Resume Next
            ilRet = GetPrivateProfileString(slName, "IndexBalancing", "Not Found", slReturn, 128, slFieldSet)
            If Left$(slReturn, ilRet) = "Yes" Then
                ilFileFlags = ilFileFlags + &H20
            End If
            On Error Resume Next
            ilRet = GetPrivateProfileString(slName, "FreeSpace", "Not Found", slReturn, 128, slFieldSet)
            If Left$(slReturn, ilRet) = "10" Then
                ilFileFlags = ilFileFlags + &H40
            ElseIf Left$(slReturn, ilRet) = "20" Then
                ilFileFlags = ilFileFlags + &H80
            ElseIf Left$(slReturn, ilRet) = "30" Then
                ilFileFlags = ilFileFlags + &HC0
            End If
            On Error Resume Next
            ilRet = GetPrivateProfileString(slName, "PageSize", "Not Found", slReturn, 128, slFieldSet)
            If Left$(slReturn, ilRet) <> "Not Found" Then
                ilPageSize = Val(Left$(slReturn, ilRet))
            Else
                If lmNewBaseRecLen < 256 Then
                    ilPages = 0
                ElseIf lmNewBaseRecLen < 512 Then
                    ilPages = 1
                ElseIf lmNewBaseRecLen < 1024 Then
                    ilPages = 2
                ElseIf lmNewBaseRecLen < 2048 Then
                    ilPages = 3
                Else
                    ilPages = 0
                End If
                'ilPageSize = imPageSize(ilPages)    '4096
                'Because of ast issue, always use 4096 as the page size
                ilPageSize = imPageSize(3)
            End If
            If ilPageSize = imPageSize(1) Then
                ilPages = 1
            ElseIf ilPageSize = imPageSize(2) Then
                ilPages = 2
            ElseIf ilPageSize = imPageSize(3) Then
                ilPages = 3
            Else
                ilPages = 0
            End If
            ilAltCollateReq = False
            For ilLoop = 0 To UBound(tmNewDDFIndex) - 1 Step 1
                'Find field
                For ilNewField = 0 To UBound(tmNewDDFField) - 1 Step 1
                    If tmNewDDFField(ilNewField).iFieldID = tmNewDDFIndex(ilLoop).iFieldID Then
                        If (tmNewDDFIndex(ilLoop).iFlag And BTRV_KF_ALT_COL_SEQ) = BTRV_KF_ALT_COL_SEQ Then
                            ilAltCollateReq = True
                            Exit For
                        End If
                    End If
                Next 'ilNewField
                If ilAltCollateReq Then
                    Exit For
                End If
            Next 'ilLoop
            slAltCollateFile = ""
            slAltCollateSeq = ""
            If ilAltCollateReq Then
                On Error Resume Next
                ilRet = GetPrivateProfileString(slName, "AltCollateFile", "Not Found", slReturn, 128, slFieldSet)
                If Left$(slReturn, ilRet) <> "Not Found" Then
                    slAltCollateFile = Left$(slReturn, ilRet)
                Else
                    On Error Resume Next
                    ilRet = GetPrivateProfileString(slName, "AltCollateSeq", "Not Found", slReturnAlt, 256, slFieldSet)
                    If Left$(slReturn, ilRet) <> "Not Found" Then
                        slAltCollateSeq = Left$(slReturn, ilRet)
                    Else
                        slAltCollateFile = "Upper.Alt"  'default
                    End If
                End If
            End If
            If rbcFile(0).Value Then
                Do
                    On Error Resume Next
                    'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
                    'Kill smCreateDBPath & slNewFile
                    mDeleteFile smCreateDBPath & slNewFile
                    On Error GoTo 0
                    hmNewFile = CBtrvTable()    'CBtrvObj()
                    ilNoKeys = 1
                    For ilLoop = 0 To UBound(tmNewDDFIndex) - 1 Step 1
                        If tmNewDDFIndex(ilLoop).iNumber + 1 > ilNoKeys Then
                            ilNoKeys = tmNewDDFIndex(ilLoop).iNumber + 1
                        End If
                    Next ilLoop
                    ilRet = btrCreFile(hmNewFile, lmNewBaseRecLen, ilPageSize, ilNoKeys, ilFileFlags, 0)
                    If ilRet <> BTRV_ERR_NONE Then
                        Print #hmTo, "Create File Failed for " & smCreateDBPath & slNewFile & " conversion aborted"
                        Close hmTo
                        Screen.MousePointer = vbDefault
                        MsgBox "Create Failed for " & smCreateDBPath & slNewFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                        bErrorFound = True
                        lacNoFiles.Caption = ""
                        lncNoFiles.Visible = False
                        lncNoFiles.X2 = lncNoFiles.X1
                        lacFileName.Caption = ""
                        lacTotal.Caption = ""
                        lncGauge.Visible = False
                        lncGauge.X2 = lncGauge.X1
                        Exit Sub
                    End If
                    For ilLoop = 0 To UBound(tmNewDDFIndex) - 1 Step 1
                        'Find field
                        For ilNewField = 0 To UBound(tmNewDDFField) - 1 Step 1
                            If tmNewDDFField(ilNewField).iFieldID = tmNewDDFIndex(ilLoop).iFieldID Then
                                ilRet = btrCreKey(hmNewFile, tmNewDDFField(ilNewField).iOffset + 1, tmNewDDFField(ilNewField).iSize, tmNewDDFIndex(ilLoop).iFlag, Asc(tmNewDDFField(ilNewField).sDataType), 0)
                                If ilRet <> BTRV_ERR_NONE Then
                                    Print #hmTo, "Create Key Failed for " & smCreateDBPath & slNewFile & " key " & Trim$(tmNewDDFField(ilNewField).sName) & " conversion aborted"
                                    Close hmTo
                                    Screen.MousePointer = vbDefault
                                    MsgBox "Create Key Failed for " & smCreateDBPath & slNewFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                                    bErrorFound = True
                                    lacNoFiles.Caption = ""
                                    lncNoFiles.Visible = False
                                    lncNoFiles.X2 = lncNoFiles.X1
                                    lacFileName.Caption = ""
                                    lacTotal.Caption = ""
                                    lncGauge.Visible = False
                                    lncGauge.X2 = lncGauge.X1
                                    Exit Sub
                                End If
                                Exit For
                            End If
                        Next ilNewField
                    Next ilLoop
                    If ilAltCollateReq Then
                        If Len(slAltCollateFile) > 0 Then
                            ilRet = btrCreCollateFile(hmNewFile, slAltCollateFile)
                            If ilRet <> BTRV_ERR_NONE Then
                                Print #hmTo, "Create Collate File Failed for " & smCreateDBPath & slNewFile & "(" & slAltCollateFile & ") conversion aborted"
                                Close hmTo
                                Screen.MousePointer = vbDefault
                                MsgBox "Create Failed for " & smCreateDBPath & slNewFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                                bErrorFound = True
                                lacNoFiles.Caption = ""
                                lncNoFiles.Visible = False
                                lncNoFiles.X2 = lncNoFiles.X1
                                lacFileName.Caption = ""
                                lacTotal.Caption = ""
                                lncGauge.Visible = False
                                lncGauge.X2 = lncGauge.X1
                                Exit Sub
                            End If
                        Else
                            If Len(slAltCollateSeq) > 0 Then
                                slSeqName = slNewFile
                                ilPos = InStr(slSeqName, ".")
                                If ilPos > 0 Then
                                    slSeqName = Left$(slSeqName, ilPos - 1)
                                End If
                                ilRet = btrCreCollate(hmNewFile, slSeqName, slAltCollateSeq)
                                If ilRet <> BTRV_ERR_NONE Then
                                    Print #hmTo, "Create Collate Sequence Failed for " & smCreateDBPath & slNewFile & " conversion aborted"
                                    Close hmTo
                                    Screen.MousePointer = vbDefault
                                    MsgBox "Create Failed for " & smCreateDBPath & slNewFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                                    bErrorFound = True
                                    lacNoFiles.Caption = ""
                                    lncNoFiles.Visible = False
                                    lncNoFiles.X2 = lncNoFiles.X1
                                    lacFileName.Caption = ""
                                    lacTotal.Caption = ""
                                    lncGauge.Visible = False
                                    lncGauge.X2 = lncGauge.X1
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                    ilRet = btrCreCreate(hmNewFile, smCreateDBPath & slNewFile, 0)
                    btrDestroy hmNewFile
                    If ilRet = BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    ilPages = ilPages + 1
                    If ilPages > UBound(imPageSize) Then
                        Print #hmTo, "Create Failed for " & smCreateDBPath & slNewFile & " conversion aborted"
                        Close hmTo
                        Screen.MousePointer = vbDefault
                        MsgBox "Create Failed for " & smCreateDBPath & slNewFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                        bErrorFound = True
                        lacNoFiles.Caption = ""
                        lncNoFiles.Visible = False
                        lncNoFiles.X2 = lncNoFiles.X1
                        lacFileName.Caption = ""
                        lacTotal.Caption = ""
                        lncGauge.Visible = False
                        lncGauge.X2 = lncGauge.X1
                        Exit Sub
                    End If
                    ilPageSize = imPageSize(ilPages)
                Loop While ilPages <= UBound(imPageSize)
                Print #hmTo, "    Created New File: " & smCreateDBPath & slNewFile
            Else
                Print #hmTo, "    Copying to New File: " & smCreateDBPath & slNewFile
            End If
            'Open old
            '10/30/13: Bypass those files that only want to create blank images
            blCopyRecords = True
            If IsArray(smCreateOnlyArray) Then
                For ilField = 0 To UBound(smCreateOnlyArray) Step 1
                    If UCase(slOldFile) = UCase(smCreateOnlyArray(ilField)) Then
                        blCopyRecords = False
                        Exit For
                    End If
                Next ilField
            End If
            If (slOldFile <> "") And (blCopyRecords) Then
                mInitValues slFieldSet, slName
                lmOldVarLenRec = False
                imOldVarType = 0
                lmOldBaseRecLen = 0
                For ilField = 0 To UBound(tmOldDDFField) - 1 Step 1
                    If (Asc(tmOldDDFField(ilField).sDataType) = 12) Then
                        lmOldVarLenRec = True
                        imOldVarType = 12
                    ElseIf (Asc(tmOldDDFField(ilField).sDataType) = 13) Then
                        lmOldVarLenRec = True
                        imOldVarType = 13
                    Else
                        lmOldBaseRecLen = lmOldBaseRecLen + tmOldDDFField(ilField).iSize
                    End If
                Next ilField
                hmOldFile = CBtrvTable()    'CBtrvObj()
                ilRet = btrOpen(hmOldFile, "", smDBPath & slOldFile, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet <> BTRV_ERR_FILE_NOT_FOUND Then
                        ilCRet = btrClose(hmOldFile)
                        btrDestroy hmOldFile
                        ilCRet = btrClose(hmNewFile)
                        btrDestroy hmNewFile
                        Print #hmTo, "Open File Failed for " & smDBPath & slOldFile & " conversion aborted"
                        Close hmTo
                        Screen.MousePointer = vbDefault
                        MsgBox "Open Failed for " & smDBPath & slOldFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                        bErrorFound = True
                        lacNoFiles.Caption = ""
                        lncNoFiles.Visible = False
                        lncNoFiles.X2 = lncNoFiles.X1
                        lacFileName.Caption = ""
                        lacTotal.Caption = ""
                        lncGauge.Visible = False
                        lncGauge.X2 = lncGauge.X1
                        Exit Sub
                    Else
                        btrDestroy hmOldFile
                        lmMaxNoRec = 0
                        Print #hmTo, "File didn't exist " & smDBPath & slOldFile & " conversion continued"
                    End If
                Else
                    lmMaxNoRec = btrRecords(hmOldFile)
                End If
                hmNewFile = CBtrvTable()    'CBtrvObj()
                ilRet = btrOpen(hmNewFile, "", smCreateDBPath & slNewFile, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                If ilRet <> BTRV_ERR_NONE Then
                    ilCRet = btrClose(hmOldFile)
                    btrDestroy hmOldFile
                    ilCRet = btrClose(hmNewFile)
                    btrDestroy hmNewFile
                    Print #hmTo, "Open File Failed for " & smCreateDBPath & slNewFile & " conversion aborted"
                    Close hmTo
                    Screen.MousePointer = vbDefault
                    MsgBox "Open Failed for " & smCreateDBPath & slNewFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                    bErrorFound = True
                    lacNoFiles.Caption = ""
                    lncNoFiles.Visible = False
                    lncNoFiles.X2 = lncNoFiles.X1
                    lacFileName.Caption = ""
                    lacTotal.Caption = ""
                    lncGauge.Visible = False
                    lncGauge.X2 = lncGauge.X1
                    Exit Sub
                End If
                lacFileName.Caption = "Processing: " & slName
                llNoProcessed = 0
                lacTotal.Caption = "0 of" & Str(lmMaxNoRec) & " records"
                'plcGauge.Value = 0
                lncGauge.X2 = lncGauge.X1
            
                If lmMaxNoRec > 0 Then
                    If lmOldVarLenRec Then
                        lmOldActualRecLen = Len(tmOldRec)
                        'lmNewBaseRecLen = lmNewActualRecLen
                        ilRet = btrGetFirst(hmOldFile, tmOldRec, lmOldActualRecLen, INDEXKEY0, BTRV_LOCK_NONE)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilCRet = btrClose(hmOldFile)
                            btrDestroy hmOldFile
                            ilCRet = btrClose(hmNewFile)
                            btrDestroy hmNewFile
                            Print #hmTo, "Get First Failed for " & smDBPath & slOldFile & " conversion aborted"
                            Close hmTo
                            Screen.MousePointer = vbDefault
                            MsgBox "Get First Failed for " & smDBPath & slOldFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                            bErrorFound = True
                            lacNoFiles.Caption = ""
                            lncNoFiles.Visible = False
                            lncNoFiles.X2 = lncNoFiles.X1
                            lacFileName.Caption = ""
                            lacTotal.Caption = ""
                            lncGauge.Visible = False
                            lncGauge.X2 = lncGauge.X1
                            Exit Sub
                        End If
                        Do While (ilRet = BTRV_ERR_NONE)
                            DoEvents
                            If imTerminate Then
                                Exit Do
                            End If
                            'Copy fields
                            mCopyFields
                            'Insert record
                            If (imNewVarLenRec) And (imNewVarType = 12) Then
                                For ilField = 0 To UBound(tmNewDDFField) - 1 Step 1
                                    If (Asc(tmNewDDFField(ilField).sDataType) = 12) Then
                                        lmNewActualRecLen = lmNewBaseRecLen + imVarLenNote  'tmNewDDFField(ilField).iSize
                                        ilRet = btrInsert(hmNewFile, tgNewRec, lmNewActualRecLen, INDEXKEY0)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            ilCRet = btrClose(hmOldFile)
                                            btrDestroy hmOldFile
                                            ilCRet = btrClose(hmNewFile)
                                            btrDestroy hmNewFile
                                            Print #hmTo, "Insert Failed for " & smCreateDBPath & slNewFile & " conversion aborted"
                                            Close hmTo
                                            Screen.MousePointer = vbDefault
                                            MsgBox "Insert Failed for " & smCreateDBPath & slNewFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                                            bErrorFound = True
                                            lacNoFiles.Caption = ""
                                            lncNoFiles.Visible = False
                                            lncNoFiles.X2 = lncNoFiles.X1
                                            lacFileName.Caption = ""
                                            lacTotal.Caption = ""
                                            lncGauge.Visible = False
                                            lncGauge.X2 = lncGauge.X1
                                            Exit Sub
                                        End If
                                    End If
                                Next ilField
                            ElseIf (imNewVarLenRec) And (imNewVarType = 13) Then
                                For ilField = 0 To UBound(tmNewDDFField) - 1 Step 1
                                    If (Asc(tmNewDDFField(ilField).sDataType) = 13) Then
                                        ilIndex = tmNewDDFField(ilField).iOffset + 1
                                        HMemCpy ilValue, tgNewRec.sChar(ilIndex), 2
                                        lmNewActualRecLen = lmNewBaseRecLen + ilValue + 4 'lmNewBaseRecLen + Len(smFieldValue) + 4
                                        ilRet = btrInsert(hmNewFile, tgNewRec, lmNewActualRecLen, INDEXKEY0)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            ilCRet = btrClose(hmOldFile)
                                            btrDestroy hmOldFile
                                            ilCRet = btrClose(hmNewFile)
                                            btrDestroy hmNewFile
                                            Print #hmTo, "Insert Failed for " & smCreateDBPath & slNewFile & " conversion aborted"
                                            Close hmTo
                                            Screen.MousePointer = vbDefault
                                            MsgBox "Insert Failed for " & smCreateDBPath & slNewFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                                            bErrorFound = True
                                            lacNoFiles.Caption = ""
                                            lncNoFiles.Visible = False
                                            lncNoFiles.X2 = lncNoFiles.X1
                                            lacFileName.Caption = ""
                                            lacTotal.Caption = ""
                                            lncGauge.Visible = False
                                            lncGauge.X2 = lncGauge.X1
                                            Exit Sub
                                        End If
                                        Exit For
                                    End If
                                Next ilField
                            Else
                                lmNewActualRecLen = lmNewBaseRecLen
                                ilRet = btrInsert(hmNewFile, tgNewRec, lmNewActualRecLen, INDEXKEY0)
                                If ilRet <> BTRV_ERR_NONE Then
                                    ilCRet = btrClose(hmOldFile)
                                    btrDestroy hmOldFile
                                    ilCRet = btrClose(hmNewFile)
                                    btrDestroy hmNewFile
                                    Print #hmTo, "Insert Failed for " & smCreateDBPath & slNewFile & " conversion aborted"
                                    Close hmTo
                                    Screen.MousePointer = vbDefault
                                    MsgBox "Insert Failed for " & smCreateDBPath & slNewFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                                    bErrorFound = True
                                    lacNoFiles.Caption = ""
                                    lncNoFiles.Visible = False
                                    lncNoFiles.X2 = lncNoFiles.X1
                                    lacFileName.Caption = ""
                                    lacTotal.Caption = ""
                                    lncGauge.Visible = False
                                    lncGauge.X2 = lncGauge.X1
                                    Exit Sub
                                End If
                            End If
    
                            llNoProcessed = llNoProcessed + 1
                            lacTotal.Caption = Trim$(Str(llNoProcessed)) & " of" & Str(lmMaxNoRec) & " records"
                            llPercent = 100 * (llNoProcessed / lmMaxNoRec)
                            If llPercent >= 100 Then
                                llPercent = 100
                            End If
                            lncGauge.X2 = ((pbcGauge.Width - 2 * lncGauge.X1) * llPercent) / 100
                            If llPercent > 0 Then
                                lncGauge.Visible = True
                            End If
                            DoEvents
                            lmOldActualRecLen = Len(tmOldRec)
                            ilRet = btrGetNext(hmOldFile, tmOldRec, lmOldActualRecLen, BTRV_LOCK_NONE)
                        Loop
                    Else
                        lmOldActualRecLen = lmOldBaseRecLen
                        llExtLen = lmOldActualRecLen
                        llNoRec = gExtNoRec(CInt(llExtLen)) 'btrRecords(hlAdf) 'Obtain number of records
                        btrExtClear hmFile   'Clear any previous extend operation
                        ilRet = btrGetFirst(hmOldFile, tmOldRec, lmOldActualRecLen, INDEXKEY0, BTRV_LOCK_NONE)
                        If ilRet <> BTRV_ERR_NONE Then
                            ilCRet = btrClose(hmOldFile)
                            btrDestroy hmOldFile
                            ilCRet = btrClose(hmNewFile)
                            btrDestroy hmNewFile
                            Print #hmTo, "Get First Failed for " & smDBPath & slOldFile & " conversion aborted"
                            Close hmTo
                            Screen.MousePointer = vbDefault
                            MsgBox "Get First Failed for " & smDBPath & slOldFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                            bErrorFound = True
                            lacNoFiles.Caption = ""
                            lncNoFiles.Visible = False
                            lncNoFiles.X2 = lncNoFiles.X1
                            lacFileName.Caption = ""
                            lacTotal.Caption = ""
                            lncGauge.Visible = False
                            lncGauge.X2 = lncGauge.X1
                            Exit Sub
                        End If
                        Call btrExtSetBounds(hmOldFile, llNoRec, -1, "UC") 'Set extract limits (all records)
                        ilRet = btrExtAddField(hmOldFile, 0, llExtLen)  'Extract iCode field
                        If ilRet <> BTRV_ERR_NONE Then
                            ilCRet = btrClose(hmOldFile)
                            btrDestroy hmOldFile
                            ilCRet = btrClose(hmNewFile)
                            btrDestroy hmNewFile
                            Print #hmTo, "Ext Set Bounds Failed for " & smDBPath & slOldFile & " conversion aborted"
                            Close hmTo
                            Screen.MousePointer = vbDefault
                            MsgBox "Ext Set Bounds Failed for " & smDBPath & slOldFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                            bErrorFound = True
                            lacNoFiles.Caption = ""
                            lncNoFiles.Visible = False
                            lncNoFiles.X2 = lncNoFiles.X1
                            lacFileName.Caption = ""
                            lacTotal.Caption = ""
                            lncGauge.Visible = False
                            lncGauge.X2 = lncGauge.X1
                            Exit Sub
                        End If
                        ilRet = btrExtGetNext(hmOldFile, tmOldRec, llExtLen, llRecPos)
                        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                            If ilRet <> BTRV_ERR_NONE Then
                                ilCRet = btrClose(hmOldFile)
                                btrDestroy hmOldFile
                                ilCRet = btrClose(hmNewFile)
                                btrDestroy hmNewFile
                                Print #hmTo, "Ext Get Next Failed for " & smDBPath & slOldFile & " conversion aborted"
                                Close hmTo
                                Screen.MousePointer = vbDefault
                                MsgBox "Ext Get Next Failed for " & smDBPath & slOldFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                                bErrorFound = True
                                lacNoFiles.Caption = ""
                                lncNoFiles.Visible = False
                                lncNoFiles.X2 = lncNoFiles.X1
                                lacFileName.Caption = ""
                                lacTotal.Caption = ""
                                lncGauge.Visible = False
                                lncGauge.X2 = lncGauge.X1
                                Exit Sub
                            End If
                            llExtLen = lmOldActualRecLen  'Extract operation record size
                            'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(imRowUpperBound), ilExtLen, llRecPos)
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hmOldFile, tmOldRec, llExtLen, llRecPos)
                            Loop
                            Do While ilRet = BTRV_ERR_NONE
                                DoEvents
                                If imTerminate Then
                                    Exit Do
                                End If
                                'Copy fields
                                mCopyFields
                                'Insert record
                                If (imNewVarLenRec) And (imNewVarType = 12) Then
                                    For ilField = 0 To UBound(tmNewDDFField) - 1 Step 1
                                        If (Asc(tmNewDDFField(ilField).sDataType) = 12) Then
                                            lmNewActualRecLen = lmNewBaseRecLen + tmNewDDFField(ilField).iSize
                                            ilRet = btrInsert(hmNewFile, tgNewRec, lmNewActualRecLen, INDEXKEY0)
                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilCRet = btrClose(hmOldFile)
                                                btrDestroy hmOldFile
                                                ilCRet = btrClose(hmNewFile)
                                                btrDestroy hmNewFile
                                                Print #hmTo, "Insert Failed for " & smCreateDBPath & slNewFile & " conversion aborted"
                                                Close hmTo
                                                Screen.MousePointer = vbDefault
                                                MsgBox "Insert Failed for " & smCreateDBPath & slNewFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                                                bErrorFound = True
                                                lacNoFiles.Caption = ""
                                                lncNoFiles.Visible = False
                                                lncNoFiles.X2 = lncNoFiles.X1
                                                lacFileName.Caption = ""
                                                lacTotal.Caption = ""
                                                lncGauge.Visible = False
                                                lncGauge.X2 = lncGauge.X1
                                                Exit Sub
                                            End If
                                        End If
                                    Next ilField
                                ElseIf (imNewVarLenRec) And (imNewVarType = 13) Then
                                    For ilField = 0 To UBound(tmNewDDFField) - 1 Step 1
                                        If (Asc(tmNewDDFField(ilField).sDataType) = 13) Then
                                            ilIndex = tmNewDDFField(ilField).iOffset + 1
                                            HMemCpy ilValue, tgNewRec.sChar(ilIndex), 2
                                            lmNewActualRecLen = lmNewBaseRecLen + ilValue + 4 'lmNewBaseRecLen + Len(smFieldValue) + 4
                                            ilRet = btrInsert(hmNewFile, tgNewRec, lmNewActualRecLen, INDEXKEY0)
                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilCRet = btrClose(hmOldFile)
                                                btrDestroy hmOldFile
                                                ilCRet = btrClose(hmNewFile)
                                                btrDestroy hmNewFile
                                                Print #hmTo, "Insert Failed for " & smCreateDBPath & slNewFile & " conversion aborted"
                                                Close hmTo
                                                Screen.MousePointer = vbDefault
                                                MsgBox "Insert Failed for " & smCreateDBPath & slNewFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                                                bErrorFound = True
                                                lacNoFiles.Caption = ""
                                                lncNoFiles.Visible = False
                                                lncNoFiles.X2 = lncNoFiles.X1
                                                lacFileName.Caption = ""
                                                lacTotal.Caption = ""
                                                lncGauge.Visible = False
                                                lncGauge.X2 = lncGauge.X1
                                                Exit Sub
                                            End If
                                            Exit For
                                        End If
                                    Next ilField
                                Else
                                    ilRet = btrInsert(hmNewFile, tgNewRec, lmNewBaseRecLen, INDEXKEY0)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        ilCRet = btrClose(hmOldFile)
                                        btrDestroy hmOldFile
                                        ilCRet = btrClose(hmNewFile)
                                        btrDestroy hmNewFile
                                        Print #hmTo, "Insert Failed for " & smCreateDBPath & slNewFile & " conversion aborted"
                                        Close hmTo
                                        Screen.MousePointer = vbDefault
                                        MsgBox "Insert Failed for " & smCreateDBPath & slNewFile & " error # " & Str(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Error"
                                        bErrorFound = True
                                        lacNoFiles.Caption = ""
                                        lncNoFiles.Visible = False
                                        lncNoFiles.X2 = lncNoFiles.X1
                                        lacFileName.Caption = ""
                                        lacTotal.Caption = ""
                                        lncGauge.Visible = False
                                        lncGauge.X2 = lncGauge.X1
                                        Exit Sub
                                    End If
                                End If
                                llNoProcessed = llNoProcessed + 1
                                lacTotal.Caption = Trim$(Str(llNoProcessed)) & " of" & Str(lmMaxNoRec) & " records"
                                llPercent = 100 * (llNoProcessed / lmMaxNoRec)
                                If llPercent >= 100 Then
                                    llPercent = 100
                                End If
                                lncGauge.X2 = ((pbcGauge.Width - 2 * lncGauge.X1) * llPercent) / 100
                                If llPercent > 0 Then
                                    lncGauge.Visible = True
                                End If
                                DoEvents
                                llExtLen = lmOldActualRecLen  'Extract operation record size
                                ilRet = btrExtGetNext(hmOldFile, tmOldRec, llExtLen, llRecPos)
                                Do While ilRet = BTRV_ERR_REJECT_COUNT
                                    ilRet = btrExtGetNext(hmOldFile, tmOldRec, llExtLen, llRecPos)
                                Loop
                            Loop
                        End If
                    End If
                Else
                    llPercent = 100
                    lncGauge.X2 = ((pbcGauge.Width - 2 * lncGauge.X1) * llPercent) / 100
                    If llPercent > 0 Then
                        lncGauge.Visible = True
                    End If
                    DoEvents
                End If
                slLine = Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
                Print #hmTo, "    Moved " & Str(llNoProcessed) & " records out of" & Str(lmMaxNoRec) & " for " & slName & " on " & slLine
            End If
        End If
    Next ilNewLoop
    DoEvents
    ilRet = btrClose(hmNewFile)
    btrDestroy hmNewFile
    ilRet = btrClose(hmOldFile)
    btrDestroy hmOldFile
    'Rename files
    slLine = Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmTo, "DDFReorg: Rename files on " & slLine
    For ilNewLoop = 0 To lbcNewFile.ListCount - 1 Step 1
        If lbcNewFile.Selected(ilNewLoop) Then
            slName = Trim$(lbcNewFile.List(ilNewLoop))
            For ilFile = 0 To UBound(tmNewDDFFile) - 1 Step 1
                DoEvents
                If StrComp(Trim$(tmNewDDFFile(ilFile).sName), slName, 1) = 0 Then
                    'Extract File Name C:\xxx\xxx.btr or c:xxx.btr
                    slNewFile = Trim$(gRemoveZero(tmNewDDFFile(ilFile).sLocation))
                    ilPos = InStr(slNewFile, ":")
                    If ilPos > 0 Then
                        slNewFile = Mid$(slNewFile, ilPos + 1)
                    End If
                    ilPos = InStr(slNewFile, "\")
                    Do While ilPos > 0
                        slNewFile = Mid$(slNewFile, ilPos + 1)
                        ilPos = InStr(slNewFile, "\")
                    Loop
                    ilPos = InStr(slNewFile, ".")
                    slExtName = Mid$(slNewFile, ilPos)
                    slNewFile = Left(slNewFile, ilPos - 1)
                    On Error Resume Next
                    'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
                    'Kill smDBPath & slNewFile & ".old"
                    mDeleteFile smDBPath & slNewFile & ".old"
                    'Determine if btr or mkd extension
                    On Error GoTo gGetDateTimeErr:
                    ilRet = 0
                    slStamp = FileDateTime(smDBPath & slNewFile & slExtName)
                    
                    On Error GoTo 0
                            
                    'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
                    mRenameFile smDBPath & slNewFile & slExtName, smDBPath & slNewFile & ".old"
                    
                    'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
                    mRenameFile smCreateDBPath & slNewFile & "_.New", smDBPath & slNewFile & slExtName
                    
                    'Check for extensions
                    slExtension = ".^01"
                    ilExtNumber = 1
                    Do
                        On Error Resume Next
                        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
                        'Kill smDBPath & slNewFile & "_old" & slExtension
                        mDeleteFile smDBPath & slNewFile & "_old" & slExtension
                        On Error GoTo gGetDateTimeErr:
                        ilRet = 0
                        slStamp = FileDateTime(smDBPath & slNewFile & slExtension)
                        If ilRet = 0 Then
                            ilRet = 0
                            'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
                            mRenameFile smDBPath & slNewFile & slExtension, smDBPath & slNewFile & "_old" & slExtension
                            
                            On Error GoTo 0
                            ilExtNumber = ilExtNumber + 1
                            
                            If ilExtNumber <= 15 Then
                                slExtension = ".^0" & Trim$(Hex$(ilExtNumber)) 'Trim$(Str$(ilExtNumber))
                            Else
                                slExtension = ".^" & Trim$(Hex$(ilExtNumber)) 'Trim$(Str$(ilExtNumber))
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                    slExtension = ".^01"
                    ilExtNumber = 1
                    Do
                        On Error GoTo gGetDateTimeErr:
                        ilRet = 0
                        slStamp = FileDateTime(smCreateDBPath & slNewFile & "_" & slExtension)
                        If ilRet = 0 Then
                            ilRet = 0
                            'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
                            mRenameFile smCreateDBPath & slNewFile & "_" & slExtension, smDBPath & slNewFile & slExtension
                            
                            On Error GoTo 0
                            ilExtNumber = ilExtNumber + 1
                            
                            If ilExtNumber <= 15 Then
                                slExtension = ".^0" & Trim$(Hex$(ilExtNumber)) 'Trim$(Str$(ilExtNumber))
                            Else
                                slExtension = ".^" & Trim$(Hex$(ilExtNumber)) 'Trim$(Str$(ilExtNumber))
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                    On Error GoTo 0
                    Exit For
                End If
            Next ilFile
        End If
    Next ilNewLoop
    
    If StrComp(smNewDDFPath, smOldDDFPath, 1) <> 0 Then
        lacRename.Caption = "Copying DDF's: " & smNewDDFPath & " To " & smOldDDFPath
        DoEvents
        slLine = Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
        Print #hmTo, "DDFReorg: Rename DDF's on " & slLine
        mMoveDDFs
    End If
    lacRename.Caption = "Finished Converting and Renaming Files"
    slLine = Format$(Now, "m/d/yy") & " at " & Format$(Now, "h:mm:ssAM/PM")
    Print #hmTo, "DDFReorg: Finished on " & slLine
    Close hmTo
    lacNoFiles.Caption = ""
    lncNoFiles.Visible = False
    lncNoFiles.X2 = lncNoFiles.X1
    lacFileName.Caption = ""
    lacTotal.Caption = ""
    lncGauge.Visible = False
    lncGauge.X2 = lncGauge.X1
    lbcNewFile.Clear
    lbcOldFile.Clear
    
    lbcOldPath_Change
    lbcNewPath_Change
    Screen.MousePointer = 0 'DEFAULT
    Exit Sub
gGetDateTimeErr:
    ilRet = Err
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPutField                       *
'*                                                     *
'*             Created:7/05/95       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Put field into record           *
'*                                                     *
'*******************************************************
Private Sub mPutField(tlDDFField As DDFFIELD)
'Taken from ShowFix for vb6
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilSize As Integer
    Dim ilValue As Integer
    Dim llValue As Long
    Dim ilDyMn As Integer
    Dim ilYear As Integer
    Dim ilHsSec As Integer
    Dim ilMinHr As Integer
    Dim slMoney As String
    Dim slStr As String
    Dim slChar As String
    ilIndex = tlDDFField.iOffset + 1
    ilSize = tlDDFField.iSize
    Select Case Asc(tlDDFField.sDataType)
        Case 0  'String
            For ilLoop = 1 To ilSize Step 1
                If Len(smFieldValue) > 0 Then
                    slChar = Left$(smFieldValue, 1)
                    If ((Asc(slChar) >= Asc(" ")) And (Asc(slChar) <= Asc("~"))) Or (imProcURF) Then
                        tgNewRec.sChar(ilIndex) = Asc(slChar)
                    Else
                        tgNewRec.sChar(ilIndex) = Asc(" ")
                    End If
                Else
                    tgNewRec.sChar(ilIndex) = Asc(" ")
                End If
                smFieldValue = Mid(smFieldValue, 2)
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 1  'Integer
            'Remove any non-integer digits
            slStr = ""
            For ilLoop = 1 To Len(smFieldValue) Step 1
                slChar = Mid$(smFieldValue, ilLoop, 1)
                If ilLoop = 1 Then
                    If (Asc(slChar) >= Asc("0")) And (Asc(slChar) <= Asc("9")) Or (Asc(slChar) = Asc(".")) Or (Asc(slChar) = Asc("-")) Then
                        slStr = slStr & slChar
                    End If
                Else
                    If (Asc(slChar) >= Asc("0")) And (Asc(slChar) <= Asc("9")) Or (Asc(slChar) = Asc(".")) Then
                        slStr = slStr & slChar
                    End If
                End If
            Next ilLoop
            smFieldValue = slStr
            If ilSize = 1 Then
                ilValue = Val(smFieldValue)
                'tgNewRec.sChar(ilIndex) = Chr(ilValue)
                tgNewRec.sChar(ilIndex) = ilValue And &HFF
            ElseIf ilSize = 2 Then
                ilValue = Val(smFieldValue)
                HMemCpy tgNewRec.sChar(ilIndex), ilValue, 2
            ElseIf ilSize = 4 Then
                llValue = Val(smFieldValue)
                HMemCpy tgNewRec.sChar(ilIndex), llValue, 4
            End If
        Case 2  'IEEE Float
        Case 3  'Btrieve Date
            gPackDate smFieldValue, ilDyMn, ilYear
            tgNewRec.sChar(ilIndex) = ilDyMn And &HFF
            tgNewRec.sChar(ilIndex + 1) = (ilDyMn \ 256) And &HFF
            tgNewRec.sChar(ilIndex + 2) = ilYear And &HFF
            tgNewRec.sChar(ilIndex + 3) = (ilYear \ 256) And &HFF
        Case 4  'Btrieve Time
            gPackTime smFieldValue, ilHsSec, ilMinHr
            tgNewRec.sChar(ilIndex) = ilHsSec And &HFF
            tgNewRec.sChar(ilIndex + 1) = (ilHsSec \ 256) And &HFF
            tgNewRec.sChar(ilIndex + 2) = ilMinHr And &HFF
            tgNewRec.sChar(ilIndex + 3) = (ilMinHr \ 256) And &HFF
        Case 5  'COBOL Decimal COMP-3
            gStrToPDN smFieldValue, Asc(tlDDFField.sDec), ilSize, slMoney
            For ilLoop = 1 To ilSize Step 1
                tgNewRec.sChar(ilIndex) = Asc(Left$(slMoney, 1))
                slMoney = Mid(slMoney, 2)
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 6  'COBOL Money
            gStrToPDN smFieldValue, Asc(tlDDFField.sDec), ilSize, slMoney
            For ilLoop = 1 To ilSize Step 1
                tgNewRec.sChar(ilIndex) = Asc(Left$(slMoney, 1))
                slMoney = Mid(slMoney, 2)
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 7  'Logical
            If ilSize = 1 Then
                ilValue = Val(smFieldValue)
                tgNewRec.sChar(ilIndex) = ilValue And &HFF
            ElseIf ilSize = 2 Then
                ilValue = Val(smFieldValue)
                HMemCpy tgNewRec.sChar(ilIndex), ilValue, 2
            End If
        Case 8  'COBOL Numeric
        Case 9  'BASIC Float
        Case 10 'Pascal LString
            ilValue = Len(smFieldValue)
            tgNewRec.sChar(ilIndex) = ilValue And &HFF
            ilIndex = ilIndex + 1
            For ilLoop = 1 To ilValue Step 1
                tgNewRec.sChar(ilIndex) = Asc(Left$(smFieldValue, 1))
                smFieldValue = Mid(smFieldValue, 2)
                ilIndex = ilIndex + 1
            Next ilLoop
        Case 11 'C ZString
            ilValue = Len(smFieldValue)
            For ilLoop = 1 To ilValue Step 1
                tgNewRec.sChar(ilIndex) = Asc(Left$(smFieldValue, 1))
                smFieldValue = Mid(smFieldValue, 2)
                ilIndex = ilIndex + 1
            Next ilLoop
            tgNewRec.sChar(ilIndex) = 0
        Case 12 'Variable Length Note
        Case 13 'LVar (Length, Characters, 0, 0); Length does not include zeros
            ilValue = Len(smFieldValue)
            HMemCpy tgNewRec.sChar(ilIndex), ilValue, 2
            ilIndex = ilIndex + 2
            For ilLoop = 1 To ilValue Step 1
                'tgNewRec.sChar(ilIndex) = left$(smFieldValue, 1)
                tgNewRec.sChar(ilIndex) = Asc(Left$(smFieldValue, 1))
                smFieldValue = Mid(smFieldValue, 2)
                ilIndex = ilIndex + 1
            Next ilLoop
            tgNewRec.sChar(ilIndex) = 0
            tgNewRec.sChar(ilIndex + 1) = 0
        Case 14 'Unsigned Binary
            llValue = Val(smFieldValue)
            ilValue = gLongToUnsignInt(llValue)
            HMemCpy tgNewRec.sChar(ilIndex), ilValue, 2
        Case 15 'AutoIncrement
            If ilSize = 2 Then
                ilValue = Val(smFieldValue)
                HMemCpy tgNewRec.sChar(ilIndex), ilValue, 2
            ElseIf ilSize = 4 Then
                llValue = Val(smFieldValue)
                HMemCpy tgNewRec.sChar(ilIndex), llValue, 4
            End If
        Case 16 'Bit
        Case 17 'COBOL Numeric STS
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetCommands                    *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set Buttons                     *
'*                                                     *
'*******************************************************
Private Sub mSetCommands()
    Dim ilNewLoop As Integer
    cmcProcess.Enabled = False
    For ilNewLoop = 0 To lbcNewFile.ListCount - 1 Step 1
        If lbcNewFile.Selected(ilNewLoop) Then
            lacRename.Caption = "Changes Found.  DDF needs processing!"
            cmcProcess.Enabled = True
            Exit For
        End If
    Next ilNewLoop
End Sub

Private Sub mSetFiles()
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer
    Dim ilNewLoop As Integer
    DoEvents
    If (lbcNewFile.ListCount > 0) And (lbcOldFile.ListCount > 0) Then
        ilValue = False
        llRg = CLng(lbcNewFile.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcNewFile.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        llRg = CLng(lbcOldFile.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcOldFile.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        mTestFieldDiff
        mTestIndexDiff
        mTestCreateOnly
        mSetCommands
    Else
        cmcProcess.Enabled = False
    End If
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
    Dim ilLoop As Integer
    Erase tmInitValues
    Erase tmAdjustValues
    ilRet = btrClose(hmFile)
    btrDestroy hmFile
    Unload DDFReorg
    'Set DDFReorg = Nothing   'Remove data segment
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFieldDiff                  *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if field changed      *
'*                                                     *
'*******************************************************
Private Sub mTestFieldDiff()
    Dim slName As String
    Dim ilOldLoop As Integer
    Dim ilFile As Integer
    Dim ilNewLoop As Integer
    Dim ilOldField As Integer
    Dim ilNewField As Integer
    Dim ilFound As Integer
    Dim ilCompared As Integer
    Dim slFileName As String
    Dim slStamp As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    For ilOldLoop = 0 To lbcOldFile.ListCount - 1 Step 1
        slName = Trim$(lbcOldFile.List(ilOldLoop))
        For ilFile = 0 To UBound(tmOldDDFFile) - 1 Step 1
            If StrComp(Trim$(tmOldDDFFile(ilFile).sName), slName, 1) = 0 Then
                mFieldBuild smOldDDFPath, ".DDF", tmOldDDFFile(ilFile).iFileID, tmOldDDFField()
                Exit For
            End If
        Next ilFile
        For ilNewLoop = 0 To lbcNewFile.ListCount - 1 Step 1
            If StrComp(Trim$(lbcNewFile.List(ilNewLoop)), slName, 1) = 0 Then
                For ilFile = 0 To UBound(tmNewDDFFile) - 1 Step 1
                    If StrComp(Trim$(tmNewDDFFile(ilFile).sName), slName, 1) = 0 Then
                        If imRU Then
                            'Test if file exist
                            slFileName = Trim$(gRemoveZero(tmNewDDFFile(ilFile).sLocation))
                            ilPos = InStr(slFileName, ":")
                            If ilPos > 0 Then
                                slFileName = Mid$(slFileName, ilPos + 1)
                            End If
                            ilPos = InStr(slFileName, "\")
                            Do While ilPos > 0
                                slFileName = Mid$(slFileName, ilPos + 1)
                                ilPos = InStr(slFileName, "\")
                            Loop
                            On Error GoTo mTestFieldDiffErr:
                            ilRet = 0
                            slStamp = FileDateTime(smDBPath & slFileName)
                            If ilRet = 0 Then
                                ilFound = True
                            Else
                                ilFound = False
                            End If
                        Else
                            ilFound = True
                        End If
                        If ilFound Then
                            mFieldBuild smNewDDFPath, ".DDF", tmNewDDFFile(ilFile).iFileID, tmNewDDFField()
                            'Compare fields
                            For ilOldField = 0 To UBound(tmOldDDFField) - 1 Step 1
                                ilCompared = True
                                ilFound = False
                                For ilNewField = 0 To UBound(tmNewDDFField) - 1 Step 1
                                    If StrComp(Trim$(tmOldDDFField(ilOldField).sName), Trim$(tmNewDDFField(ilNewField).sName), 1) = 0 Then
                                        ilFound = True
                                        If StrComp(Trim$(tmOldDDFField(ilOldField).sDataType), Trim$(tmNewDDFField(ilNewField).sDataType), 1) <> 0 Then
                                            ilCompared = False
                                        End If
                                        If tmOldDDFField(ilOldField).iOffset <> tmNewDDFField(ilNewField).iOffset Then
                                            ilCompared = False
                                        End If
                                        If tmOldDDFField(ilOldField).iSize <> tmNewDDFField(ilNewField).iSize Then
                                            ilCompared = False
                                        End If
                                        If StrComp(Trim$(tmOldDDFField(ilOldField).sDec), Trim$(tmNewDDFField(ilNewField).sDec), 1) <> 0 Then
                                            ilCompared = False
                                        End If
                                        If tmOldDDFField(ilOldField).iFlags <> tmNewDDFField(ilNewField).iFlags Then
                                            ilCompared = False
                                        End If
                                        Exit For
                                    End If
                                Next ilNewField
                                If ilFound Then
                                    If Not ilCompared Then
                                        lbcOldFile.Selected(ilOldLoop) = True
                                        lbcNewFile.Selected(ilNewLoop) = True
                                        Exit For
                                    End If
                                Else
                                    lbcOldFile.Selected(ilOldLoop) = True
                                    lbcNewFile.Selected(ilNewLoop) = True
                                    Exit For
                                End If
                            Next ilOldField
                            'Any new field
                            For ilNewField = 0 To UBound(tmNewDDFField) - 1 Step 1
                                ilFound = False
                                For ilOldField = 0 To UBound(tmOldDDFField) - 1 Step 1
                                    If StrComp(Trim$(tmNewDDFField(ilNewField).sName), Trim$(tmOldDDFField(ilOldField).sName), 1) = 0 Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilOldField
                                If Not ilFound Then
                                    lbcOldFile.Selected(ilOldLoop) = True
                                    lbcNewFile.Selected(ilNewLoop) = True
                                    Exit For
                                End If
                            Next ilNewField
                        End If
                        Exit For
                    End If
                Next ilFile
            End If
        Next ilNewLoop
    Next ilOldLoop

    If Not imRU Then
        For ilNewLoop = 0 To lbcNewFile.ListCount - 1 Step 1
            slName = lbcNewFile.List(ilNewLoop)
            ilFound = False
            For ilOldLoop = 0 To lbcOldFile.ListCount - 1 Step 1
                If StrComp(Trim$(lbcOldFile.List(ilOldLoop)), slName, 1) = 0 Then
                    ilFound = True
                    Exit For
                End If
            Next ilOldLoop
            If Not ilFound Then
                lbcNewFile.Selected(ilNewLoop) = True
            End If
        Next ilNewLoop
    End If
    Exit Sub
mTestFieldDiffErr:
    ilRet = Err
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestFieldDiff                  *
'*                                                     *
'*             Created:4/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if field changed      *
'*                                                     *
'*******************************************************
Private Sub mTestIndexDiff()
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilOldLoop As Integer
    Dim ilFile As Integer
    Dim ilNewLoop As Integer
    Dim ilOldIndex As Integer
    Dim ilNewIndex As Integer
    Dim ilFound As Integer
    Dim ilCompared As Integer
    Dim slOldFieldName As String
    Dim slNewFieldName As String
    Dim slFileName As String
    Dim slStamp As String
    Dim ilPos As Integer
    Dim ilRet As Integer
    For ilOldLoop = 0 To lbcOldFile.ListCount - 1 Step 1
        slName = Trim$(lbcOldFile.List(ilOldLoop))
        For ilFile = 0 To UBound(tmOldDDFFile) - 1 Step 1
            If StrComp(Trim$(tmOldDDFFile(ilFile).sName), slName, 1) = 0 Then
                mFieldBuild smOldDDFPath, ".DDF", tmOldDDFFile(ilFile).iFileID, tmOldDDFField()
                mIndexBuild smOldDDFPath, ".DDF", tmOldDDFFile(ilFile).iFileID, tmOldDDFIndex()
                Exit For
            End If
        Next ilFile
        For ilNewLoop = 0 To lbcNewFile.ListCount - 1 Step 1
            If StrComp(Trim$(lbcNewFile.List(ilNewLoop)), slName, 1) = 0 Then
                For ilFile = 0 To UBound(tmNewDDFFile) - 1 Step 1
                    If StrComp(Trim$(tmNewDDFFile(ilFile).sName), slName, 1) = 0 Then
                        If imRU Then
                            'Test if file exist
                            slFileName = Trim$(gRemoveZero(tmNewDDFFile(ilFile).sLocation))
                            ilPos = InStr(slFileName, ":")
                            If ilPos > 0 Then
                                slFileName = Mid$(slFileName, ilPos + 1)
                            End If
                            ilPos = InStr(slFileName, "\")
                            Do While ilPos > 0
                                slFileName = Mid$(slFileName, ilPos + 1)
                                ilPos = InStr(slFileName, "\")
                            Loop
                            On Error GoTo mTestIndexDiffErr:
                            ilRet = 0
                            slStamp = FileDateTime(smDBPath & slFileName)
                            If ilRet = 0 Then
                                ilFound = True
                            Else
                                ilFound = False
                            End If
                        Else
                            ilFound = True
                        End If
                        If ilFound Then
                            mFieldBuild smNewDDFPath, ".DDF", tmNewDDFFile(ilFile).iFileID, tmNewDDFField()
                            mIndexBuild smNewDDFPath, ".DDF", tmNewDDFFile(ilFile).iFileID, tmNewDDFIndex()
                            
                            'Compare fields
                            For ilOldIndex = 0 To UBound(tmOldDDFIndex) - 1 Step 1
                                'Get field name
                                For ilLoop = 0 To UBound(tmOldDDFField) - 1 Step 1
                                    If tmOldDDFField(ilLoop).iFieldID = tmOldDDFIndex(ilOldIndex).iFieldID Then
                                        slOldFieldName = Trim$(tmOldDDFField(ilLoop).sName)
                                        Exit For
                                    End If
                                Next ilLoop
                                ilCompared = True
                                ilFound = False
                                For ilNewIndex = 0 To UBound(tmNewDDFIndex) - 1 Step 1
                                    'Get field name
                                    For ilLoop = 0 To UBound(tmNewDDFField) - 1 Step 1
                                        If tmNewDDFField(ilLoop).iFieldID = tmNewDDFIndex(ilNewIndex).iFieldID Then
                                            slNewFieldName = Trim$(tmNewDDFField(ilLoop).sName)
                                            Exit For
                                        End If
                                    Next ilLoop
    
                                    If (StrComp(slOldFieldName, slNewFieldName, 1) = 0) And (tmOldDDFIndex(ilOldIndex).iNumber = tmNewDDFIndex(ilNewIndex).iNumber) Then
                                        ilFound = True
                                        If tmOldDDFIndex(ilOldIndex).iPart <> tmNewDDFIndex(ilNewIndex).iPart Then
                                            ilCompared = False
                                        End If
                                        If tmOldDDFIndex(ilOldIndex).iFlag <> tmNewDDFIndex(ilNewIndex).iFlag Then
                                            ilCompared = False
                                        End If
                                        Exit For
                                    End If
                                Next ilNewIndex
                                If ilFound Then
                                    If Not ilCompared Then
                                        lbcOldFile.Selected(ilOldLoop) = True
                                        lbcNewFile.Selected(ilNewLoop) = True
                                        Exit For
                                    End If
                                Else
                                    lbcOldFile.Selected(ilOldLoop) = True
                                    lbcNewFile.Selected(ilNewLoop) = True
                                    Exit For
                                End If
                            Next ilOldIndex
                            'Any new field
                            For ilNewIndex = 0 To UBound(tmNewDDFIndex) - 1 Step 1
                                For ilLoop = 0 To UBound(tmNewDDFField) - 1 Step 1
                                    If tmNewDDFField(ilLoop).iFieldID = tmNewDDFIndex(ilNewIndex).iFieldID Then
                                        slNewFieldName = Trim$(tmNewDDFField(ilLoop).sName)
                                        Exit For
                                    End If
                                Next ilLoop
                                ilFound = False
                                For ilOldIndex = 0 To UBound(tmOldDDFIndex) - 1 Step 1
                                    For ilLoop = 0 To UBound(tmOldDDFField) - 1 Step 1
                                        If tmOldDDFField(ilLoop).iFieldID = tmOldDDFIndex(ilOldIndex).iFieldID Then
                                            slOldFieldName = Trim$(tmOldDDFField(ilLoop).sName)
                                            Exit For
                                        End If
                                    Next ilLoop
                                    
                                    If (StrComp(slNewFieldName, slOldFieldName, 1) = 0) And (tmNewDDFIndex(ilNewIndex).iNumber = tmOldDDFIndex(ilOldIndex).iNumber) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilOldIndex
                                If Not ilFound Then
                                    lbcOldFile.Selected(ilOldLoop) = True
                                    lbcNewFile.Selected(ilNewLoop) = True
                                    Exit For
                                End If
                            Next ilNewIndex
                        End If
                        Exit For
                    End If
                Next ilFile
            End If
        Next ilNewLoop
    Next ilOldLoop
    Exit Sub
mTestIndexDiffErr:
    ilRet = Err
    Resume Next
End Sub

Private Sub rbcExt_Click(Index As Integer, Value As Integer)
End Sub

Private Sub plcFile_Paint()
    plcFile.CurrentX = 0
    plcFile.CurrentY = 0
    plcFile.Print "File"
End Sub

Private Sub mMoveDDFs()
    If StrComp(smNewDDFPath, smOldDDFPath, 1) <> 0 Then
        'Rename DDF's
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'Kill smOldDDFPath & "File.OLD"
        mDeleteFile smOldDDFPath & "File.OLD"
        
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'Name smOldDDFPath & "File.DDF" As smOldDDFPath & "File.OLD"
        mRenameFile smOldDDFPath & "File.DDF", smOldDDFPath & "File.OLD"
        
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'FileCopy smNewDDFPath & "File.DDF", smOldDDFPath & "File.DDF"
        mCopyFile smNewDDFPath & "File.DDF", smOldDDFPath & "File.DDF"
        
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'Kill smOldDDFPath & "Field.OLD"
        mDeleteFile smOldDDFPath & "Field.OLD"
        
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'Name smOldDDFPath & "Field.DDF" As smOldDDFPath & "Field.OLD"
        mRenameFile smOldDDFPath & "Field.DDF", smOldDDFPath & "Field.OLD"
        
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'FileCopy smNewDDFPath & "Field.DDF", smOldDDFPath & "Field.DDF"
        mCopyFile smNewDDFPath & "Field.DDF", smOldDDFPath & "Field.DDF"
       
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'Kill smOldDDFPath & "Index.OLD"
        mDeleteFile smOldDDFPath & "Index.OLD"
        
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'Name smOldDDFPath & "Index.DDF" As smOldDDFPath & "Index.OLD"
        mRenameFile smOldDDFPath & "Index.DDF", smOldDDFPath & "Index.OLD"
        
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'FileCopy smNewDDFPath & "Index.DDF", smOldDDFPath & "Index.DDF"
        mCopyFile smNewDDFPath & "Index.DDF", smOldDDFPath & "Index.DDF"
        
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'Kill smOldDDFPath & "Comment.OLD"
        mDeleteFile smOldDDFPath & "Comment.OLD"
        
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'Name smOldDDFPath & "Comment.DDF" As smOldDDFPath & "Comment.OLD"
        mRenameFile smOldDDFPath & "Comment.DDF", smOldDDFPath & "Comment.OLD"

        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'FileCopy smNewDDFPath & "Comment.DDF", smOldDDFPath & "Comment.DDF"
        mCopyFile smNewDDFPath & "Comment.DDF", smOldDDFPath & "Comment.DDF"
        
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'Kill smOldDDFPath & "FileComments.OLD"
        mDeleteFile smOldDDFPath & "FileComments.OLD"
        
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'Name smOldDDFPath & "FileComments.DDF" As smOldDDFPath & "FileComments.OLD"
        mRenameFile smOldDDFPath & "FileComments.DDF", smOldDDFPath & "FileComments.OLD"
        
        'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
        'FileCopy smNewDDFPath & "FileComments.DDF", smOldDDFPath & "FileComments.DDF"
        mCopyFile smNewDDFPath & "FileComments.DDF", smOldDDFPath & "FileComments.DDF"
        
    End If
End Sub

Private Sub mTestCreateOnly()
    Dim ilCreate As Integer
    Dim ilNewLoop As Integer
    Dim ilFile As Integer
    Dim ilPos As Integer
    Dim slFileName As String
    Dim slName As String
    
    If IsArray(smCreateOnlyArray) Then
        If UBound(smCreateOnlyArray) >= LBound(smCreateOnlyArray) Then
            For ilNewLoop = 0 To lbcNewFile.ListCount - 1 Step 1
                If Not lbcNewFile.Selected(ilNewLoop) Then
                    slName = Trim$(lbcNewFile.List(ilNewLoop))
                    For ilFile = 0 To UBound(tmNewDDFFile) - 1 Step 1
                        If StrComp(Trim$(tmNewDDFFile(ilFile).sName), slName, 1) = 0 Then
                            slFileName = Trim$(gRemoveZero(tmNewDDFFile(ilFile).sLocation))
                            ilPos = InStr(slFileName, ":")
                            If ilPos > 0 Then
                                slFileName = Mid$(slFileName, ilPos + 1)
                            End If
                            ilPos = InStr(slFileName, "\")
                            Do While ilPos > 0
                                slFileName = Mid$(slFileName, ilPos + 1)
                                ilPos = InStr(slFileName, "\")
                            Loop
                            For ilCreate = 0 To UBound(smCreateOnlyArray) Step 1
                                If UCase(slFileName) = UCase(smCreateOnlyArray(ilCreate)) Then
                                    lbcNewFile.Selected(ilNewLoop) = True
                                End If
                            Next ilCreate
                            Exit For
                        End If
                    Next ilFile
                End If
            Next ilNewLoop
        End If
    End If
End Sub

'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
Function mRenameFile(sOldFile, sNewFile) As Boolean
    On Error GoTo RenameError
    lacRename.Caption = "Renaming: " & sOldFile & " To " & sNewFile
    lacRename.Refresh
    Print #hmTo, "  Renaming: " & sOldFile & " To " & sNewFile
    Dim slStamp As String
    
    If mFileExists(sOldFile) Then
        'TTP 10334 - last file to get processed doesn't remove the _.NEW file
        Sleep 100
        slStamp = FileDateTime(sOldFile)
        If DateDiff("S", slStamp, Now()) < 5 Then
            Sleep 3000
        End If
        If Not mFileExists(sNewFile) Then
            Name sOldFile As sNewFile
            Print #hmTo, "   -> Renaming: Successfully Renamed: " & sOldFile & " To " & sNewFile
        Else
            Sleep 100
            Print #hmTo, "   -> Renaming Error: Target File already Exists: " & sNewFile
            mDeleteFile (sNewFile)
            
            Name sOldFile As sNewFile
            Print #hmTo, "   -> Renaming: Successfully Renamed: " & sOldFile & " To " & sNewFile
        End If
    Else
        Print #hmTo, "   -> Renaming: File Does not exit to rename: " & sOldFile
    End If
    
    mRenameFile = True
    On Error GoTo 0
    Exit Function

RenameError:
    mRenameFile = False
    bErrorFound = True
    Print #hmTo, "   -> Error # " & Err & " - " & Error(Err) & "; Renaming: " & sOldFile & " To " & sNewFile
    On Error GoTo 0
End Function

'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
Function mDeleteFile(sFileName) As Boolean
    'On Error GoTo DeleteError
    On Error Resume Next
    Dim slStamp As String
    Dim ilTry As Integer
    Dim lbSuccess As Boolean
    
    lacRename.Caption = "Deleting: " & sFileName
    lacRename.Refresh
    Print #hmTo, "  Deleting: " & sFileName
    ilTry = 0
    If mFileExists(sFileName) Then
        'TTP 10334 - last file to get processed doesn't remove the _.NEW file
        'TTP 10650 - DDF Reorg: errors during processing when .OLD files are in the Data folder (When Read-Only .OLD files are in the data folder)
        Do
            mRemoveReadOnlyAttr sFileName
            If ilTry > 0 Then Sleep 1000
            slStamp = FileDateTime(sFileName)
            If DateDiff("S", slStamp, Now()) < 5 Then
                Sleep 500
            End If
            Err = 0
            Kill sFileName
            If Err <> 0 Then
                ilTry = ilTry + 1
            Else
                lbSuccess = True
            End If
        Loop While lbSuccess = False And ilTry < 3
        If ilTry >= 3 Then
            GoTo DeleteError
        End If
        Print #hmTo, "   -> Successfully Deleted: " & sFileName
    Else
        Print #hmTo, "   -> File to Delete doesnt exist: " & sFileName
    End If
    mDeleteFile = True
    On Error GoTo 0
    Exit Function
    
DeleteError:
    mDeleteFile = False
    bErrorFound = True
    Print #hmTo, "   -> Error # " & Err & " - " & Error(Err) & "; Deleting: "; sFileName
    On Error GoTo 0
End Function

'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
Function mCopyFile(sOldFile, sNewFile) As Boolean
    On Error GoTo CopyError
    lacRename.Caption = "Copy: " & sOldFile & " To " & sNewFile
    lacRename.Refresh
    Print #hmTo, "  Copy: " & sOldFile & " To " & sNewFile
    
    If mFileExists(sNewFile) Then
        mDeleteFile (sNewFile)
        'TTP 10334 - last file to get processed doesn't remove the _.NEW file
        Sleep 500
    End If
    
    FileCopy sOldFile, sNewFile
    mCopyFile = True
    Print #hmTo, "   -> Succesfully Copied:" & sOldFile & " To " & sNewFile
    On Error GoTo 0
    Exit Function
    
CopyError:
    mCopyFile = False
    bErrorFound = True
    Print #hmTo, "   -> Error # " & Err & " - " & Error(Err) & "; Copying: " & sOldFile & " To " & sNewFile
    On Error GoTo 0
End Function

'TTP 10334 - 11/1/21 - JW - Error 75 during re-org
Public Function mFileExists(ByVal Fname As String) As Boolean
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        mFileExists = True
    Else
        mFileExists = False
    End If
    On Error GoTo 0
    Exit Function

FileExistsError:
    mFileExists = False
    bErrorFound = True
    Print #hmTo, "   -> Error # " & Err & " - " & Error(Err) & "; Checking File Exist: " & Fname
    On Error GoTo 0
End Function

'TTP 10650 - DDF Reorg: errors during processing when .OLD files are in the Data folder
Public Function mRemoveReadOnlyAttr(ByVal Fname As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fl As File
    If fso.FileExists(Fname) Then
        Set fl = fso.GetFile(Fname)
        If (fl.Attributes And ReadOnly) Then
          fl.Attributes = fl.Attributes - ReadOnly
          Print #hmTo, "   -> Removed 'Read Only' Attribute: " & Fname
        End If
    End If
    Set fso = Nothing
    On Error GoTo 0
End Function
