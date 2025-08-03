VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.ocx"
Begin VB.Form frmImportUpdateStations 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Aired Spots"
   ClientHeight    =   5805
   ClientLeft      =   4440
   ClientTop       =   4920
   ClientWidth     =   9525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkDisableCleanup 
      Caption         =   "DISABLE cleanup of unused DMA/MSA"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1605
      Width           =   3465
   End
   Begin VB.CheckBox ckcHawaiiToPacific 
      Caption         =   "Convert Hawaii Time Zone to Pacific"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.CheckBox ckcAlaskaToPacific 
      Caption         =   "Convert Alaska Time Zone to Pacific"
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Top             =   555
      Value           =   1  'Checked
      Width           =   2955
   End
   Begin VB.CheckBox ckcMissingStations 
      Caption         =   "Report Import Stations not in System"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1335
      Width           =   3255
   End
   Begin VB.CheckBox chkReportStationsNotUpdated 
      Caption         =   "Report stations not included in the Import file."
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1065
      Width           =   3720
   End
   Begin VB.TextBox txtFile 
      Height          =   300
      Left            =   990
      TabIndex        =   5
      Top             =   150
      Width           =   6285
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse"
      Height          =   300
      Left            =   7425
      TabIndex        =   4
      Top             =   150
      Width           =   1065
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   2595
      Width           =   9135
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   5250
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   5805
      FormDesignWidth =   9525
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   2745
      TabIndex        =   2
      Top             =   5295
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4860
      TabIndex        =   3
      Top             =   5295
      Width           =   1575
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   210
      Left            =   2985
      TabIndex        =   7
      Top             =   4950
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CheckBox ckcChangeToInvalidIfLetterReassigned 
      Caption         =   "Allow import to change call letters to ""invalid"" if call letters are being reassigned to a different station ID"
      Height          =   435
      Left            =   4200
      TabIndex        =   14
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label lacBands 
      Caption         =   "Allowed Bands:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   1935
      Width           =   7620
   End
   Begin VB.Label lbcFile 
      Caption         =   "Import File"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   165
      Width           =   780
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   195
      TabIndex        =   0
      Top             =   2235
      Width           =   9045
   End
End
Attribute VB_Name = "frmImportUpdateStations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmImportUpdateStations
'*
'*  Created August 14, 2006 by Jeff Dutschke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text

Private imImporting As Integer
Private bmImported As Boolean
Private imTerminate As Integer
Private hmFrom As Integer
Private lmTotalRecords As Long
Private lmProcessedRecords As Long
Private lmPercent As Long
Private bmMatchOnPermStationID As Boolean
Private bmStationPreviouslyDefined As Boolean
Private tmUpdateStation() As UPDATESTATION
Private tmSvUpdateStation() As UPDATESTATION
Private smPass0LinesBypassed() As String
Private tmFormatLinkInfo() As FORMATLINKINFO

Private tmReportInfo() As BIAREPORTINFO
Private tmUsedMarkets() As MARKETINFO
Private tmUsedMSAMarkets() As MARKETINFO
Private tmUsedOwners() As OWNERINFO
Private tmUsedFormats() As FORMATINFO
Private tmMonikerInfo() As MNTINFO
Private tmCityInfo() As MNTINFO
Private tmCountyInfo() As MNTINFO
Private tmTerritoryInfo() As MNTINFO
Private tmAreaInfo() As MNTINFO
Private tmOperatorInfo() As MNTINFO
Private smCallLetters As String
'Private smBand As String
'Private smMarketName As String
'Private smRank As String
'Private smOwnerName As String
'Private smFormat As String
Private bmUpdateDatabase As Boolean
Private smReportPathFileName As String
Private imMap(0 To 89) As Integer
Private bmIgnoreBlanks(0 To 89) As Boolean

Private lmAttCode() As Long

Private bmAdjPledge As Boolean
Private bmNotAddedMsg As Boolean
Private smTimeZone As String
Private imTztCode As Integer
Private smMailCity As String
Private smCityLic As String
Private lmCityLicMntCode As Long
Private smCountyLic As String
Private lmCountyLicMntCode As Long
Private smMailState As String
Private smMoniker As String
Private lmMonikerMntCode As Long
Private lmDMAMktCode As Long
Private lmDMAMktIdx As Long
Private smDMAMarket As String
Private lmMSAMktCode As Long
Private lmMSAMktIdx As Long
Private smMSAMarket As String
Private lmOwnerCode As Long
Private lmOwnerIdx As Long
Private smOwner As String
Private imFormatCode As Integer
Private lmFormatIdx As Long
Private smFormat As String
Private lmCityMntCode As Long
Private smTerritory As String
Private lmTerritoryMntCode As Long
Private smArea As String
Private lmAreaMntCode As Long
Private smStateLic As String
Private imDaylight As Integer   '0=Yes; 1=No
Private smUsedAgreement As String * 1
Private smUsedXDS As String * 1
Private smUsedWegener As String * 1
Private smUsedOLA As String * 1
Private smOperator As String
Private lmOperatorMntCode As Long
Private smMarketRep As String
Private imMarketRepUstCode As Integer
Private smServiceRep As String
Private imServiceRepUstCode As Integer
Private smOnAir As String * 1
Private smCommercial As String * 1
Private smHistoricalDate As String
Private lmPhysicalCityMntCode As Long
Private smPhysicalCity As String
Private smPhysicalState As String
Private imPersonTitles As Integer

Private smBandFields() As String



Private rst_Shtt As ADODB.Recordset
Private rst_artt As ADODB.Recordset
Private rst_cmt As ADODB.Recordset

Const CALLLETTERS = 1
Const ID = 2
Const FREQUENCY = 3
Const TERRITORY = 4
Const AREA = 5
Const STATIONFORMAT = 6
Const DMARANK = 7
Const DMANAME = 8
Const CITYLIC = 9
Const COUNTYLIC = 10
Const STATELIC = 11
Const OWNER = 12
Const OPERATOR = 13
Const MSARANK = 14
Const MSANAME = 15
Const MARKETREP = 16
Const SERVICEREP = 17
Const ZONE = 18
Const OnAir = 19
Const COMMERCIAL = 20
Const DAYLIGHT = 21
Const XDSID = 22
Const IPUMPID = 23
Const SERIAL1 = 24
Const SERIAL2 = 25
Const USEAGREEMENT = 26
Const USEXDS = 27
Const USEWEGENER = 28
Const USEOLA = 29
Const MONIKER = 30
Const WATTS = 31
Const HISTORICALDATE = 32
Const P12PLUS = 33
Const WEBADDR = 34
Const WEBPW = 35
Const ENTERPRISEID = 36
Const MAILADDR1 = 37
Const MAILADDR2 = 38
Const MAILCITY = 39
Const MAILSTATE = 40
Const MAILZIP = 41
Const MAILCOUNTRY = 42
Const PHYSICALADDR1 = 43
Const PHYSICALADDR2 = 44
Const PHYSICALCITY = 45
Const PHYSICALSTATE = 46
Const PHYSICALZIP = 47
Const PHONE = 48
Const FAX = 49
Const PERSON = 50
Const PNAME = 0
Const PTITLE = 1
Const PPHONE = 2
Const PFAX = 3
Const PEMAIL = 4
Const PAFFLABEL = 5
Const PISCIEXPORT = 6
Const PAFFEMAIL = 7
Const SHTTINDEX = 0

'***************************************************************************
'
'***************************************************************************
Private Sub cmdImport_Click()
    Dim iLoop As Integer
    Dim sFileName As String
    Dim iRet As Integer
    Dim iVef As Integer
    Dim iZone As Integer
    Dim sToFile As String
    Dim sDateTime As String
    Dim sMsgFileName As String
    Dim sMoDate As String
    Dim ilPass As Integer
    Dim llLoop As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    lbcMsg.Clear
    bmUpdateDatabase = True
    bmNotAddedMsg = False
    'If chkReportOnly.Value Then
    '    ' User wants to view the report without actually making the changes to the database.
    '    bmUpdateDatabase = False
    '    gLogMsg "SHOW REPORT ONLY. NO CHANGES ARE BEING MADE TO THE DATABASE.", "AffErrorLog.Txt", False
    'End If
    Kill smReportPathFileName
    ReDim tmReportInfo(0 To 0) As BIAREPORTINFO
    ReDim tmUsedMarkets(0 To 0) As MARKETINFO
    ReDim tmUsedMSAMarkets(0 To 0) As MARKETINFO
    ReDim tmUsedOwners(0 To 0) As OWNERINFO
    ReDim tmUsedFormats(0 To 0) As FORMATINFO

    Screen.MousePointer = vbHourglass
    If sgUsingStationID = "Y" Then
        bmMatchOnPermStationID = True
    Else
        bmMatchOnPermStationID = False
        For iLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(iLoop).lPermStationID <> 0 Then
                Screen.MousePointer = vbDefault
                Call mFailImportMsg("Terminated - Station ID exist but Site is Not Set")
                Exit Sub
            End If
        Next iLoop
    End If
    If UBound(tgStationInfo) <= LBound(tgStationInfo) Then
        bmStationPreviouslyDefined = False
    Else
        bmStationPreviouslyDefined = True
    End If
    
    If Not gPopMarkets() Then
        Screen.MousePointer = vbDefault
        Call mFailImportMsg("Unable to Load Existing DMA Market Names.")
        Exit Sub
    End If
    
    If Not gPopMSAMarkets() Then
        Screen.MousePointer = vbDefault
        Call mFailImportMsg("Unable to Load Existing MSA Market Names.")
        Exit Sub
    End If
    
    If Not gPopOwnerNames() Then
        Screen.MousePointer = vbDefault
        Call mFailImportMsg("Unable to Load Existing Owner Names.")
        Exit Sub
    End If

    If Not gPopFormats() Then
        Screen.MousePointer = vbDefault
        Call mFailImportMsg("Unable to Load Existing Format Names.")
        Exit Sub
    End If

    If Not gPopStations() Then
        Screen.MousePointer = vbDefault
        Call mFailImportMsg("Unable to Load Existing Station Names.")
        Exit Sub
    End If
    If Not mPopFormatLink() Then
        Screen.MousePointer = vbDefault
        Call mFailImportMsg("Unable to Load Existing External Format Names.")
        Exit Sub
    End If
    
    If Not gPopTimeZones() Then
        Screen.MousePointer = vbDefault
        Call mFailImportMsg("Unable to Load Time Zone Names.")
        Exit Sub
    End If
    
    If Not gPopTitleNames() Then
        Screen.MousePointer = vbDefault
        Call mFailImportMsg("Unable to Load Title Names.")
        Exit Sub
    End If
    
    gPopMntInfo "M", tmMonikerInfo()
    gPopMntInfo "C", tmCityInfo()
    gPopMntInfo "Y", tmCountyInfo()
    gPopMntInfo "T", tmTerritoryInfo()
    gPopMntInfo "A", tmAreaInfo()
    gPopMntInfo "O", tmOperatorInfo()
    
    gPopRepInfo "M", tgMarketRepInfo()
    gPopRepInfo "S", tgServiceRepInfo()
    
    imImporting = True
    bmImported = True
    cmcBrowse.Enabled = False
    cmdImport.Enabled = False
    txtFile.Enabled = False
    chkReportStationsNotUpdated.Enabled = False
    cmdCancel.Caption = "Abort"
    '6/11/12:  Changed message to reflect what is truely happening
    'mSetResults "Importing Update Station Information", RGB(0, 0, 0)
    If sgUsingStationID = "Y" Then
        mSetResults "Importing Station Adds and Updates", RGB(0, 0, 0)
    ElseIf sgUsingStationID = "A" Then
        If UBound(tgStationInfo) > LBound(tgStationInfo) Then
            mSetResults "Importing Additinal Station Adds", RGB(0, 0, 0)
        Else
            mSetResults "Importing Inital Station Adds", RGB(0, 0, 0)
        End If
    Else
        If UBound(tgStationInfo) > LBound(tgStationInfo) Then
            mSetResults "Importing Station Updates", RGB(0, 0, 0)
        Else
            mSetResults "Importing Inital Station Adds", RGB(0, 0, 0)
        End If
    End If

    '9/15/14: Add second pass to process those records that match call letters of a different station
    For ilPass = 0 To 1 Step 1
        iRet = mReadStationImportFile(ilPass)
        If iRet = False Then
            Screen.MousePointer = vbDefault
            Call mFailImportMsg("Terminated - mReadStationImportFile returned False")
            Exit Sub
        End If
        
        If Not mCheckFormats() Then
            Screen.MousePointer = vbDefault
            Call mFailImportMsg("Import Terminated by User")
            Exit Sub
        End If
        
        If Not mCheckDMAMarkets() Then
            Screen.MousePointer = vbDefault
            Call mFailImportMsg("Import Terminated by User")
            Exit Sub
        End If
        
        If Not mCheckMSAMarkets() Then
            Screen.MousePointer = vbDefault
            Call mFailImportMsg("Import Terminated by User")
            Exit Sub
        End If
        
        If Not mCheckOwners() Then
            Screen.MousePointer = vbDefault
            Call mFailImportMsg("Import Terminated by User")
            Exit Sub
        End If
        
        If bmUpdateDatabase Then
            If Not mRemoveDuplicateMarketNames() Then
                Screen.MousePointer = vbDefault
                Call mFailImportMsg("Unable to Remove Dupliacte Market Names.")
                Exit Sub
            End If
            
            If Not gPopMarkets() Then
                Screen.MousePointer = vbDefault
                Call mFailImportMsg("Unable to Load Existing Market Names.")
                Exit Sub
            End If
        End If
        'If UBound(tgMarketInfo) > 0 Then
        '    ArraySortTyp fnAV(tgMarketInfo(), 0), UBound(tgMarketInfo), 0, LenB(tgMarketInfo(0)), 2, LenB(tgMarketInfo(0).sName), 1
        'End If
    
        If lmTotalRecords > 0 Then
            plcGauge.Value = 0
            plcGauge.Visible = True
        End If
        lmPercent = 0
        lmProcessedRecords = 0
        If (bmMatchOnPermStationID) Then
            mSetResults "Processing Import and bypassing Duplicated Call Letters.", RGB(0, 0, 0)
        Else
            mSetResults "Processing Duplicated Call Letters.", RGB(0, 0, 0)
        End If
        iRet = mProcessInfo()
        plcGauge.Visible = False
        If imTerminate Then
            Screen.MousePointer = vbDefault
            Call mFailImportMsg("User Terminated")
            Exit Sub
        End If
        If iRet = False Then
            Screen.MousePointer = vbDefault
            Call mFailImportMsg("Terminated - mProcessInfo returned False")
            Exit Sub
        End If
        If Not gPopMarkets() Then
            Screen.MousePointer = vbDefault
            Call mFailImportMsg("Unable to Load Existing DMA Market Names.")
            Exit Sub
        End If
        If Not gPopMSAMarkets() Then
            Screen.MousePointer = vbDefault
            Call mFailImportMsg("Unable to Load Existing MSA Market Names.")
            Exit Sub
        End If
        If Not gPopOwnerNames() Then
            Screen.MousePointer = vbDefault
            Call mFailImportMsg("Unable to Load Existing Owner Names.")
            Exit Sub
        End If
        If Not gPopFormats() Then
            Screen.MousePointer = vbDefault
            Call mFailImportMsg("Unable to Load Existing Format Names.")
            Exit Sub
        End If
        'Force re-populate
        sgShttTimeStamp = ""
        '11/26/17: Set Changed date/time
        gFileChgdUpdate "shtt.mkd", True
        If Not gPopStations() Then
            Screen.MousePointer = vbDefault
            Call mFailImportMsg("Unable to Load Existing Station Names.")
            Exit Sub
        End If
        If (bmMatchOnPermStationID) And (((UBound(smPass0LinesBypassed) > 0) And (ilPass = 0)) Or (ilPass = 1)) Then
            If ilPass = 0 Then
                ReDim tmSvUpdateStation(0 To UBound(tmUpdateStation)) As UPDATESTATION
                For llLoop = 0 To UBound(tmUpdateStation) - 1 Step 1
                    tmSvUpdateStation(llLoop) = tmUpdateStation(llLoop)
                Next llLoop
            Else
                For llLoop = 0 To UBound(tmUpdateStation) - 1 Step 1
                    tmSvUpdateStation(UBound(tmSvUpdateStation)) = tmUpdateStation(llLoop)
                    ReDim Preserve tmSvUpdateStation(0 To UBound(tmSvUpdateStation) + 1) As UPDATESTATION
                Next llLoop
                ReDim tmUpdateStation(0 To UBound(tmSvUpdateStation)) As UPDATESTATION
                For llLoop = 0 To UBound(tmSvUpdateStation) - 1 Step 1
                    tmUpdateStation(llLoop) = tmSvUpdateStation(llLoop)
                Next llLoop
                lmTotalRecords = UBound(tmUpdateStation)
                If UBound(tmUpdateStation) > 0 Then
                    'Sort by call letters, offset is two
                    ArraySortTyp fnAV(tmUpdateStation(), 0), UBound(tmUpdateStation), 0, LenB(tmUpdateStation(0)), 2, LenB(tmUpdateStation(0).sCallLetters), 0
                End If
            End If
        Else
            Exit For
        End If
    Next ilPass
    ' Call TEST_FixBadStationNames
    iRet = mSetUsedArrays()
    
   ' TTP 10592 JJB 2023-04-27
    If chkDisableCleanup.Value = vbUnchecked Then
        Call mRemoveUnusedDMAMarkets
        Call mRemoveUnusedMSAMarkets
    End If
    
    Call mRemoveUnusedOwners
    Call mRemoveUnusedFormats

    If UBound(tmReportInfo) < 1 Then
        mSetResults "All station information was 100% up to date. No records were changed.", RGB(0, 0, 0)
    End If

    'fix v81 TTP 10984 per Jason email: Fri 4/19/24 11:13 AM
    'The Global list of Station names needs to be reloaded
    sgShttTimeStamp = ""
    tgFctChgdInfo(SHTTINDEX).lLastDateChgd = 0
    ilRet = gPopStations()
    
    If chkReportStationsNotUpdated.Value = vbChecked Then
        Call mReportStationsNotUpdated     ' Not found in the BIA file.
    End If

    Call mCreateStatusReport

    If imTerminate Then
        Screen.MousePointer = vbDefault
        Call mFailImportMsg("User Terminated")
        Exit Sub
    End If
    imImporting = False
    mSetResults "Operation Completed Successfully.", RGB(0, 200, 0)
    
    Screen.MousePointer = vbDefault
    mShowReport
    cmdCancel.Caption = "&Done"

    Exit Sub

ErrHand:
    Resume Next
End Sub

'***************************************************************************
'
'***************************************************************************
Private Function mReadStationImportFile(ilPass As Integer) As Integer
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim slCallLetters As String
    Dim slChar As String
    Dim slFields(0 To 89) As String
    Dim llMaxRecords As Long
    Dim ilLoop As Integer
    Dim blTitleLineFound As Boolean
    Dim llIndex As Long
    Dim ilPerson As Integer
    Dim ilUpdateType As Integer '0=Add; 1=Update, 2=Bypass
    Dim llLoop As Long
    Dim llStationID As Long
    Dim llLineIndex As Long
    
    On Error GoTo mReadStationImportFileErr:
    mReadStationImportFile = False
    lmTotalRecords = 0
    llMaxRecords = 10000
    ReDim tmUpdateStation(0 To llMaxRecords) As UPDATESTATION
    If ilPass = 0 Then
        ReDim smPass0LinesBypassed(0 To 0) As String
        lbcMsg.Clear
        blTitleLineFound = False
        mSetResults "Importing Update Station Information...", RGB(0, 0, 0)
        slFromFile = txtFile.Text
        'ilRet = 0
        'hmFrom = FreeFile
        'Open slFromFile For Input Access Read As hmFrom
        ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
        If ilRet <> 0 Then
            mSetResults "Unable to open file. Error = " & Trim$(Str$(ilRet)), RGB(255, 0, 0)
            Exit Function
        End If
        For ilLoop = LBound(imMap) To UBound(imMap) Step 1
            imMap(ilLoop) = 0
            bmIgnoreBlanks(ilLoop) = True
            slFields(ilLoop) = ""
        Next ilLoop
        
        Do While Not EOF(hmFrom)
            ilRet = 0
            'Line Input #hmFrom, slLine
            slLine = ""
            Do While Not EOF(hmFrom)
                slChar = Input(1, #hmFrom)
                If slChar = sgLF Then
                    Exit Do
                ElseIf slChar <> sgCR Then
                    slLine = slLine & slChar
                End If
            Loop
            If ilRet <> 0 Then
                mSetResults "Unable to read from file. Error = " & Trim$(Str$(ilRet)), RGB(255, 0, 0)
                Exit Function
            End If
            
            gParseCDFields slLine, False, slFields()
            For llLoop = UBound(slFields) - 1 To LBound(slFields) Step -1
                slFields(llLoop + 1) = slFields(llLoop)
            Next llLoop
            slFields(0) = ""
            If (Not blTitleLineFound) Then
                If (StrComp(Trim$(slFields(CALLLETTERS)), "Call Letters", vbTextCompare) = 0) Then
                    If Not mCreateMap(slFields()) Then
                        MsgBox "Field Names don't match the required names"
                        Exit Function
                    End If
                    'Test for required titles: DMA; Time Zone; and maybe Perm ID
                    If (imMap(DMANAME) <= 0) Then
                        MsgBox "Required column 'DMA Name' missing, Import terminated"
                        Exit Function
                    End If
                    If (imMap(ZONE) <= 0) Then
                        MsgBox "Required column 'Zone' missing, Import terminated"
                        Exit Function
                    End If
                    If bmMatchOnPermStationID And (imMap(ID) <= 0) Then
                        MsgBox "Required column 'ID#' missing, Import terminated"
                        Exit Function
                    End If
                    blTitleLineFound = True
                End If
            Else
                If Trim(slLine) <> "" Then
                    'Debug.Print "process Line:" & slLine
                    ilRet = mProcessLine(ilPass, llMaxRecords, slLine, slFields())
                End If
            End If
        Loop
        Close hmFrom
    Else
        For llLineIndex = 0 To UBound(smPass0LinesBypassed) - 1 Step 1
            slLine = smPass0LinesBypassed(llLineIndex)
            gParseCDFields slLine, False, slFields()
            For llLoop = UBound(slFields) - 1 To LBound(slFields) Step -1
                slFields(llLoop + 1) = slFields(llLoop)
            Next llLoop
            slFields(0) = ""
            'Debug.Print "process Line:" & llLineIndex + 1 & " - " & slLine
            ilRet = mProcessLine(ilPass, llMaxRecords, slLine, slFields())
        Next llLineIndex
    End If
    If lmTotalRecords <> 0 Then
        ReDim Preserve tmUpdateStation(0 To lmTotalRecords) As UPDATESTATION
    
        ' Sort the array.
        If UBound(tmUpdateStation) > 0 Then
            'Sort by call letters, offset is two
            ArraySortTyp fnAV(tmUpdateStation(), 0), UBound(tmUpdateStation), 0, LenB(tmUpdateStation(0)), 2, LenB(tmUpdateStation(0).sCallLetters), 0
        End If
    Else
        ReDim Preserve tmUpdateStation(0 To 0) As UPDATESTATION
    End If
    mSetResults Trim(Str(lmTotalRecords)) & " records imported.", RGB(0, 0, 0)
    mReadStationImportFile = True
    Exit Function
mReadStationImportFileErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mReadStationImportFile"
    mReadStationImportFile = False
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mProcessInfo() As Integer
    Dim llUpdateStationIndex As Long
    Dim llStationIDX As Long
    Dim llMktIdx As Long
    Dim llOwnerIdx As Long
    Dim llFormatIdx As Long
    Dim llIdx As Long
    Dim slCallLetters As String
    Dim llStationID As Long
    Dim ilRet As Integer
    Dim ilOldRank As Integer
    Dim sOriginalOwnerName As String
    Dim llTempIdx As Long
    Dim ilUpdateType As Integer '0=Add; 1=Update, 2=Bypass
    Dim llLoop As Long
    
    On Error GoTo ErrHandler:
    mProcessInfo = False
    
    ' Process all records loaded in the BIA list
    mSetResults "Processing Station Update information.", RGB(0, 0, 0)
    For llUpdateStationIndex = 0 To lmTotalRecords - 1
        DoEvents
        If imTerminate Then
            Exit Function
        End If
'        If llUpdateStationIndex > 0 And llUpdateStationIndex Mod 1000 = 0 Then
'            mSetResults Trim(Str(llUpdateStationIndex)) & " records processed.", RGB(0, 0, 0)
'
''            mProcessInfo = True
''            Exit Function
'        End If
         If lmTotalRecords > 0 Then
            lmProcessedRecords = lmProcessedRecords + 1
            lmPercent = (lmProcessedRecords * CSng(100)) / lmTotalRecords
            If lmPercent >= 100 Then
                If lmProcessedRecords + 1 < lmTotalRecords Then
                    lmPercent = 99
                Else
                    lmPercent = 100
                End If
            End If
            If plcGauge.Value <> lmPercent Then
                plcGauge.Value = lmPercent
                DoEvents
            End If
        End If
        slCallLetters = Trim(tmUpdateStation(llUpdateStationIndex).sCallLetters)
        llStationID = tmUpdateStation(llUpdateStationIndex).lID
        ' Find this station in the stations array. If not found then ignore this BIA record.
        llStationIDX = mLookupStation(llStationID, slCallLetters)
'Moved ro mReadStationImportFile
'        If llStationIDX <> -1 Then
'            ilUpdateType = 1
'            If bmMatchOnPermStationID Then
'                For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
'                    If StrComp(UCase$(Trim(tgStationInfo(llLoop).sCallLetters)), UCase$(Trim(slCallLetters)), vbTextCompare) = 0 Then
'                        If llStationIDX <> llLoop Then
'                            ilUpdateType = 2
'                            Call mUpdateReport(-1, "WARNING: Import Station previously defined within System: " & Trim(slCallLetters) & " ID " & llStationID)
'                        End If
'                        Exit For
'                    End If
'                Next llLoop
'            End If
'        Else
'            If bmMatchOnPermStationID Then
'                ilUpdateType = 0
'                For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
'                    If StrComp(UCase$(Trim(tgStationInfo(llLoop).sCallLetters)), UCase$(Trim(slCallLetters)), vbTextCompare) = 0 Then
'                        If tgStationInfo(llLoop).lPermStationID = 0 Then
'                            ilUpdateType = 1
'                            llStationIDX = llLoop
'                        Else
'                            ilUpdateType = 2
'                            Call mUpdateReport(-1, "WARNING: Import Station previously defined within System: " & Trim(slCallLetters) & " ID " & llStationID)
'                        End If
'                        Exit For
'                    End If
'                Next llLoop
'            Else
'                'Update only if not stations existed
'                If Not bmStationPreviouslyDefined Then
'                    ilUpdateType = 0
'                Else
'                    ilUpdateType = 2
'                    Call mUpdateReport(-1, "WARNING: Import Station not in System: " & Trim(slCallLetters))
'                End If
'            End If
'        End If
        If tmUpdateStation(llUpdateStationIndex).iCode > 0 Then
            ilUpdateType = 2
            For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                If tgStationInfo(llLoop).iCode = tmUpdateStation(llUpdateStationIndex).iCode Then
                    ilUpdateType = 1
                    llStationIDX = llLoop
                    Exit For
                End If
            Next llLoop
        Else
            ilUpdateType = 0
        End If
        If ilUpdateType = 0 Then
            If Not mAddStation(llUpdateStationIndex) Then
                Call mUpdateReport(llUpdateStationIndex, "WARNING: Station not Added. (" & Trim(slCallLetters) & ")")
                If Not bmNotAddedMsg Then
                    mSetResults "Some Stations not Added", RGB(0, 0, 0)
                    bmNotAddedMsg = True
                End If
            Else
                Call mUpdateReport(llUpdateStationIndex, "Station Added. (" & Trim(slCallLetters) & ")")
            End If
        ElseIf ilUpdateType = 1 Then
            If bmMatchOnPermStationID Then
                If StrComp(Trim(tgStationInfo(llStationIDX).sCallLetters), Trim(slCallLetters), vbTextCompare) <> 0 Then
                    'Update Call letters and place Original call letters into history
                    If Not mUpdateCallLetters(llStationIDX, slCallLetters, llUpdateStationIndex) Then
                        Exit Function
                    End If
                End If
            End If
            If Not mUpdateStationInfo(llStationIDX, llUpdateStationIndex) Then
                Call mUpdateReport(llUpdateStationIndex, "WARNING: Station not Updated. (" & Trim(slCallLetters) & ")")
            End If

        End If
    Next
    mSetResults Trim(Str(lmTotalRecords)) & " Records Processed.", RGB(0, 0, 0)

    mProcessInfo = True
    Exit Function
    
ErrHandler:
    Screen.MousePointer = vbDefault
    gMsg = "A general error has occured in mProcessInfo: "
    gLogMsg "A general error has occured in mProcessInfo: ", "AffErrorLog.Txt", False
    gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mUpdateStationsMarket(lStationIDX As Long, lMktIdx As Long, lUpdateStationIndex As Long) As Integer
    Dim NewIdx As Long
    Dim slSQLQuery As String
    
    On Error GoTo ErrHandler:
    mUpdateStationsMarket = False
    tgStationInfo(lStationIDX).iMktCode = tgMarketInfo(lMktIdx).lCode
    slSQLQuery = "Update shtt Set shttMktCode = " & tgStationInfo(lStationIDX).iMktCode & " Where shttCode = " & tgStationInfo(lStationIDX).iCode
    If bmUpdateDatabase Then
        'cnn.Execute slSQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHandler:
            gHandleError "AffErrorLog.txt", "ImportUpdateStations-mUpdateStationsMarket"
            mUpdateStationsMarket = False
            Exit Function
        End If
    End If

    Call mUpdateReport(lUpdateStationIndex, "DMA Market was assigned (" & Trim(tmUpdateStation(lUpdateStationIndex).sDMAMarket) & ")")
    mUpdateStationsMarket = True
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mUpdateStationsMarket"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mUpdateStationsMSAMarket(lStationIDX As Long, lMktIdx As Long, lUpdateStationIndex As Long) As Integer
    Dim NewIdx As Long
    Dim slSQLQuery As String
    
    On Error GoTo ErrHandler:
    mUpdateStationsMSAMarket = False
    tgStationInfo(lStationIDX).iMSAMktCode = tgMSAMarketInfo(lMktIdx).lCode
    slSQLQuery = "Update shtt Set shttMetCode = " & tgStationInfo(lStationIDX).iMSAMktCode & " Where shttCode = " & tgStationInfo(lStationIDX).iCode
    If bmUpdateDatabase Then
        'cnn.Execute slSQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHandler:
            gHandleError "AffErrorLog.txt", "ImportUpdateStations-mUpdateStationsMSAMarket"
            mUpdateStationsMSAMarket = False
            Exit Function
        End If
    End If

    Call mUpdateReport(lUpdateStationIndex, "MSA Market was assigned (" & Trim(tmUpdateStation(lUpdateStationIndex).sMSAMarket) & ")")
    mUpdateStationsMSAMarket = True
    Exit Function

ErrHandler:
    gMsg = ""
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mUpdateStationsMSAMarket"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mUpdateStationsOwner(StationIDX As Long, OwnerIdx As Long, lUpdateStationIndex As Long) As Integer
    Dim NewIdx As Long
    Dim slSQLQuery As String

    On Error GoTo ErrHandler:
    mUpdateStationsOwner = False
    tgStationInfo(StationIDX).lOwnerCode = tgOwnerInfo(OwnerIdx).lCode
    slSQLQuery = "Update shtt Set shttOwnerArttCode = " & tgStationInfo(StationIDX).lOwnerCode & " Where shttCode = " & tgStationInfo(StationIDX).iCode
    If bmUpdateDatabase Then
        'cnn.Execute slSQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHandler:
            gHandleError "AffErrorLog.txt", "ImportUpdateStations-mUpdateStationsOwner"
            mUpdateStationsOwner = False
            Exit Function
        End If
    End If

    Call mUpdateReport(lUpdateStationIndex, "Owner was assigned (" & Trim(tmUpdateStation(lUpdateStationIndex).sOwner) & ")")
    mUpdateStationsOwner = True
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mUpdateStationsOwner"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mUpdateStationsFormat(llStationIDX As Long, llFormatIdx As Long, llUpdateStationIndex As Long) As Integer
    Dim ilIndex As Integer
    Dim slSQLQuery As String

    On Error GoTo ErrHandler:
    mUpdateStationsFormat = False
    tgStationInfo(llStationIDX).iFormatCode = tmFormatLinkInfo(llFormatIdx).iIntFmtCode
    slSQLQuery = "Update shtt Set shttfmtCode = " & tgStationInfo(llStationIDX).iFormatCode & " Where shttCode = " & tgStationInfo(llStationIDX).iCode
    If bmUpdateDatabase Then
        'cnn.Execute slSQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHandler:
            gHandleError "AffErrorLog.txt", "ImportUpdateStations-mUpdateStationsFormat"
            mUpdateStationsFormat = False
            Exit Function
        End If
    End If
    ilIndex = gBinarySearchFmt(CLng(tmFormatLinkInfo(llFormatIdx).iIntFmtCode))
    If ilIndex <> -1 Then
        Call mUpdateReport(llUpdateStationIndex, "Format was assigned (" & Trim(tgFormatInfo(ilIndex).sName) & ")")
    Else
        Call mUpdateReport(llUpdateStationIndex, "Format was assigned")
    End If
    mUpdateStationsFormat = True
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mUpdateStationsFormat"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mUpdateRank(lMktIdx As Long, lUpdateStationIndex As Long) As Integer
    Dim slSQLQuery As String
    
    On Error GoTo ErrHandler:
    mUpdateRank = False
    slSQLQuery = "Update MKT Set mktRank = '" & tmUpdateStation(lUpdateStationIndex).iDMARank & "' Where mktCode = " & tgMarketInfo(lMktIdx).lCode
    If bmUpdateDatabase Then
        'cnn.Execute slSQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHandler:
            gHandleError "AffErrorLog.txt", "frmImportUpdateStations-mUpdateRank"
            mUpdateRank = False
            Exit Function
        End If
    End If
    
    Call mUpdateReport(lUpdateStationIndex, "DMA Market rank was changed from " & Trim(Str(tgMarketInfo(lMktIdx).iRank)) & " to " & Trim(Str(tmUpdateStation(lUpdateStationIndex).iDMARank)))
    tgMarketInfo(lMktIdx).iRank = tmUpdateStation(lUpdateStationIndex).iDMARank
    mUpdateRank = True
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mUpdateRank"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mUpdateMSARank(lMktIdx As Long, lUpdateStationIndex As Long) As Integer
    Dim slSQLQuery As String

    On Error GoTo ErrHandler:
    mUpdateMSARank = False
    slSQLQuery = "Update MET Set metRank = '" & tmUpdateStation(lUpdateStationIndex).iMSARank & "' Where metCode = " & tgMSAMarketInfo(lMktIdx).lCode
    If bmUpdateDatabase Then
        'cnn.Execute slSQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHandler:
            gHandleError "AffErrorLog.txt", "ImportUpdateStations-mUpdateMSARank"
            mUpdateMSARank = False
            Exit Function
        End If
    End If
    
    Call mUpdateReport(lUpdateStationIndex, "MSA Market rank was changed from " & Trim(Str(tgMSAMarketInfo(lMktIdx).iRank)) & " to " & Trim(Str(tmUpdateStation(lUpdateStationIndex).iMSARank)))
    tgMSAMarketInfo(lMktIdx).iRank = tmUpdateStation(lUpdateStationIndex).iMSARank
    mUpdateMSARank = True
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mUpdateMSARank"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Sub mAddMarketToUsedArray(llMktIdx As Long)
    Dim llNewIdx As Long
    Dim llLoop As Long
    
    For llLoop = 0 To UBound(tmUsedMarkets) - 1 Step 1
        If tgMarketInfo(llMktIdx).lCode = tmUsedMarkets(llLoop).lCode Then
            Exit Sub
        End If
    Next llLoop
    llNewIdx = UBound(tmUsedMarkets)
    ReDim Preserve tmUsedMarkets(0 To llNewIdx + 1) As MARKETINFO
    tmUsedMarkets(llNewIdx).lCode = tgMarketInfo(llMktIdx).lCode
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub mAddMSAMarketToUsedArray(llMktIdx As Long)
    Dim llNewIdx As Long
    Dim llLoop As Long
    
    For llLoop = 0 To UBound(tmUsedMSAMarkets) - 1 Step 1
        If tgMSAMarketInfo(llMktIdx).lCode = tmUsedMSAMarkets(llLoop).lCode Then
            Exit Sub
        End If
    Next llLoop
    llNewIdx = UBound(tmUsedMSAMarkets)
    ReDim Preserve tmUsedMSAMarkets(0 To llNewIdx + 1) As MARKETINFO
    tmUsedMSAMarkets(llNewIdx).lCode = tgMSAMarketInfo(llMktIdx).lCode
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub mAddOwnerToUsedArray(llOwnerIdx As Long)
    Dim llNewIdx As Long
    Dim llLoop As Long
    
    For llLoop = 0 To UBound(tmUsedOwners) - 1 Step 1
        If tgOwnerInfo(llOwnerIdx).lCode = tmUsedOwners(llLoop).lCode Then
            Exit Sub
        End If
    Next llLoop
    llNewIdx = UBound(tmUsedOwners)
    ReDim Preserve tmUsedOwners(0 To llNewIdx + 1) As OWNERINFO
    tmUsedOwners(llNewIdx).lCode = tgOwnerInfo(llOwnerIdx).lCode
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub mAddFormatToUsedArray(llFormatIdx As Long)
    Dim llNewIdx As Long
    Dim llLoop As Long
    
    For llLoop = 0 To UBound(tmUsedFormats) - 1 Step 1
        If tgFormatInfo(llFormatIdx).lCode = tmUsedFormats(llLoop).lCode Then
            Exit Sub
        End If
    Next llLoop
    llNewIdx = UBound(tmUsedFormats)
    ReDim Preserve tmUsedFormats(0 To llNewIdx + 1) As FORMATINFO
    tmUsedFormats(llNewIdx).lCode = tgFormatInfo(llFormatIdx).lCode
End Sub

'***************************************************************************
'
'***************************************************************************
Private Function mRemoveUnusedDMAMarkets() As Integer
    Dim llStationIDX As Long
    Dim llMktIdx As Long
    Dim blLinkIsOK As Boolean
    Dim llIdx As Long
    Dim llSef As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHandler
    mRemoveUnusedDMAMarkets = True
    mSetResults "Verifying station DMA Market information...", RGB(0, 0, 0)

    If UBound(tgStationInfo) < 1 Then
        Exit Function
    End If
    If UBound(tgMarketInfo) < 1 Then
        Exit Function
    End If
    
    ' At this point the arrary tmUsedMarkets contains a list of valid markets. Any not found in this
    ' list will now be deleted.
    If UBound(tmUsedMarkets) < 1 Then
        Exit Function
    End If
    For llMktIdx = 0 To UBound(tgMarketInfo) - 1 Step 1
        blLinkIsOK = False
        For llIdx = 0 To UBound(tmUsedMarkets) - 1 Step 1
            If tmUsedMarkets(llIdx).lCode = tgMarketInfo(llMktIdx).lCode Then
                llIdx = UBound(tmUsedMarkets) ' Exit out of this loop.
                blLinkIsOK = True
                Exit For
            End If
        Next
        If Not blLinkIsOK Then
            ' This market name is not being used. Remove it.
            If gUpdateRegions("M", tgMarketInfo(llMktIdx).lCode, -1, "AffErrorLog.Txt") Then
                SQLQuery = "Delete From MKT Where mktCode = " & tgMarketInfo(llMktIdx).lCode
                If bmUpdateDatabase Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHandler:
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mRemoveUnusedDMAMarkets"
                        mRemoveUnusedDMAMarkets = False
                        Exit Function
                    End If
                End If
                Call mUpdateReport(-1, "Unused DMA Market entry (" & Trim$(tgMarketInfo(llMktIdx).sName) & ") was removed.")
            End If
        End If
    Next
    Erase tmUsedMarkets
    ilRet = gPopMarkets()
    Exit Function
    
ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mRemoveUnusedDMAMarkets"
    mRemoveUnusedDMAMarkets = False
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mRemoveUnusedMSAMarkets() As Integer
    Dim llStationIDX As Long
    Dim llMktIdx As Long
    Dim blLinkIsOK As Boolean
    Dim llIdx As Long
    Dim llSef As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHandler
    mRemoveUnusedMSAMarkets = True
    mSetResults "Verifying station MSA Market information...", RGB(0, 0, 0)

    If UBound(tgStationInfo) < 1 Then
        Exit Function
    End If
    If UBound(tgMSAMarketInfo) < 1 Then
        Exit Function
    End If
    
    
    ' At this point the arrary tmUsedMarkets contains a list of valid markets. Any not found in this
    ' list will now be deleted.
    If UBound(tmUsedMSAMarkets) < 1 Then
        Exit Function
    End If
    For llMktIdx = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
        blLinkIsOK = False
        For llIdx = 0 To UBound(tmUsedMSAMarkets) - 1 Step 1
            If tmUsedMSAMarkets(llIdx).lCode = tgMSAMarketInfo(llMktIdx).lCode Then
                llIdx = UBound(tmUsedMSAMarkets) ' Exit out of this loop.
                blLinkIsOK = True
                Exit For
            End If
        Next
        If Not blLinkIsOK Then
            If gUpdateRegions("A", tgMSAMarketInfo(llMktIdx).lCode, -1, "AffErrorLog.Txt") Then
                ' This market name is not being used. Remove it.
                SQLQuery = "Delete From MET Where metCode = " & tgMSAMarketInfo(llMktIdx).lCode
                If bmUpdateDatabase Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHandler:
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mRemoveUnusedMSAMarkets"
                        mRemoveUnusedMSAMarkets = False
                        Exit Function
                    End If
                End If
                Call mUpdateReport(-1, "Unused MSA Market entry (" & Trim$(tgMSAMarketInfo(llMktIdx).sName) & ") was removed.")
            End If
        End If
    Next
    Erase tmUsedMSAMarkets
    ilRet = gPopMSAMarkets()
    Exit Function
    
ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mRemoveUnusedMSAMarkets"
    mRemoveUnusedMSAMarkets = False
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mRemoveUnusedOwners() As Integer
    Dim llStationIDX As Long
    Dim llOwnerIdx As Long
    Dim blLinkIsOK As Boolean
    Dim llIdx As Long
    Dim llSef As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHandler
    mRemoveUnusedOwners = True
    mSetResults "Verifying station owner information...", RGB(0, 0, 0)

    If UBound(tgStationInfo) < 1 Then
        Exit Function
    End If
    If UBound(tgOwnerInfo) < 1 Then
        Exit Function
    End If
    

    ' At this point the arrary tmUsedOwners contains a list of valid owners. Any not found in this
    ' list will now be deleted.
    If UBound(tmUsedOwners) < 1 Then
        Exit Function
    End If
    For llOwnerIdx = 0 To UBound(tgOwnerInfo) - 1 Step 1
        blLinkIsOK = False
        For llIdx = 0 To UBound(tmUsedOwners) - 1 Step 1
            If tmUsedOwners(llIdx).lCode = tgOwnerInfo(llOwnerIdx).lCode Then
                llIdx = UBound(tmUsedOwners) ' Exit out of this loop.
                blLinkIsOK = True
                Exit For
            End If
        Next
        If Not blLinkIsOK Then
            ' This market name is not being used. Remove it.
            SQLQuery = "Delete From artt Where arttCode = " & tgOwnerInfo(llOwnerIdx).lCode & " And arttType = 'O'"
            If bmUpdateDatabase Then
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/11/16: Replaced GoSub
                    'GoSub ErrHandler:
                    gHandleError "AffErrorLog.txt", "ImportUpdateStations-mRemoveUnusedOwners"
                    mRemoveUnusedOwners = False
                    Exit Function
                End If
            End If
            Call mUpdateReport(-1, "Unused owner entry (" & Trim$(tgOwnerInfo(llOwnerIdx).sName) & ") was removed.")
        End If
    Next
    Erase tmUsedOwners
    ilRet = gPopOwnerNames()
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mRemoveUnusedOwners"
    mRemoveUnusedOwners = False
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mRemoveUnusedFormats() As Integer
    Dim llStationIDX As Long
    Dim llFmtIdx As Long
    Dim blLinkIsOK As Boolean
    Dim llIdx As Long
    Dim llSef As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHandler
    mRemoveUnusedFormats = True
    mSetResults "Verifying station format information...", RGB(0, 0, 0)

    If UBound(tgStationInfo) < 1 Then
        Exit Function
    End If
    If UBound(tgFormatInfo) < 1 Then
        Exit Function
    End If

    ' At this point the arrary tmUsedFormats contains a list of valid formats. Any not found in this
    ' list will now be deleted.
    If UBound(tmUsedFormats) < 1 Then
        Exit Function
    End If
    For llFmtIdx = 0 To UBound(tgFormatInfo) - 1 Step 1
        blLinkIsOK = False
        For llIdx = 0 To UBound(tmUsedFormats) - 1 Step 1
            If tmUsedFormats(llIdx).lCode = tgFormatInfo(llFmtIdx).lCode Then
                llIdx = UBound(tmUsedFormats) ' Exit out of this loop.
                blLinkIsOK = True
                Exit For
            End If
        Next
        If Not blLinkIsOK Then
            If gUpdateRegions("F", tgFormatInfo(llFmtIdx).lCode, -1, "AffErrorLog.Txt") Then
                ' This format name is not being used. Remove it.
                SQLQuery = "Delete From FMT_Station_Format Where fmtCode = " & tgFormatInfo(llFmtIdx).lCode & ""
                If bmUpdateDatabase Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHandler:
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mRemoveUnusedFormats"
                        mRemoveUnusedFormats = False
                        Exit Function
                    End If
                End If
                Call mUpdateReport(-1, "Unused format entry (" & Trim$(tgFormatInfo(llFmtIdx).sName) & ") was removed.")
            End If
        End If
    Next
    Erase tmUsedFormats
    ilRet = gPopFormats()
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mRemoveUnusedFormats"
    mRemoveUnusedFormats = False
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mReportStationsNotUpdated() As Long
    Dim llLoop As Long
    Dim lllUpdateStationIndex As Long
    Dim slCallLetters As String

    If UBound(tmUpdateStation) > 0 Then
        'Sort by call letters, offset is two
        ArraySortTyp fnAV(tmUpdateStation(), 0), UBound(tmUpdateStation), 0, LenB(tmUpdateStation(0)), 2, LenB(tmUpdateStation(0).sCallLetters), 0
    End If
    mSetResults "Reporting stations not in Import file...", RGB(0, 0, 0)
    Call mUpdateReport(-1, "Stations not updated report.")
    Call mUpdateReport(-1, "This list shows stations not in the Import file and therefore could not be validated")
    Call mUpdateReport(-1, "---------------------------------------------------------------------------------")
    For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
        DoEvents
        If imTerminate Then
            Exit Function
        End If
        slCallLetters = Trim(tgStationInfo(llLoop).sCallLetters)
        lllUpdateStationIndex = mBinarySearchStation(slCallLetters)
        If lllUpdateStationIndex = -1 Then
            Call mUpdateReport(-1, slCallLetters)
        End If
    Next
    Call mUpdateReport(-1, "---------------------------------------------------------------------------------")
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mBinarySearchStation(slCallLetters As String) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim ilResult As Integer
    
    On Error GoTo ErrHand
    
    mBinarySearchStation = -1    ' Start out as not found.
    llMin = LBound(tmUpdateStation)
    llMax = UBound(tmUpdateStation) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        'ilResult = StrComp(Trim(tmUpdateStation(llMiddle).sCallLetters), slCallLetters, vbTextCompare)
        ilResult = StrComp(UCase(Trim(tmUpdateStation(llMiddle).sCallLetters)), Trim(UCase(slCallLetters)), vbBinaryCompare)
        Select Case ilResult
            Case 0:
                mBinarySearchStation = llMiddle  ' Found it !
                Exit Function
            Case 1:
                llMax = llMiddle - 1
            Case -1:
                llMin = llMiddle + 1
        End Select
    Loop
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in mBinarySearchStation: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mLookupStation(llStationID As Long, slCallLetters As String) As Long
    Dim llLoop As Long
    
    mLookupStation = -1
    On Error GoTo ErrHandler
    If bmMatchOnPermStationID Then
        If llStationID <= 0 Then
            Exit Function
        End If
        For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(llLoop).lPermStationID = llStationID Then
                mLookupStation = llLoop
                Exit Function
            End If
        Next
    Else
        For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If StrComp(UCase$(Trim(tgStationInfo(llLoop).sCallLetters)), UCase$(Trim(slCallLetters)), vbTextCompare) = 0 Then
                mLookupStation = llLoop
                Exit Function
            End If
        Next
    End If
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mLookupMarket(mktCode As Integer) As Long
    Dim llLoop As Long
    
    mLookupMarket = -1
    On Error GoTo ErrHandler
    If UBound(tgMarketInfo) < 1 Then
        Exit Function
    End If
    For llLoop = 0 To UBound(tgMarketInfo) - 1 Step 1
        If tgMarketInfo(llLoop).lCode = mktCode Then
            mLookupMarket = llLoop
            Exit Function
        End If
    Next
    Exit Function
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mLookupOwner(llOwnerCode As Long) As Long
    Dim llLoop As Long
    
    mLookupOwner = -1
    On Error GoTo ErrHandler
    If UBound(tgOwnerInfo) < 1 Then
        Exit Function
    End If
    For llLoop = 0 To UBound(tgOwnerInfo) - 1 Step 1
        If tgOwnerInfo(llLoop).lCode = llOwnerCode Then
            mLookupOwner = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mLookupDMAMarketByName(slMarketName As String) As Long
    Dim llLoop As Long
    Dim llIdx As Long

    mLookupDMAMarketByName = -1
    On Error GoTo ErrHandler
    llIdx = UBound(tgMarketInfo)
    If llIdx < 1 Then
        Exit Function
    End If
    For llLoop = 0 To llIdx - 1 Step 1
        If StrComp(Trim(tgMarketInfo(llLoop).sName), Trim(slMarketName), vbTextCompare) = 0 Then
            mLookupDMAMarketByName = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mLookupMSAMarketByName(slMarketName As String) As Long
    Dim llLoop As Long
    Dim llIdx As Long

    mLookupMSAMarketByName = -1
    On Error GoTo ErrHandler
    llIdx = UBound(tgMSAMarketInfo)
    If llIdx < 1 Then
        Exit Function
    End If
    For llLoop = 0 To llIdx - 1 Step 1
        If StrComp(Trim(tgMSAMarketInfo(llLoop).sName), Trim(slMarketName), vbTextCompare) = 0 Then
            mLookupMSAMarketByName = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mLookupOwnerByName(sOwnerName As String) As Long
    Dim llLoop As Long
    Dim llIdx As Long

    mLookupOwnerByName = -1
    On Error GoTo ErrHandler
    llIdx = UBound(tgOwnerInfo)
    If llIdx < 1 Then
        Exit Function
    End If
    For llLoop = 0 To llIdx
        If StrComp(Trim(tgOwnerInfo(llLoop).sName), Trim(sOwnerName), vbTextCompare) = 0 Then
            mLookupOwnerByName = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mCreateStatusReport()
    Dim llBIAReportIdx As Long
    Dim ilLoop As Long
    Dim hlFile As Integer
    Dim slStatus As String
    Dim slStartCallLetters As String
    Dim slCurrentCallLetters As String
    Dim ilRet As Integer
    
    mCreateStatusReport = False
    mSetResults "Creating Status Report...", RGB(0, 0, 0)
    llBIAReportIdx = UBound(tmReportInfo)
    If llBIAReportIdx < 1 Then
        mCreateStatusReport = True
        Exit Function
    End If
    On Error GoTo IgnoreError
    Kill smReportPathFileName
    On Error GoTo ErrHandler
    'hlFile = FreeFile
    'Open smReportPathFileName For Append As hlFile
    ilRet = gFileOpen(smReportPathFileName, "Append", hlFile)
    If Not bmUpdateDatabase Then
        Print #hlFile, "*** SHOW REPORT ONLY WAS CHECKED."
        Print #hlFile, "*** NO CHANGES WERE MADE TO THE DATABASE"
        Print #hlFile, "*** THE FOLLOWING INFORMATION SHOWS WHAT WOULD HAVE OCCURRED"
        Print #hlFile, "***"
    End If
    slStartCallLetters = ""
    For ilLoop = 0 To llBIAReportIdx
'        slCurrentCallLetters = tmReportInfo(ilLoop).sCallLetters
'        If slStartCallLetters <> slCurrentCallLetters Then
'            slStartCallLetters = slCurrentCallLetters
'            Print #hlFile, slStartCallLetters
'        End If
'        slStatus = "    " & Trim(tmReportInfo(ilLoop).sReportInfo)
'        Print #hlFile, slStatus

        slStatus = Trim(tmReportInfo(ilLoop).sReportInfo)
        If Len(Trim(tmReportInfo(ilLoop).sCallLetters)) > 0 Then
            If Trim(tmReportInfo(ilLoop).sCallLetters) <> "" And Trim(slStatus) <> "" And tmReportInfo(ilLoop).lPermStationNo <> 0 Then
                Print #hlFile, Trim(tmReportInfo(ilLoop).sCallLetters) & ", ", slStatus
            End If
        Else
            Print #hlFile, slStatus
        End If
    Next
    Close hlFile
    mCreateStatusReport = True
    Exit Function

IgnoreError:
    Resume Next
ErrHandler:
    mSetResults "Error creating status report.", RGB(255, 0, 0)
End Function

'***************************************************************************
'
'***************************************************************************
Private Sub mUpdateReport(llUpdateStationIndex As Long, slMsg As String)
    Dim llReportIndex As Long
    
    llReportIndex = UBound(tmReportInfo)
    ReDim Preserve tmReportInfo(0 To llReportIndex + 1) As BIAREPORTINFO
    If llUpdateStationIndex <> -1 Then
        tmReportInfo(llReportIndex).sCallLetters = Trim(tmUpdateStation(llUpdateStationIndex).sCallLetters)
        tmReportInfo(llReportIndex).lPermStationNo = tgStationInfo(llUpdateStationIndex).lPermStationID
    Else
        tmReportInfo(llReportIndex).sCallLetters = ""
    End If
    tmReportInfo(llReportIndex).sReportInfo = slMsg
End Sub
'***************************************************************************
'
'***************************************************************************
Private Sub mUpdateReport2(sCallLetters As String, lPermStationID As Long, slMsg As String)
    Dim llReportIndex As Long
    
    llReportIndex = UBound(tmReportInfo)
    ReDim Preserve tmReportInfo(0 To llReportIndex + 1) As BIAREPORTINFO
    'If llUpdateStationIndex <> -1 Then
    tmReportInfo(llReportIndex).sCallLetters = Trim(sCallLetters)
    tmReportInfo(llReportIndex).lPermStationNo = lPermStationID
    'Else
    '    tmReportInfo(llReportIndex).sCallLetters = ""
    'End If
    tmReportInfo(llReportIndex).sReportInfo = slMsg
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub mSetResults(Msg As String, FGC As Long)
    gLogMsg Msg, "AffErrorLog.Txt", False
    lbcMsg.AddItem Msg
    lbcMsg.ListIndex = lbcMsg.ListCount - 1
    lbcMsg.ForeColor = FGC
    DoEvents
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub cmcBrowse_Click()
    Dim slCurDir As String
    
    slCurDir = CurDir
    
    sgBrowseMaskFile = "*.csv"
    igBrowseType = 1
    frmBrowse.Show vbModal
    If igBrowseReturn = 1 Then
        txtFile.Text = Trim$(sgBrowseFile)
    End If
    ChDir slCurDir
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub cmdCancel_Click()
    Dim ilResp As Integer
    
    If imImporting Then
        ilResp = gMsgBox("Are you sure you want to abort?", vbYesNo)
        If ilResp = vbYes Then
            imTerminate = True
        End If
        Exit Sub
    End If
    If (bmImported) And (sgUsingStationID = "A") Then
        On Error GoTo ErrHand:
        SQLQuery = "Update site Set "
        SQLQuery = SQLQuery & "siteUsingStationID = '" & "N" & "' "
        SQLQuery = SQLQuery & " Where siteCode = " & 1
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHandler:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "ImportUpdateStations-cmdCancel_Click"
            Exit Sub
        End If
        sgUsingStationID = "N"
        frmMain!mnuImportStation.Caption = "Update Existing Stations"
    End If
    Unload frmImportUpdateStations
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-cmdCancel"
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = vbNormal
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imImporting Then
        Screen.MousePointer = vbHourglass
    End If
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub Form_Load()
    Dim iLoop As Integer
    Dim iZone As Integer
    
    Screen.MousePointer = vbHourglass
    'frmImportUpdateStations.Caption = "Update Station Information - " & sgClientName
    bmImported = False
    'TTP 10984 - Station Information Import: when using Station IDs, add method that will allow call letters to be swapped
    ckcChangeToInvalidIfLetterReassigned.Enabled = False
    If sgUsingStationID = "Y" Then
        frmImportUpdateStations.Caption = "Update/Add Stations: " & sgClientName
        ckcChangeToInvalidIfLetterReassigned.Enabled = True
    ElseIf sgUsingStationID = "A" Then
        If UBound(tgStationInfo) > LBound(tgStationInfo) Then
            frmImportUpdateStations.Caption = "Continue Adding Stations: " & sgClientName
        Else
            frmImportUpdateStations.Caption = "Add Initial Stations: " & sgClientName
        End If
    Else
        If UBound(tgStationInfo) > LBound(tgStationInfo) Then
            frmImportUpdateStations.Caption = "Update Existing Stations: " & sgClientName
        Else
            frmImportUpdateStations.Caption = "Add Initial Stations: " & sgClientName
        End If
    End If
    imTerminate = False
    imImporting = False
    bmAdjPledge = False
    chkReportStationsNotUpdated.Value = vbChecked
    ckcMissingStations.Value = vbChecked
    
    txtFile.Text = ""
    smReportPathFileName = sgMsgDirectory & "UpdateStationStatus.txt"
    mBuildImportTitles
    mGetBands
    Screen.MousePointer = vbDefault
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not bmUpdateDatabase Then
        ' If the database was not updated, then reload these arrays.
        If Not gPopMarkets() Then
            Call mFailImportMsg("Unable to Load Existing DMA Market Names.")
            Exit Sub
        End If
        If Not gPopMSAMarkets() Then
            Call mFailImportMsg("Unable to Load Existing MSA Market Names.")
            Exit Sub
        End If
        If Not gPopOwnerNames() Then
            Call mFailImportMsg("Unable to Load Existing Owner Names.")
            Exit Sub
        End If
        If Not gPopFormats() Then
            Call mFailImportMsg("Unable to Load Existing Format Names.")
            Exit Sub
        End If
    End If
    Erase lmAttCode
    Erase tmReportInfo
    Erase tmUpdateStation
    Erase tmSvUpdateStation
    Erase smPass0LinesBypassed
    Erase tmFormatLinkInfo
    Erase tmMonikerInfo
    Erase tmCityInfo
    Erase tmCountyInfo
    Erase tmTerritoryInfo
    Erase tmAreaInfo
    Erase tmOperatorInfo
    
    Erase tmUsedMarkets
    Erase tmUsedMSAMarkets
    Erase tmUsedOwners
    Erase tmUsedFormats
    rst_Shtt.Close
    rst_artt.Close
    rst_cmt.Close
    Set frmImportUpdateStations = Nothing
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub mFailImportMsg(sMsg As String)
    mSetResults sMsg, RGB(255, 0, 0)
    imImporting = False
    'cmdViewReport.Enabled = True
    cmdCancel.Caption = "&Done"
    cmdCancel.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Function mRemoveDuplicateMarketNames() As Integer
    Dim llOutsideLoop As Long
    Dim llInsideLoop As Long
    Dim ilRet As Integer
    
    mRemoveDuplicateMarketNames = True
    On Error GoTo ErrHand:
    mSetResults "Checking and removing duplicated Market Names...", RGB(0, 0, 0)
    For llOutsideLoop = LBound(tgMarketInfo) To UBound(tgMarketInfo) - 1 Step 1
        If tgMarketInfo(llOutsideLoop).lCode > 0 Then
            For llInsideLoop = llOutsideLoop + 1 To UBound(tgMarketInfo) - 1 Step 1
                If StrComp(tgMarketInfo(llInsideLoop).sName, tgMarketInfo(llOutsideLoop).sName, vbTextCompare) = 0 Then
                    SQLQuery = "Update shtt Set shttMktCode = " & tgMarketInfo(llOutsideLoop).lCode & " Where shttMktCode = " & tgMarketInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mRemoveDuplicateMarketNames"
                        mRemoveDuplicateMarketNames = False
                        Exit Function
                    End If
                    SQLQuery = "Update mat Set matMktCode = " & tgMarketInfo(llOutsideLoop).lCode & " Where matMktCode = " & tgMarketInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mRemoveDuplicateMarketNames"
                        mRemoveDuplicateMarketNames = False
                        Exit Function
                    End If
                    SQLQuery = "Update mgt Set mgtMktCode = " & tgMarketInfo(llOutsideLoop).lCode & " Where mgtMktCode = " & tgMarketInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mRemoveDuplicateMarketNames"
                        mRemoveDuplicateMarketNames = False
                        Exit Function
                    End If
                    'SQLQuery = "SELECT * FROM raf_region_area WHERE ((rafCategory = 'M') and (rafType = 'C' OR rafType = 'N'))"
                    'Set rst_Raf = gSQLSelectCall(SQLQuery)
                    'Do While Not rst_Raf.EOF
                    '    SQLQuery = "SELECT * FROM sef_Split_Entity WHERE sefIntCode = " & tgMarketInfo(llOutsideLoop).iCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
                    '    Set rst_Sef = gSQLSelectCall(SQLQuery)
                    '    If rst_Sef.EOF Then
                    '        SQLQuery = "Update sef_Split_Entity Set sefIntCode = " & tgMarketInfo(llOutsideLoop).iCode & " Where sefIntCode = " & tgMarketInfo(llInsideLoop).iCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
                    '        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '            GoSub ErrHand:
                    '        End If
                    '    Else
                    '        SQLQuery = "DELETE FROM sef_Split_Entity WHERE sefIntCode = " & tgMarketInfo(llInsideLoop).iCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
                    '        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '            GoSub ErrHand:
                    '        End If
                    '    End If
                    '    rst_Raf.MoveNext
                    'Loop
                    ilRet = gUpdateRegions("M", tgMarketInfo(llInsideLoop).lCode, tgMarketInfo(llOutsideLoop).lCode, "AffErrorLog.Txt")
                    If Not ilRet Then
                        mRemoveDuplicateMarketNames = False
                        Exit Function
                    End If
                    SQLQuery = "DELETE FROM mkt WHERE mktCode = " & tgMarketInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mRemoveDuplicateMarketNames"
                        mRemoveDuplicateMarketNames = False
                        Exit Function
                    End If
                    tgMarketInfo(llInsideLoop).lCode = -1
                End If
            Next llInsideLoop
        End If
    Next llOutsideLoop
    
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mRemoveDuplicateMarketNames"
    mRemoveDuplicateMarketNames = False
    Exit Function
End Function

Private Function mSetUsedArrays() As Integer
    Dim llStationIDX As Long
    Dim llMktIdx As Long
    Dim llOwnerIdx As Long
    Dim llFmtIdx As Long
    Dim blLinkIsOK As Boolean
    
    mSetUsedArrays = True
    On Error GoTo ErrHandler
    
    If UBound(tgStationInfo) < 1 Then
        Exit Function
    End If

    For llStationIDX = 0 To UBound(tgStationInfo) - 1 Step 1
        DoEvents
        If imTerminate Then
            Exit Function
        End If
        blLinkIsOK = False
        If (tgStationInfo(llStationIDX).iMktCode > 0) And (UBound(tgMarketInfo) > 0) Then
            For llMktIdx = 0 To UBound(tgMarketInfo) - 1 Step 1
                If tgStationInfo(llStationIDX).iMktCode = tgMarketInfo(llMktIdx).lCode Then
                    ' If this market name is blank, don't add it. It will get deleted later.
                    If Len(Trim(tgMarketInfo(llMktIdx).sName)) > 0 Then
                        mAddMarketToUsedArray llMktIdx
                        blLinkIsOK = True
                        Exit For
                    End If
                End If
            Next
            If Not blLinkIsOK Then
                ' The station is not pointing to a valid market name.
                SQLQuery = "Update shtt Set shttMktCode = 0 Where shttCode = " & tgStationInfo(llStationIDX).iCode
                If bmUpdateDatabase Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHandler:
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mSetUsedArrays"
                        mSetUsedArrays = False
                        Exit Function
                    End If
                End If
                Call mUpdateReport(-1, Trim(tgStationInfo(llStationIDX).sCallLetters) & " Bad DMA Market pointer was removed. Station now has no DMA Market.")
            End If
        End If
    
        blLinkIsOK = False
        If (tgStationInfo(llStationIDX).iMSAMktCode > 0) And (UBound(tgMSAMarketInfo) > 0) Then
            For llMktIdx = 0 To UBound(tgMSAMarketInfo) - 1 Step 1
                If tgStationInfo(llStationIDX).iMSAMktCode = tgMSAMarketInfo(llMktIdx).lCode Then
                    ' If this market name is blank, don't add it. It will get deleted later.
                    If Len(Trim(tgMSAMarketInfo(llMktIdx).sName)) > 0 Then
                        mAddMSAMarketToUsedArray llMktIdx
                        blLinkIsOK = True
                        Exit For
                    End If
                End If
            Next
            If Not blLinkIsOK Then
                ' The station is not pointing to a valid market name.
                SQLQuery = "Update shtt Set shttMetCode = 0 Where shttCode = " & tgStationInfo(llStationIDX).iCode
                If bmUpdateDatabase Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHandler:
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mSetUsedArrays"
                        mSetUsedArrays = False
                        Exit Function
                    End If
                End If
                Call mUpdateReport(-1, Trim(tgStationInfo(llStationIDX).sCallLetters) & " Bad MSA Market pointer was removed. Station now has no MSA Market.")
            End If
        End If
    
        blLinkIsOK = False
        If (tgStationInfo(llStationIDX).lOwnerCode > 0) And (UBound(tgOwnerInfo) > 0) Then ' Look only when a link exist.
            For llOwnerIdx = 0 To UBound(tgOwnerInfo) - 1 Step 1
                If tgStationInfo(llStationIDX).lOwnerCode = tgOwnerInfo(llOwnerIdx).lCode Then
                    mAddOwnerToUsedArray llOwnerIdx
                    blLinkIsOK = True
                    Exit For
                End If
            Next
            If Not blLinkIsOK Then
                ' The station is not pointing to a valid owner name.
                SQLQuery = "Update shtt Set shttOwnerArttCode = 0 Where shttCode = " & tgStationInfo(llStationIDX).iCode
                If bmUpdateDatabase Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHandler:
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mSetUsedArrays"
                        mSetUsedArrays = False
                        Exit Function
                    End If
                End If
                Call mUpdateReport(-1, Trim(tgStationInfo(llStationIDX).sCallLetters) & " Bad owner pointer was removed. Station now has no owner.")
            End If
        End If
    
        blLinkIsOK = False
        If (tgStationInfo(llStationIDX).iFormatCode > 0) And (UBound(tgFormatInfo) > 0) Then ' Look only when a link exist.
            For llFmtIdx = 0 To UBound(tgFormatInfo) - 1 Step 1
                If tgStationInfo(llStationIDX).iFormatCode = tgFormatInfo(llFmtIdx).lCode Then
                    mAddFormatToUsedArray llFmtIdx
                    blLinkIsOK = True
                    Exit For
                End If
            Next
            If Not blLinkIsOK Then
                ' The station is not pointing to a valid format name.
                SQLQuery = "Update shtt Set shttFmtCode = 0 Where shttCode = " & tgStationInfo(llStationIDX).iCode
                If bmUpdateDatabase Then
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHandler:
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mSetUsedArrays"
                        mSetUsedArrays = False
                        Exit Function
                    End If
                End If
                Call mUpdateReport(-1, Trim(tgStationInfo(llStationIDX).sCallLetters) & " Bad format pointer was removed. Station now has no format.")
            End If
        End If
    Next
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mSetUsedArrays"
    mSetUsedArrays = False
    Exit Function
End Function

Private Function mUpdateRegions(slCategory As String, ilFromCode As Integer, ilToCode As Integer) As Integer
    Dim slSQLQuery As String
    Dim rst_Raf As ADODB.Recordset
    Dim rst_Sef As ADODB.Recordset

    mUpdateRegions = True
    On Error GoTo ErrHand:
    slSQLQuery = "SELECT * FROM raf_region_area WHERE ((rafCategory = '" & slCategory & "') and (rafType = 'C' OR rafType = 'N'))"
    Set rst_Raf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_Raf.EOF
        slSQLQuery = "SELECT * FROM sef_Split_Entity WHERE sefIntCode = " & ilToCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
        Set rst_Sef = gSQLSelectCall(slSQLQuery)
        If rst_Sef.EOF Then
            slSQLQuery = "Update sef_Split_Entity Set sefIntCode = " & ilToCode & " Where sefIntCode = " & ilFromCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
            If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "ImportUpdateStations-mUpdateRegions"
                mUpdateRegions = False
                Exit Function
            End If
        Else
            slSQLQuery = "DELETE FROM sef_Split_Entity WHERE sefIntCode = " & ilFromCode & " AND " & "sefRafCode = " & rst_Raf!rafCode
            If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "ImportUpdateStations-mUpdateRegions"
                mUpdateRegions = False
                Exit Function
            End If
        End If
        rst_Raf.MoveNext
    Loop
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mUpdateRegions"
    mUpdateRegions = False
    Exit Function
End Function

Private Function mCheckFormats() As Integer
    Dim llLoop As Long
    Dim llNameCheck As Long
    Dim ilFound As Integer
    Dim ilRet As Integer
    Dim ilMax As Integer
    Dim blBypassMatch As Boolean
    
    On Error GoTo ErrHand
    
    ReDim tgNewNamesImported(0 To 0) As NEWNAMESIMPORTED
    For llLoop = LBound(tmUpdateStation) To UBound(tmUpdateStation) - 1 Step 1
        If Len(Trim$(tmUpdateStation(llLoop).sFormat)) > 0 Then
            If mLookupFormatLinkByName(tmUpdateStation(llLoop).sFormat) < 0 Then
                ilFound = False
                For llNameCheck = LBound(tgNewNamesImported) To UBound(tgNewNamesImported) - 1 Step 1
                    If StrComp(Trim$(tgNewNamesImported(llNameCheck).sNewName), Trim$(tmUpdateStation(llLoop).sFormat), vbTextCompare) = 0 Then
                        tgNewNamesImported(llNameCheck).lUpdateStationIndex = -1
                        tgNewNamesImported(llNameCheck).iCount = tgNewNamesImported(llNameCheck).iCount + 1
                        ilFound = True
                        Exit For
                    End If
                Next llNameCheck
                If Not ilFound Then
                    tgNewNamesImported(UBound(tgNewNamesImported)).sNewName = tmUpdateStation(llLoop).sFormat
                    tgNewNamesImported(UBound(tgNewNamesImported)).lUpdateStationIndex = llLoop
                    tgNewNamesImported(UBound(tgNewNamesImported)).lReplaceCode = 0
                    tgNewNamesImported(UBound(tgNewNamesImported)).iCount = 1
                    ReDim Preserve tgNewNamesImported(0 To UBound(tgNewNamesImported) + 1) As NEWNAMESIMPORTED
                End If
            End If
        End If
    Next llLoop
    If UBound(tgNewNamesImported) > LBound(tgNewNamesImported) Then
        'Determine if records should be added:  No records exist or no reference to records exist
        blBypassMatch = True
        If UBound(tgFormatInfo) > LBound(tgFormatInfo) Then
            For llLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
                If tgStationInfo(llLoop).iFormatCode > 0 Then
                    blBypassMatch = False
                    Exit For
                End If
            Next llLoop
        End If
        If Not blBypassMatch Then
            'Ask user how the new names should be handled
            igNewNamesImportedType = 4
            frmCategoryMatching.Show vbModal
            If igNewNamesImportedReturn = 0 Then
                mCheckFormats = False
                Exit Function
            End If
        End If
        For llLoop = LBound(tgNewNamesImported) To UBound(tgNewNamesImported) - 1 Step 1
            If tgNewNamesImported(llLoop).lReplaceCode <= 0 Then
                SQLQuery = "Select MAX(fmtCode) from FMT_Station_Format"
                Set rst = gSQLSelectCall(SQLQuery)
                If IsNull(rst(0).Value) = True Then
                    ilMax = 1
                Else
                    ilMax = rst(0).Value + 1
                End If
                
                SQLQuery = "INSERT INTO FMT_Station_Format (fmtCode, fmtName, fmtUstCode, fmtGroupName, fmtDftCode, fmtUnused) "
                SQLQuery = SQLQuery & " VALUES ( " & ilMax & ", '" & gFixQuote(Trim$(tgNewNamesImported(llLoop).sNewName)) & "', " & igUstCode & ", " & "''" & ", " & 0 & ",''" & ")"
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/11/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "ImportUpdateStations-mCheckFormats"
                    mCheckFormats = False
                    Exit Function
                End If
                tgNewNamesImported(llLoop).lReplaceCode = ilMax
            End If
            SQLQuery = "Insert into flt "
            SQLQuery = SQLQuery & "(fltExtFormatName, fltIntFmtCode, fltUnused) "
            SQLQuery = SQLQuery & " VALUES ('" & gFixQuote(Trim$(tgNewNamesImported(llLoop).sNewName)) & "'," & CInt(tgNewNamesImported(llLoop).lReplaceCode) & "," & "''" & ")"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "ImportUpdateStations-mCheckFormats"
                mCheckFormats = False
                Exit Function
            End If
        Next llLoop
        ilRet = gPopFormats()
        ilRet = mPopFormatLink()
    End If
    mCheckFormats = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mCheckFormats"
    mCheckFormats = False
End Function

Private Function mCheckDMAMarkets() As Integer
    Dim llLoop As Long
    Dim slReplacedName As String
    Dim ilFound As Integer
    Dim llNameCheck As Long
    Dim llIndex As Long
    Dim ilRet As Integer
    Dim blBypassMatch As Boolean
    
    On Error GoTo ErrHand
    
    ReDim tgNewNamesImported(0 To 0) As NEWNAMESIMPORTED
    For llLoop = LBound(tmUpdateStation) To UBound(tmUpdateStation) - 1 Step 1
        If Len(Trim$(tmUpdateStation(llLoop).sDMAMarket)) > 0 Then
            If mLookupDMAMarketByName(tmUpdateStation(llLoop).sDMAMarket) < 0 Then
                ilFound = False
                For llNameCheck = LBound(tgNewNamesImported) To UBound(tgNewNamesImported) - 1 Step 1
                    If StrComp(Trim$(tgNewNamesImported(llNameCheck).sNewName), Trim$(tmUpdateStation(llLoop).sDMAMarket), vbTextCompare) = 0 Then
                        tgNewNamesImported(llNameCheck).lUpdateStationIndex = -1
                        tgNewNamesImported(llNameCheck).iCount = tgNewNamesImported(llNameCheck).iCount + 1
                        ilFound = True
                        Exit For
                    End If
                Next llNameCheck
                If Not ilFound Then
                    tgNewNamesImported(UBound(tgNewNamesImported)).sNewName = tmUpdateStation(llLoop).sDMAMarket
                    tgNewNamesImported(UBound(tgNewNamesImported)).iRank = tmUpdateStation(llLoop).iDMARank
                    tgNewNamesImported(UBound(tgNewNamesImported)).lUpdateStationIndex = llLoop
                    tgNewNamesImported(UBound(tgNewNamesImported)).lReplaceCode = 0
                    tgNewNamesImported(UBound(tgNewNamesImported)).iCount = 1
                    ReDim Preserve tgNewNamesImported(0 To UBound(tgNewNamesImported) + 1) As NEWNAMESIMPORTED
                End If
            End If
        End If
    Next llLoop
    If UBound(tgNewNamesImported) > LBound(tgNewNamesImported) Then
        'Determine if records should be added:  No records exist or no reference to records exist
        blBypassMatch = True
        If UBound(tgMarketInfo) > LBound(tgMarketInfo) Then
            For llLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
                If tgStationInfo(llLoop).iMktCode > 0 Then
                    blBypassMatch = False
                    Exit For
                End If
            Next llLoop
        End If
        If Not blBypassMatch Then
            'Ask user how the new names should be handled
            igNewNamesImportedType = 1
            frmCategoryMatching.Show vbModal
            If igNewNamesImportedReturn = 0 Then
                mCheckDMAMarkets = False
                Exit Function
            End If
        End If
        For llLoop = LBound(tgNewNamesImported) To UBound(tgNewNamesImported) - 1 Step 1
            If tgNewNamesImported(llLoop).lReplaceCode <= 0 Then
                'Add Name
                SQLQuery = "Insert into mkt "
                SQLQuery = SQLQuery & "(mktName, mktRank, mktUsfCode, mktGroupName, mktUnused) "
                SQLQuery = SQLQuery & " VALUES ('" & gFixQuote(Trim$(tgNewNamesImported(llLoop).sNewName)) & "'," & tgNewNamesImported(llLoop).iRank & "," & igUstCode & ",'" & "" & "'," & "''" & ")"
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/11/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "ImportUpdateStations-mCheckDMAMarkets"
                    mCheckDMAMarkets = False
                    Exit Function
                End If
                mUpdateReport -1, "DMA Market Name Added: " & Trim$(tgNewNamesImported(llLoop).sNewName) & " on " & tgNewNamesImported(llLoop).iCount & " stations"
            Else
                'Replace Name
                llIndex = gBinarySearchMkt(tgNewNamesImported(llLoop).lReplaceCode)
                If llIndex <> -1 Then
                    slReplacedName = Trim$(tgMarketInfo(llIndex).sName)
                End If
                SQLQuery = "UPDATE mkt"
                SQLQuery = SQLQuery & " SET mktUsfCode = " & igUstCode & ","
                SQLQuery = SQLQuery & "mktName = '" & gFixQuote(Trim$(tgNewNamesImported(llLoop).sNewName)) & "'"
                SQLQuery = SQLQuery & " WHERE mktCode = " & tgNewNamesImported(llLoop).lReplaceCode
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/11/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "ImportUpdateStations-mCheckDMAMarkets"
                    mCheckDMAMarkets = False
                    Exit Function
                End If
                If llIndex <> -1 Then
                    mUpdateReport -1, "DMA Market Name: " & slReplacedName & " Replaced by: " & Trim$(tgNewNamesImported(llLoop).sNewName) & " on " & tgNewNamesImported(llLoop).iCount & " stations"
                End If
            End If
        Next llLoop
        ilRet = gPopMarkets()
    End If
    mCheckDMAMarkets = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mCheckDMAMarkets"
    mCheckDMAMarkets = False
End Function

Private Function mCheckMSAMarkets() As Integer
    Dim llLoop As Long
    Dim slReplacedName As String
    Dim ilFound As Integer
    Dim llNameCheck As Long
    Dim llIndex As Long
    Dim ilRet As Integer
    Dim blBypassMatch As Boolean
    
    On Error GoTo ErrHand
    
    ReDim tgNewNamesImported(0 To 0) As NEWNAMESIMPORTED
    For llLoop = LBound(tmUpdateStation) To UBound(tmUpdateStation) - 1 Step 1
        If Len(Trim$(tmUpdateStation(llLoop).sMSAMarket)) > 0 Then
            If mLookupMSAMarketByName(tmUpdateStation(llLoop).sMSAMarket) < 0 Then
                ilFound = False
                For llNameCheck = LBound(tgNewNamesImported) To UBound(tgNewNamesImported) - 1 Step 1
                    If StrComp(Trim$(tgNewNamesImported(llNameCheck).sNewName), Trim$(tmUpdateStation(llLoop).sMSAMarket), vbTextCompare) = 0 Then
                        tgNewNamesImported(llNameCheck).lUpdateStationIndex = -1
                        tgNewNamesImported(llNameCheck).iCount = tgNewNamesImported(llNameCheck).iCount + 1
                        ilFound = True
                        Exit For
                    End If
                Next llNameCheck
                If Not ilFound Then
                    tgNewNamesImported(UBound(tgNewNamesImported)).sNewName = tmUpdateStation(llLoop).sMSAMarket
                    tgNewNamesImported(UBound(tgNewNamesImported)).iRank = tmUpdateStation(llLoop).iMSARank
                    tgNewNamesImported(UBound(tgNewNamesImported)).lUpdateStationIndex = llLoop
                    tgNewNamesImported(UBound(tgNewNamesImported)).lReplaceCode = 0
                    tgNewNamesImported(UBound(tgNewNamesImported)).iCount = 1
                    ReDim Preserve tgNewNamesImported(0 To UBound(tgNewNamesImported) + 1) As NEWNAMESIMPORTED
                End If
            End If
        End If
    Next llLoop
    If UBound(tgNewNamesImported) > LBound(tgNewNamesImported) Then
        'Determine if records should be added:  No records exist or no reference to records exist
        blBypassMatch = True
        If UBound(tgMSAMarketInfo) > LBound(tgMSAMarketInfo) Then
            For llLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
                If tgStationInfo(llLoop).iMSAMktCode > 0 Then
                    blBypassMatch = False
                    Exit For
                End If
            Next llLoop
        End If
        If Not blBypassMatch Then
            'Ask user how the new names should be handled
            igNewNamesImportedType = 2
            frmCategoryMatching.Show vbModal
            If igNewNamesImportedReturn = 0 Then
                mCheckMSAMarkets = False
                Exit Function
            End If
        End If
        For llLoop = LBound(tgNewNamesImported) To UBound(tgNewNamesImported) - 1 Step 1
            If tgNewNamesImported(llLoop).lReplaceCode <= 0 Then
                'Add Name
                SQLQuery = "Insert into met "
                SQLQuery = SQLQuery & "(metName, metRank, metUstCode, metGroupName, metUnused) "
                SQLQuery = SQLQuery & " VALUES ('" & gFixQuote(Trim$(tgNewNamesImported(llLoop).sNewName)) & "'," & tgNewNamesImported(llLoop).iRank & "," & igUstCode & ",'" & "" & "'," & "''" & ")"
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/11/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "ImportUpdateStations-mCheckMSAMarkets"
                    mCheckMSAMarkets = False
                    Exit Function
                End If
                mUpdateReport -1, "MSA Market Name Added: " & Trim$(tgNewNamesImported(llLoop).sNewName) & " on " & tgNewNamesImported(llLoop).iCount & " stations"
            Else
                'Replace Name
                llIndex = gBinarySearchMSAMkt(tgNewNamesImported(llLoop).lReplaceCode)
                If llIndex <> -1 Then
                    slReplacedName = Trim$(tgMSAMarketInfo(llIndex).sName)
                End If
                SQLQuery = "UPDATE met"
                SQLQuery = SQLQuery & " SET metUstCode = " & igUstCode & ","
                SQLQuery = SQLQuery & "metName = '" & gFixQuote(Trim$(tgNewNamesImported(llLoop).sNewName)) & "'"
                SQLQuery = SQLQuery & " WHERE metCode = " & tgNewNamesImported(llLoop).lReplaceCode
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/11/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "ImportUpdateStations-mCheckMSAMarkets"
                    mCheckMSAMarkets = False
                    Exit Function
                End If
                If llIndex <> -1 Then
                    mUpdateReport -1, "MSA Market Name: " & slReplacedName & " Replaced by: " & Trim$(tgNewNamesImported(llLoop).sNewName) & " on " & tgNewNamesImported(llLoop).iCount & " stations"
                End If
            End If
        Next llLoop
        ilRet = gPopMSAMarkets()
    End If
    mCheckMSAMarkets = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mCheckMSAMarkets"
    mCheckMSAMarkets = False
End Function

Private Function mCheckOwners() As Integer
    Dim llLoop As Long
    Dim slReplacedName As String
    Dim ilFound As Integer
    Dim llNameCheck As Long
    Dim llIndex As Long
    Dim llStation As Long
    Dim ilRet As Integer
    Dim blBypassMatch As Boolean
    Dim blFirst As Boolean
    
    On Error GoTo ErrHand
    
    ReDim tgNewNamesImported(0 To 0) As NEWNAMESIMPORTED
    For llLoop = LBound(tmUpdateStation) To UBound(tmUpdateStation) - 1 Step 1
        If Len(Trim$(tmUpdateStation(llLoop).sOwner)) > 0 Then
            If mLookupOwnerByName(tmUpdateStation(llLoop).sOwner) < 0 Then
                ilFound = False
                For llNameCheck = LBound(tgNewNamesImported) To UBound(tgNewNamesImported) - 1 Step 1
                    If StrComp(Trim$(tgNewNamesImported(llNameCheck).sNewName), Trim$(tmUpdateStation(llLoop).sOwner), vbTextCompare) = 0 Then
                        tgNewNamesImported(llNameCheck).lUpdateStationIndex = -1
                        tgNewNamesImported(llNameCheck).iCount = tgNewNamesImported(llNameCheck).iCount + 1
                        ilFound = True
                        Exit For
                    End If
                Next llNameCheck
                If Not ilFound Then
                    tgNewNamesImported(UBound(tgNewNamesImported)).sNewName = tmUpdateStation(llLoop).sOwner
                    tgNewNamesImported(UBound(tgNewNamesImported)).lUpdateStationIndex = llLoop
                    tgNewNamesImported(UBound(tgNewNamesImported)).lReplaceCode = 0
                    tgNewNamesImported(UBound(tgNewNamesImported)).iCount = 1
                    ReDim Preserve tgNewNamesImported(0 To UBound(tgNewNamesImported) + 1) As NEWNAMESIMPORTED
                End If
            End If
        End If
    Next llLoop
    If UBound(tgNewNamesImported) > LBound(tgNewNamesImported) Then
        'Determine if records should be added:  New records exist or no reference to records exist
        blBypassMatch = True
        If UBound(tgOwnerInfo) > LBound(tgOwnerInfo) Then
            For llLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
                If tgStationInfo(llLoop).lOwnerCode > 0 Then
                    blBypassMatch = False
                    Exit For
                End If
            Next llLoop
        End If
        If Not blBypassMatch Then
            'Ask user how the new names should be handled
            igNewNamesImportedType = 3
            frmCategoryMatching.Show vbModal
            If igNewNamesImportedReturn = 0 Then
                mCheckOwners = False
                Exit Function
            End If
        End If
        
        If igNewNamesImportedReturn = 2 Then
            For llLoop = LBound(tgNewNamesImported) To UBound(tgNewNamesImported) - 1 Step 1
                blFirst = True
                For llStation = LBound(tmUpdateStation) To UBound(tmUpdateStation) - 1 Step 1
                    If StrComp(Trim$(tgNewNamesImported(llLoop).sNewName), Trim$(tmUpdateStation(llStation).sOwner), vbTextCompare) = 0 Then
                        If blFirst Then
                            'Add Name
                            SQLQuery = "Insert into Artt "
                            SQLQuery = SQLQuery & "(arttType, arttLastName, arttAddress1, arttAddress2, arttCity, "
                            SQLQuery = SQLQuery & "arttAddressState, arttCountry, arttZip, arttPhone, "
                            SQLQuery = SQLQuery & "arttFax, arttEmail, arttEMailRights)"
                            SQLQuery = SQLQuery & " VALUES ('O', '" & gFixQuote(Trim$(tgNewNamesImported(llLoop).sNewName)) & "', '" & "" & "', '" & "" & "', '" & "" & "', "
                            SQLQuery = SQLQuery & "'" & "" & "', '" & "" & "', '" & "" & "', '" & "" & "', "
                            SQLQuery = SQLQuery & "'" & "" & "', '" & "" & "', '" & "N" & "'" & ")"
                            'cnn.Execute SQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/11/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "ImportUpdateStations-mCheckOwners"
                                mCheckOwners = False
                                Exit Function
                            End If
                            ilRet = gPopOwnerNames()
                            mUpdateReport -1, "Owner Name Added :" & Trim$(tgNewNamesImported(llLoop).sNewName) & " on " & tgNewNamesImported(llLoop).iCount & " stations"
                            blFirst = False
                        End If
                        llIndex = mLookupOwnerByName(tgNewNamesImported(llLoop).sNewName)
                        If llIndex <> -1 Then
                            SQLQuery = "Update shtt Set shttOwnerArttCode = " & tgOwnerInfo(llIndex).lCode & " Where shttCode = " & tmUpdateStation(llStation).iCode
                        End If
                    End If
                Next llStation
            Next llLoop
        Else
            For llLoop = LBound(tgNewNamesImported) To UBound(tgNewNamesImported) - 1 Step 1
                If tgNewNamesImported(llLoop).lReplaceCode <= 0 Then
                    'Add Name
                    SQLQuery = "Insert into Artt "
                    SQLQuery = SQLQuery & "(arttType, arttLastName, arttAddress1, arttAddress2, arttCity, "
                    SQLQuery = SQLQuery & "arttAddressState, arttCountry, arttZip, arttPhone, "
                    SQLQuery = SQLQuery & "arttFax, arttEmail, arttEMailRights)"
                    SQLQuery = SQLQuery & " VALUES ('O', '" & gFixQuote(Trim$(tgNewNamesImported(llLoop).sNewName)) & "', '" & "" & "', '" & "" & "', '" & "" & "', "
                    SQLQuery = SQLQuery & "'" & "" & "', '" & "" & "', '" & "" & "', '" & "" & "', "
                    SQLQuery = SQLQuery & "'" & "" & "', '" & "" & "', '" & "N" & "'" & ")"
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mCheckOwners"
                        mCheckOwners = False
                        Exit Function
                    End If
                    mUpdateReport -1, "Owner Name Added :" & Trim$(tgNewNamesImported(llLoop).sNewName) & " on " & tgNewNamesImported(llLoop).iCount & " stations"
                Else
                    'Replace Name
                    llIndex = mLookupOwner(tgNewNamesImported(llLoop).lReplaceCode)
                    If llIndex <> -1 Then
                        slReplacedName = Trim$(tgOwnerInfo(llIndex).sName)
                    End If
                    SQLQuery = "UPDATE artt"
                    SQLQuery = SQLQuery & " SET arttLastName = '" & gFixQuote(Trim$(tgNewNamesImported(llLoop).sNewName)) & "'"
                    SQLQuery = SQLQuery & " WHERE arttCode = " & tgNewNamesImported(llLoop).lReplaceCode
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mCheckOwners"
                        mCheckOwners = False
                        Exit Function
                    End If
                    If llIndex <> -1 Then
                        mUpdateReport -1, "Owner Name: " & slReplacedName & " Replaced by: " & Trim$(tgNewNamesImported(llLoop).sNewName) & " on " & tgNewNamesImported(llLoop).iCount & " stations"
                    End If
                End If
            Next llLoop
            ilRet = gPopOwnerNames()
        End If
    End If
    mCheckOwners = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mCheckOwners"
    mCheckOwners = False
End Function

Private Function mPopFormatLink()
    Dim flt_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim llMax As Long
    Dim ilFmt As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(fltCode) from flt"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim tmFormatLinkInfo(0 To 0) As FORMATLINKINFO
        mPopFormatLink = True
        Exit Function
    End If
    
    llMax = rst(0).Value
    ReDim tmFormatLinkInfo(0 To llMax) As FORMATLINKINFO
    
    SQLQuery = "Select fltCode, fltExtFormatName, fltIntFmtCode from flt "
    Set flt_rst = gSQLSelectCall(SQLQuery)
    ilUpper = 0
    While Not flt_rst.EOF
        tmFormatLinkInfo(ilUpper).iCode = flt_rst!fltCode
        tmFormatLinkInfo(ilUpper).sExtFormatName = flt_rst!fltExtFormatName
        tmFormatLinkInfo(ilUpper).iIntFmtCode = flt_rst!fltIntFmtCode
        ilUpper = ilUpper + 1
        flt_rst.MoveNext
    Wend

    ReDim Preserve tmFormatLinkInfo(0 To ilUpper) As FORMATLINKINFO

    'Add Standard formats
    For ilFmt = 0 To UBound(tgFormatInfo) - 1 Step 1
        tmFormatLinkInfo(ilUpper).iCode = 0
        tmFormatLinkInfo(ilUpper).sExtFormatName = Trim$(tgFormatInfo(ilFmt).sName)
        tmFormatLinkInfo(ilUpper).iIntFmtCode = tgFormatInfo(ilFmt).lCode
        ilUpper = ilUpper + 1
        ReDim Preserve tmFormatLinkInfo(0 To ilUpper) As FORMATLINKINFO
    Next ilFmt
   
    mPopFormatLink = True
    flt_rst.Close
    Exit Function

ErrHand:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mPopFormatLink"
    mPopFormatLink = False
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mLookupFormatLinkByName(slExtFormatName As String) As Long
    Dim llLoop As Long
    Dim llIdx As Long

    mLookupFormatLinkByName = -1
    On Error GoTo ErrHandler
    llIdx = UBound(tmFormatLinkInfo)
    If llIdx < 1 Then
        Exit Function
    End If
    For llLoop = 0 To llIdx - 1 Step 1
        If StrComp(Trim(tmFormatLinkInfo(llLoop).sExtFormatName), Trim(slExtFormatName), vbTextCompare) = 0 Then
            mLookupFormatLinkByName = llLoop
            Exit Function
        End If
    Next
ErrHandler:
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mUpdateCallLetters(llStationIDX As Long, slCallLetters As String, llUpdateStationIndex As Long) As Integer
    Dim slOldCallLetters As String
    Dim slSQLQuery As String
    Dim ilRet As Integer
    
    On Error GoTo ErrHandler:
    mUpdateCallLetters = False
    slOldCallLetters = Trim$(tgStationInfo(llStationIDX).sCallLetters)
    slSQLQuery = "INSERT INTO clt (cltShfCode, cltCallLetters, cltEndDate) "
    slSQLQuery = slSQLQuery & " VALUES ( " & tgStationInfo(llStationIDX).iCode & ", '" & tgStationInfo(llStationIDX).sCallLetters & "', '" & Format$(gNow(), sgSQLDateForm) & "')"

    If bmUpdateDatabase Then
        'cnn.Execute slSQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHandler:
            gHandleError "AffErrorLog.txt", "ImportUpdateStations-mUpdateCallLetters"
            mUpdateCallLetters = False
            Exit Function
        End If
    End If
    slSQLQuery = "Update shtt Set shttCallLetters = '" & Trim$(slCallLetters) & "'" & " Where shttCode = " & tgStationInfo(llStationIDX).iCode
    If bmUpdateDatabase Then
        'cnn.Execute slSQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHandler:
            gHandleError "AffErrorLog.txt", "ImportUpdateStations-mUpdateCallLetters"
            mUpdateCallLetters = False
            Exit Function
        End If
    End If
    
    '07-13-15
    'Update station call letters
    'Add station to EDS and link to network - look to see if any vehicle names match the call letters and then test
    'vehicle options on insertionsif yes the update the link between network and station
    'not sure how to run this
    If gGetEMailDistribution Then
        'ilRet = gChangeStationName(slOldCallLetters, slCallLetters)
        'ilRet = gAddSingleStation(slCallLetters)
    End If
       
    Call mUpdateReport(llUpdateStationIndex, slOldCallLetters & " changed to " & Trim(slCallLetters))
    mUpdateCallLetters = True
    
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mUpdateCallLetters"
    Exit Function
End Function

Private Function mUpdateStationInfo(llStationIDX As Long, llUpdateStationIndex As Long) As Integer
    Dim ilLoop As Integer
'    Dim slTimeZone As String
'    Dim ilTztCode As Integer
'    Dim slMoniker As String
'    Dim llMonikerMntCode As Long
    Dim slWebAddress As String
    Dim slAddr1 As String
    Dim slAddr2 As String
'    Dim slMailCity As String
'    Dim llCityMntCode As Long
'    Dim slMailState As String
    Dim slZip As String
    Dim slCountry As String
    Dim slPhone As String
    Dim slFax As String
'    Dim slCityLic As String
'    Dim llCityLicMntCode As Long
'    Dim slCountyLic As String
'    Dim llCountyLicMntCode As Long
'    Dim slDMAMarket As String
'    Dim llDMAMktCode As Long
'    Dim slMSAMarket As String
'    Dim llMSAMktCode As Long
'    Dim slOwner As String
'    Dim llOwnerCode As Long
'    Dim slFormat As String
'    Dim ilFormatCode As Integer
'    Dim llMktIdx As Long
'    Dim llOwnerIdx As Long
'    Dim llFormatIdx As Long
    Dim ilRet As Integer
    Dim slOldTimeZone As String
    Dim slSQLQuery As String
    
    On Error GoTo ErrHandler:
    mUpdateStationInfo = False
    mSetFields llUpdateStationIndex
    'slSQLQuery = "Update shtt Set "
    slSQLQuery = ""
'    'Time Zone
'    slTimeZone = ""
'    ilTztCode = 0
'    For ilLoop = LBound(tgTimeZoneInfo) To UBound(tgTimeZoneInfo) - 1 Step 1
'        If tmUpdateStation(llUpdateStationIndex).sZone = Left$(tgTimeZoneInfo(ilLoop).sCSIName, 1) Then
'            slTimeZone = tgTimeZoneInfo(ilLoop).sCSIName
'            ilTztCode = tgTimeZoneInfo(ilLoop).iCode
'            Exit For
'        End If
'    Next ilLoop
    '4/1/16: Only update time zone if changed and Log the message
    'If (imTztCode > 0) Or ((imTztCode = 0) And (bmIgnoreBlanks(imMap(ZONE)) = False)) Then
    '    slSQLQuery = slSQLQuery & ", " & "shttTimeZone = '" & Trim$(smTimeZone) & "', "
    '    slSQLQuery = slSQLQuery & "shttTztCode = " & imTztCode
    'End If
    If (imTztCode > 0) Or ((imTztCode = 0) And (bmIgnoreBlanks(imMap(ZONE)) = False)) Then
        If UCase(Left(tgStationInfo(llStationIDX).sZone, 1)) <> UCase(Left(smTimeZone, 1)) Then
            slOldTimeZone = tgStationInfo(llStationIDX).sZone
            ilRet = mZoneChange(tgStationInfo(llStationIDX).iCode, llUpdateStationIndex, UCase(Left(tgStationInfo(llStationIDX).sZone, 1)), UCase(Left(smTimeZone, 1)))
            Call mUpdateReport(llUpdateStationIndex, "Time Zone Changed from " & slOldTimeZone & " to " & Trim(smTimeZone))
            slSQLQuery = slSQLQuery & ", " & "shttTimeZone = '" & Trim$(smTimeZone) & "', "
            slSQLQuery = slSQLQuery & "shttTztCode = " & imTztCode
        End If
    End If
    'Mailing City
'    slMailCity = ""
'    llCityMntCode = 0
'    For ilLoop = LBound(tmCityInfo) To UBound(tmCityInfo) - 1 Step 1
'        If UCase(Trim$(tmUpdateStation(llUpdateStationIndex).sMailCity)) = UCase(Trim$(tmCityInfo(ilLoop).sName)) Then
'            slMailCity = tmCityInfo(ilLoop).sName
'            llCityMntCode = tmCityInfo(ilLoop).lCode
'            Exit For
'        End If
'    Next ilLoop
'    If (slMailCity = "") And (Trim$(tmUpdateStation(llUpdateStationIndex).sMailCity) <> "") Then
'        'Add City
'        llCityMntCode = mAddMultiName("C", Trim$(tmUpdateStation(llUpdateStationIndex).sMailCity))
'        If llCityMntCode = -1 Then
'            llCityMntCode = 0
'            slMailCity = ""
'            Call mUpdateReport(llUpdateStationIndex, "WARNING: City name from Import Not Found. (" & Trim(tmUpdateStation(llUpdateStationIndex).sMailCity) & ")")
'        Else
'            slMailCity = Trim$(tmUpdateStation(llUpdateStationIndex).sMailCity)
'            tmCityInfo(UBound(tmCityInfo)).lCode = llCityMntCode
'            tmCityInfo(UBound(tmCityInfo)).sName = slMailCity
'            tmCityInfo(UBound(tmCityInfo)).sState = "A"
'            ReDim Preserve tmCityInfo(LBound(tmCityInfo) To UBound(tmCityInfo) + 1) As MNTINFO
'        End If
'    End If
    If (smMailCity <> "") Or ((smMailCity = "") And (bmIgnoreBlanks(imMap(MAILCITY)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttCity = '" & gFixQuote(Trim$(smMailCity)) & "' "
        slSQLQuery = slSQLQuery & ", " & "shttCityMntCode = " & lmCityMntCode
    End If
    'City License
'    slCityLic = ""
'    llCityLicMntCode = 0
'    For ilLoop = LBound(tmCityInfo) To UBound(tmCityInfo) - 1 Step 1
'        If UCase(Trim$(tmUpdateStation(llUpdateStationIndex).sCityLicense)) = UCase(Trim$(tmCityInfo(ilLoop).sName)) Then
'            slCityLic = tmCityInfo(ilLoop).sName
'            llCityLicMntCode = tmCityInfo(ilLoop).lCode
'            Exit For
'        End If
'    Next ilLoop
'    If (slCityLic = "") And (Trim$(tmUpdateStation(llUpdateStationIndex).sCityLicense) <> "") Then
'        'Add City
'        llCityLicMntCode = mAddMultiName("C", Trim$(tmUpdateStation(llUpdateStationIndex).sCityLicense))
'        If llCityLicMntCode = -1 Then
'            llCityLicMntCode = 0
'            slCityLic = ""
'            Call mUpdateReport(llUpdateStationIndex, "WARNING: City License name from Import Not Found. (" & Trim(tmUpdateStation(llUpdateStationIndex).sCityLicense) & ")")
'        Else
'            slCityLic = Trim$(tmUpdateStation(llUpdateStationIndex).sCityLicense)
'            tmCityInfo(UBound(tmCityInfo)).lCode = llCityLicMntCode
'            tmCityInfo(UBound(tmCityInfo)).sName = Trim$(tmUpdateStation(llUpdateStationIndex).sCityLicense)
'            tmCityInfo(UBound(tmCityInfo)).sState = "A"
'            ReDim Preserve tmCityInfo(LBound(tmCityInfo) To UBound(tmCityInfo) + 1) As MNTINFO
'        End If
'    End If
    If (smCityLic <> "") Or ((smCityLic = "") And (bmIgnoreBlanks(imMap(CITYLIC)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttCityLic = '" & gFixQuote(Trim$(smCityLic)) & "' "
        slSQLQuery = slSQLQuery & ", " & "shttCityLicMntCode = " & lmCityLicMntCode
    End If
    'County License
'    slCountyLic = ""
'    llCountyLicMntCode = 0
'    For ilLoop = LBound(tmCountyInfo) To UBound(tmCountyInfo) - 1 Step 1
'        If UCase(Trim$(tmUpdateStation(llUpdateStationIndex).sCountyLicense)) = UCase(Trim$(tmCountyInfo(ilLoop).sName)) Then
'            slCountyLic = tmCountyInfo(ilLoop).sName
'            llCountyLicMntCode = tmCountyInfo(ilLoop).lCode
'            Exit For
'        End If
'    Next ilLoop
'    If (slCountyLic = "") And (Trim$(tmUpdateStation(llUpdateStationIndex).sCountyLicense) <> "") Then
'        'Add County
'        llCountyLicMntCode = mAddMultiName("Y", Trim$(tmUpdateStation(llUpdateStationIndex).sCountyLicense))
'        If llCountyLicMntCode = -1 Then
'            llCountyLicMntCode = 0
'            slCountyLic = ""
'            Call mUpdateReport(llUpdateStationIndex, "WARNING: County License name from Import Not Found. (" & Trim(tmUpdateStation(llUpdateStationIndex).sCountyLicense) & ")")
'        Else
'            slCountyLic = Trim$(tmUpdateStation(llUpdateStationIndex).sCountyLicense)
'            tmCountyInfo(UBound(tmCountyInfo)).lCode = llCountyLicMntCode
'            tmCountyInfo(UBound(tmCountyInfo)).sName = slCountyLic
'            tmCountyInfo(UBound(tmCountyInfo)).sState = "A"
'            ReDim Preserve tmCountyInfo(LBound(tmCountyInfo) To UBound(tmCountyInfo) + 1) As MNTINFO
'        End If
'    End If
    If (smCountyLic <> "") Or ((smCountyLic = "") And (bmIgnoreBlanks(imMap(COUNTYLIC)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttCountyLicMntCode = " & lmCountyLicMntCode
    End If
    'Mail State
'    slMailState = ""
'    For ilLoop = LBound(tgStateInfo) To UBound(tgStateInfo) - 1 Step 1
'        If StrComp(Trim$(tgStateInfo(ilLoop).sPostalName), Trim$(tmUpdateStation(llUpdateStationIndex).sMailState), vbTextCompare) = 0 Then
'            slMailState = Trim$(tmUpdateStation(llUpdateStationIndex).sMailState)
'            Exit For
'        End If
'    Next ilLoop
    If (Trim$(smMailState) <> "") Or ((Trim$(smMailState) = "") And (bmIgnoreBlanks(imMap(MAILSTATE)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttState = '" & Trim$(smMailState) & "' "
    End If
    'Moniker
'    slMoniker = ""
'    llMonikerMntCode = 0
'    For ilLoop = LBound(tmMonikerInfo) To UBound(tmMonikerInfo) - 1 Step 1
'        If UCase(Trim$(tmUpdateStation(llUpdateStationIndex).sMoniker)) = UCase(Trim$(tmMonikerInfo(ilLoop).sName)) Then
'            slMoniker = tmMonikerInfo(ilLoop).sName
'            llMonikerMntCode = tmMonikerInfo(ilLoop).lCode
'            Exit For
'        End If
'    Next ilLoop
'    If (slMoniker = "") And (Trim$(tmUpdateStation(llUpdateStationIndex).sMoniker) <> "") Then
'        'Add Moniker
'        llMonikerMntCode = mAddMultiName("M", Trim$(tmUpdateStation(llUpdateStationIndex).sMoniker))
'        If llMonikerMntCode = -1 Then
'            llMonikerMntCode = 0
'            slMoniker = ""
'            Call mUpdateReport(llUpdateStationIndex, "WARNING: Moniker name from Import Not Found. (" & Trim(tmUpdateStation(llUpdateStationIndex).sMoniker) & ")")
'        Else
'            slMoniker = Trim$(tmUpdateStation(llUpdateStationIndex).sMoniker)
'            tmMonikerInfo(UBound(tmMonikerInfo)).lCode = llMonikerMntCode
'            tmMonikerInfo(UBound(tmMonikerInfo)).sName = slMoniker
'            tmMonikerInfo(UBound(tmMonikerInfo)).sState = "A"
'            ReDim Preserve tmMonikerInfo(LBound(tmMonikerInfo) To UBound(tmMonikerInfo) + 1) As MNTINFO
'        End If
'    End If
    If (lmMonikerMntCode > 0) Or ((lmMonikerMntCode = 0) And (bmIgnoreBlanks(imMap(MONIKER)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttMonikerMntCode = " & lmMonikerMntCode
    End If
    'DMA Market
'    slDMAMarket = ""
'    llDMAMktCode = 0
'    llMktIdx = 0
'    If Len(Trim(tmUpdateStation(llUpdateStationIndex).sDMAMarket)) > 0 Then  ' Don't look this up if blank.
'        llMktIdx = mLookupDMAMarketByName(tmUpdateStation(llUpdateStationIndex).sDMAMarket)
'        If llMktIdx <> -1 Then
'            slDMAMarket = Trim$(tmUpdateStation(llUpdateStationIndex).sDMAMarket)
'            llDMAMktCode = tgMarketInfo(llMktIdx).lCode
'        Else
'            Call mUpdateReport(llUpdateStationIndex, "WARNING: DMA Market name from Import missing from Affiliate. (" & Trim(tmUpdateStation(llUpdateStationIndex).sDMAMarket) & ")")
'        End If
'    End If
    If (smDMAMarket <> "") Or ((smDMAMarket = "") And (bmIgnoreBlanks(imMap(DMANAME)) = False)) Then
        If lmDMAMktIdx <> -1 Then
            slSQLQuery = slSQLQuery & ", " & "shttMarket = '" & gFixQuote(Trim$(smDMAMarket)) & "' "
            slSQLQuery = slSQLQuery & ", " & "shttMktCode = " & lmDMAMktCode
            If (Trim$(tmUpdateStation(llUpdateStationIndex).sDMARank) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sDMARank) = "") And (bmIgnoreBlanks(imMap(DMARANK)) = False)) Then
                slSQLQuery = slSQLQuery & ", " & "shttRank = " & tmUpdateStation(llUpdateStationIndex).iDMARank
                If lmDMAMktIdx > 0 Then
                    If tgMarketInfo(lmDMAMktIdx).iRank <> tmUpdateStation(llUpdateStationIndex).iDMARank Then
                        ilRet = mUpdateRank(lmDMAMktIdx, llUpdateStationIndex)
                    End If
                End If
            End If
        End If
    End If
    'MSA Market
'    slMSAMarket = ""
'    llMSAMktCode = 0
'    llMktIdx = 0
'    If Len(Trim(tmUpdateStation(llUpdateStationIndex).sMSAMarket)) > 0 Then  ' Don't look this up if blank.
'        llMktIdx = mLookupMSAMarketByName(tmUpdateStation(llUpdateStationIndex).sMSAMarket)
'        If llMktIdx <> -1 Then
'            slMSAMarket = Trim$(tmUpdateStation(llUpdateStationIndex).sMSAMarket)
'            llMSAMktCode = tgMarketInfo(llMktIdx).lCode
'        Else
'            Call mUpdateReport(llUpdateStationIndex, "WARNING: MSA Market name from Import missing from Affiliate. (" & Trim(tmUpdateStation(llUpdateStationIndex).sMSAMarket) & ")")
'        End If
'    End If
    If (smMSAMarket <> "") Or ((smMSAMarket = "") And (bmIgnoreBlanks(imMap(MSANAME)) = False)) Then
        If lmMSAMktIdx <> -1 Then
            slSQLQuery = slSQLQuery & ", " & "shttMetCode = " & lmMSAMktCode
            If (Trim$(tmUpdateStation(llUpdateStationIndex).sMSARank) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sMSARank) = "") And (bmIgnoreBlanks(imMap(MSARANK)) = False)) Then
                If lmMSAMktIdx > 0 Then
                    If tgMSAMarketInfo(lmMSAMktIdx).iRank <> tmUpdateStation(llUpdateStationIndex).iMSARank Then
                        ilRet = mUpdateMSARank(lmMSAMktIdx, llUpdateStationIndex)
                    End If
                End If
            End If
        End If
    End If
'    slOwner = ""
'    llOwnerCode = 0
'    llOwnerIdx = 0
'    If Len(Trim(tmUpdateStation(llUpdateStationIndex).sOwner)) > 0 Then
'        llOwnerIdx = mLookupOwnerByName(tmUpdateStation(llUpdateStationIndex).sOwner)
'        If llOwnerIdx <> -1 Then
'            slOwner = Trim$(tmUpdateStation(llUpdateStationIndex).sOwner)
'            llOwnerCode = tgOwnerInfo(llOwnerIdx).lCode
'        Else
'            Call mUpdateReport(llUpdateStationIndex, "WARNING: Owner name from Import missing from Affiliate. (" & Trim(tmUpdateStation(llUpdateStationIndex).sOwner) & ")")
'        End If
'    End If
    If (smOwner <> "") Or ((smOwner = "") And (bmIgnoreBlanks(imMap(OWNER)) = False)) Then
        If lmOwnerIdx <> -1 Then
            slSQLQuery = slSQLQuery & ", " & "shttOwnerArttCode = " & lmOwnerCode
        End If
    End If
'    slFormat = ""
'    ilFormatCode = 0
'    llFormatIdx = 0
'    If Len(Trim(tmUpdateStation(llUpdateStationIndex).sFormat)) > 0 Then
'        llFormatIdx = mLookupFormatLinkByName(tmUpdateStation(llUpdateStationIndex).sFormat)
'        If llFormatIdx <> -1 Then
'            slFormat = Trim$(tmUpdateStation(llUpdateStationIndex).sFormat)
'            ilFormatCode = tmFormatLinkInfo(llFormatIdx).iIntFmtCode
'        Else
'            Call mUpdateReport(llUpdateStationIndex, "WARNING: Format name from Import missing from Affiliate. (" & Trim(tmUpdateStation(llUpdateStationIndex).sFormat) & ")")
'        End If
'    End If
    If (smFormat <> "") Or ((smFormat = "") And (bmIgnoreBlanks(imMap(STATIONFORMAT)) = False)) Then
        If lmFormatIdx <> -1 Then
            slSQLQuery = slSQLQuery & ", " & "shttFmtCode = " & imFormatCode
        End If
    End If
    
    'Transact ID
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sTransactID) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sTransactID) = "") And (bmIgnoreBlanks(imMap(ENTERPRISEID)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttVieroID = '" & Trim$(tmUpdateStation(llUpdateStationIndex).sTransactID) & "' "
    End If
    'Frequency
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sFrequency) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sFrequency) = "") And (bmIgnoreBlanks(imMap(FREQUENCY)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttFrequency = '" & Trim$(tmUpdateStation(llUpdateStationIndex).sFrequency) & "' "
    End If
    'Station ID
    slSQLQuery = slSQLQuery & ", " & "shttPermStationID = " & tmUpdateStation(llUpdateStationIndex).lID
    'Watts
    If (tmUpdateStation(llUpdateStationIndex).lWatts > 0) Or ((tmUpdateStation(llUpdateStationIndex).lWatts = 0) And (bmIgnoreBlanks(imMap(WATTS)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttWatts = " & tmUpdateStation(llUpdateStationIndex).lWatts
    End If
    'Person 12 plus
    If (tmUpdateStation(llUpdateStationIndex).lP12Plus > 0) Or ((tmUpdateStation(llUpdateStationIndex).lP12Plus = 0) And (bmIgnoreBlanks(imMap(P12PLUS)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttAudP12Plus = " & tmUpdateStation(llUpdateStationIndex).lP12Plus
    End If
    'Web Address
    slWebAddress = gFixQuote(Trim$(tmUpdateStation(llUpdateStationIndex).sWebAddress))
    If (Trim$(slWebAddress) <> "") Or ((Trim$(slWebAddress) = "") And (bmIgnoreBlanks(imMap(WEBADDR)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttWebAddress = '" & Trim$(slWebAddress) & "'"
    End If
    'Mailing Address 1
    slAddr1 = gFixQuote(Trim$(tmUpdateStation(llUpdateStationIndex).sMailAddress1))
    If (Trim$(slAddr1) <> "") Or ((Trim$(slAddr1) = "") And (bmIgnoreBlanks(imMap(MAILADDR1)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttAddress1 = '" & Trim$(slAddr1) & "'"
    End If
    'Mailing address 2
    slAddr2 = gFixQuote(Trim$(tmUpdateStation(llUpdateStationIndex).sMailAddress2))
    If (Trim$(slAddr2) <> "") Or ((Trim$(slAddr2) = "") And (bmIgnoreBlanks(imMap(MAILADDR2)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttAddress2 = '" & Trim$(slAddr2) & "'"
    End If
    'Mailing Zip
    slZip = Trim$(tmUpdateStation(llUpdateStationIndex).sMailZip)
    If (Trim$(slZip) <> "") Or ((Trim$(slZip) = "") And (bmIgnoreBlanks(imMap(MAILZIP)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttZip = '" & Trim$(slZip) & "'"
    End If
    'Country
    slCountry = gFixQuote(Trim$(tmUpdateStation(llUpdateStationIndex).sMailCountry))
    If (Trim$(slCountry) <> "") Or ((Trim$(slCountry) = "") And (bmIgnoreBlanks(imMap(MAILCOUNTRY)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttCountry = '" & Trim$(slCountry) & "'"
    End If
    'Phone
    slPhone = Trim$(tmUpdateStation(llUpdateStationIndex).sPhone)
    If (Trim$(slPhone) <> "") Or ((Trim$(slPhone) = "") And (bmIgnoreBlanks(imMap(PHONE)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttPhone = '" & Trim$(slPhone) & "'"
    End If
    'Fax
    slFax = Trim$(tmUpdateStation(llUpdateStationIndex).sFax)
    If (Trim$(slFax) <> "") Or ((Trim$(slFax) = "") And (bmIgnoreBlanks(imMap(FAX)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttFax = '" & Trim$(slFax) & "'"
    End If
    
    'Territory
    If (lmTerritoryMntCode > 0) Or ((lmTerritoryMntCode = 0) And (bmIgnoreBlanks(imMap(TERRITORY)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttMntCode = " & lmTerritoryMntCode
    End If
    'Area
    If (lmAreaMntCode > 0) Or ((lmAreaMntCode = 0) And (bmIgnoreBlanks(imMap(AREA)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttAreaMntCode = " & lmAreaMntCode
    End If
    'State License
    If (smStateLic <> "") Or ((smStateLic = "") And (bmIgnoreBlanks(imMap(STATELIC)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttStateLic = '" & gFixQuote(Trim$(smStateLic)) & "' "
    End If
    'Operator
    If (lmOperatorMntCode > 0) Or ((lmOperatorMntCode = 0) And (bmIgnoreBlanks(imMap(OPERATOR)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttOperatorMntCode = " & lmOperatorMntCode
    End If
    'Market Rep
    If (imMarketRepUstCode > 0) Or ((imMarketRepUstCode = 0) And (bmIgnoreBlanks(imMap(MARKETREP)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttMktRepUstCode = " & imMarketRepUstCode
    End If
    'Service Rep
    If (imServiceRepUstCode > 0) Or ((imServiceRepUstCode = 0) And (bmIgnoreBlanks(imMap(SERVICEREP)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttServRepUstCode = " & imServiceRepUstCode
    End If
    'Daylight
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sDaylight) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sDaylight) = "") And (bmIgnoreBlanks(imMap(DAYLIGHT)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttAckDaylight = " & imDaylight
    End If
    'XDS ID
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sXDSStationID) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sXDSStationID) = "") And (bmIgnoreBlanks(imMap(XDSID)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttStationID = " & tmUpdateStation(llUpdateStationIndex).lXDSStationID
    End If
    'iPump ID
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sIPumpID) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sIPumpID) = "") And (bmIgnoreBlanks(imMap(IPUMPID)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttIPumpID = '" & tmUpdateStation(llUpdateStationIndex).sIPumpID & "' "
    End If
    'Serial #1
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sSerial1) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sSerial1) = "") And (bmIgnoreBlanks(imMap(SERIAL1)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttSerialNo1 = '" & Trim$(tmUpdateStation(llUpdateStationIndex).sSerial1) & "'"
    End If
    'Serial #2
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sSerial2) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sSerial2) = "") And (bmIgnoreBlanks(imMap(SERIAL2)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttSerialNo2 = '" & Trim$(tmUpdateStation(llUpdateStationIndex).sSerial2) & "'"
    End If
    'On Air
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sOnAir) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sOnAir) = "") And (bmIgnoreBlanks(imMap(OnAir)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttOnAir = '" & smOnAir & "'"
    End If
    'Commercial
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sCommercial) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sCommercial) = "") And (bmIgnoreBlanks(imMap(COMMERCIAL)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttStationType = '" & smCommercial & "'"
    End If
    'Used for XDS
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sUsedXDS) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sUsedXDS) = "") And (bmIgnoreBlanks(imMap(USEXDS)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttUsedForXDigital = '" & smUsedXDS & "'"
    End If
    'Used for Wegener
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sUsedWegener) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sUsedWegener) = "") And (bmIgnoreBlanks(imMap(USEWEGENER)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttUsedForWegener = '" & smUsedWegener & "'"
    End If
    'Used for OLA
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sUsedOLA) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sUsedOLA) = "") And (bmIgnoreBlanks(imMap(USEOLA)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttUsedForOLA = '" & smUsedOLA & "'"
    End If
    'Web Password
    If (Trim$(tmUpdateStation(llUpdateStationIndex).sWebPassword) <> "") Or ((Trim$(tmUpdateStation(llUpdateStationIndex).sWebPassword) = "") And (bmIgnoreBlanks(imMap(WEBPW)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttWebPW = '" & Trim$(tmUpdateStation(llUpdateStationIndex).sWebPassword) & "'"
    End If
    'Physical Address 1
    slAddr1 = gFixQuote(Trim$(tmUpdateStation(llUpdateStationIndex).sPhysicalAddress1))
    If (Trim$(slAddr1) <> "") Or ((Trim$(slAddr1) = "") And (bmIgnoreBlanks(imMap(PHYSICALADDR1)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttONAddress1 = '" & Trim$(slAddr1) & "'"
    End If
    'Physical address 2
    slAddr2 = gFixQuote(Trim$(tmUpdateStation(llUpdateStationIndex).sPhysicalAddress2))
    If (Trim$(slAddr2) <> "") Or ((Trim$(slAddr2) = "") And (bmIgnoreBlanks(imMap(PHYSICALADDR2)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttONAddress2 = '" & Trim$(slAddr2) & "'"
    End If
    'Physical City
    If (smPhysicalCity <> "") Or ((smPhysicalCity = "") And (bmIgnoreBlanks(imMap(PHYSICALCITY)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttOnCity = '" & gFixQuote(Trim$(smPhysicalCity)) & "' "
        slSQLQuery = slSQLQuery & ", " & "shttOnCityMntCode = " & lmPhysicalCityMntCode
    End If
    'Physical State
    If (Trim$(smPhysicalState) <> "") Or ((Trim$(smPhysicalState) = "") And (bmIgnoreBlanks(imMap(PHYSICALSTATE)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttOnState = '" & Trim$(smPhysicalState) & "' "
    End If
    'Physical Zip
    slZip = Trim$(tmUpdateStation(llUpdateStationIndex).sPhysicalZip)
    If (Trim$(slZip) <> "") Or ((Trim$(slZip) = "") And (bmIgnoreBlanks(imMap(PHYSICALZIP)) = False)) Then
        slSQLQuery = slSQLQuery & ", " & "shttONZip = '" & Trim$(slZip) & "'"
    End If
    
    'User
    slSQLQuery = slSQLQuery & ", " & "shttUsfCode = " & igUstCode

    If Left(slSQLQuery, 1) = "," Then
        slSQLQuery = Trim$(Mid(slSQLQuery, 2))
    End If
    slSQLQuery = "Update shtt Set " & slSQLQuery
    slSQLQuery = slSQLQuery & " Where shttCode = " & tgStationInfo(llStationIDX).iCode
    If bmUpdateDatabase Then
        'cnn.Execute slSQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHandler:
            gHandleError "AffErrorLog.txt", "ImportUpdateStations-mUpdateStationInfo"
            mUpdateStationInfo = False
            Exit Function
        End If
        ilRet = mPersonnel(llUpdateStationIndex)
        '10/3/18: Dan- In the routine gVatSetToGoToWebByShttCode I will ignore the VendorID  Dan no longer ignored
        'TTP 8824 reopened
        '7942
        gVatSetToGoToWebByShttCode tgStationInfo(llStationIDX).iCode, Vendors.XDS_Break
    End If
    mUpdateStationInfo = True
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mUpdateStationInfo"
    Exit Function
End Function

Private Function mRemoveStationMSAMarket(llStationIDX As Long, llUpdateStationIndex As Long) As Integer
    On Error GoTo ErrHandler:
    mRemoveStationMSAMarket = False
    SQLQuery = "Select shttMetCode From shtt Where shttCode = " & tgStationInfo(llStationIDX).iCode
    Set rst_Shtt = gSQLSelectCall(SQLQuery)
    If rst_Shtt!shttMetCode > 0 Then
    
        SQLQuery = "Update shtt Set shttMetCode = " & 0 & " Where shttCode = " & tgStationInfo(llStationIDX).iCode
        If bmUpdateDatabase Then
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHandler:
                gHandleError "AffErrorLog.txt", "ImportUpdateStations-mRemoveStationMSAMarket"
                mRemoveStationMSAMarket = False
                Exit Function
            End If
        End If
    
        Call mUpdateReport(llUpdateStationIndex, "MSA Market was removed as name missing from Import file")
    End If
    mRemoveStationMSAMarket = True
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mRemoveStationMSAMarket"
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Sub mShowReport()
    Dim slCmd As String
    Dim slDateTime As String
    Dim ilRet As Integer
    
    'On Error GoTo ErrHandler
    ilRet = 0
    'slDateTime = FileDateTime(smReportPathFileName)
    ilRet = gFileExist(smReportPathFileName)
    If ilRet <> 0 Then
        Exit Sub
    End If
    ilRet = MsgBox("View Result File?", vbApplicationModal + vbInformation + vbYesNo, "Question")
    If ilRet = vbNo Then
        Exit Sub
    End If
    slCmd = "Notepad.exe " & smReportPathFileName
    Call Shell(slCmd, vbNormalFocus)
    Exit Sub
    
'ErrHandler:
'    ilRet = -1
'    Resume Next
End Sub

Private Function mCreateMap(slFields() As String) As Boolean
    Dim ilLoop As Integer
    Dim ilPerson As Integer
    
    For ilLoop = CALLLETTERS To PERSON + 5 * 8 - 1 Step 1
        imMap(ilLoop) = mFindFieldName(sgStationImportTitles(ilLoop), slFields())
    Next ilLoop
    mCreateMap = True
    For ilLoop = LBound(imMap) To UBound(imMap) Step 1
        If imMap(ilLoop) = -1 Then
            mCreateMap = False
            Exit Function
        End If
    Next ilLoop
    imPersonTitles = 0
    For ilPerson = 1 To 5 Step 1
        If imMap(PERSON + PNAME + 8 * (ilPerson - 1)) > 0 Then
            imPersonTitles = ilPerson
        End If
        If imMap(PERSON + PPHONE + 8 * (ilPerson - 1)) > 0 Then
            imPersonTitles = ilPerson
        End If
        If imMap(PERSON + PEMAIL + 8 * (ilPerson - 1)) > 0 Then
            imPersonTitles = ilPerson
        End If
    Next ilPerson
End Function

Private Function mFindFieldName(slFindName As String, slFields() As String) As Integer
    Dim ilLoop As Integer

    For ilLoop = LBound(slFields) + 1 To UBound(slFields) Step 1
        If StrComp(Trim$(UCase(slFields(ilLoop))), Trim$(UCase(slFindName)), vbTextCompare) = 0 Then
            mFindFieldName = ilLoop
            Exit Function
        End If
    Next ilLoop
    mFindFieldName = 0
    Exit Function
        
End Function

Private Function mCheckStationName(slCallLetters As String) As Boolean
    Dim slSQLQuery As String
    
    On Error GoTo ErrHandler:
    mCheckStationName = False
    slSQLQuery = "Select shttCode From shtt Where shttCallLetters = " & slCallLetters
    Set rst_Shtt = gSQLSelectCall(slSQLQuery)
    If Not rst_Shtt.EOF Then
        mCheckStationName = True
    End If
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mCheckStationName"
    Exit Function
End Function

Private Function mAddStation(llAddStationIndex As Long) As Integer
'    Dim slTimeZone As String
'    Dim ilTztCode As Integer
'    Dim slMailCity As String
'    Dim llCityMntCode As Long
'    Dim slCityLic As String
'    Dim llCityLicMntCode As Long
'    Dim slCountyLic As String
'    Dim llCountyLicMntCode As Long
'    Dim slMoniker As String
'    Dim llMonikerMntCode As Long
    Dim ilLoop As Integer
    Dim ilCode As Integer
'    Dim llDMAMktCode As Long
'    Dim llMktIdx As Long
'    Dim llMSAMktCode As Long
'    Dim llOwnerCode As Long
'    Dim llOwnerIdx As Long
'    Dim ilFormatCode As Integer
'    Dim llFormatIdx As Long
'    Dim slMailState As String
    Dim ilRet As Integer
    Dim ilField As Integer
    Dim blBandAllowed As Boolean
    
'    Dim slTerritory As String
'    Dim llTerritoryMntCode As Long
'    Dim slArea As String
'    Dim llAreaMntCode As Long
'    Dim slStateLic As String
'    Dim slOperator As String
'    Dim llOperatorMntCode As Long
'    Dim slMarketRep As String
'    Dim ilMarketRepUstCode As Integer
'    Dim slServiceRep As String
'    Dim ilServiceRepUstCode As Integer
'    Dim slOnAir As String * 1
'    Dim slCommercial As String * 1
'    Dim ilDaylight As Integer   '0=Yes; 1=No
'    Dim slUsedAgreement As String * 1
'    Dim slUsedXDS As String * 1
'    Dim slUsedWegener As String * 1
'    Dim slUsedOLA As String * 1
'    Dim slHistoricalDate As String
'    Dim llPhysicalMntCode As Long
'    Dim slPhysicalCity As String
'    Dim slPhysicalState As String
    
    On Error GoTo ErrHandler:
    
    mAddStation = False

    mSetFields llAddStationIndex

'    'Time zone
'    slTimeZone = ""
'    ilTztCode = 0
'    For ilLoop = LBound(tgTimeZoneInfo) To UBound(tgTimeZoneInfo) - 1 Step 1
'        If tmUpdateStation(llAddStationIndex).sZone = Left$(tgTimeZoneInfo(ilLoop).sCSIName, 1) Then
'            slTimeZone = tgTimeZoneInfo(ilLoop).sCSIName
'            ilTztCode = tgTimeZoneInfo(ilLoop).iCode
'            Exit For
'        End If
'    Next ilLoop
    blBandAllowed = False
    If IsArray(smBandFields) Then
        For ilField = 0 To UBound(smBandFields)
            If InStr(1, smBandFields(ilField), "-") > 0 Then
                If InStr(1, UCase$(tmUpdateStation(llAddStationIndex).sCallLetters), smBandFields(ilField), vbTextCompare) > 0 Then
                    blBandAllowed = True
                    Exit For
                End If
            Else
                If InStr(1, UCase$(tmUpdateStation(llAddStationIndex).sCallLetters), "-" & smBandFields(ilField), vbTextCompare) > 0 Then
                    blBandAllowed = True
                    Exit For
                End If
            End If
        Next ilField
    End If
    'If (InStr(1, UCase$(tmUpdateStation(llAddStationIndex).sCallLetters), "-AM", vbTextCompare) <= 0) And (InStr(1, UCase$(tmUpdateStation(llAddStationIndex).sCallLetters), "-FM", vbTextCompare) <= 0) Then
    If (blBandAllowed = False) Then
        Call mUpdateReport(llAddStationIndex, "WARNING: Station band is not allowed, Station not added. (" & Trim(tmUpdateStation(llAddStationIndex).sCallLetters) & ")")
        If Not bmNotAddedMsg Then
            mSetResults "Some Stations not Added", RGB(0, 0, 0)
            bmNotAddedMsg = True
        End If
        Exit Function
    End If
    If bmMatchOnPermStationID And (tmUpdateStation(llAddStationIndex).lID <= 0) Then
        Call mUpdateReport(llAddStationIndex, "WARNING: Station ID missing or zero, Station not added. (" & Trim(tmUpdateStation(llAddStationIndex).sCallLetters) & ")")
        If Not bmNotAddedMsg Then
            mSetResults "Some Stations not Added", RGB(0, 0, 0)
            bmNotAddedMsg = True
        End If
        Exit Function
    End If
    If smTimeZone = "" Then
        Call mUpdateReport(llAddStationIndex, "WARNING: Time zone missing or not found, Station not added. (" & Trim(tmUpdateStation(llAddStationIndex).sCallLetters) & ")")
        If Not bmNotAddedMsg Then
            mSetResults "Some Stations not Added", RGB(0, 0, 0)
            bmNotAddedMsg = True
        End If
        Exit Function
    End If

    If lmDMAMktCode = 0 Then
        Call mUpdateReport(llAddStationIndex, "WARNING: DMA Market missing or not found, Station not added. (" & Trim(tmUpdateStation(llAddStationIndex).sCallLetters) & ")")
        If Not bmNotAddedMsg Then
            mSetResults "Some Stations not Added", RGB(0, 0, 0)
            bmNotAddedMsg = True
        End If
        Exit Function
    End If
    
    'Debug.Print " -> AddStation: " & Trim(tmUpdateStation(llAddStationIndex).sCallLetters)
    SQLQuery = "Insert Into shtt ( "
    SQLQuery = SQLQuery & "shttCode, "
    SQLQuery = SQLQuery & "shttCallLetters, "
    SQLQuery = SQLQuery & "shttAddress1, "
    SQLQuery = SQLQuery & "shttAddress2, "
    SQLQuery = SQLQuery & "shttCity, "
    SQLQuery = SQLQuery & "shttState, "
    SQLQuery = SQLQuery & "shttCountry, "
    SQLQuery = SQLQuery & "shttZip, "
    SQLQuery = SQLQuery & "shttSelected, "
    SQLQuery = SQLQuery & "shttEmail, "
    SQLQuery = SQLQuery & "shttFax, "
    SQLQuery = SQLQuery & "shttPhone, "
    SQLQuery = SQLQuery & "shttTimeZone, "
    SQLQuery = SQLQuery & "shttHomePage, "
    SQLQuery = SQLQuery & "shttPDName, "
    SQLQuery = SQLQuery & "shttPDPhone, "
    SQLQuery = SQLQuery & "shttIPumpID, "
    SQLQuery = SQLQuery & "shttTDName, "
    SQLQuery = SQLQuery & "shttTDPhone, "
    SQLQuery = SQLQuery & "shttOnCityMntCode, "
    SQLQuery = SQLQuery & "shttOnCountry, "
    SQLQuery = SQLQuery & "shttCityMntCode, "
    SQLQuery = SQLQuery & "shttCityLicMntCode, "
    SQLQuery = SQLQuery & "shttCountyLicMntCode, "
    SQLQuery = SQLQuery & "shttMDName, "
    SQLQuery = SQLQuery & "shttAgreementExist, "
    SQLQuery = SQLQuery & "shttCommentExist, "
    SQLQuery = SQLQuery & "shttMktRepUstCode, "
    SQLQuery = SQLQuery & "shttServRepUstCode, "
    SQLQuery = SQLQuery & "shttOperatorMntCode, "
    SQLQuery = SQLQuery & "shttAreaMntCode, "
    SQLQuery = SQLQuery & "shttHistStartDate, "
    SQLQuery = SQLQuery & "shttAudP12Plus, "
    SQLQuery = SQLQuery & "shttMonikerMntCode, "
    SQLQuery = SQLQuery & "shttMultiCastGroupID, "
    SQLQuery = SQLQuery & "shttClusterGroupID, "
    SQLQuery = SQLQuery & "shttMasterCluster, "
    SQLQuery = SQLQuery & "shttVieroID, "
    SQLQuery = SQLQuery & "shttFrequency, "
    SQLQuery = SQLQuery & "shttPermStationID, "
    SQLQuery = SQLQuery & "shttOnAir, "
    SQLQuery = SQLQuery & "shttStationType, "
    SQLQuery = SQLQuery & "shttWatts, "
    SQLQuery = SQLQuery & "shttUnused, "
    SQLQuery = SQLQuery & "shttACName, "
    SQLQuery = SQLQuery & "shttACPhone, "
    SQLQuery = SQLQuery & "shttMntCode, "
    SQLQuery = SQLQuery & "shttChecked, "
    SQLQuery = SQLQuery & "shttMarket, "
    SQLQuery = SQLQuery & "shttRank, "
    SQLQuery = SQLQuery & "shttUsfCode, "
    SQLQuery = SQLQuery & "shttEnterDate, "
    SQLQuery = SQLQuery & "shttEnterTime, "
    SQLQuery = SQLQuery & "shttType, "
    SQLQuery = SQLQuery & "shttONAddress1, "
    SQLQuery = SQLQuery & "shttONAddress2, "
    SQLQuery = SQLQuery & "shttONCity, "
    SQLQuery = SQLQuery & "shttONState, "
    SQLQuery = SQLQuery & "shttONZip, "
    SQLQuery = SQLQuery & "shttStationID, "
    SQLQuery = SQLQuery & "shttCityLic, "
    SQLQuery = SQLQuery & "shttStateLic, "
    SQLQuery = SQLQuery & "shttAckDaylight, "
    SQLQuery = SQLQuery & "shttWebEmail, "
    SQLQuery = SQLQuery & "shttWebPW, "
    SQLQuery = SQLQuery & "shttOwnerArttCode, "
    SQLQuery = SQLQuery & "shttMktCode, "
    SQLQuery = SQLQuery & "shttWebAddress, "
    SQLQuery = SQLQuery & "shttfmtCode, "
    SQLQuery = SQLQuery & "shttSerialNo1, "
    SQLQuery = SQLQuery & "shttSerialNo2, "
    SQLQuery = SQLQuery & "shttTztCode, "
    SQLQuery = SQLQuery & "shttWebNumber, "
    SQLQuery = SQLQuery & "shttUsedForAtt, "
    SQLQuery = SQLQuery & "shttUsedForXDigital, "
    SQLQuery = SQLQuery & "shttUsedForWegener, "
    SQLQuery = SQLQuery & "shttUsedForOLA, "
    SQLQuery = SQLQuery & "shttPort, "
    SQLQuery = SQLQuery & "shttMetCode, "
    SQLQuery = SQLQuery & "shttSpotsPerWebPage "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "  'shttCode
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sCallLetters)) & "', "  'shttCallLetters
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sMailAddress1)) & "', " 'shttAddress1
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sMailAddress2)) & "', " 'shttAddress2
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(smMailCity)) & "', "     'shttCity
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(smMailState)) & "', "    'shttState
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sMailCountry)) & "', "  'shttCountry
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sMailZip)) & "', "      'shttZip
    SQLQuery = SQLQuery & -1 & ", " 'shttSelected
    SQLQuery = SQLQuery & "'" & "" & "', "  'shttEmail
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sFax)) & "', "  'shttFax
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sPhone)) & "', "    'shttPhone
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(smTimeZone)) & "', "   'shttTimeZone
    SQLQuery = SQLQuery & "'" & "" & "', "  'shttHomePage
    SQLQuery = SQLQuery & "'" & "" & "', "  'shttPDName
    SQLQuery = SQLQuery & "'" & "" & "', "  'shttPDPhone
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sIPumpID)) & "', "   'shttIPumpID
    SQLQuery = SQLQuery & "'" & "" & "', "  'shttTDName
    SQLQuery = SQLQuery & "'" & "" & "', "  'shttTDPhone
    SQLQuery = SQLQuery & lmPhysicalCityMntCode & ", "          'shttOnCityMntCode
    SQLQuery = SQLQuery & "'" & "" & "', "  'shttOnCountry
    SQLQuery = SQLQuery & lmCityMntCode & ", "  'shttCityMntCode
    SQLQuery = SQLQuery & lmCityLicMntCode & ", "   'shttCityLicMntCode
    SQLQuery = SQLQuery & lmCountyLicMntCode & ", " 'shttCountyLicMntCode
    SQLQuery = SQLQuery & "'" & gFixQuote("") & "', "   'shttMDName
    SQLQuery = SQLQuery & "'" & "N" & "', "             'shttAgreementExist
    SQLQuery = SQLQuery & "'" & "N" & "', "             'shttCommentExist
    SQLQuery = SQLQuery & imMarketRepUstCode & ", "              'shttMktRepUstCode
    SQLQuery = SQLQuery & imServiceRepUstCode & ", "              'shttServRepUstCode
    SQLQuery = SQLQuery & lmOperatorMntCode & ", "              'shttOperatorMntVode
    SQLQuery = SQLQuery & lmAreaMntCode & ", "              'shttAreaMntCode
    SQLQuery = SQLQuery & "'" & Format$(smHistoricalDate, sgSQLDateForm) & "', "  'shttHistStartDate
    SQLQuery = SQLQuery & tmUpdateStation(llAddStationIndex).lP12Plus & ", "    'shttAudP12Plus
    SQLQuery = SQLQuery & lmMonikerMntCode & ", "   'shttMonikerMntCode
    SQLQuery = SQLQuery & 0 & ", "                  'shttMultiCastGroupID
    SQLQuery = SQLQuery & 0 & ", "                  'shttClusterGroupID
    SQLQuery = SQLQuery & "'" & "N" & "', "         'shttMasterCluster
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sTransactID)) & "', "          'shttVieroID
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sFrequency)) & "', "    'shttFrequency
    SQLQuery = SQLQuery & tmUpdateStation(llAddStationIndex).lID & ", "     'shttPermStatrtID
    SQLQuery = SQLQuery & "'" & smOnAir & "', "     'shttOnAir
    SQLQuery = SQLQuery & "'" & smCommercial & "', "     'shttStationType
    SQLQuery = SQLQuery & tmUpdateStation(llAddStationIndex).lWatts & ", "      'shttWatts
    SQLQuery = SQLQuery & "'" & "" & "', "      'shttUnused
    SQLQuery = SQLQuery & "'" & "" & "', "      'shttACName
    SQLQuery = SQLQuery & "'" & "" & "', "      'shttACPhone
    SQLQuery = SQLQuery & lmTerritoryMntCode & ", "              'shttMntCode
    SQLQuery = SQLQuery & -1 & ", "             'shttChecked
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sDMAMarket)) & "', "  'shttMarket
    SQLQuery = SQLQuery & tmUpdateStation(llAddStationIndex).iDMARank & ", "    'shttRank
    SQLQuery = SQLQuery & igUstCode & ", "  'shttUsfCode
    SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "  'shttEnterDate
    SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLTimeForm) & "', "  'shttEnterTime
    SQLQuery = SQLQuery & 0 & ", "  'shttType
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sPhysicalAddress1)) & "', "      'shttOnAddress1
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sPhysicalAddress2)) & "', "      'shttOnAddress2
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(smPhysicalCity)) & "', "      'shttOnCity
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(smPhysicalState)) & "', "      'shttOnState
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sPhysicalZip)) & "', "      'shttOnZip
    SQLQuery = SQLQuery & tmUpdateStation(llAddStationIndex).lXDSStationID & ", "              'shttStationID
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(smCityLic)) & "', "    'shttCityLic
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(smStateLic)) & "', "   'State Lic
    SQLQuery = SQLQuery & imDaylight & ", "  'shttAckDaylight
    SQLQuery = SQLQuery & "'" & "" & "', "  'WebEmail
    SQLQuery = SQLQuery & "'" & Trim$(tmUpdateStation(llAddStationIndex).sWebPassword) & "', "  'WebPW
    SQLQuery = SQLQuery & lmOwnerCode & ", "    'shttOwnerArttCode
    SQLQuery = SQLQuery & lmDMAMktCode & ", "   'shttMktCode
    SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(tmUpdateStation(llAddStationIndex).sWebAddress)) & "', "      'shttWebAddress
    SQLQuery = SQLQuery & imFormatCode & ", "   'shttFmtCode
    SQLQuery = SQLQuery & "'" & Trim$(tmUpdateStation(llAddStationIndex).sSerial1) & "', "      'shttSerialNo1
    SQLQuery = SQLQuery & "'" & Trim$(tmUpdateStation(llAddStationIndex).sSerial2) & "', "      'shttSerialNo2
    SQLQuery = SQLQuery & imTztCode & ", "      'shttTztCode
    '4/14/21: TTP 9052
    If Trim$(sgWebNumber) = "" Then
        sgWebNumber = "1"
    End If
    SQLQuery = SQLQuery & "'" & sgWebNumber & "', "     '"1" & "', "     'shttWebNumber
    SQLQuery = SQLQuery & "'" & smUsedAgreement & "', "     'shttUsedForAtt
    SQLQuery = SQLQuery & "'" & smUsedXDS & "', "     'shttUsedForXDigital
    SQLQuery = SQLQuery & "'" & smUsedWegener & "', "     'shttUsedForWegener
    SQLQuery = SQLQuery & "'" & smUsedOLA & "', "     'shttUsedForOLA
    SQLQuery = SQLQuery & "'" & "" & "', "      'shttPort
    SQLQuery = SQLQuery & lmMSAMktCode & ", "                     'shttMetCode
    SQLQuery = SQLQuery & 0 'Spots per Web Page
    SQLQuery = SQLQuery & ") "
        
    ilCode = CInt(gInsertAndReturnCode(SQLQuery, "shtt", "shttCode", "Replace"))
    If ilCode > 0 Then
        tmUpdateStation(llAddStationIndex).iCode = ilCode
        tgStationInfo(UBound(tgStationInfo)).sCallLetters = tmUpdateStation(llAddStationIndex).sCallLetters
        tgStationInfo(UBound(tgStationInfo)).lPermStationID = tmUpdateStation(llAddStationIndex).lID
        ReDim Preserve tgStationInfo(0 To UBound(tgStationInfo) + 1) As STATIONINFO
        ilRet = mPersonnel(llAddStationIndex)
        mAddStation = True
    End If
    
    '07-13-15
    'Add station to EDS and link to network - look to see if any vehicle names match the call letters and then test
    'vehicle options on insertionsif yes the update the link between network and station
    If gGetEMailDistribution Then
        'ilRet = gAddSingleStation(tmUpdateStation(llAddStationIndex).sCallLetters)
    End If
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mAddStation"
    Exit Function
End Function

Private Function mAddMultiName(slType As String, slName As String) As Long
    Dim llCode As Long
    Dim slSQLQuery As String
    
    On Error GoTo ErrHandler:
    mAddMultiName = -1
    slSQLQuery = "Insert Into mnt ( "
    slSQLQuery = slSQLQuery & "mntCode, "
    slSQLQuery = slSQLQuery & "mntType, "
    slSQLQuery = slSQLQuery & "mntName, "
    slSQLQuery = slSQLQuery & "mntState, "
    slSQLQuery = slSQLQuery & "mntUnused "
    slSQLQuery = slSQLQuery & ") "
    slSQLQuery = slSQLQuery & "Values ( "
    slSQLQuery = slSQLQuery & "Replace" & ", "
    slSQLQuery = slSQLQuery & "'" & slType & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(slName) & "', "
    slSQLQuery = slSQLQuery & "'" & "A" & "', "
    slSQLQuery = slSQLQuery & "'" & "" & "' "
    slSQLQuery = slSQLQuery & ") "
    llCode = CInt(gInsertAndReturnCode(slSQLQuery, "mnt", "mntCode", "Replace"))
    If llCode > 0 Then
        mAddMultiName = llCode
    End If
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mAddMultiName"
    Exit Function
End Function

Private Sub mBuildImportTitles()
    Dim ilLoop As Integer
    ReDim sgStationImportTitles(0 To 0) As String
    mAddImportTitles CALLLETTERS, "Call Letters"
    mAddImportTitles ID, "ID#"
    mAddImportTitles FREQUENCY, "Frequency"
    mAddImportTitles TERRITORY, "Territory"
    mAddImportTitles AREA, "Area"
    mAddImportTitles STATIONFORMAT, "Format"
    mAddImportTitles DMARANK, "DMA Rank"
    mAddImportTitles DMANAME, "DMA Name"
    mAddImportTitles CITYLIC, "City of License"
    mAddImportTitles COUNTYLIC, "County of License"
    mAddImportTitles STATELIC, "State of License"
    mAddImportTitles OWNER, "Owner"
    mAddImportTitles OPERATOR, "Operator"
    mAddImportTitles MSARANK, "MSA Rank"
    mAddImportTitles MSANAME, "MSA Name"
    mAddImportTitles MARKETREP, "Market Rep"
    mAddImportTitles SERVICEREP, "Service Rep"
    mAddImportTitles ZONE, "Zone"
    mAddImportTitles OnAir, "On air"
    mAddImportTitles COMMERCIAL, "Commercial"
    mAddImportTitles DAYLIGHT, "Honor Daylight Saving"
    mAddImportTitles XDSID, "XDS Station ID"
    mAddImportTitles IPUMPID, "iPump Station ID"
    mAddImportTitles SERIAL1, "Serial # 1"
    mAddImportTitles SERIAL2, "Serial # 2"
    mAddImportTitles USEAGREEMENT, "Used for Agreements"
    mAddImportTitles USEXDS, "Used for XDS"
    mAddImportTitles USEWEGENER, "Used for Wegener"
    mAddImportTitles USEOLA, "Used for OLA"
    mAddImportTitles MONIKER, "Moniker"
    mAddImportTitles WATTS, "Watts"
    mAddImportTitles HISTORICALDATE, "Historical Start Date"
    mAddImportTitles P12PLUS, "P12+"
    mAddImportTitles WEBADDR, "Web Address"
    mAddImportTitles WEBPW, "Web Password"
    mAddImportTitles ENTERPRISEID, "Transact Enterprise ID"
    mAddImportTitles MAILADDR1, "Mailing Address 1"
    mAddImportTitles MAILADDR2, "Mailing Address 2"
    mAddImportTitles MAILCITY, "Mailing City"
    mAddImportTitles MAILSTATE, "Mailing State"
    mAddImportTitles MAILZIP, "Mailing Zip"
    mAddImportTitles MAILCOUNTRY, "Mailing Country"
    mAddImportTitles PHYSICALADDR1, "Physical Address 1"
    mAddImportTitles PHYSICALADDR2, "Physical Address 2"
    mAddImportTitles PHYSICALCITY, "Physical City"
    mAddImportTitles PHYSICALSTATE, "Physical State"
    mAddImportTitles PHYSICALZIP, "Physical Zip"
    mAddImportTitles PHONE, "Phone"
    mAddImportTitles FAX, "Fax"
    For ilLoop = 1 To 5 Step 1
        mAddImportTitles PERSON + 8 * (ilLoop - 1), "Person " & Trim(Str(ilLoop)) & " Name"
        mAddImportTitles PERSON + 8 * (ilLoop - 1) + 1, "Person " & Trim(Str(ilLoop)) & " Title"
        mAddImportTitles PERSON + 8 * (ilLoop - 1) + 2, "Person " & Trim(Str(ilLoop)) & " Phone"
        mAddImportTitles PERSON + 8 * (ilLoop - 1) + 3, "Person " & Trim(Str(ilLoop)) & " Fax"
        mAddImportTitles PERSON + 8 * (ilLoop - 1) + 4, "Person " & Trim(Str(ilLoop)) & " EMail"
        mAddImportTitles PERSON + 8 * (ilLoop - 1) + 5, "Person " & Trim(Str(ilLoop)) & " Aff-Label"
        mAddImportTitles PERSON + 8 * (ilLoop - 1) + 6, "Person " & Trim(Str(ilLoop)) & " ISCI Export"
        mAddImportTitles PERSON + 8 * (ilLoop - 1) + 7, "Person " & Trim(Str(ilLoop)) & " Aff-Email"
    Next ilLoop
End Sub

Private Sub mAddImportTitles(ilIndex As Integer, slTitle As String)
    If ilIndex >= UBound(sgStationImportTitles) Then
        ReDim Preserve sgStationImportTitles(0 To ilIndex + 1) As String
    End If
    sgStationImportTitles(ilIndex) = slTitle
End Sub

Private Function mPersonnel(llIndex As Long) As Boolean
    Dim ilPerson As Integer
    Dim llCode As Long
    Dim slName As String
    Dim slFirstName As String
    Dim slLastName As String
    Dim slTitle As String
    Dim slPhone As String
    Dim slFax As String
    Dim slEMail As String
    Dim ilWebRefID As Integer
    Dim slAffLabel As String
    Dim slISCIExport As String
    Dim slAffEmail As String
    Dim ilLen As Integer
    Dim ilPos As Integer
    Dim ilUpper As Integer
    Dim slAddChg As String
    Dim iltntCode As Integer
    Dim ilLoop As Integer
    Dim slSQLQuery As String
    Dim llArttRemoveTitle As Long
    Dim ilFound As Integer
    Dim blTagForWebUpdate As Boolean
    Dim ilRet As Integer
    
    On Error GoTo ErrHandler:
    mPersonnel = False
    
    ReDim tlPersonInfo(0 To 0) As AFFAEINFO
    slSQLQuery = "SELECT * FROM artt"
    slSQLQuery = slSQLQuery + " WHERE ("
    slSQLQuery = slSQLQuery & " arttType = 'P'"
    slSQLQuery = slSQLQuery & " AND arttShttCode = " & tmUpdateStation(llIndex).iCode & ")"
    slSQLQuery = slSQLQuery & " ORDER BY arttFirstName, arttLastName"
    Set rst_artt = gSQLSelectCall(slSQLQuery)
    Do While Not rst_artt.EOF
        ilUpper = UBound(tlPersonInfo)
        tlPersonInfo(ilUpper).sFirstName = rst_artt!arttFirstName
        tlPersonInfo(ilUpper).sLastName = rst_artt!arttLastName
        tlPersonInfo(ilUpper).sName = tlPersonInfo(ilUpper).sFirstName & " " & tlPersonInfo(ilUpper).sLastName
        tlPersonInfo(ilUpper).sEmail = rst_artt!arttEmail
        tlPersonInfo(ilUpper).iTntCode = rst_artt!arttTntCode
        tlPersonInfo(ilUpper).lCode = rst_artt!arttCode
        ReDim Preserve tlPersonInfo(0 To ilUpper + 1) As AFFAEINFO
        rst_artt.MoveNext
    Loop
    For ilPerson = 1 To imPersonTitles Step 1
        llCode = 0
        slFirstName = ""
        slLastName = ""
        slName = Trim$(tmUpdateStation(llIndex).sPersonName(ilPerson - 1))
        slPhone = Trim$(tmUpdateStation(llIndex).sPersonPhone(ilPerson - 1))
        slFax = Trim$(tmUpdateStation(llIndex).sPersonFax(ilPerson - 1))
        slEMail = Trim$(tmUpdateStation(llIndex).sPersonEMail(ilPerson - 1))
        If Trim$(tmUpdateStation(llIndex).sPersonAffLabel(ilPerson - 1)) <> "" Then
            If UCase(Left$(Trim$(tmUpdateStation(llIndex).sPersonAffLabel(ilPerson - 1)), 1)) = "Y" Then
                slAffLabel = "1"
            Else
                slAffLabel = ""
            End If
        Else
            slAffLabel = ""
        End If
        If Trim$(tmUpdateStation(llIndex).sPersonISCIExport(ilPerson - 1)) <> "" Then
            If UCase(Left$(Trim$(tmUpdateStation(llIndex).sPersonISCIExport(ilPerson - 1)), 1)) = "Y" Then
                slISCIExport = "1"
            Else
                slISCIExport = ""
            End If
        Else
            slISCIExport = ""
        End If
        If Trim$(tmUpdateStation(llIndex).sPersonAffEMail(ilPerson - 1)) <> "" Then
            If UCase(Left$(Trim$(tmUpdateStation(llIndex).sPersonAffEMail(ilPerson - 1)), 1)) = "Y" Then
                slAffEmail = "Y"
            Else
                slAffEmail = "N"
            End If
        Else
            slAffEmail = "N"
        End If
        If (slName <> "") Then
            ilLen = Len(slName)
            If ilLen > 0 Then
                ilPos = InStrRev(slName, " ")
                If ilPos > 0 Then
                    slFirstName = gFixQuote(Left(slName, ilPos - 1))
                    slLastName = gFixQuote(Trim(right(slName, ilLen - ilPos)))
                Else
                    slLastName = gFixQuote(Trim(slName))
                End If
            End If
                    
            slTitle = ""
            iltntCode = 0
            For ilLoop = LBound(tgTitleInfo) To UBound(tgTitleInfo) - 1 Step 1
                If UCase(Trim$(tmUpdateStation(llIndex).sPersonTitle(ilPerson - 1))) = UCase(Trim$(tgTitleInfo(ilLoop).sTitle)) Then
                    slTitle = Trim$(tgTitleInfo(ilLoop).sTitle)
                    iltntCode = tgTitleInfo(ilLoop).iCode
                    Exit For
                End If
            Next ilLoop
            If (slTitle = "") And (Trim$(tmUpdateStation(llIndex).sPersonTitle(ilPerson - 1)) <> "") Then
                'Add City
                iltntCode = gAddTitleName(Trim$(tmUpdateStation(llIndex).sPersonTitle(ilPerson - 1)))
                If iltntCode = -1 Then
                    iltntCode = 0
                    slTitle = ""
                    Call mUpdateReport(llIndex, "WARNING: Title name from Import Not Found. (" & Trim(tmUpdateStation(llIndex).sMailCity) & ")")
                Else
                    slTitle = Trim$(tmUpdateStation(llIndex).sPersonTitle(ilPerson - 1))
                    tgTitleInfo(UBound(tgTitleInfo)).iCode = iltntCode
                    tgTitleInfo(UBound(tgTitleInfo)).sTitle = slTitle
                    ReDim Preserve tgTitleInfo(LBound(tgTitleInfo) To UBound(tgTitleInfo) + 1) As TITLEINFO
                End If
            End If
        
            slAddChg = "A"
            llArttRemoveTitle = -1
            'Multi-personnel can have the same title so this test is not valid
            'It was used to know to update the user name
            'If iltntCode > 0 Then
            '    For ilLoop = 0 To UBound(tlPersonInfo) - 1 Step 1
            '        If tlPersonInfo(ilLoop).iTntCode = iltntCode Then
            '            slAddChg = "C"
            '            llCode = tlPersonInfo(ilLoop).lCode
            '            Exit For
            '        End If
            '    Next ilLoop
            'End If
            If slAddChg = "A" Then
                For ilLoop = 0 To UBound(tlPersonInfo) - 1 Step 1
                    If StrComp(UCase(Trim$(tlPersonInfo(ilLoop).sFirstName)), UCase(slFirstName), vbTextCompare) = 0 Then
                        If StrComp(UCase(Trim$(tlPersonInfo(ilLoop).sLastName)), UCase(slLastName), vbTextCompare) = 0 Then
                            slAddChg = "C"
                            llCode = tlPersonInfo(ilLoop).lCode
                            Exit For
                        End If
                    End If
                    If slAddChg = "A" Then
                        If StrComp(UCase(Trim$(tlPersonInfo(ilLoop).sName)), UCase(slName), vbTextCompare) = 0 Then
                            slAddChg = "C"
                            llCode = tlPersonInfo(ilLoop).lCode
                            Exit For
                        End If
                    End If
                Next ilLoop
            Else
                For ilLoop = 0 To UBound(tlPersonInfo) - 1 Step 1
                    If StrComp(UCase(Trim$(tlPersonInfo(ilLoop).sFirstName)), UCase(slFirstName), vbTextCompare) = 0 Then
                        If StrComp(UCase(Trim$(tlPersonInfo(ilLoop).sLastName)), UCase(slLastName), vbTextCompare) = 0 Then
                            If llCode <> tlPersonInfo(ilLoop).lCode Then
                                llArttRemoveTitle = llCode
                                llCode = tlPersonInfo(ilLoop).lCode
                            Else
                                Exit For
                            End If
                        End If
                        If llArttRemoveTitle = -1 Then
                            If StrComp(UCase(Trim$(tlPersonInfo(ilLoop).sName)), UCase(slName), vbTextCompare) = 0 Then
                                If llCode <> tlPersonInfo(ilLoop).lCode Then
                                    llArttRemoveTitle = llCode
                                    llCode = tlPersonInfo(ilLoop).lCode
                                Else
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next ilLoop
            End If
        Else
            iltntCode = 0
            slAddChg = ""
            If Trim$(slEMail) <> "" Then
                slPhone = ""
                slFax = ""
                slAffLabel = ""
                slISCIExport = ""
                slAffEmail = ""
                iltntCode = 0
                slAddChg = "A"
                For ilLoop = 0 To UBound(tlPersonInfo) - 1 Step 1
                    If UCase$(Trim$(tlPersonInfo(ilLoop).sEmail)) = UCase(Trim$(slEMail)) Then
                        slAddChg = ""
                    End If
                Next ilLoop
            End If
        End If
        
        If slAffLabel = "1" Then
            slSQLQuery = "Update artt Set "
            slSQLQuery = slSQLQuery & "arttAffContact = '" & "" & "' "    'arttAffContact
            slSQLQuery = slSQLQuery & " WHERE arttType = '" & "P" & "'"
            slSQLQuery = slSQLQuery & " AND arttShttCode = " & tmUpdateStation(llIndex).iCode
            If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHandler:
                gHandleError "AffErrorLog.txt", "ImportUpdateStations-mPersonnel"
            End If
        End If
        
        If slAddChg = "A" Then
            slSQLQuery = "SELECT MAX(arttWebEMailRefID) from artt WHERE arttShttCode = " & tmUpdateStation(llIndex).iCode
            Set rst_artt = gSQLSelectCall(slSQLQuery)
            If IsNull(rst_artt(0).Value) Then
                ilWebRefID = 1
            Else
                If Not rst_artt.EOF Then
                    ilWebRefID = rst_artt(0).Value + 1
                Else
                    ilWebRefID = 1
                End If
            End If
            
            slSQLQuery = "Insert Into artt ( "
            slSQLQuery = slSQLQuery & "arttCode, "
            slSQLQuery = slSQLQuery & "arttFirstName, "
            slSQLQuery = slSQLQuery & "arttLastName, "
            slSQLQuery = slSQLQuery & "arttPhone, "
            slSQLQuery = slSQLQuery & "arttFax, "
            slSQLQuery = slSQLQuery & "arttEmail, "
            slSQLQuery = slSQLQuery & "arttEmailRights, "
            slSQLQuery = slSQLQuery & "arttState, "
            slSQLQuery = slSQLQuery & "arttUsfCode, "
            slSQLQuery = slSQLQuery & "arttAddress1, "
            slSQLQuery = slSQLQuery & "arttAddress2, "
            slSQLQuery = slSQLQuery & "arttCity, "
            slSQLQuery = slSQLQuery & "arttAddressState, "
            slSQLQuery = slSQLQuery & "arttZip, "
            slSQLQuery = slSQLQuery & "arttCountry, "
            slSQLQuery = slSQLQuery & "arttType, "
            slSQLQuery = slSQLQuery & "arttTntCode, "
            slSQLQuery = slSQLQuery & "arttShttCode, "
            slSQLQuery = slSQLQuery & "arttAffContact, "
            slSQLQuery = slSQLQuery & "arttISCI2Contact, "
            slSQLQuery = slSQLQuery & "arttWebEMail, "
            slSQLQuery = slSQLQuery & "arttEMailToWeb, "
            slSQLQuery = slSQLQuery & "arttWebEMailRefID, "
            slSQLQuery = slSQLQuery & "arttUnused "
            slSQLQuery = slSQLQuery & ") "
            slSQLQuery = slSQLQuery & "Values ( "
            slSQLQuery = slSQLQuery & "Replace" & ", "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(slFirstName) & "', "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(slLastName) & "', "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(slPhone) & "', "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(slFax) & "', "
            slSQLQuery = slSQLQuery & "'" & gFixQuote(slEMail) & "', "
            slSQLQuery = slSQLQuery & "'" & "N" & "', "    'arttEMailRights
            slSQLQuery = slSQLQuery & 0 & ", "
            slSQLQuery = slSQLQuery & igUstCode & ", "
            slSQLQuery = slSQLQuery & "'" & "" & "', "    'arttAddress1
            slSQLQuery = slSQLQuery & "'" & "" & "', "    'arttAddress2
            slSQLQuery = slSQLQuery & "'" & "" & "', "    'arttCity
            slSQLQuery = slSQLQuery & "'" & "" & "', "    'arttAddressState
            slSQLQuery = slSQLQuery & "'" & "" & "', "    'arttZip
            slSQLQuery = slSQLQuery & "'" & "" & "', "    'arttCountry
            slSQLQuery = slSQLQuery & "'" & "P" & "', "   'arttTypef
            slSQLQuery = slSQLQuery & iltntCode & ", "    'arttTntCode
            slSQLQuery = slSQLQuery & tmUpdateStation(llIndex).iCode & ", "      'arttShttCode
            slSQLQuery = slSQLQuery & "'" & slAffLabel & "', "    'arttAffContact
            slSQLQuery = slSQLQuery & "'" & slISCIExport & "', "    'arttISCI2Contact
            slSQLQuery = slSQLQuery & "'" & slAffEmail & "', "   'arttWebEMail
            slSQLQuery = slSQLQuery & "'" & "I" & "', "   'arttEMailToWeb
            slSQLQuery = slSQLQuery & ilWebRefID & ", "   'arttWebEMailRefID
            slSQLQuery = slSQLQuery & "'" & "" & "' "
            slSQLQuery = slSQLQuery & ") "
            llCode = gInsertAndReturnCode(slSQLQuery, "artt", "arttCode", "Replace")
            
            ' JD TTP 10860
            ilRet = gWebInsertEmail(llCode, tmUpdateStation(llIndex).iCode, ilWebRefID, gFixQuote(slEMail), gFixQuote(slFirstName), gFixQuote(slLastName), iltntCode)
            
        ElseIf (slAddChg = "C") Then
            blTagForWebUpdate = False
            slSQLQuery = "Update artt Set "
            If (slFirstName <> "") Or ((slFirstName = "") And (bmIgnoreBlanks(imMap(PERSON + PNAME + 8 * (ilPerson - 1))) = False)) Then
                slSQLQuery = slSQLQuery & "arttFirstName = '" & gFixQuote(slFirstName) & "', "
                blTagForWebUpdate = True
            End If
            If (slLastName <> "") Or ((slLastName = "") And (bmIgnoreBlanks(imMap(PERSON + PNAME + 8 * (ilPerson - 1))) = False)) Then
                slSQLQuery = slSQLQuery & "arttLastName = '" & gFixQuote(slLastName) & "', "
                blTagForWebUpdate = True
            End If
            If (slPhone <> "") Or ((slPhone = "") And (bmIgnoreBlanks(imMap(PERSON + PPHONE + 8 * (ilPerson - 1))) = False)) Then
                slSQLQuery = slSQLQuery & "arttPhone = '" & gFixQuote(slPhone) & "', "
            End If
            If (slFax <> "") Or ((slFax = "") And (bmIgnoreBlanks(imMap(PERSON + PFAX + 8 * (ilPerson - 1))) = False)) Then
                slSQLQuery = slSQLQuery & "arttFax = '" & gFixQuote(slFax) & "', "
            End If
            If (slEMail <> "") Or ((slEMail = "") And (bmIgnoreBlanks(imMap(PERSON + PEMAIL + 8 * (ilPerson - 1))) = False)) Then
                slSQLQuery = slSQLQuery & "arttEmail = '" & gFixQuote(slEMail) & "', "
                blTagForWebUpdate = True
            End If
            If (slTitle <> "") Or ((slTitle = "") And (bmIgnoreBlanks(imMap(PERSON + PTITLE + 8 * (ilPerson - 1))) = False)) Then
                slSQLQuery = slSQLQuery & "arttTntCode = " & iltntCode & ","
            End If
            If (slAffLabel <> "") Or ((slAffLabel = "") And (bmIgnoreBlanks(imMap(PERSON + PAFFLABEL + 8 * (ilPerson - 1))) = False)) Then
                slSQLQuery = slSQLQuery & "arttAffContact = '" & gFixQuote(slAffLabel) & "', "
            End If
            If (slISCIExport <> "") Or ((slISCIExport = "") And (bmIgnoreBlanks(imMap(PERSON + PISCIEXPORT + 8 * (ilPerson - 1))) = False)) Then
                slSQLQuery = slSQLQuery & "arttISCI2Contact = '" & gFixQuote(slISCIExport) & "', "
            End If
            If (slAffEmail <> "") Or ((slAffEmail = "") And (bmIgnoreBlanks(imMap(PERSON + PAFFEMAIL + 8 * (ilPerson - 1))) = False)) Then
                slSQLQuery = slSQLQuery & "arttWebEMail = '" & gFixQuote(slAffEmail) & "', "
            End If
            If blTagForWebUpdate Then
                slSQLQuery = slSQLQuery & " arttEMailToWeb = 'W', "
            End If
            slSQLQuery = Trim$(slSQLQuery)
            If right(slSQLQuery, 1) = "," Then
                If llCode > 0 Then
                    slSQLQuery = Left(slSQLQuery, Len(slSQLQuery) - 1)
                    slSQLQuery = slSQLQuery & " WHERE arttCode = " & llCode
                    
                    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHandler:
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mPersonnel"
                    End If
                End If
            End If
            If llArttRemoveTitle > 0 Then
                slSQLQuery = "Update artt Set "
                slSQLQuery = slSQLQuery & "arttTntCode = " & 0
                slSQLQuery = slSQLQuery & " WHERE arttCode = " & llArttRemoveTitle
                If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                    '6/11/16: Replaced GoSub
                    'GoSub ErrHandler:
                    gHandleError "AffErrorLog.txt", "ImportUpdateStations-mPersonnel"
                End If
            End If
        End If
    Next ilPerson
    ilRet = mUpdateWebSite()    ' JD TTP 10860
    mPersonnel = True
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mPersonnel"
    Exit Function
End Function

'Private Function gAddTitleName(slTitle As String) As Long
'    Dim ilCode As Integer
'    Dim slSQLQuery As String
'
'    On Error GoTo ErrHandler:
'    gAddTitleName = -1
'    slSQLQuery = "Insert Into tnt ( "
'    slSQLQuery = slSQLQuery & "tntCode, "
'    slSQLQuery = slSQLQuery & "tntTitle, "
'    slSQLQuery = slSQLQuery & "tntUsfCode, "
'    slSQLQuery = slSQLQuery & "tntUnused "
'    slSQLQuery = slSQLQuery & ") "
'    slSQLQuery = slSQLQuery & "Values ( "
'    slSQLQuery = slSQLQuery & "Replace" & ", "
'    slSQLQuery = slSQLQuery & "'" & gFixQuote(slTitle) & "', "
'    slSQLQuery = slSQLQuery & igUstCode & ", "
'    slSQLQuery = slSQLQuery & "'" & "" & "' "
'    slSQLQuery = slSQLQuery & ") "
'    ilCode = CInt(gInsertAndReturnCode(slSQLQuery, "tnt", "tntCode", "Replace"))
'    If ilCode > 0 Then
'        gAddTitleName = ilCode
'    End If
'
'    Exit Function

Private Sub mSetFields(llIndex As Long)
    Dim ilLoop As Integer
    Dim llMktIdx As Long
    Dim llOwnerIdx As Long
    Dim llFormatIdx As Long
    
    'Territory
    smTerritory = ""
    lmTerritoryMntCode = 0
    For ilLoop = LBound(tgTerritoryInfo) To UBound(tgTerritoryInfo) - 1 Step 1
        If UCase(Trim$(tmUpdateStation(llIndex).sTerritory)) = UCase(Trim$(tgTerritoryInfo(ilLoop).sName)) Then
            smTerritory = tgTerritoryInfo(ilLoop).sName
            lmTerritoryMntCode = tgTerritoryInfo(ilLoop).lCode
            Exit For
        End If
    Next ilLoop
    If (smTerritory = "") And (Trim$(tmUpdateStation(llIndex).sTerritory) <> "") Then
        'Add Territory
        lmTerritoryMntCode = mAddMultiName("T", Trim$(tmUpdateStation(llIndex).sTerritory))
        If lmTerritoryMntCode = -1 Then
            lmTerritoryMntCode = 0
            smTerritory = ""
            Call mUpdateReport(llIndex, "WARNING: Territory name from Import Not Found. (" & Trim(tmUpdateStation(llIndex).sTerritory) & ")")
        Else
            smTerritory = Trim$(tmUpdateStation(llIndex).sTerritory)
            tgTerritoryInfo(UBound(tgTerritoryInfo)).lCode = lmTerritoryMntCode
            tgTerritoryInfo(UBound(tgTerritoryInfo)).sName = smTerritory
            tgTerritoryInfo(UBound(tgTerritoryInfo)).sState = "A"
            ReDim Preserve tgTerritoryInfo(LBound(tgTerritoryInfo) To UBound(tgTerritoryInfo) + 1) As MNTINFO
        End If
    End If
    'Area
    smArea = ""
    lmAreaMntCode = 0
    For ilLoop = LBound(tmAreaInfo) To UBound(tmAreaInfo) - 1 Step 1
        If UCase(Trim$(tmUpdateStation(llIndex).sArea)) = UCase(Trim$(tmAreaInfo(ilLoop).sName)) Then
            smArea = tmAreaInfo(ilLoop).sName
            lmAreaMntCode = tmAreaInfo(ilLoop).lCode
            Exit For
        End If
    Next ilLoop
    If (smArea = "") And (Trim$(tmUpdateStation(llIndex).sArea) <> "") Then
        'Add Area
        lmAreaMntCode = mAddMultiName("A", Trim$(tmUpdateStation(llIndex).sArea))
        If lmAreaMntCode = -1 Then
            lmAreaMntCode = 0
            smArea = ""
            Call mUpdateReport(llIndex, "WARNING: Area name from Import Not Found. (" & Trim(tmUpdateStation(llIndex).sArea) & ")")
        Else
            smArea = Trim$(tmUpdateStation(llIndex).sArea)
            tmAreaInfo(UBound(tmAreaInfo)).lCode = lmAreaMntCode
            tmAreaInfo(UBound(tmAreaInfo)).sName = smArea
            tmAreaInfo(UBound(tmAreaInfo)).sState = "A"
            ReDim Preserve tmAreaInfo(LBound(tmAreaInfo) To UBound(tmAreaInfo) + 1) As MNTINFO
        End If
    End If
    'Format
    smFormat = ""
    imFormatCode = 0
    lmFormatIdx = 0
    If Len(Trim(tmUpdateStation(llIndex).sFormat)) > 0 Then
        lmFormatIdx = mLookupFormatLinkByName(tmUpdateStation(llIndex).sFormat)
        If lmFormatIdx <> -1 Then
            smFormat = Trim$(tmUpdateStation(llIndex).sFormat)
            imFormatCode = tmFormatLinkInfo(lmFormatIdx).iIntFmtCode
        Else
            Call mUpdateReport(llIndex, "WARNING: Format name from Import missing from Affiliate. (" & Trim(tmUpdateStation(llIndex).sFormat) & ")")
        End If
    End If
    'DMA Market
    smDMAMarket = ""
    lmDMAMktCode = 0
    lmDMAMktIdx = 0
    If Len(Trim(tmUpdateStation(llIndex).sDMAMarket)) > 0 Then  ' Don't look this up if blank.
        lmDMAMktIdx = mLookupDMAMarketByName(tmUpdateStation(llIndex).sDMAMarket)
        If lmDMAMktIdx <> -1 Then
            smDMAMarket = Trim$(tmUpdateStation(llIndex).sDMAMarket)
            lmDMAMktCode = tgMarketInfo(lmDMAMktIdx).lCode
        End If
    End If
    'City License
    smCityLic = ""
    lmCityLicMntCode = 0
    For ilLoop = LBound(tmCityInfo) To UBound(tmCityInfo) - 1 Step 1
        If UCase(Trim$(tmUpdateStation(llIndex).sCityLicense)) = UCase(Trim$(tmCityInfo(ilLoop).sName)) Then
            smCityLic = tmCityInfo(ilLoop).sName
            lmCityLicMntCode = tmCityInfo(ilLoop).lCode
            Exit For
        End If
    Next ilLoop
    If (smCityLic = "") And (Trim$(tmUpdateStation(llIndex).sCityLicense) <> "") Then
        'Add City
        lmCityLicMntCode = mAddMultiName("C", Trim$(tmUpdateStation(llIndex).sCityLicense))
        If lmCityLicMntCode = -1 Then
            lmCityLicMntCode = 0
            smCityLic = ""
            Call mUpdateReport(llIndex, "WARNING: City License name from Import Not Found. (" & Trim(tmUpdateStation(llIndex).sCityLicense) & ")")
        Else
            smCityLic = Trim$(tmUpdateStation(llIndex).sCityLicense)
            tmCityInfo(UBound(tmCityInfo)).lCode = lmCityLicMntCode
            tmCityInfo(UBound(tmCityInfo)).sName = Trim$(tmUpdateStation(llIndex).sCityLicense)
            tmCityInfo(UBound(tmCityInfo)).sState = "A"
            ReDim Preserve tmCityInfo(LBound(tmCityInfo) To UBound(tmCityInfo) + 1) As MNTINFO
        End If
    End If
    'County License
    smCountyLic = ""
    lmCountyLicMntCode = 0
    For ilLoop = LBound(tmCountyInfo) To UBound(tmCountyInfo) - 1 Step 1
        If UCase(Trim$(tmUpdateStation(llIndex).sCountyLicense)) = UCase(Trim$(tmCountyInfo(ilLoop).sName)) Then
            smCountyLic = tmCountyInfo(ilLoop).sName
            lmCountyLicMntCode = tmCountyInfo(ilLoop).lCode
            Exit For
        End If
    Next ilLoop
    If (smCountyLic = "") And (Trim$(tmUpdateStation(llIndex).sCountyLicense) <> "") Then
        'Add City
        lmCountyLicMntCode = mAddMultiName("Y", Trim$(tmUpdateStation(llIndex).sCountyLicense))
        If lmCountyLicMntCode = -1 Then
            lmCountyLicMntCode = 0
            smCountyLic = ""
            Call mUpdateReport(llIndex, "WARNING: County License name from Import Not Found. (" & Trim(tmUpdateStation(llIndex).sCountyLicense) & ")")
        Else
            smCountyLic = Trim$(tmUpdateStation(llIndex).sCountyLicense)
            tmCountyInfo(UBound(tmCountyInfo)).lCode = lmCountyLicMntCode
            tmCountyInfo(UBound(tmCountyInfo)).sName = smCountyLic
            tmCountyInfo(UBound(tmCountyInfo)).sState = "A"
            ReDim Preserve tmCountyInfo(LBound(tmCountyInfo) To UBound(tmCountyInfo) + 1) As MNTINFO
        End If
    End If
    'State License
    smStateLic = ""
    If Trim$(tmUpdateStation(llIndex).sStateLicense) <> "" Then
        For ilLoop = LBound(tgStateInfo) To UBound(tgStateInfo) - 1 Step 1
            If StrComp(Trim$(tgStateInfo(ilLoop).sPostalName), Trim$(tmUpdateStation(llIndex).sStateLicense), vbTextCompare) = 0 Then
                smStateLic = Trim$(tmUpdateStation(llIndex).sStateLicense)
                Exit For
            End If
        Next ilLoop
    End If
    'Owner
    lmOwnerCode = 0
    lmOwnerIdx = 0
    smOwner = ""
    If Len(Trim(tmUpdateStation(llIndex).sOwner)) > 0 Then
        lmOwnerIdx = mLookupOwnerByName(tmUpdateStation(llIndex).sOwner)
        If lmOwnerIdx <> -1 Then
            smOwner = Trim$(tmUpdateStation(llIndex).sOwner)
            lmOwnerCode = tgOwnerInfo(lmOwnerIdx).lCode
        Else
            Call mUpdateReport(llIndex, "WARNING: Owner name from Import missing from Affiliate. (" & Trim(tmUpdateStation(llIndex).sOwner) & ")")
        End If
    End If
    'Operator
    smOperator = ""
    lmOperatorMntCode = 0
    For ilLoop = LBound(tmOperatorInfo) To UBound(tmOperatorInfo) - 1 Step 1
        If UCase(Trim$(tmUpdateStation(llIndex).sOperator)) = UCase(Trim$(tmOperatorInfo(ilLoop).sName)) Then
            smOperator = tmOperatorInfo(ilLoop).sName
            lmOperatorMntCode = tmOperatorInfo(ilLoop).lCode
            Exit For
        End If
    Next ilLoop
    If (smOperator = "") And (Trim$(tmUpdateStation(llIndex).sOperator) <> "") Then
        'Add Operator
        lmOperatorMntCode = mAddMultiName("O", Trim$(tmUpdateStation(llIndex).sOperator))
        If lmOperatorMntCode = -1 Then
            lmOperatorMntCode = 0
            smOperator = ""
            Call mUpdateReport(llIndex, "WARNING: Operator name from Import Not Found. (" & Trim(tmUpdateStation(llIndex).sOperator) & ")")
        Else
            smOperator = Trim$(tmUpdateStation(llIndex).sOperator)
            tmOperatorInfo(UBound(tmOperatorInfo)).lCode = lmOperatorMntCode
            tmOperatorInfo(UBound(tmOperatorInfo)).sName = smOperator
            tmOperatorInfo(UBound(tmOperatorInfo)).sState = "A"
            ReDim Preserve tmOperatorInfo(LBound(tmOperatorInfo) To UBound(tmOperatorInfo) + 1) As MNTINFO
        End If
    End If
    'MSA Market
    lmMSAMktCode = 0
    lmMSAMktIdx = 0
    smMSAMarket = ""
    If Len(Trim(tmUpdateStation(llIndex).sMSAMarket)) > 0 Then  ' Don't look this up if blank.
        lmMSAMktIdx = mLookupMSAMarketByName(tmUpdateStation(llIndex).sMSAMarket)
        If lmMSAMktIdx <> -1 Then
            smMSAMarket = Trim$(tmUpdateStation(llIndex).sMSAMarket)
            lmMSAMktCode = tgMSAMarketInfo(lmMSAMktIdx).lCode
        Else
            Call mUpdateReport(llIndex, "WARNING: MSA Market name from Import missing from Affiliate. (" & Trim(tmUpdateStation(llIndex).sMSAMarket) & ")")
        End If
    End If
    'Market Rep
    smMarketRep = ""
    imMarketRepUstCode = 0
    If (Trim$(tmUpdateStation(llIndex).sMarketRep) <> "") Or (bmIgnoreBlanks(imMap(MARKETREP)) = False) Then
        For ilLoop = LBound(tgMarketRepInfo) To UBound(tgMarketRepInfo) - 1 Step 1
            If tmUpdateStation(llIndex).sMarketRep = tgMarketRepInfo(ilLoop).sReportName Then
                smMarketRep = tgMarketRepInfo(ilLoop).sReportName
                imMarketRepUstCode = tgMarketRepInfo(ilLoop).iUstCode
                Exit For
            End If
        Next ilLoop
        If smMarketRep = "" Then
            For ilLoop = LBound(tgMarketRepInfo) To UBound(tgMarketRepInfo) - 1 Step 1
                If tmUpdateStation(llIndex).sMarketRep = tgMarketRepInfo(ilLoop).sLogInName Then
                    smMarketRep = tgMarketRepInfo(ilLoop).sLogInName
                    imMarketRepUstCode = tgMarketRepInfo(ilLoop).iUstCode
                    Exit For
                End If
            Next ilLoop
        End If
        If imMarketRepUstCode = 0 Then
            Call mUpdateReport(llIndex, "WARNING: Market Rep missing or not found (" & Trim(tmUpdateStation(llIndex).sCallLetters) & ")")
        End If
    End If
    'Service Rep
    smServiceRep = ""
    imServiceRepUstCode = 0
    If (Trim$(tmUpdateStation(llIndex).sServiceRep) <> "") Or (bmIgnoreBlanks(imMap(SERVICEREP)) = False) Then
        For ilLoop = LBound(tgServiceRepInfo) To UBound(tgServiceRepInfo) - 1 Step 1
            If tmUpdateStation(llIndex).sServiceRep = tgServiceRepInfo(ilLoop).sReportName Then
                smServiceRep = tgServiceRepInfo(ilLoop).sReportName
                imServiceRepUstCode = tgServiceRepInfo(ilLoop).iUstCode
                Exit For
            End If
        Next ilLoop
        If smServiceRep = "" Then
            For ilLoop = LBound(tgServiceRepInfo) To UBound(tgServiceRepInfo) - 1 Step 1
                If tmUpdateStation(llIndex).sServiceRep = tgServiceRepInfo(ilLoop).sLogInName Then
                    smServiceRep = tgServiceRepInfo(ilLoop).sLogInName
                    imServiceRepUstCode = tgServiceRepInfo(ilLoop).iUstCode
                    Exit For
                End If
            Next ilLoop
        End If
        If imServiceRepUstCode = 0 Then
            Call mUpdateReport(llIndex, "WARNING: Service Rep missing or not found (" & Trim(tmUpdateStation(llIndex).sCallLetters) & ")")
        End If
    End If
    'Time zone
    smTimeZone = ""
    imTztCode = 0
    If (tmUpdateStation(llIndex).sZone = "A") And (ckcAlaskaToPacific.Value = vbChecked) Then
        tmUpdateStation(llIndex).sZone = "P"
    End If
    If (tmUpdateStation(llIndex).sZone = "H") And (ckcHawaiiToPacific.Value = vbChecked) Then
        tmUpdateStation(llIndex).sZone = "P"
    End If
    For ilLoop = LBound(tgTimeZoneInfo) To UBound(tgTimeZoneInfo) - 1 Step 1
        If tmUpdateStation(llIndex).sZone = "E" Then
            ilLoop = ilLoop
        End If
        If tmUpdateStation(llIndex).sZone = Left$(tgTimeZoneInfo(ilLoop).sCSIName, 1) Then
            smTimeZone = tgTimeZoneInfo(ilLoop).sCSIName
            imTztCode = tgTimeZoneInfo(ilLoop).iCode
            Exit For
        End If
    Next ilLoop
    'On Air
    smOnAir = "Y"
    If Trim$(tmUpdateStation(llIndex).sOnAir) <> "" Then
        If UCase$(Left$(tmUpdateStation(llIndex).sOnAir, 1)) = "N" Then
            smOnAir = "N"
        End If
    End If
    'Commercial
    smCommercial = "C"
    If Trim$(tmUpdateStation(llIndex).sCommercial) <> "" Then
        If UCase$(Left$(tmUpdateStation(llIndex).sCommercial, 1)) = "N" Then
            smCommercial = "N"
        End If
    End If
    'Honor Daylight
    imDaylight = 0
    If Trim$(tmUpdateStation(llIndex).sDaylight) <> "" Then
        If UCase$(Left$(tmUpdateStation(llIndex).sDaylight, 1)) = "N" Then
            imDaylight = 1
        End If
    End If
    'Used for Agreement
    smUsedAgreement = "Y"
    If Trim$(tmUpdateStation(llIndex).sUsedAgreement) <> "" Then
        If UCase$(Left$(tmUpdateStation(llIndex).sUsedAgreement, 1)) = "N" Then
            smUsedAgreement = "N"
        End If
    End If
    'Used for XDS
    smUsedXDS = "N"
    If Trim$(tmUpdateStation(llIndex).sUsedXDS) <> "" Then
        If UCase$(Left$(tmUpdateStation(llIndex).sUsedXDS, 1)) = "Y" Then
            smUsedXDS = "Y"
        End If
    End If
    'Used for Wegener
    smUsedWegener = "N"
    If Trim$(tmUpdateStation(llIndex).sUsedWegener) <> "" Then
        If UCase$(Left$(tmUpdateStation(llIndex).sUsedWegener, 1)) = "Y" Then
            smUsedWegener = "Y"
        End If
    End If
    'Used for OLA
    smUsedOLA = "N"
    If Trim$(tmUpdateStation(llIndex).sUsedOLA) <> "" Then
        If UCase$(Left$(tmUpdateStation(llIndex).sUsedOLA, 1)) = "Y" Then
            smUsedOLA = "Y"
        End If
    End If
    'Moniker
    smMoniker = ""
    lmMonikerMntCode = 0
    For ilLoop = LBound(tmMonikerInfo) To UBound(tmMonikerInfo) - 1 Step 1
        If UCase(Trim$(tmUpdateStation(llIndex).sMoniker)) = UCase(Trim$(tmMonikerInfo(ilLoop).sName)) Then
            smMoniker = tmMonikerInfo(ilLoop).sName
            lmMonikerMntCode = tmMonikerInfo(ilLoop).lCode
            Exit For
        End If
    Next ilLoop
    If (smMoniker = "") And (Trim$(tmUpdateStation(llIndex).sMoniker) <> "") Then
        'Add Moniker
        lmMonikerMntCode = mAddMultiName("M", Trim$(tmUpdateStation(llIndex).sMoniker))
        If lmMonikerMntCode = -1 Then
            lmMonikerMntCode = 0
            smMoniker = ""
            Call mUpdateReport(llIndex, "WARNING: Moniker name from Import Not Found. (" & Trim(tmUpdateStation(llIndex).sMoniker) & ")")
        Else
            smMoniker = Trim$(tmUpdateStation(llIndex).sMoniker)
            tmMonikerInfo(UBound(tmMonikerInfo)).lCode = lmMonikerMntCode
            tmMonikerInfo(UBound(tmMonikerInfo)).sName = smMoniker
            tmMonikerInfo(UBound(tmMonikerInfo)).sState = "A"
            ReDim Preserve tmMonikerInfo(LBound(tmMonikerInfo) To UBound(tmMonikerInfo) + 1) As MNTINFO
        End If
    End If
    'Historical Date
    smHistoricalDate = "1/1/1970"
    If Trim$(tmUpdateStation(llIndex).sHistoricalDate) <> "" Then
        If gIsDate(Trim$(tmUpdateStation(llIndex).sHistoricalDate)) Then
            smHistoricalDate = Trim$(tmUpdateStation(llIndex).sHistoricalDate)
        End If
    End If
    'Mailing city
    smMailCity = ""
    lmCityMntCode = 0
    For ilLoop = LBound(tmCityInfo) To UBound(tmCityInfo) - 1 Step 1
        If UCase(Trim$(tmUpdateStation(llIndex).sMailCity)) = UCase(Trim$(tmCityInfo(ilLoop).sName)) Then
            smMailCity = tmCityInfo(ilLoop).sName
            lmCityMntCode = tmCityInfo(ilLoop).lCode
            Exit For
        End If
    Next ilLoop
    If (smMailCity = "") And (Trim$(tmUpdateStation(llIndex).sMailCity) <> "") Then
        'Add City
        lmCityMntCode = mAddMultiName("C", Trim$(tmUpdateStation(llIndex).sMailCity))
        If lmCityMntCode = -1 Then
            lmCityMntCode = 0
            smMailCity = ""
            Call mUpdateReport(llIndex, "WARNING: City name from Import Not Found. (" & Trim(tmUpdateStation(llIndex).sMailCity) & ")")
        Else
            smMailCity = Trim$(tmUpdateStation(llIndex).sMailCity)
            tmCityInfo(UBound(tmCityInfo)).lCode = lmCityMntCode
            tmCityInfo(UBound(tmCityInfo)).sName = smMailCity
            tmCityInfo(UBound(tmCityInfo)).sState = "A"
            ReDim Preserve tmCityInfo(LBound(tmCityInfo) To UBound(tmCityInfo) + 1) As MNTINFO
        End If
    End If
    'Mail State
    smMailState = ""
    If Trim$(tmUpdateStation(llIndex).sMailState) <> "" Then
        For ilLoop = LBound(tgStateInfo) To UBound(tgStateInfo) - 1 Step 1
            If StrComp(Trim$(tgStateInfo(ilLoop).sPostalName), Trim$(tmUpdateStation(llIndex).sMailState), vbTextCompare) = 0 Then
                smMailState = Trim$(tmUpdateStation(llIndex).sMailState)
                Exit For
            End If
        Next ilLoop
    End If
    'Physical City
    smPhysicalCity = ""
    lmPhysicalCityMntCode = 0
    For ilLoop = LBound(tmCityInfo) To UBound(tmCityInfo) - 1 Step 1
        If UCase(Trim$(tmUpdateStation(llIndex).sPhysicalCity)) = UCase(Trim$(tmCityInfo(ilLoop).sName)) Then
            smPhysicalCity = tmCityInfo(ilLoop).sName
            lmPhysicalCityMntCode = tmCityInfo(ilLoop).lCode
            Exit For
        End If
    Next ilLoop
    If (smPhysicalCity = "") And (Trim$(tmUpdateStation(llIndex).sPhysicalCity) <> "") Then
        'Add City
        lmPhysicalCityMntCode = mAddMultiName("C", Trim$(tmUpdateStation(llIndex).sPhysicalCity))
        If lmPhysicalCityMntCode = -1 Then
            lmPhysicalCityMntCode = 0
            smPhysicalCity = ""
            Call mUpdateReport(llIndex, "WARNING: Physical City name from Import Not Found. (" & Trim(tmUpdateStation(llIndex).sMailCity) & ")")
        Else
            smPhysicalCity = Trim$(tmUpdateStation(llIndex).sPhysicalCity)
            tmCityInfo(UBound(tmCityInfo)).lCode = lmPhysicalCityMntCode
            tmCityInfo(UBound(tmCityInfo)).sName = smPhysicalCity
            tmCityInfo(UBound(tmCityInfo)).sState = "A"
            ReDim Preserve tmCityInfo(LBound(tmCityInfo) To UBound(tmCityInfo) + 1) As MNTINFO
        End If
    End If
    'Physical State
    smPhysicalState = ""
    If Trim$(tmUpdateStation(llIndex).sPhysicalState) <> "" Then
        For ilLoop = LBound(tgStateInfo) To UBound(tgStateInfo) - 1 Step 1
            If StrComp(Trim$(tgStateInfo(ilLoop).sPostalName), Trim$(tmUpdateStation(llIndex).sPhysicalState), vbTextCompare) = 0 Then
                smPhysicalState = Trim$(tmUpdateStation(llIndex).sPhysicalState)
                Exit For
            End If
        Next ilLoop
    End If
End Sub

Private Function mProcessLine(ilPass As Integer, llMaxRecords As Long, slLine As String, slFields() As String) As Integer
    Dim slCallLetters As String
    Dim llStationID As Long
    Dim llIndex As Integer
    Dim ilUpdateType As Integer
    Dim llLoop As Long
    Dim ilLoop As Integer
    Dim ilPerson As Integer
    Dim ilInvalidNo As Integer
    
    mProcessLine = True
    slCallLetters = Trim$(slFields(imMap(CALLLETTERS)))
    llStationID = Val(slFields(imMap(ID)))
    llIndex = mLookupStation(llStationID, slCallLetters)
    If llIndex = -1 Then
        '9/25/15: Check if station entered twice or more times within the import file
        For llLoop = 0 To lmTotalRecords - 1 Step 1
            If UCase(Trim$(tmUpdateStation(llLoop).sCallLetters)) = UCase(slCallLetters) Then
                Exit Function
            End If
        Next llLoop
    End If
    'If llIndex <> -1 Then
    '    If Trim$(UCase(tgStationInfo(llIndex).sCallLetters)) <> Trim$(UCase$(slFields(imMap(CALLLETTERS)))) Then
    '        'Test if call letters currently defined for a different station
    '        If mCheckStationName(slFields(imMap(CALLLETTERS))) Then
    '            Call mUpdateReport(-1, "WARNING: Import Station " & Trim(slFields(CALLLETTERS)) & " ID " & Val(slFields(imMap(ID))) & " Not Updated because Call Letters already exist with another station")
    '            llIndex = -1
    '        End If
    '    End If
    '    If llIndex <> -1 Then
    If llIndex <> -1 Then
        ilUpdateType = 1
        If bmMatchOnPermStationID Then
            For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                If StrComp(UCase$(Trim(tgStationInfo(llLoop).sCallLetters)), UCase$(Trim(slCallLetters)), vbTextCompare) = 0 Then
                    If llIndex <> llLoop Then
                        'ilUpdateType = 2
                        If ilPass = 1 Or ckcChangeToInvalidIfLetterReassigned.Value = vbChecked Then
                            'TTP 10984 - Station Information Import: when using Station IDs, add method that will allow call letters to be swapped
                            If ckcChangeToInvalidIfLetterReassigned.Value = vbChecked Then
                                ilInvalidNo = mGetAvailableInvalidNo()
                                Debug.Print " -> Rename:" & llLoop & ",ID:" & tgStationInfo(llLoop).lPermStationID & " - " & Trim(tgStationInfo(llLoop).sCallLetters) & " to " & "Invalid" & Trim(Str(ilInvalidNo)) & "; " & slLine
                                mRenameCallLetters llLoop, "Invalid" & Trim(Str(ilInvalidNo))
                                ilUpdateType = 1
                                'tgStationInfo(llLoop).sCallLetters = "Invalid" & Trim(Str(ilInvalidNo))
                            Else
                                Debug.Print " -> Preventing:" & llLoop & ",ID:" & tgStationInfo(llLoop).lPermStationID & " - " & Trim(tgStationInfo(llLoop).sCallLetters) & " to " & "Invalid" & Trim(Str(ilInvalidNo)) & "; " & slLine
                                ilUpdateType = 2
                                Call mUpdateReport(-1, "WARNING: Call Letters currently in use, can't change " & Trim(tgStationInfo(llLoop).sCallLetters) & " to " & Trim(slCallLetters) & " ID " & llStationID)
                            End If
                        Else
                            'Fix TTP 10984
                            Debug.Print " -> Bypassing:" & llLoop & "; " & slLine
                            ilUpdateType = 2
                            smPass0LinesBypassed(UBound(smPass0LinesBypassed)) = slLine
                            ReDim Preserve smPass0LinesBypassed(0 To UBound(smPass0LinesBypassed) + 1) As String
                        End If
                        Exit For
                    End If
                End If
            Next llLoop
        Else
            'Disallow update if in add mode
            If sgUsingStationID = "A" Then
                ilUpdateType = 2
                Call mUpdateReport(-1, "WARNING: Import Station previously defined within System: " & Trim(slCallLetters))
            End If
        End If
    Else
        If bmMatchOnPermStationID Then
            ilUpdateType = 0
            For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
                If StrComp(UCase$(Trim(tgStationInfo(llLoop).sCallLetters)), UCase$(Trim(slCallLetters)), vbTextCompare) = 0 Then
                    If tgStationInfo(llLoop).lPermStationID = 0 Then
                        llIndex = llLoop
                        ilUpdateType = 1
                    Else
                        ilUpdateType = 2
                        If ilPass = 1 Then
                            Call mUpdateReport(-1, "WARNING: Call Letters currently in use, can't add " & Trim(slCallLetters) & " ID " & llStationID)
                        Else
                            smPass0LinesBypassed(UBound(smPass0LinesBypassed)) = slLine
                            ReDim Preserve smPass0LinesBypassed(0 To UBound(smPass0LinesBypassed) + 1) As String
                        End If
                        Exit For
                    End If
                End If
            Next llLoop
        Else
            'Add only if no stations existed or in Add mode
            If (Not bmStationPreviouslyDefined) Or (sgUsingStationID = "A") Then
                ilUpdateType = 0
            Else
                ilUpdateType = 2
                Call mUpdateReport(-1, "WARNING: Import Station not in System: " & Trim(slCallLetters))
            End If
        End If
    End If
    If ilUpdateType <> 2 Then
        For ilLoop = LBound(slFields) + 1 To UBound(slFields) Step 1
            If Trim$(slFields(ilLoop)) <> "" Then
                '3/31/1Disallow time zone and DMA to be set to blank
                'bmIgnoreBlanks(ilLoop) = False
                If (ilLoop <> imMap(ZONE)) And (ilLoop <> imMap(DMANAME)) Then
                    bmIgnoreBlanks(ilLoop) = False
                End If
            End If
            If StrComp(Trim$(slFields(ilLoop)), "N/A", vbTextCompare) = 0 Then
                slFields(ilLoop) = ""
            End If
        Next ilLoop
        If llIndex <> -1 Then
            tmUpdateStation(lmTotalRecords).iCode = tgStationInfo(llIndex).iCode
        Else
            tmUpdateStation(lmTotalRecords).iCode = 0
        End If
                                        
        tmUpdateStation(lmTotalRecords).sCallLetters = Trim$(slFields(imMap(CALLLETTERS)))
        If bmMatchOnPermStationID Then
            Debug.Print " -> Update:" & lmTotalRecords & ",ID:" & Val(slFields(imMap(ID))) & " to " & Trim$(slFields(imMap(CALLLETTERS)))
            tmUpdateStation(lmTotalRecords).lID = Val(slFields(imMap(ID)))
        Else
            Debug.Print " -> Update:" & lmTotalRecords & ",ID:0 to " & Trim$(slFields(imMap(CALLLETTERS)))
            tmUpdateStation(lmTotalRecords).lID = 0
        End If
        tmUpdateStation(lmTotalRecords).sFrequency = Trim$(slFields(imMap(FREQUENCY)))
        tmUpdateStation(lmTotalRecords).sTerritory = Trim$(slFields(imMap(TERRITORY)))
        tmUpdateStation(lmTotalRecords).sArea = Trim$(slFields(imMap(AREA)))
        tmUpdateStation(lmTotalRecords).sFormat = Trim(slFields(imMap(STATIONFORMAT)))
        tmUpdateStation(lmTotalRecords).iDMARank = Val(slFields(imMap(DMARANK)))
        tmUpdateStation(lmTotalRecords).sDMARank = Trim$(slFields(imMap(DMARANK)))
        If Left$(Trim(slFields(imMap(DMANAME))), 1) = "<" Then
            slFields(imMap(DMANAME)) = Mid$(Trim$(slFields(imMap(DMANAME))), 2)
        End If
        If right$(Trim(slFields(imMap(DMANAME))), 1) = ">" Then
            slFields(imMap(DMANAME)) = Left$(Trim$(slFields(imMap(DMANAME))), Len(Trim$(slFields(imMap(DMANAME)))) - 1)
        End If
        tmUpdateStation(lmTotalRecords).sDMAMarket = Trim(slFields(imMap(DMANAME)))
        If Left$(Trim(slFields(imMap(CITYLIC))), 1) = "<" Then
            slFields(imMap(CITYLIC)) = Mid$(Trim$(slFields(imMap(CITYLIC))), 2)
        End If
        If right$(Trim(slFields(imMap(CITYLIC))), 1) = ">" Then
            slFields(imMap(CITYLIC)) = Left$(Trim$(slFields(imMap(CITYLIC))), Len(Trim$(slFields(imMap(CITYLIC)))) - 1)
        End If
        tmUpdateStation(lmTotalRecords).sCityLicense = Trim$(slFields(imMap(CITYLIC)))
        If Left$(Trim(slFields(imMap(COUNTYLIC))), 1) = "<" Then
            slFields(imMap(COUNTYLIC)) = Mid$(Trim$(slFields(imMap(COUNTYLIC))), 2)
        End If
        If right$(Trim(slFields(imMap(COUNTYLIC))), 1) = ">" Then
            slFields(imMap(COUNTYLIC)) = Left$(Trim$(slFields(imMap(COUNTYLIC))), Len(Trim$(slFields(imMap(COUNTYLIC)))) - 1)
        End If
        tmUpdateStation(lmTotalRecords).sCountyLicense = Trim$(slFields(imMap(COUNTYLIC)))
        If Left$(Trim(slFields(imMap(STATELIC))), 1) = "<" Then
            slFields(imMap(STATELIC)) = Mid$(Trim$(slFields(imMap(STATELIC))), 2)
        End If
        If right$(Trim(slFields(imMap(STATELIC))), 1) = ">" Then
            slFields(imMap(STATELIC)) = Left$(Trim$(slFields(imMap(STATELIC))), Len(Trim$(slFields(imMap(STATELIC)))) - 1)
        End If
        tmUpdateStation(lmTotalRecords).sStateLicense = Trim$(slFields(imMap(STATELIC)))
        If Left$(Trim$(slFields(imMap(OWNER))), 1) = "&" Then
            slFields(imMap(OWNER)) = Mid$(Trim(slFields(imMap(OWNER))), 2)
        End If
        tmUpdateStation(lmTotalRecords).sOwner = Trim(slFields(imMap(OWNER)))
        If Left$(Trim$(slFields(imMap(OPERATOR))), 1) = "&" Then
            slFields(imMap(OPERATOR)) = Mid$(Trim(slFields(imMap(OPERATOR))), 2)
        End If
        tmUpdateStation(lmTotalRecords).sOperator = Trim(slFields(imMap(OPERATOR)))
        tmUpdateStation(lmTotalRecords).iMSARank = Val(slFields(imMap(MSARANK)))
        tmUpdateStation(lmTotalRecords).sMSARank = Trim$(slFields(imMap(MSARANK)))
        If Left$(Trim(slFields(imMap(MSANAME))), 1) = "<" Then
            slFields(imMap(MSANAME)) = Mid$(Trim$(slFields(imMap(MSANAME))), 2)
        End If
        If right$(Trim(slFields(imMap(MSANAME))), 1) = ">" Then
            slFields(imMap(MSANAME)) = Left$(Trim$(slFields(imMap(MSANAME))), Len(Trim$(slFields(imMap(MSANAME)))) - 1)
        End If
        tmUpdateStation(lmTotalRecords).sMSAMarket = Trim(slFields(imMap(MSANAME)))
        tmUpdateStation(lmTotalRecords).sMarketRep = Trim(slFields(imMap(MARKETREP)))
        tmUpdateStation(lmTotalRecords).sServiceRep = Trim(slFields(imMap(SERVICEREP)))
        tmUpdateStation(lmTotalRecords).sZone = Trim(slFields(imMap(ZONE)))
        tmUpdateStation(lmTotalRecords).sOnAir = Trim(slFields(imMap(OnAir)))
        tmUpdateStation(lmTotalRecords).sCommercial = Trim(slFields(imMap(COMMERCIAL)))
        tmUpdateStation(lmTotalRecords).sDaylight = Trim(slFields(imMap(DAYLIGHT)))
        tmUpdateStation(lmTotalRecords).lXDSStationID = Val(slFields(imMap(XDSID)))
        tmUpdateStation(lmTotalRecords).sXDSStationID = Trim$(slFields(imMap(XDSID)))
        tmUpdateStation(lmTotalRecords).sIPumpID = Trim$(slFields(imMap(IPUMPID)))
        tmUpdateStation(lmTotalRecords).sSerial1 = Trim(slFields(imMap(SERIAL1)))
        tmUpdateStation(lmTotalRecords).sSerial2 = Trim(slFields(imMap(SERIAL2)))
        tmUpdateStation(lmTotalRecords).sUsedAgreement = Trim(slFields(imMap(USEAGREEMENT)))
        tmUpdateStation(lmTotalRecords).sUsedXDS = Trim(slFields(imMap(USEXDS)))
        tmUpdateStation(lmTotalRecords).sUsedWegener = Trim(slFields(imMap(USEWEGENER)))
        tmUpdateStation(lmTotalRecords).sUsedOLA = Trim(slFields(imMap(USEOLA)))
        tmUpdateStation(lmTotalRecords).sMoniker = Trim(slFields(imMap(MONIKER)))
        tmUpdateStation(lmTotalRecords).lWatts = Val(gRemoveChar(slFields(imMap(WATTS)), ","))
        tmUpdateStation(lmTotalRecords).sWatts = Trim$(slFields(imMap(WATTS)))
        tmUpdateStation(lmTotalRecords).sHistoricalDate = Trim$(slFields(imMap(HISTORICALDATE)))
        tmUpdateStation(lmTotalRecords).sTransactID = Trim$(slFields(imMap(ENTERPRISEID)))
        tmUpdateStation(lmTotalRecords).lP12Plus = Val(gRemoveChar(slFields(imMap(P12PLUS)), ","))
        tmUpdateStation(lmTotalRecords).sP12Plus = Trim$(slFields(imMap(P12PLUS)))
        tmUpdateStation(lmTotalRecords).sWebAddress = Trim(slFields(imMap(WEBADDR)))
        tmUpdateStation(lmTotalRecords).sWebPassword = Trim(slFields(imMap(WEBPW)))
        tmUpdateStation(lmTotalRecords).sMailAddress1 = Trim(slFields(imMap(MAILADDR1)))
        tmUpdateStation(lmTotalRecords).sMailAddress2 = Trim(slFields(imMap(MAILADDR2)))
        tmUpdateStation(lmTotalRecords).sMailCity = Trim(slFields(imMap(MAILCITY)))
        tmUpdateStation(lmTotalRecords).sMailState = Trim(slFields(imMap(MAILSTATE)))
        tmUpdateStation(lmTotalRecords).sMailZip = Trim(slFields(imMap(MAILZIP)))
        tmUpdateStation(lmTotalRecords).sMailCountry = Trim(slFields(imMap(MAILCOUNTRY)))
        tmUpdateStation(lmTotalRecords).sPhysicalAddress1 = Trim(slFields(imMap(PHYSICALADDR1)))
        tmUpdateStation(lmTotalRecords).sPhysicalAddress2 = Trim(slFields(imMap(PHYSICALADDR2)))
        tmUpdateStation(lmTotalRecords).sPhysicalCity = Trim(slFields(imMap(PHYSICALCITY)))
        tmUpdateStation(lmTotalRecords).sPhysicalState = Trim(slFields(imMap(PHYSICALSTATE)))
        tmUpdateStation(lmTotalRecords).sPhysicalZip = Trim(slFields(imMap(PHYSICALZIP)))
        tmUpdateStation(lmTotalRecords).sPhone = Trim(slFields(imMap(PHONE)))
        tmUpdateStation(lmTotalRecords).sFax = Trim(slFields(imMap(FAX)))
        For ilPerson = 1 To imPersonTitles Step 1
            tmUpdateStation(lmTotalRecords).sPersonName(ilPerson - 1) = Trim$(slFields(imMap(PERSON + PNAME + 8 * (ilPerson - 1))))
            tmUpdateStation(lmTotalRecords).sPersonTitle(ilPerson - 1) = Trim$(slFields(imMap(PERSON + PTITLE + 8 * (ilPerson - 1))))
            tmUpdateStation(lmTotalRecords).sPersonPhone(ilPerson - 1) = Trim$(slFields(imMap(PERSON + PPHONE + 8 * (ilPerson - 1))))
            tmUpdateStation(lmTotalRecords).sPersonFax(ilPerson - 1) = Trim$(slFields(imMap(PERSON + PFAX + 8 * (ilPerson - 1))))
            tmUpdateStation(lmTotalRecords).sPersonEMail(ilPerson - 1) = Trim$(slFields(imMap(PERSON + PEMAIL + 8 * (ilPerson - 1))))
            tmUpdateStation(lmTotalRecords).sPersonAffLabel(ilPerson - 1) = Trim$(slFields(imMap(PERSON + PAFFLABEL + 8 * (ilPerson - 1))))
            tmUpdateStation(lmTotalRecords).sPersonISCIExport(ilPerson - 1) = Trim$(slFields(imMap(PERSON + PISCIEXPORT + 8 * (ilPerson - 1))))
            tmUpdateStation(lmTotalRecords).sPersonAffEMail(ilPerson - 1) = Trim$(slFields(imMap(PERSON + PAFFEMAIL + 8 * (ilPerson - 1))))
        Next ilPerson
        
        lmTotalRecords = lmTotalRecords + 1
        If lmTotalRecords >= llMaxRecords Then
            ' Allocate another 1000 entries.
            llMaxRecords = llMaxRecords + 5000
            ReDim Preserve tmUpdateStation(0 To llMaxRecords) As UPDATESTATION
        End If
    End If
End Function

Private Function mZoneChange(ilShttCode As Integer, llUpdateStationIndex As Long, slOldTimeZone As String, slNewTimeZone As String) As Integer
    'D.S. 1/7/07
    'This function replaces the mRemapTimes function
    'Purpose: When the time zone is changed for a station it needs to update the AST, DAT and the web site with the new
    'time and date values for all of the agreements that are for that station and are current.
    
    Dim ilRet As Integer
    Dim ilDay As Integer
    Dim ilLDay As Integer
    Dim ilTDay As Integer
    Dim ilZone As Integer
    Dim ilOldTimeAdj As Integer
    Dim ilNewTimeAdj As Integer
    Dim ilFinalTimeAdj As Integer
    Dim llVefCode As Long
    Dim llAtt As Long
    Dim llFdTime As Long
    Dim llPdTime As Long
    Dim slFileName As String
    Dim slCurDtTime As String
    Dim slCurDate As String
    Dim slCurTime As String
    Dim slVefName As String
    Dim slFdStTime As String
    Dim slFdEdTime As String
    Dim tmpStr As String
    
    Dim slFdTime As String
    Dim slFdDate As String
    Dim slPdStTime As String
    Dim slPdEdTime As String
    Dim slPdDate As String
    Dim slAirTime As String
    Dim slAirDate As String
    Dim slLastPostedDate As String
    Dim slDelay As String
    Dim slToFile As String
    
    Dim attrst As ADODB.Recordset
    Dim astrst As ADODB.Recordset
    Dim DATRST As ADODB.Recordset

    ReDim ilFdDay(0 To 6) As Integer
    ReDim ilPdDay(0 To 6) As Integer
    '7701
    Dim slAttExportToUnivision As String
    Dim slattExportToMarketron As String
    Dim slattExportToCBS As String
    Dim slattExportToClearCh As String
    'problem with date/time format Dan 10/15/15
    Dim slDate As String
    Dim slTime As String
    Dim slSQLQuery As String
    Dim llLoop As Long
    Dim slCDStartTime As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim llSpotCount As Long
    Dim llRowsEffected As Long
        
    On Error GoTo ErrHand
    
    mZoneChange = False
    '10/3/18: Dan- In the routine gVatSetToGoToWebByShttCode I will ignore the VendorID. Dan no longer ignored
    'TTP 8824 reopened
    ''7941 time zone change? Update web on next export
    gVatSetToGoToWebByShttCode ilShttCode, Vendors.XDS_Break
    
    ReDim lmAttCode(0 To 0) As Long
    
    slCurDtTime = Format(Now(), "ddddd ttttt")
    slCurDate = Format(slCurDtTime, sgShowDateForm)
    slCurTime = Format(slCurDtTime, "hh:mm:ss")
    
    ilDay = Weekday(slCurDate, vbMonday) - 1
    
    'D.S. Gather the Avail agreements that are current as of today. Do not get the CD/Tape or Daypart agreements
    slSQLQuery = "SELECT *"
    slSQLQuery = slSQLQuery + " FROM att"
    slSQLQuery = slSQLQuery + " WHERE (attShfCode = " & ilShttCode
    slSQLQuery = slSQLQuery + " AND attOffAir >= '" & Format$(slCurDtTime, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery + " AND attDropDate >= '" & Format$(slCurDtTime, sgSQLDateForm) & "'" & ")"
    slSQLQuery = slSQLQuery + " AND (attTimeType = 1 " & ")"
    
    Set attrst = gSQLSelectCall(slSQLQuery)
    'D.S. Build an array of the attcodes we got from above call
    While Not attrst.EOF
        lmAttCode(UBound(lmAttCode)) = attrst!attCode
        ReDim Preserve lmAttCode(0 To UBound(lmAttCode) + 1) As Long
        attrst.MoveNext
    Wend

    For llAtt = 0 To UBound(lmAttCode) - 1 Step 1
        DoEvents
        slSQLQuery = "SELECT *"
        slSQLQuery = slSQLQuery + " FROM att "
        slSQLQuery = slSQLQuery + " WHERE (attCode = " & lmAttCode(llAtt) & ")"
        Set attrst = gSQLSelectCall(slSQLQuery)
        If Not attrst.EOF Then
            ilOldTimeAdj = 0
            ilNewTimeAdj = 0
            llVefCode = gBinarySearchVef(CLng(attrst!attvefCode))
            
            'Get the Zone offsets
            If llVefCode <> -1 Then
                slVefName = tgVehicleInfo(llVefCode).sVehicle
                For ilZone = LBound(tgVehicleInfo(llVefCode).sZone) To UBound(tgVehicleInfo(llVefCode).sZone) Step 1
                    If StrComp(slOldTimeZone, Left(tgVehicleInfo(llVefCode).sZone(ilZone), 1), 1) = 0 Then
                        ilOldTimeAdj = tgVehicleInfo(llVefCode).iVehLocalAdj(ilZone)
                        Exit For
                    End If
                Next ilZone
                
                For ilZone = LBound(tgVehicleInfo(llVefCode).sZone) To UBound(tgVehicleInfo(llVefCode).sZone) Step 1
                    If StrComp(slNewTimeZone, Left(tgVehicleInfo(llVefCode).sZone(ilZone), 1), 1) = 0 Then
                        ilNewTimeAdj = tgVehicleInfo(llVefCode).iVehLocalAdj(ilZone)
                        Exit For
                    End If
                Next ilZone
            End If
           
            ilFinalTimeAdj = ilNewTimeAdj - ilOldTimeAdj
           
        '********************************* Update DAT *********************************
           
            slSQLQuery = "SELECT * "
            slSQLQuery = slSQLQuery + " FROM dat"
            slSQLQuery = slSQLQuery + " WHERE datAtfCode= " & lmAttCode(llAtt)
            slSQLQuery = slSQLQuery & " ORDER BY datFdStTime"
            Set DATRST = gSQLSelectCall(slSQLQuery)
            DoEvents
            While Not DATRST.EOF
                For ilLDay = 0 To 6 Step 1
                    ilFdDay(ilLDay) = 0
                Next ilLDay
                If (DATRST!datFdMon = 1) Then
                    ilFdDay(0) = 1
                End If
                If (DATRST!datFdTue = 1) Then
                    ilFdDay(1) = 1
                End If
                If (DATRST!datFdWed = 1) Then
                    ilFdDay(2) = 1
                End If
                If (DATRST!datFdThu = 1) Then
                    ilFdDay(3) = 1
                End If
                If (DATRST!datFdFri = 1) Then
                    ilFdDay(4) = 1
                End If
                If (DATRST!datFdSat = 1) Then
                    ilFdDay(5) = 1
                End If
                If (DATRST!datFdSun = 1) Then
                    ilFdDay(6) = 1
                End If
            
                llFdTime = gTimeToLong(DATRST!datFdStTime, False) + 3600 * (ilNewTimeAdj - ilOldTimeAdj)
                If llFdTime < 0 Then
                    llFdTime = llFdTime + 86400
                    ilTDay = ilFdDay(0)
                    For ilLDay = 0 To 5 Step 1
                        ilFdDay(ilLDay) = ilFdDay(ilLDay + 1)
                    Next ilLDay
                    ilFdDay(6) = ilTDay
                ElseIf llFdTime > 86400 Then
                    llFdTime = llFdTime - 86400
                    ilTDay = ilFdDay(6)
                    For ilLDay = 5 To 0 Step -1
                        ilFdDay(ilLDay + 1) = ilFdDay(ilLDay)
                    Next ilLDay
                    ilFdDay(0) = ilTDay
                End If
                slFdStTime = Format$(gLongToTime(llFdTime), "hh:mm:ss")
                llFdTime = gTimeToLong(DATRST!datFdEdTime, False) + 3600 * (ilNewTimeAdj - ilOldTimeAdj)
                If llFdTime < 0 Then
                    llFdTime = llFdTime + 86400
                ElseIf llFdTime > 86400 Then
                    llFdTime = llFdTime - 86400
                End If
                slFdEdTime = Format$(gLongToTime(llFdTime), "hh:mm:ss")
                If DATRST!datFdStatus = 0 Then
                    slPdStTime = slFdStTime
                    slPdEdTime = slFdEdTime
                'Else
                    If ilNewTimeAdj <> ilOldTimeAdj Then
                        slDelay = "Check Delays"
                    End If
                    For ilLDay = 0 To 6 Step 1
                        ilPdDay(ilLDay) = 0
                    Next ilLDay
                    If (DATRST!datPdMon = 1) Then
                        ilPdDay(0) = 1
                    End If
                    If (DATRST!datPdTue = 1) Then
                        ilPdDay(1) = 1
                    End If
                    If (DATRST!datPdWed = 1) Then
                        ilPdDay(2) = 1
                    End If
                    If (DATRST!datPdThu = 1) Then
                        ilPdDay(3) = 1
                    End If
                    If (DATRST!datPdFri = 1) Then
                        ilPdDay(4) = 1
                    End If
                    If (DATRST!datPdSat = 1) Then
                        ilPdDay(5) = 1
                    End If
                    If (DATRST!datPdSun = 1) Then
                        ilPdDay(6) = 1
                    End If
                    llPdTime = gTimeToLong(DATRST!datPdStTime, False) + 3600 * (ilNewTimeAdj - ilOldTimeAdj)
                    If llPdTime < 0 Then
                        llPdTime = llPdTime + 86400
                        ilTDay = ilPdDay(0)
                        For ilLDay = 0 To 5 Step 1
                            ilPdDay(ilLDay) = ilPdDay(ilLDay + 1)
                        Next ilLDay
                        ilPdDay(6) = ilTDay
                    ElseIf llPdTime > 86400 Then
                        llPdTime = llPdTime - 86400
                        ilTDay = ilPdDay(6)
                        For ilLDay = 5 To 0 Step -1
                            ilPdDay(ilLDay + 1) = ilPdDay(ilLDay)
                        Next ilLDay
                        ilPdDay(0) = ilTDay
                    End If
                    slPdStTime = Format$(gLongToTime(llPdTime), "hh:mm:ss")
                    llPdTime = gTimeToLong(DATRST!datPdEdTime, False) + 3600 * (ilNewTimeAdj - ilOldTimeAdj)
                    If llPdTime < 0 Then
                        llPdTime = llPdTime + 86400
                    ElseIf llPdTime > 86400 Then
                        llPdTime = llPdTime - 86400
                    End If
                    slPdEdTime = Format$(gLongToTime(llPdTime), "hh:mm:ss")
                End If
                
                slSQLQuery = "UPDATE dat"
                slSQLQuery = slSQLQuery & " SET datFdMon = " & ilFdDay(0) & ","
                slSQLQuery = slSQLQuery & "datFdTue = " & ilFdDay(1) & ","
                slSQLQuery = slSQLQuery & "datFdWed = " & ilFdDay(2) & ","
                slSQLQuery = slSQLQuery & "datFdThu = " & ilFdDay(3) & ","
                slSQLQuery = slSQLQuery & "datFdFri = " & ilFdDay(4) & ","
                slSQLQuery = slSQLQuery & "datFdSat = " & ilFdDay(5) & ","
                slSQLQuery = slSQLQuery & "datFdSun = " & ilFdDay(6) & ","
                
                If DATRST!datFdStatus = 0 Then
                    slSQLQuery = slSQLQuery & "datPdMon = " & ilPdDay(0) & ","
                    slSQLQuery = slSQLQuery & "datPdTue = " & ilPdDay(1) & ","
                    slSQLQuery = slSQLQuery & "datPdWed = " & ilPdDay(2) & ","
                    slSQLQuery = slSQLQuery & "datPdThu = " & ilPdDay(3) & ","
                    slSQLQuery = slSQLQuery & "datPdFri = " & ilPdDay(4) & ","
                    slSQLQuery = slSQLQuery & "datPdSat = " & ilPdDay(5) & ","
                    slSQLQuery = slSQLQuery & "datPdSun = " & ilPdDay(6) & ","
                End If
                
                If bmAdjPledge Then
                    slSQLQuery = slSQLQuery & "datFdStTime = " & "'" & Format$(DateAdd("h", ilFinalTimeAdj, DATRST!datFdStTime), sgSQLTimeForm) & "', "
                    slSQLQuery = slSQLQuery & "datFdEdTime = " & "'" & Format$(DateAdd("h", ilFinalTimeAdj, DATRST!datFdEdTime), sgSQLTimeForm) & "', "
                    slSQLQuery = slSQLQuery & "datPdStTime = " & "'" & Format$(DateAdd("h", ilFinalTimeAdj, DATRST!datPdStTime), sgSQLTimeForm) & "', "
                    slSQLQuery = slSQLQuery & "datPdEdTime = " & "'" & Format$(DateAdd("h", ilFinalTimeAdj, DATRST!datPdEdTime), sgSQLTimeForm) & "' "
                Else
                    slSQLQuery = slSQLQuery & "datFdStTime = " & "'" & Format$(DateAdd("h", ilFinalTimeAdj, DATRST!datFdStTime), sgSQLTimeForm) & "', "
                    slSQLQuery = slSQLQuery & "datFdEdTime = " & "'" & Format$(DateAdd("h", ilFinalTimeAdj, DATRST!datFdEdTime), sgSQLTimeForm) & "' "
                End If
                slSQLQuery = slSQLQuery & " WHERE (datCode = " & DATRST!datCode & ")"
                'cnn.Execute slSQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                    '6/11/16: Replaced GoSub
                    'GoSub ErrHand:
                    gHandleError "AffErrorLog.txt", "ImportUpdateStations-mZoneChange"
                    mZoneChange = False
                    Exit Function
                End If
                DATRST.MoveNext
            Wend
            
            '********************************* Update AST *********************************
            '7701
            slAttExportToUnivision = ""
            slattExportToMarketron = ""
            slattExportToCBS = ""
            slattExportToClearCh = ""
            If gIsVendorWithAgreement(lmAttCode(llAtt), Vendors.cBs) Then
                slattExportToCBS = "Y"
            End If
            If gIsVendorWithAgreement(lmAttCode(llAtt), Vendors.iheart) Then
                slattExportToClearCh = "Y"
            End If
            If gIsVendorWithAgreement(lmAttCode(llAtt), Vendors.NetworkConnect) Then
                slattExportToMarketron = "Y"
            End If
            '7701
            slLastPostedDate = gGetLastPostedDate(lmAttCode(llAtt), attrst!attExportType, attrst!attExportToWeb, slAttExportToUnivision, slattExportToMarketron, slattExportToCBS, slattExportToClearCh)
            Screen.MousePointer = vbHourglass
            'slLastPostedDate = gGetLastPostedDate(lmAttCode(llAtt), attrst!attExportType, attrst!attExportToWeb, attrst!attExportToUnivision, attrst!attExportToMarketron, attrst!attExportToCBS, attrst!attExportToClearCh)
           
            slSQLQuery = "Select * FROM ast WHERE "
            slSQLQuery = slSQLQuery + " astAtfCode = " & lmAttCode(llAtt)
            Set astrst = gSQLSelectCall(slSQLQuery)
            While Not astrst.EOF
                                 
                'Feed Date and Time
                'Dan 10/15/15 issue with format
                slDate = Format(astrst!astFeedDate, sgShowDateForm)
                slTime = Format(astrst!astFeedTime, sgShowTimeWSecForm)
                tmpStr = slDate & " " & slTime
                'tmpStr = Trim$(astrst!astFeedDate) & " " & Trim$(astrst!astFeedTime)
                tmpStr = DateAdd("h", ilFinalTimeAdj, tmpStr)
                slFdDate = Format(tmpStr, sgSQLDateForm)
                slFdTime = Format(tmpStr, sgSQLTimeForm)
                
                '12/13/13: Pledge information removed (DAT used instead)
                'Pledge Date & Start and End Time
                'tmpStr = Trim$(astrst!astPledgeDate) & " " & Trim$(astrst!astPledgeStartTime)
                'tmpStr = DateAdd("h", ilFinalTimeAdj, tmpStr)
                'slPdDate = Format(tmpStr, sgSQLDateForm)
                'slPdStTime = Format(tmpStr, sgSQLTimeForm)
                'tmpStr = Trim$(astrst!astPledgeDate) & " " & Trim$(astrst!astPledgeEndTime)
                'tmpStr = DateAdd("h", ilFinalTimeAdj, tmpStr)
                'slPdEdTime = Format(tmpStr, sgSQLTimeForm)
                
                'Air Date & Time
                'Dan 10/15/15 issue with format
                slDate = Format(astrst!astAirDate, sgShowDateForm)
                slTime = Format(astrst!astAirTime, sgShowTimeWSecForm)
                tmpStr = slDate & " " & slTime
               ' tmpStr = Trim$(astrst!astAirDate) & " " & Trim$(astrst!astAirTime)
                If bmAdjPledge Then
                    tmpStr = DateAdd("h", ilFinalTimeAdj, tmpStr)
                Else
                    tmpStr = DateAdd("h", 0, tmpStr)
                End If
                slAirDate = Format(tmpStr, sgSQLDateForm)
                slAirTime = Format(tmpStr, sgSQLTimeForm)
                
                'see if the feed changed due to the zone changed
                '0 = no date changed, 1 = date moved forward a day, -1 date moved back a day
                'dan m 10/15/15 more time format issues
                slTime = Format(astrst!astFeedTime, sgShowTimeWSecForm)
                ilRet = mZoneChangesDate(ilFinalTimeAdj, slTime)
               ' ilRet = mZoneChangesDate(ilFinalTimeAdj, astrst!astFeedTime)
                If ilRet = 0 Then
                    'No feed date change so we can update the record
                    slSQLQuery = "UPDATE ast SET"
                    slSQLQuery = slSQLQuery & " astFeedTime = " & "'" & slFdTime & "', "
                    slSQLQuery = slSQLQuery & " astFeedDate = " & "'" & slFdDate & "', "
                    
                    If astrst!astCPStatus <> 1 Then
                        'If the ast has been posted then don't update the airdate or airtime
                        slSQLQuery = slSQLQuery & " astAirTime = " & "'" & slAirTime & "', "
                        slSQLQuery = slSQLQuery & " astAirDate = " & "'" & slAirDate & "', "
                    End If
                    
                    '12/13/13: Pledge information removed (DAT used instead)
                    'slSQLQuery = slSQLQuery & " astPledgeStartTime = " & "'" & slPdStTime & "', "
                    'slSQLQuery = slSQLQuery & " astPledgeEndTime = " & "'" & slPdEdTime & "', "
                    'slSQLQuery = slSQLQuery & " astPledgeDate = " & "'" & slPdDate & "'"
                    slSQLQuery = slSQLQuery & " astUstCode = " & igUstCode
                    slSQLQuery = slSQLQuery & " WHERE astCode = " & astrst!astCode
                    'cnn.Execute slSQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mZoneChange"
                        mZoneChange = False
                        Exit Function
                    End If
                Else
                    'Feed date changed so, we must delete the old record and insert a now one with the same astCode
                    slSQLQuery = "DELETE FROM Ast WHERE astCode = " & astrst!astCode
                    'cnn.Execute slSQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mZoneChange"
                        mZoneChange = False
                        Exit Function
                    End If

                    slSQLQuery = "INSERT INTO ast"
                    slSQLQuery = slSQLQuery + "(astCode, astAtfCode, astShfCode, astVefCode, "
                    slSQLQuery = slSQLQuery + "astSdfCode, astLsfCode, astAirDate, astAirTime, "
                    '12/13/13: Support New AST layout
                    'slSQLQuery = slSQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, "
                    'slSQLQuery = slSQLQuery + "astPledgeStartTime, astPledgeEndTime, astPledgeStatus)"
                    slSQLQuery = slSQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, "
                    slSQLQuery = slSQLQuery + "astAdfCode, astDatCode, astCpfCode, astRsfCode, astStationCompliant, astAgencyCompliant, astAffidavitSource, astCntrNo, astLen, astLkAstCode, astMissedMnfCode, astUstCode)"
                    slSQLQuery = slSQLQuery + " VALUES "
                    slSQLQuery = slSQLQuery + "(" & astrst!astCode & ", " & astrst!astAtfCode & ", " & astrst!astShfCode & ", "
                    slSQLQuery = slSQLQuery & astrst!astVefCode & ", " & astrst!astSdfCode & ", " & astrst!astLsfCode & ", "
                    
                    If astrst!astCPStatus <> 1 Then
                        'If the ast has been posted then don't update the airdate or airtime
                        slSQLQuery = slSQLQuery + "'" & Format$(slAirDate, sgSQLDateForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "', "
                    Else
                        'Use the old posted date and time
                        slSQLQuery = slSQLQuery + "'" & Format$(astrst!astAirDate, sgSQLDateForm) & "', '" & Format$(astrst!astAirTime, sgSQLTimeForm) & "', "
                    End If
                    
                    slSQLQuery = slSQLQuery & astrst!astStatus & ", " & astrst!astCPStatus & ", '" & Format$(slFdDate, sgSQLDateForm) & "', "
                    '12/13/13: Support New AST layout
                    'slSQLQuery = slSQLQuery & "'" & Format$(slFdTime, sgSQLTimeForm) & "', '" & Format$(slPdDate, sgSQLDateForm) & "', "
                    'slSQLQuery = slSQLQuery & "'" & Format$(slPdStTime, sgSQLTimeForm) & "', '" & Format$(slPdEdTime, sgSQLTimeForm) & "', " & astrst!astPledgeStatus & ")"
                    slSQLQuery = slSQLQuery & "'" & Format$(slFdTime, sgSQLTimeForm) & "', "
                    slSQLQuery = slSQLQuery & astrst!astAdfCode & ", " & astrst!astDatCode & ", " & astrst!astCpfCode & ", " & astrst!astRsfCode & ", "
                    slSQLQuery = slSQLQuery & "'" & astrst!astStationCompliant & "', '" & astrst!astAgencyCompliant & "', '" & gRemoveIllegalChars(astrst!astAffidavitSource) & "', " & astrst!astCntrNo & ", " & astrst!astLen & ", " & astrst!astLkAstCode & ", " & astrst!astMissedMnfCode & ", " & igUstCode & ")"
                    
                    'cnn.Execute slSQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        '6/11/16: Replaced GoSub
                        'GoSub ErrHand:
                        gHandleError "AffErrorLog.txt", "ImportUpdateStations-mZoneChange"
                        mZoneChange = False
                        Exit Function
                    End If
                End If
                astrst.MoveNext
            Wend
            
            '4/2/16: Update program times
            'For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            '    If tgStationInfo(llLoop).iCode = ilShttCode Then
            '        If Len(slNewTimeZone) = 1 Then
            '            tgStationInfo(llLoop).sZone = slNewTimeZone & "ST"
            '        Else
            '            tgStationInfo(llLoop).sZone = slNewTimeZone
            '        End If
            '        slCDStartTime = ""
            '        ilRet = gDetermineAgreementTimes(ilShttCode, attrst!attvefCode, Format$(attrst!attOnAir, "m/d/yy"), Format$(attrst!attOffAir, "m/d/yy"), Format$(attrst!attDropDate, "m/d/yy"), slCDStartTime, slStartTime, slEndTime)
            '        slSQLQuery = "Update att Set "
            '        slSQLQuery = slSQLQuery & "attVehProgStartTime = '" & Format$(slStartTime, sgSQLTimeForm) & "', "
            '        slSQLQuery = slSQLQuery & "attVehProgEndTime = '" & Format$(slEndTime, sgSQLTimeForm) & "'"
            '        slSQLQuery = slSQLQuery & " Where attCode = " & attrst!attCode
            '        'cnn.Execute SQLQuery, rdExecDirect
            '        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '            GoSub ErrHand:
            '        End If
            '        Exit For
            '    End If
            'Next llLoop
            slStartTime = "1/1/2000 " & Format(attrst!attVehProgStartTime, sgShowTimeWSecForm)
            slStartTime = DateAdd("h", ilFinalTimeAdj, slStartTime)
            slEndTime = "1/1/2000 " & Format(attrst!attVehProgEndTime, sgShowTimeWSecForm)
            slEndTime = DateAdd("h", ilFinalTimeAdj, slEndTime)
            slSQLQuery = "Update att Set "
            slSQLQuery = slSQLQuery & "attVehProgStartTime = '" & Format$(slStartTime, sgSQLTimeForm) & "', "
            slSQLQuery = slSQLQuery & "attVehProgEndTime = '" & Format$(slEndTime, sgSQLTimeForm) & "'"
            slSQLQuery = slSQLQuery & " Where attCode = " & attrst!attCode
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                gHandleError "AffErrorLog.txt", "ImportUpdateStations-mZoneChange"
                mZoneChange = False
                Exit Function
            End If

            '********************************* Update Web Site *********************************
            'ilRet = gAdjustWebTimeZone(lmAttCode(llAtt), 3600 * ilFinalTimeAdj)

            If (slattExportToMarketron <> "Y") Then
                slSQLQuery = "Select Count(*) from Spots Where attCode = " & lmAttCode(llAtt)
                llSpotCount = gExecWebSQLWithRowsEffected(slSQLQuery)
                If llSpotCount > 0 Then
                    slSQLQuery = "Update Spots Set FeedTime = DateAdd(HOUR, " & ilFinalTimeAdj & ", FeedTime) Where attCode = " & lmAttCode(llAtt)
                    llRowsEffected = gExecWebSQLWithRowsEffected(slSQLQuery)
                    If llRowsEffected = -1 Then
                        mZoneChange = False
                        'gLogMsg "Error: Failed to Update Web: " & SQLQuery, "USRNAdjustmentLog.txt", False
                        Call mUpdateReport(llUpdateStationIndex, "Error: Failed to Update Web: " & slSQLQuery)
                        Exit Function
                    End If
                End If
                slSQLQuery = "Select Count(*) from Spot_History Where attCode = " & lmAttCode(llAtt)
                llSpotCount = gExecWebSQLWithRowsEffected(slSQLQuery)
                If llSpotCount > 0 Then
                    slSQLQuery = "Update Spot_History Set FeedTime = DateAdd(HOUR, " & ilFinalTimeAdj & ", FeedTime) Where attCode = " & lmAttCode(llAtt)
                    llRowsEffected = gExecWebSQLWithRowsEffected(slSQLQuery)
                    If llRowsEffected = -1 Then
                        mZoneChange = False
                        'gLogMsg "Error: Failed to Update Web: " & slSQLQuery, "USRNAdjustmentLog.txt", False
                        Call mUpdateReport(llUpdateStationIndex, "Error: Failed to Update Web: " & slSQLQuery)
                        Exit Function
                    End If
                End If
            End If
            
            If slattExportToMarketron = "Y" Then
                Call mUpdateReport(llUpdateStationIndex, "This was a Marketron agreement.  You need to re-export the spots.")
            End If

        End If
    Next llAtt
    
    mZoneChange = True
    Exit Function
ErrHand:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mZoneChange"
End Function

Private Function mZoneChangesDate(iOffSet As Integer, sDatTime As String) As Integer
    Dim slOrigDtTime As String
    Dim slNewDtTime As String
    
    slOrigDtTime = "1/1/2000 " & sDatTime
    slNewDtTime = DateAdd("h", iOffSet, slOrigDtTime)
    
    'did we back up one day?
    If DateValue(slOrigDtTime) < DateValue(slNewDtTime) Then
        mZoneChangesDate = -1
        Exit Function
    End If
    
    'did we go forward one day?
    If DateValue(slOrigDtTime) > DateValue(slNewDtTime) Then
        mZoneChangesDate = 1
        Exit Function
    End If
    
    'No change in date
    mZoneChangesDate = 0
End Function

Private Sub mGetBands()
    Dim slSQLQuery As String
    Dim ilField As Integer
    Dim slStr As String
    
    lacBands.Caption = ""
    slSQLQuery = "SELECT cmtPart1,cmtPart2, cmtPart3 , cmtPart4 FROM Site Left Outer Join cmt on siteBandCmtCode = cmtCode Where siteCode = 1"
    Set rst_cmt = gSQLSelectCall(slSQLQuery)
    If Not rst_cmt.EOF Then
        lacBands.Caption = "AM, FM" ', " & Trim$(rst_cmt!cmtPart1 & rst_cmt!cmtPart2 & rst_cmt!cmtPart3 & rst_cmt!cmtPart4)
        If IsNull(rst_cmt!cmtPart1) = False Then
            If Trim$(rst_cmt!cmtPart1) <> "" Then
                slStr = slStr & ", " & Trim$(rst_cmt!cmtPart1)
            End If
        End If
        If IsNull(rst_cmt!cmtPart2) = False Then
            If Trim$(rst_cmt!cmtPart2) <> "" Then
                slStr = slStr & ", " & Trim$(rst_cmt!cmtPart2)
            End If
        End If
        If IsNull(rst_cmt!cmtPart3) = False Then
            If Trim$(rst_cmt!cmtPart3) <> "" Then
                slStr = slStr & ", " & Trim$(rst_cmt!cmtPart3)
            End If
        End If
        If IsNull(rst_cmt!cmtPart4) = False Then
            If Trim$(rst_cmt!cmtPart4) <> "" Then
                slStr = slStr & ", " & Trim$(rst_cmt!cmtPart4)
            End If
        End If
        If slStr <> "" Then
            lacBands.Caption = lacBands.Caption & slStr
        End If
    Else
        lacBands.Caption = "AM, FM"
    End If
    smBandFields = Split(lacBands.Caption, ",")
    If IsArray(smBandFields) Then
        For ilField = 0 To UBound(smBandFields)
            smBandFields(ilField) = Trim$(UCase$(smBandFields(ilField)))
        Next ilField
    End If
    lacBands.Caption = "Allowed Station Bands: " & lacBands.Caption
End Sub

' JD TTP 10860
Function mUpdateWebSite()
    Dim slSQLQuery As String
    Dim ilRowsEffected As Integer
    Dim ilTotalRecordsUpdated As Integer

    On Error GoTo err1
    ilTotalRecordsUpdated = 0
    slSQLQuery = "Select arttCode, arttFirstName, arttLastName, arttEmail "
    slSQLQuery = slSQLQuery & " FROM artt Where arttEMailToWeb = 'W' And arttWebEmail = 'Y'"
    Set rst_artt = gSQLSelectCall(slSQLQuery)
    
    Do While Not rst_artt.EOF
        slSQLQuery = "Update WebEMT set FirstName = '" & Trim$(rst_artt!arttFirstName) & "', "
        slSQLQuery = slSQLQuery & "LastName = '" & Trim$(rst_artt!arttLastName) & "', "
        slSQLQuery = slSQLQuery & "Email = '" & Trim$(rst_artt!arttEmail) & "' "
        slSQLQuery = slSQLQuery & " Where Code = " & rst_artt!arttCode
        
        ilRowsEffected = gExecWebSQLWithRowsEffected(slSQLQuery)
        rst_artt.MoveNext
        ilTotalRecordsUpdated = ilTotalRecordsUpdated + 1
    Loop
    
    If ilTotalRecordsUpdated > 0 Then
        ' Turn them all back off
        slSQLQuery = "Update artt Set arttEMailToWeb = ' ' Where arttEMailToWeb = 'W'"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            gHandleError "AffErrorLog.txt", "ImportUpdateStations-mUpdateWebSite"
        End If
    End If
    Exit Function
err1:
    gLogMsg "A general error has occured in mUpdateWebSite: ", "AffErrorLog.Txt", False
End Function

'***************************************************************************
'
'***************************************************************************
Private Function mRenameCallLetters(llStationIDX As Long, slCallLetters As String) As Integer
    Dim slOldCallLetters As String
    Dim slSQLQuery As String
    Dim ilRet As Integer
    
    On Error GoTo ErrHandler:
    mRenameCallLetters = False
    slOldCallLetters = Trim$(tgStationInfo(llStationIDX).sCallLetters)

    slSQLQuery = "Update shtt Set shttCallLetters = '" & Trim$(slCallLetters) & "'" & " Where shttCode = " & tgStationInfo(llStationIDX).iCode
    If bmUpdateDatabase Then
        'cnn.Execute slSQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHandler:
            gHandleError "AffErrorLog.txt", "ImportUpdateStations-mUpdateCallLetters"
            mRenameCallLetters = False
            Exit Function
        End If
    End If
    'tmUpdateStation(llStationIDX).sCallLetters = slOldCallLetters
    'Call mUpdateReport(llStationIDX, slOldCallLetters & " changed to " & Trim(slCallLetters))
    Call mUpdateReport2(slOldCallLetters, tgStationInfo(llStationIDX).lPermStationID, slOldCallLetters & " changed to " & Trim(slCallLetters))
    tgStationInfo(llStationIDX).sCallLetters = Trim(slCallLetters)
    
    mRenameCallLetters = True
    Exit Function

ErrHandler:
    gHandleError "AffErorLog.txt", "frmImportUpdateStations-mUpdateCallLetters"
    Exit Function
End Function

Private Function mGetAvailableInvalidNo() As Integer
    'TTP 10984 - Station Information Import: when using Station IDs, add method that will allow call letters to be swapped
    'Find the next InvalidX number
    Dim ilInvalidNo  As Integer
    Dim ilFound As Integer
    Dim llLoop As Integer
    ilInvalidNo = 1
    Do
        ilFound = False
        For llLoop = 0 To UBound(tgStationInfo) - 1 Step 1
            If Mid(Trim(tgStationInfo(llLoop).sCallLetters), 1, 8) = "Invalid" & Trim(Str(ilInvalidNo)) Then
                ilInvalidNo = ilInvalidNo + 1
                ilFound = True
            End If
        Next llLoop
    Loop While ilFound = True
    mGetAvailableInvalidNo = ilInvalidNo
End Function
