VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportCSISpot 
   Caption         =   "Import CSI Traffic Spots"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   Icon            =   "AffImportCSISpot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   6195
   Begin VB.TextBox txtFile 
      Height          =   300
      Left            =   990
      TabIndex        =   6
      Top             =   480
      Width           =   3600
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse"
      Height          =   300
      Left            =   4845
      TabIndex        =   5
      Top             =   480
      Width           =   1065
   End
   Begin VB.ListBox lbcMsg 
      Enabled         =   0   'False
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   1395
      Width           =   5790
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5790
      Top             =   4305
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4890
      FormDesignWidth =   6195
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   1125
      TabIndex        =   2
      Top             =   4380
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3150
      TabIndex        =   3
      Top             =   4380
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5910
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "csv"
   End
   Begin VB.Label lbcFile 
      Caption         =   "Import File"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   495
      Width           =   780
   End
   Begin VB.Label lacResult 
      Height          =   480
      Left            =   150
      TabIndex        =   4
      Top             =   3765
      Width           =   5490
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5790
   End
End
Attribute VB_Name = "frmImportCSISpot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmContact - allows for selection of station/vehicle/advertiser for contact information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Option Compare Text

'Record Layout: Start
'Date:,Generation Date,Generation Time
'Vehicle:,Vehicle Name,Vehicle Station Code
'Event:,Season,GameNo,Date,Visiting Team Name,Visiting Team Abbreviation,Home Team Name,Home Team Abbreviation,Start Time,Status,Language,English,Feed Source
'Spot:,Time,Position #,Break #,Advertiser Name,Advertiser Abbreviation,Product,Spot Length,Cart,ISCI,Creative Title,Avail Name,Zone
'Copy:,Type,Advertiser Name,Advertiser Abbreviation,Cart,ISCI,Product,Creative Title,Call Letters,Station ID
'Record Layout: End


Private imImporting As Integer
Private imTerminate As Integer
'Private hmMsg As Integer
Private hmFrom As Integer
Private tmClearImportInfo() As CLEASRIMPORTINFO
Private smFileNames() As String
'4/12/15: Update CPTTAstStatus
Private imPrevLLDVefCode As Integer
Private lmPrevLLDDate As Long
Private imPrevCPTTVefCode As Integer
Private lmPrevCPTTDate As Long
Private lst_rst As ADODB.Recordset
Private att_rst As ADODB.Recordset
Dim smEventTitle1 As String
Dim smEventTitle2 As String


Private Sub cmcBrowse_Click()

    Dim slCurDir As String
    
    slCurDir = CurDir
    ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    CommonDialog1.InitDir = sgImportDirectory
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNAllowMultiselect
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    txtFile.Text = Trim$(CommonDialog1.fileName)
    ChDir slCurDir
    Exit Sub
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmdImport_Click()
    Dim iRet As Integer
    Dim ilFile As Integer

    On Error GoTo ErrHand
    
    lbcMsg.Clear
    lbcMsg.Enabled = True
    Screen.MousePointer = vbHourglass
    'Get File Names
    If Not mGetFileNames() Then
        Screen.MousePointer = vbDefault
        Beep
        gMsgBox "No File Names Specified to Import", vbCritical
        txtFile.SetFocus
        Exit Sub
    End If
    If Not mCheckFile() Then
        Beep
        lbcMsg.AddItem "Import Stopped"
        lacResult.Caption = "see CSIImportLog.Txt for list reasons import stopped"
        Screen.MousePointer = vbDefault
        txtFile.SetFocus
        Exit Sub
    End If
    lbcMsg.AddItem "Import File Structure is Ok"
    imImporting = True
    For ilFile = 0 To UBound(smFileNames) - 1 Step 1
        On Error GoTo 0
        lacResult.Caption = ""
        iRet = mImportSpots(smFileNames(ilFile))
        If (iRet = False) Then
            gLogMsg "** Error during Import **", "CSIImportLog.Txt", False
            lacResult.Caption = "see CSIImportLog.Txt for list reasons Import stopped"
            'Print #hmMsg, "** Terminated **"
            'Close #hmMsg
            'Close #hmTo
            imImporting = False
            Screen.MousePointer = vbDefault
            cmdCancel.SetFocus
            Exit Sub
        End If
        If imTerminate Then
            gLogMsg "** User Terminated Import**", "CSIImportLog.Txt", False
            lacResult.Caption = "see CSIImportLog.Txt for list reasons import stopped"
            'Print #hmMsg, "** User Terminated **"
            'Close #hmMsg
            'Close #hmTo
            imImporting = False
            Screen.MousePointer = vbDefault
            cmdCancel.SetFocus
            Exit Sub
        End If
    Next ilFile
    'Clear old aet records out
    On Error GoTo ErrHand:
    imImporting = False
    gLogMsg "** Completed Import Aired Station Spots" & " **", "CSIImportLog.Txt", False
    gLogMsg "", "CSIImportLog.Txt", False
    'Print #hmMsg, "** Completed Import Aired Station Spots: " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
    'Close #hmMsg
    lacResult.Caption = "Completed Import Aired Spots"
    cmdImport.Enabled = False
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Exit Sub
cmdImportErr:
    iRet = Err
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "ImportCSISpots-cmdImport"
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If imImporting Then
        imTerminate = True
        Exit Sub
    End If
    Unload frmImportCSISpot
End Sub


Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    imTerminate = False
    imImporting = False
    '4/12/15: Update CPTTAstStatus
    imPrevLLDVefCode = -1
    lmPrevLLDDate = 0
    imPrevCPTTVefCode = -1
    lmPrevCPTTDate = 0
    
    txtFile.Text = ""   'sgImportDirectory & "CSISpots.txt"
    ilRet = gPopAdvertisers()
    ilRet = gPopTeams()
    ilRet = gPopLangs()
    ilRet = gPopAvailNames()
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmClearImportInfo
    Erase smFileNames
    lst_rst.Close
    att_rst.Close
    Set frmImportCSISpot = Nothing
End Sub

Private Function mImportSpots(slFileName As String) As Integer
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim ilLoop As Integer
    Dim ilStartLayoutFd As Integer
    Dim ilEndLayoutFd As Integer
    Dim ilEndSpotsFd As Integer
    Dim ilDateFd As Integer
    Dim ilVehicleFd As Integer
    Dim ilGameFd As Integer
    Dim slVehicleName As String
    Dim slVehicleStationCode As String
    Dim ilVefCode As Integer
    Dim ilVef As Integer
    Dim ilTeam As Integer
    Dim ilRemoved As Integer
    Dim ilVisitMnfCode As Integer
    Dim ilHomeMnfCode As Integer
    Dim ilLang As Integer
    Dim ilLangMnfCode As Integer
    Dim slFeedSource As String
    Dim ilFirstSpot As Integer
    Dim llGhfCode As Long
    Dim llGsfCode As Long
    Dim ilGameNo As Integer
    Dim slLogDate As String
    Dim slGameStartTime As String
    Dim slGameStatus As String
    Dim ilAdf As Integer
    Dim ilAdfCode As Integer
    Dim ilAnf As Integer
    Dim ilAnfcode As Integer
    Dim llVpf As Long
    Dim slSeason As String
    Dim llLstCode As Long
    Dim ilShtt As Integer
    Dim ilShttCode As Integer
    Dim llStationID As Long
    'Dim slFields(1 To 16) As String
    Dim slFields(0 To 15) As String
    
    ilStartLayoutFd = False
    ilEndLayoutFd = False
    ilEndSpotsFd = False
    ilDateFd = False
    ilVehicleFd = False
    ilGameFd = False
    ilRemoved = False
    ilFirstSpot = True
    ReDim tmClearImportInfo(0 To 0) As CLEASRIMPORTINFO
    slFromFile = slFileName 'txtFile.Text
    gLogMsg "Importing: " & slFromFile, "CSIImportLog.Txt", False
    'ilRet = 0
    'On Error GoTo mImportSpotsErr:
    'hmFrom = FreeFile
    'Open slFromFile For Input Access Read As hmFrom
    ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
    If ilRet <> 0 Then
        Close hmFrom
        Exit Function
    End If
    Do While Not EOF(hmFrom)
        ilRet = 0
        On Error GoTo mImportSpotsErr:
        Line Input #hmFrom, slLine
        On Error GoTo 0
        If ilRet = 62 Then
            ilRet = 0
            Exit Do
        End If
        On Error GoTo ErrHand
        slLine = Trim$(slLine)
        If Len(slLine) > 0 Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            Else
                'Process Input
                gParseCDFields slLine, False, slFields()
                For ilLoop = LBound(slFields) To UBound(slFields) Step 1
                    slFields(ilLoop) = Trim$(slFields(ilLoop))
                Next ilLoop
                If Not ilStartLayoutFd Then
                    'If StrComp(slFields(1), "Record Layout: Start", vbTextCompare) = 0 Then
                    If StrComp(slFields(0), "Record Layout: Start", vbTextCompare) = 0 Then
                        ilStartLayoutFd = True
                    End If
                ElseIf Not ilEndLayoutFd Then
                    'If StrComp(slFields(1), "Record Layout: End", vbTextCompare) = 0 Then
                    If StrComp(slFields(0), "Record Layout: End", vbTextCompare) = 0 Then
                        ilEndLayoutFd = True
                    End If
                Else
                    'If StrComp(slFields(1), "Spot: End", vbTextCompare) = 0 Then
                    If StrComp(slFields(0), "Spot: End", vbTextCompare) = 0 Then
                        ilEndSpotsFd = True
                    Else
                        'Date:,Generation Date,Generation Time
                        'If StrComp(slFields(1), "Date:", vbTextCompare) = 0 Then
                        If StrComp(slFields(0), "Date:", vbTextCompare) = 0 Then
                            ilDateFd = True
                        End If
                        'Vehicle:,Vehicle Name,Vehicle Station Code
                        'If StrComp(slFields(1), "Vehicle:", vbTextCompare) = 0 Then
                        If StrComp(slFields(0), "Vehicle:", vbTextCompare) = 0 Then
                            ilVehicleFd = True
                            'slVehicleName = slFields(2)
                            slVehicleName = slFields(1)
                            'slVehicleStationCode = slFields(3)
                            slVehicleStationCode = slFields(2)
                            ilVefCode = -1
                            'If slVehicleStationCode <> "" Then
                            '    For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                            '        If StrComp(Trim$(tgVehicleInfo(ilVef).sCodeStn), slVehicleStationCode, vbTextCompare) = 0 Then
                            '            ilVefCode = tgVehicleInfo(ilVef).iCode
                            '            Exit For
                            '        End If
                            '    Next ilVef
                            'Else
                            '    For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                            '        If StrComp(Trim$(tgVehicleInfo(ilVef).sVehicle), slVehicleName, vbTextCompare) = 0 Then
                            '            ilVefCode = tgVehicleInfo(ilVef).iCode
                            '            Exit For
                            '        End If
                            '    Next ilVef
                            'End If
                            For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                                If StrComp(Trim$(tgVehicleInfo(ilVef).sVehicle), slVehicleName, vbTextCompare) = 0 Then
                                    ilVefCode = tgVehicleInfo(ilVef).iCode
                                    Exit For
                                End If
                            Next ilVef
                        End If
                        '5/11/12: Allow Game or Event (Old v5.7 used Game and later version used Event)
                        'If StrComp(slFields(1), "Event:", vbTextCompare) = 0 Then
                        ''If StrComp(slFields(0), "Event:", vbTextCompare) = 0 Then
                        'Event:,Season,GameNo,Date,Visiting Team Name,Visiting Team Abbreviation,Home Team Name,Home Team Abbreviation,Start Time,Status,Language,English,Feed Source
                        'If (StrComp(slFields(1), "Event:", vbTextCompare) = 0) Or (StrComp(slFields(1), "Game:", vbTextCompare) = 0) Then
                        If (StrComp(slFields(0), "Event:", vbTextCompare) = 0) Or (StrComp(slFields(0), "Game:", vbTextCompare) = 0) Then
                            ilFirstSpot = True
                            'If StrComp(slFields(3), "0", vbTextCompare) = 0 Then
                            If StrComp(slFields(2), "0", vbTextCompare) = 0 Then
                                ilGameFd = False
                                slSeason = ""
                                ilGameNo = 0
                                'slLogDate = slFields(4)
                                slLogDate = slFields(3)
                                slGameStartTime = ""
                                slGameStatus = ""
                            Else
                                ilGameFd = True
                                'slLogDate = slFields(4)
                                slLogDate = slFields(3)
                                'ilGameNo = Val(slFields(3))
                                ilGameNo = Val(slFields(2))
                                'slLogDate = slFields(4)
                                slLogDate = slFields(3)
                                'slGameStartTime = slFields(9)
                                slGameStartTime = slFields(8)
                                'slGameStatus = slFields(10)
                                slGameStatus = slFields(9)
                            End If
                            'Add Teams if required
                            gGetEventTitles ilVefCode, smEventTitle1, smEventTitle2
                            ilVisitMnfCode = -1
                            For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
                                'If StrComp(Trim$(tgTeamInfo(ilTeam).sName), slFields(5), vbTextCompare) = 0 Then
                                If StrComp(Trim$(tgTeamInfo(ilTeam).sName), slFields(4), vbTextCompare) = 0 Then
                                    ilVisitMnfCode = tgTeamInfo(ilTeam).iCode
                                    'If StrComp(Trim$(tgTeamInfo(ilTeam).sShortForm), slFields(6), vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(tgTeamInfo(ilTeam).sShortForm), slFields(5), vbTextCompare) <> 0 Then
                                        'ilRet = mUpdateTeam(ilVisitMnfCode, slFields(6))
                                        ilRet = mUpdateTeam(ilVisitMnfCode, slFields(5))
                                    End If
                                    Exit For
                                End If
                            Next ilTeam
                            If ilVisitMnfCode = -1 Then
                                'Add Team
                                'ilVisitMnfCode = mAddTeam(slFields(5), slFields(6))
                                ilVisitMnfCode = mAddTeam(slFields(4), slFields(5))
                                If ilVisitMnfCode = -1 Then
                                    ''lbcMsg.AddItem "Unable to add Visiting Team " & slFields(4) & ", Import Stopped"
                                    'lbcMsg.AddItem "Unable to add Visiting Team " & slFields(3) & ", Import Stopped"
                                    ''gLogMsg "Unable to add Visiting Team " & slFields(4), "CSIImportLog.Txt", False
                                    'gLogMsg "Unable to add Visiting Team " & slFields(4), "CSIImportLog.Txt", False
                                    'lbcMsg.AddItem "Unable to add " & smEventTitle1 & slFields(4) & ", Import Stopped"
                                    lbcMsg.AddItem "Unable to add " & smEventTitle1 & slFields(3) & ", Import Stopped"
                                    'gLogMsg "Unable to add " & smEventTitle1 & slFields(5), "CSIImportLog.Txt", False
                                    gLogMsg "Unable to add " & smEventTitle1 & slFields(4), "CSIImportLog.Txt", False
                                    mImportSpots = False
                                    Close hmFrom
                                    Exit Function
                                End If
                            End If
                            ilHomeMnfCode = -1
                            For ilTeam = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
                                'If StrComp(Trim$(tgTeamInfo(ilTeam).sName), slFields(7), vbTextCompare) = 0 Then
                                If StrComp(Trim$(tgTeamInfo(ilTeam).sName), slFields(6), vbTextCompare) = 0 Then
                                    ilHomeMnfCode = tgTeamInfo(ilTeam).iCode
                                    'If StrComp(Trim$(tgTeamInfo(ilTeam).sShortForm), slFields(8), vbTextCompare) <> 0 Then
                                    If StrComp(Trim$(tgTeamInfo(ilTeam).sShortForm), slFields(7), vbTextCompare) <> 0 Then
                                        'ilRet = mUpdateTeam(ilHomeMnfCode, slFields(8))
                                        ilRet = mUpdateTeam(ilHomeMnfCode, slFields(7))
                                    End If
                                    Exit For
                                End If
                            Next ilTeam
                            If ilHomeMnfCode = -1 Then
                                'Add Team
                                'ilHomeMnfCode = mAddTeam(slFields(7), slFields(8))
                                ilHomeMnfCode = mAddTeam(slFields(6), slFields(7))
                                If ilHomeMnfCode = -1 Then
                                    ''lbcMsg.AddItem "Unable to add Home Team " & slFields(6) & ", Import Stopped"
                                    'lbcMsg.AddItem "Unable to add Home Team " & slFields(5) & ", Import Stopped"
                                    ''gLogMsg "Unable to add Home Team " & slFields(6), "CSIImportLog.Txt", False
                                    'gLogMsg "Unable to add Home Team " & slFields(5), "CSIImportLog.Txt", False
                                    'lbcMsg.AddItem "Unable to add " & smEventTitle2 & slFields(7) & ", Import Stopped"
                                    lbcMsg.AddItem "Unable to add " & smEventTitle2 & slFields(6) & ", Import Stopped"
                                    'gLogMsg "Unable to add " & smEventTitle2 & slFields(7), "CSIImportLog.Txt", False
                                    gLogMsg "Unable to add " & smEventTitle2 & slFields(6), "CSIImportLog.Txt", False
                                    mImportSpots = False
                                    Close hmFrom
                                    Exit Function
                                End If
                            End If
                            'If slFields(11) <> "" Then
                            If slFields(10) <> "" Then
                                ilLangMnfCode = -1
                                For ilLang = LBound(tgLangInfo) To UBound(tgLangInfo) - 1 Step 1
                                    'If StrComp(Trim$(tgLangInfo(ilLang).sName), slFields(11), vbTextCompare) = 0 Then
                                    If StrComp(Trim$(tgLangInfo(ilLang).sName), slFields(10), vbTextCompare) = 0 Then
                                        ilLangMnfCode = tgLangInfo(ilLang).iCode
                                        'If StrComp(Trim$(tgLangInfo(ilLang).sEnglish), slFields(12), vbTextCompare) <> 0 Then
                                        If StrComp(Trim$(tgLangInfo(ilLang).sEnglish), slFields(11), vbTextCompare) <> 0 Then
                                            'ilRet = mUpdateLang(ilLangMnfCode, slFields(12))
                                            ilRet = mUpdateLang(ilLangMnfCode, slFields(11))
                                        End If
                                        Exit For
                                    End If
                                Next ilLang
                                If ilLangMnfCode = -1 Then
                                    'Add Lang
                                    'ilLangMnfCode = mAddLang(slFields(11), slFields(12))
                                    ilLangMnfCode = mAddLang(slFields(10), slFields(11))
                                    If ilLangMnfCode = -1 Then
                                        'lbcMsg.AddItem "Unable to add Language " & slFields(11) & ", Import Stopped"
                                        lbcMsg.AddItem "Unable to add Language " & slFields(10) & ", Import Stopped"
                                        'gLogMsg "Unable to add Language " & slFields(11), "CSIImportLog.Txt", False
                                        gLogMsg "Unable to add Language " & slFields(10), "CSIImportLog.Txt", False
                                        mImportSpots = False
                                        Close hmFrom
                                        Exit Function
                                    End If
                                End If
                            Else
                                ilLangMnfCode = 0
                            End If
                            'slFeedSource = slFields(13)
                            slFeedSource = slFields(12)
                        End If
                        If ilDateFd And ilVehicleFd And ilGameFd Then
                            'Spot:,Time,Position #,Break #,Advertiser Name,Advertiser Abbreviation,Product,Spot Length,Cart,ISCI,Creative Title,Avail Name,Zone
                            'If StrComp(slFields(1), "Spot:", vbTextCompare) = 0 Then
                            If StrComp(slFields(0), "Spot:", vbTextCompare) = 0 Then
                                llLstCode = -1
                                'Create spot
                                If ilFirstSpot Then
                                    If ilVefCode = -1 Then
                                        'Add vehicle
                                    '    ilVefCode = mAddVehicle(slVehicleName, slVehicleStationCode, ilGameFd)
                                    '    If ilVefCode = -1 Then
                                            lbcMsg.AddItem "Vehicle " & slVehicleName & " Not Found, Import Stopped"
                                            gLogMsg "Vehicle " & slVehicleName & " Not Found, Import Stopped", "CSIImportLog.Txt", False
                                            mImportSpots = False
                                            Close hmFrom
                                            Exit Function
                                    '    End If
                                    Else
                                        llVpf = gBinarySearchVpf(CLng(ilVefCode))
                                        If llVpf = -1 Then
                                            lbcMsg.AddItem "Vehicle Option Record for " & slVehicleName & " Not Found, Import Stopped"
                                            gLogMsg "Vehicle Option Record for " & slVehicleName & " Not Found, Import Stopped", "CSIImportLog.Txt", False
                                            mImportSpots = False
                                            Close hmFrom
                                            Exit Function
                                        Else
                                            If (Asc(tgVpfOptions(llVpf).sUsingFeatures1) And IMPORTAFFILIATESPOTS) <> IMPORTAFFILIATESPOTS Then
                                                lbcMsg.AddItem "Vehicle " & slVehicleName & " Not Defined as Affiliate Import Allowed, Import Stopped"
                                                gLogMsg "Vehicle " & slVehicleName & " Not Defined as Affiliate Import Allowed, Import Stopped", "CSIImportLog.Txt", False
                                                mImportSpots = False
                                                Close hmFrom
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                    If ilGameFd Then
                                        SQLQuery = "SELECT * FROM GHF_Game_Header WHERE ghfVefCode = " & ilVefCode
                                        SQLQuery = SQLQuery & " AND ghfSeasonName = '" & slSeason & "'"
                                        Set rst = gSQLSelectCall(SQLQuery)
                                        If rst.EOF Then
                                            llGhfCode = mAddGameHeader(ilVefCode, slSeason, ilGameNo)
                                            If llGhfCode = -1 Then
                                                lbcMsg.AddItem "Unable to add Event Header to " & slVehicleName & ", Import Stopped"
                                                gLogMsg "Unable to add Event Header to " & slVehicleName, "CSIImportLog.Txt", False
                                                mImportSpots = False
                                                Close hmFrom
                                                Exit Function
                                            End If
                                        Else
                                            llGhfCode = rst!ghfCode
                                            If ilGameNo > rst!ghfNoGames Then
                                                ilRet = mUpdateGameHeader(llGhfCode, ilGameNo)
                                                If Not ilRet Then
                                                    lbcMsg.AddItem "Unable to update Event Header for " & slVehicleName & ", Import Stopped"
                                                    gLogMsg "Unable to update Event Header for " & slVehicleName, "CSIImportLog.Txt", False
                                                    mImportSpots = False
                                                    Close hmFrom
                                                    Exit Function
                                                End If
                                            End If
                                        End If
                                
                                        SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfGhfCode = " & llGhfCode & " AND gsfVefCode = " & ilVefCode & " AND gsfGameNo = " & ilGameNo & ")"
                                        Set rst = gSQLSelectCall(SQLQuery)
                                        If rst.EOF Then
                                            llGsfCode = mAddGame(llGhfCode, ilVefCode, ilGameNo, ilVisitMnfCode, ilHomeMnfCode, slLogDate, slGameStartTime, slGameStatus, ilLangMnfCode, slFeedSource)
                                            If llGsfCode = -1 Then
                                                lbcMsg.AddItem "Unable to add Event # " & ilGameNo & " to " & slVehicleName & ", Import Stopped"
                                                gLogMsg "Unable to add Event " & ilGameNo & " to " & slVehicleName, "CSIImportLog.Txt", False
                                                mImportSpots = False
                                                Close hmFrom
                                                Exit Function
                                            End If
                                        Else
                                            llGsfCode = rst!gsfCode
                                            ilRet = mUpdateGame(llGsfCode, ilVisitMnfCode, ilHomeMnfCode, slLogDate, slGameStartTime, slGameStatus)
                                            If Not ilRet Then
                                                lbcMsg.AddItem "Unable to update Event # " & ilGameNo & " for " & slVehicleName & ", Import Stopped"
                                                gLogMsg "Unable to update Event " & ilGameNo & " for " & slVehicleName, "CSIImportLog.Txt", False
                                                mImportSpots = False
                                                Close hmFrom
                                                Exit Function
                                            End If
                                        End If
                                    Else
                                        llGsfCode = 0
                                    End If
                                    mClearPrevImport ilVefCode, slLogDate, llGsfCode
                                    ilFirstSpot = False
                                End If
                                
                                ilAdfCode = -1
                                For ilAdf = LBound(tgAdvtInfo) To UBound(tgAdvtInfo) - 1 Step 1
                                    'If StrComp(Trim$(tgAdvtInfo(ilAdf).sAdvtName), slFields(5), vbTextCompare) = 0 Then
                                    If StrComp(Trim$(tgAdvtInfo(ilAdf).sAdvtName), slFields(4), vbTextCompare) = 0 Then
                                        ilAdfCode = tgAdvtInfo(ilAdf).iCode
                                        Exit For
                                    End If
                                Next ilAdf
                                If ilAdfCode = -1 Then
                                    'Add Advertiser
                                    'ilAdfCode = mAddAdvt(slFields(5), slFields(6))
                                    ilAdfCode = mAddAdvt(slFields(4), slFields(5))
                                    If ilAdfCode = -1 Then
                                        'lbcMsg.AddItem "Unable to add Advertiser " & slFields(5) & ", Spot Not Added"
                                        lbcMsg.AddItem "Unable to add Advertiser " & slFields(4) & ", Spot Not Added"
                                        'gLogMsg "Unable to add Advertiser " & slFields(5), "CSIImportLog.Txt", False
                                        gLogMsg "Unable to add Advertiser " & slFields(4), "CSIImportLog.Txt", False
                                    End If
                                End If
                                'If slFields(12) <> "" Then
                                If slFields(11) <> "" Then
                                    ilAnfcode = -1
                                    For ilAnf = LBound(tgAvailNamesInfo) To UBound(tgAvailNamesInfo) - 1 Step 1
                                        'If StrComp(Trim$(tgAvailNamesInfo(ilAnf).sName), slFields(12), vbTextCompare) = 0 Then
                                        If StrComp(Trim$(tgAvailNamesInfo(ilAnf).sName), slFields(11), vbTextCompare) = 0 Then
                                            ilAnfcode = tgAvailNamesInfo(ilAnf).iCode
                                            Exit For
                                        End If
                                    Next ilAnf
                                    If ilAnfcode = -1 Then
                                        'Add Team
                                        'ilAnfcode = mAddAvailName(slFields(12))
                                        ilAnfcode = mAddAvailName(slFields(11))
                                        If ilAnfcode = -1 Then
                                            'lbcMsg.AddItem "Unable to add Avail Name " & slFields(12) & ", Spot Not Added"
                                            lbcMsg.AddItem "Unable to add Avail Name " & slFields(11) & ", Spot Not Added"
                                            'gLogMsg "Unable to add Avail Name " & slFields(12), "CSIImportLog.Txt", False
                                            gLogMsg "Unable to add Avail Name " & slFields(11), "CSIImportLog.Txt", False
                                        End If
                                    End If
                                Else
                                    ilAnfcode = 0
                                End If
                                If (ilAdfCode <> -1) And (ilAnfcode <> -1) Then
                                    llLstCode = mAddSpot(ilAdfCode, ilVefCode, ilAnfcode, llGsfCode, slLogDate, slFields())
                                    ilRet = mUpdateLLD(ilVefCode, slLogDate)
                                    '4/12/15: Update cpttAstStatus
                                    ilRet = mUpdateCptt(ilVefCode, slLogDate)
                                End If
                            'ElseIf StrComp(slFields(1), "Copy:", vbTextCompare) = 0 Then
                            ElseIf StrComp(slFields(0), "Copy:", vbTextCompare) = 0 Then
                                'Copy:,Type,Advertiser Name,Advertiser Abbreviation,Cart,ISCI,Product,Creative Title,Call Letters,Station ID, XDS Cue
                                If llLstCode > 0 Then
                                    ilAdfCode = -1
                                    For ilAdf = LBound(tgAdvtInfo) To UBound(tgAdvtInfo) - 1 Step 1
                                        'If StrComp(Trim$(tgAdvtInfo(ilAdf).sAdvtName), slFields(3), vbTextCompare) = 0 Then
                                        If StrComp(Trim$(tgAdvtInfo(ilAdf).sAdvtName), slFields(2), vbTextCompare) = 0 Then
                                            ilAdfCode = tgAdvtInfo(ilAdf).iCode
                                            Exit For
                                        End If
                                    Next ilAdf
                                    If ilAdfCode = -1 Then
                                        'Add Advertiser
                                        'ilAdfCode = mAddAdvt(slFields(3), slFields(4))
                                        ilAdfCode = mAddAdvt(slFields(2), slFields(3))
                                        If ilAdfCode = -1 Then
                                            'lbcMsg.AddItem "Unable to add Advertiser " & slFields(3) & ", Copy Not Added"
                                            lbcMsg.AddItem "Unable to add Advertiser " & slFields(2) & ", Copy Not Added"
                                            'gLogMsg "Unable to add Advertiser " & slFields(3), "CSIImportLog.Txt", False
                                            gLogMsg "Unable to add Advertiser " & slFields(2), "CSIImportLog.Txt", False
                                        End If
                                    End If
                                    'If (slFields(1) <> "B") Or ((slFields(1) = "B") And ilAdfCode <> -1) Then
                                    If (slFields(0) <> "B") Or ((slFields(0) = "B") And ilAdfCode <> -1) Then
                                        ilShttCode = -1
                                        'llStationID = Val(slFields(10))
                                        llStationID = Val(slFields(9))
                                        For ilShtt = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
                                            If tgStationInfo(ilShtt).lPermStationID = llStationID Then
                                                ilShttCode = tgStationInfo(ilShtt).iCode
                                            End If
                                        Next ilShtt
                                        If ilShttCode <= 0 Then
                                            'lbcMsg.AddItem "Unable to find Station ID " & slFields(10) & " for " & slFields(9) & ", Copy Not Added"
                                            lbcMsg.AddItem "Unable to find Station ID " & slFields(9) & " for " & slFields(8) & ", Copy Not Added"
                                            'gLogMsg "Unable to find Station ID " & slFields(10) & " for " & slFields(9), "CSIImportLog.Txt", False
                                            gLogMsg "Unable to find Station ID " & slFields(9) & " for " & slFields(8), "CSIImportLog.Txt", False
                                        Else
                                            'ilRet = mAddCopy(llLstCode, slFields(2), ilAdfCode, slFields(5), slFields(6), slFields(7), slFields(8), ilShttCode, slFields(11))
                                            ilRet = mAddCopy(llLstCode, slFields(1), ilAdfCode, slFields(4), slFields(5), slFields(6), slFields(7), ilShttCode, slFields(10))
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        ilRet = 0
        If ilEndSpotsFd Then
            Exit Do
        End If
    Loop
    If Not ilFirstSpot Then
        ilRet = mCheckCptt(ilVefCode, slLogDate)
    End If
    Close hmFrom
    On Error GoTo 0
    mImportSpots = True
    Exit Function
mImportSpotsErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    gHandleError "CSIImportLog.Txt", "ImportCSISpot-mImportSpots"
    mImportSpots = False
    Exit Function
ErrHand1:
    gHandleError "CSIImportLog.Txt", "ImportCSISpot-mImportSpots"
    mImportSpots = False

End Function


Private Function mCheckFile()
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim slMsg As String
    Dim slLine As String
    Dim ilLoop As Integer
    'Dim slFields(1 To 16) As String
    Dim slFields(0 To 15) As String
    Dim ilFile As Integer
    Dim ilFileStartLayoutFd As Integer
    Dim ilFileEndLayoutFd As Integer
    Dim ilFileEndSpotsFd As Integer
    Dim ilFileDateFd As Integer
    Dim ilFileVehicleFd As Integer
    Dim ilFileGameFd As Integer
    Dim ilFileSpotFd As Integer
    
    mCheckFile = True
    'slFromFile = txtFile.Text
    For ilFile = 0 To UBound(smFileNames) - 1 Step 1
        ilFileStartLayoutFd = False
        ilFileEndLayoutFd = False
        ilFileEndSpotsFd = False
        ilFileDateFd = False
        ilFileVehicleFd = False
        ilFileGameFd = False
        ilFileSpotFd = False
        slFromFile = smFileNames(ilFile)
        'ilRet = 0
        'On Error GoTo mImportSpotsErr:
        'hmFrom = FreeFile
        'Open slFromFile For Input Access Read As hmFrom
        ilRet = gFileOpen(slFromFile, "Input Access Read", hmFrom)
        If ilRet <> 0 Then
            Beep
            gMsgBox "Unable to open the Import file " & slFromFile & ", error: " & Trim$(Str$(ilRet)), vbCritical
            mCheckFile = False
            Close hmFrom
            Exit Function
        End If
        Do While Not EOF(hmFrom)
            ilRet = 0
            On Error GoTo mImportSpotsErr:
            Line Input #hmFrom, slLine
            On Error GoTo 0
            If ilRet = 62 Then
                ilRet = 0
                Exit Do
            End If
            On Error GoTo ErrHand
            slLine = Trim$(slLine)
            If Len(slLine) > 0 Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    Exit Do
                Else
                    'Process Input
                    gParseCDFields slLine, False, slFields()
                    For ilLoop = LBound(slFields) To UBound(slFields) Step 1
                        slFields(ilLoop) = Trim$(slFields(ilLoop))
                    Next ilLoop
                    If Not ilFileStartLayoutFd Then
                        'If StrComp(slFields(1), "Record Layout: Start", vbTextCompare) = 0 Then
                        If StrComp(slFields(0), "Record Layout: Start", vbTextCompare) = 0 Then
                            ilFileStartLayoutFd = True
                        End If
                    ElseIf Not ilFileEndLayoutFd Then
                        'If StrComp(slFields(1), "Record Layout: End", vbTextCompare) = 0 Then
                        If StrComp(slFields(0), "Record Layout: End", vbTextCompare) = 0 Then
                            ilFileEndLayoutFd = True
                        End If
                    Else
                        'If StrComp(slFields(1), "Spot: End", vbTextCompare) = 0 Then
                        If StrComp(slFields(0), "Spot: End", vbTextCompare) = 0 Then
                            ilFileEndSpotsFd = True
                        Else
                            'If StrComp(slFields(1), "Date:", vbTextCompare) = 0 Then
                            If StrComp(slFields(0), "Date:", vbTextCompare) = 0 Then
                                ilFileDateFd = True
                            End If
                            'If StrComp(slFields(1), "Vehicle:", vbTextCompare) = 0 Then
                            If StrComp(slFields(0), "Vehicle:", vbTextCompare) = 0 Then
                                ilFileVehicleFd = True
                            End If
                            '5/11/12: Allow Game or Event (Old v5.7 used Game and later version used Event)
                            ''If StrComp(slFields(1), "Event:", vbTextCompare) = 0 Then
                            'If StrComp(slFields(0), "Event:", vbTextCompare) = 0 Then
                            If (StrComp(slFields(0), "Event:", vbTextCompare) = 0) Or (StrComp(slFields(0), "Game:", vbTextCompare) = 0) Then
                                ilFileGameFd = True
                            End If
                            If ilFileDateFd And ilFileVehicleFd And ilFileGameFd Then
                                'If StrComp(slFields(1), "Spot:", vbTextCompare) = 0 Then
                                If StrComp(slFields(0), "Spot:", vbTextCompare) = 0 Then
                                    ilFileSpotFd = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            ilRet = 0
        Loop
        Close hmFrom
        If Not ilFileStartLayoutFd Then
            lbcMsg.AddItem "No Start Layout Record found in the Import file " & slFromFile
            gLogMsg "No Start Layout Record found in the Import file " & slFromFile, "CSIImportLog.Txt", False
            mCheckFile = False
        End If
        If Not ilFileEndLayoutFd Then
            lbcMsg.AddItem "No End Layout Record found in the Import file " & slFromFile
            gLogMsg "No End Layout Record found in the Import file " & slFromFile, "CSIImportLog.Txt", False
            mCheckFile = False
        End If
        If Not ilFileEndSpotsFd Then
            lbcMsg.AddItem "No End Spots Record found in the Import file " & slFromFile
            gLogMsg "No End Spots Record found in the Import file " & slFromFile, "CSIImportLog.Txt", False
            mCheckFile = False
        End If
        If Not ilFileDateFd Then
            lbcMsg.AddItem "No Date Record found in the Import file " & slFromFile
            gLogMsg "No Date Record found in the Import file " & slFromFile, "CSIImportLog.Txt", False
            mCheckFile = False
        End If
        If Not ilFileVehicleFd Then
            lbcMsg.AddItem "No Vehicle Record found in the Import file " & slFromFile
            gLogMsg "No Vehicle Record found in the Import file " & slFromFile, "CSIImportLog.Txt", False
            mCheckFile = False
        End If
        If Not ilFileGameFd Then
            lbcMsg.AddItem "No Event Record found in the Import file " & slFromFile
            gLogMsg "No Event Record found in the Import file " & slFromFile, "CSIImportLog.Txt", False
            mCheckFile = False
        End If
    Next ilFile
    On Error GoTo 0
    Exit Function
mImportSpotsErr:
    ilRet = Err.Number
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "ImportCSISpots-mCheckFile"
    mCheckFile = False
End Function

Private Function mAddVehicle(slVehicleName As String, slVehicleStationCode As String, ilGame As Integer) As Integer
    Dim tlVef As VEF
    Dim ilVefCode As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer

    On Error GoTo ErrHand
    
    tlVef.iCode = 0
    tlVef.sName = slVehicleName
    tlVef.sAddr1 = ""
    tlVef.sAddr2 = ""
    tlVef.sAddr3 = ""
    tlVef.sPhone = ""
    tlVef.sFax = ""
    tlVef.sUnused1 = ""
    tlVef.sDialPos = ""
    tlVef.lPvfCode = 0
    tlVef.iReallDnfCode = 0
    tlVef.sUpdateRvf1 = ""
    tlVef.sUpdateRvf2 = ""
    tlVef.sUpdateRvf3 = ""
    tlVef.sUpdateRvf4 = ""
    tlVef.sUpdateRvf5 = ""
    tlVef.sUpdateRvf6 = ""
    tlVef.sUpdateRvf7 = ""
    tlVef.sUpdateRvf8 = ""
    tlVef.iCombineVefCode = 0
    tlVef.iMnfHubCode = 0
    tlVef.iTrfCode = 0
    tlVef.sType = "I"
    tlVef.sCodeStn = slVehicleStationCode
    tlVef.iVefCode = 0
    tlVef.iUnused2 = 0
    tlVef.iProdPct = 0
    tlVef.iProdPct2 = 0
    tlVef.iProdPct3 = 0
    tlVef.iProdPct4 = 0
    tlVef.iProdPct5 = 0
    tlVef.iProdPct6 = 0
    tlVef.iProdPct7 = 0
    tlVef.iProdPct8 = 0
    tlVef.sState = "A"
    tlVef.imnfGroup = 0
    tlVef.imnfGroup2 = 0
    tlVef.imnfGroup3 = 0
    tlVef.imnfGroup4 = 0
    tlVef.imnfGroup5 = 0
    tlVef.imnfGroup6 = 0
    tlVef.imnfGroup7 = 0
    tlVef.imnfGroup8 = 0
    tlVef.iSort = 0
    tlVef.idnfCode = 0
    tlVef.imnfDemo = 0
    tlVef.imnfSSCode1 = 0
    tlVef.imnfSSCode2 = 0
    tlVef.imnfSSCode3 = 0
    tlVef.imnfSSCode4 = 0
    tlVef.imnfSSCode5 = 0
    tlVef.imnfSSCode6 = 0
    tlVef.imnfSSCode7 = 0
    tlVef.imnfSSCode8 = 0
    tlVef.sExportRAB = "N"
    tlVef.lVsfCode = 0
    tlVef.lRateAud = 0
    tlVef.lCPPCPM = 0
    tlVef.lYearAvails = 0
    tlVef.iPctSellout = 0
    tlVef.iMnfVehGp2 = 0
    tlVef.iMnfVehGp3Mkt = 0
    tlVef.iMnfVehGp4Fmt = 0
    tlVef.iMnfVehGp5Rsch = 0
    tlVef.iMnfVehGp6Sub = 0
    tlVef.iMnfVehGp7 = 0
    tlVef.iSSMnfCode = 0
    tlVef.sStdPrice = ""
    tlVef.sStdInvTime = ""
    tlVef.sStdAlter = ""
    tlVef.iStdIndex = 0
    tlVef.sStdAlterName = ""
    tlVef.iRemoteID = 0
    tlVef.iAutoCode = 0
    tlVef.sExtUpdateRvf1 = ""
    tlVef.sExtUpdateRvf2 = ""
    tlVef.sExtUpdateRvf3 = ""
    tlVef.sExtUpdateRvf4 = ""
    tlVef.sExtUpdateRvf5 = ""
    tlVef.sExtUpdateRvf6 = ""
    tlVef.sExtUpdateRvf7 = ""
    tlVef.sExtUpdateRvf8 = ""
    tlVef.sStdSelCriteria = ""
    tlVef.sStdOverrideFlag = ""
    tlVef.sContact = ""
    
    
    SQLQuery = "Insert Into VEF_Vehicles ( "
    SQLQuery = SQLQuery & "vefCode, "
    SQLQuery = SQLQuery & "vefName, "
    SQLQuery = SQLQuery & "vefAddr1, "
    SQLQuery = SQLQuery & "vefAddr2, "
    SQLQuery = SQLQuery & "vefAddr3, "
    SQLQuery = SQLQuery & "vefPhone, "
    SQLQuery = SQLQuery & "vefFax, "
    SQLQuery = SQLQuery & "vefUnused1, "
    SQLQuery = SQLQuery & "vefDialPos, "
    SQLQuery = SQLQuery & "vefPvfCode, "
    SQLQuery = SQLQuery & "vefReallDnfCode, "
    SQLQuery = SQLQuery & "vefUpdateRvf1, "
    SQLQuery = SQLQuery & "vefUpdateRvf2, "
    SQLQuery = SQLQuery & "vefUpdateRvf3, "
    SQLQuery = SQLQuery & "vefUpdateRvf4, "
    SQLQuery = SQLQuery & "vefUpdateRvf5, "
    SQLQuery = SQLQuery & "vefUpdateRvf6, "
    SQLQuery = SQLQuery & "vefUpdateRvf7, "
    SQLQuery = SQLQuery & "vefUpdateRvf8, "
    SQLQuery = SQLQuery & "vefCombineVefCode, "
    SQLQuery = SQLQuery & "vefMnfHubCode, "
    SQLQuery = SQLQuery & "vefTrfCode, "
    SQLQuery = SQLQuery & "vefType, "
    SQLQuery = SQLQuery & "vefCodeStn, "
    SQLQuery = SQLQuery & "vefvefCode, "
    SQLQuery = SQLQuery & "vefUnused2, "
    SQLQuery = SQLQuery & "vefProdPct, "
    SQLQuery = SQLQuery & "vefProdPct2, "
    SQLQuery = SQLQuery & "vefProdPct3, "
    SQLQuery = SQLQuery & "vefProdPct4, "
    SQLQuery = SQLQuery & "vefProdPct5, "
    SQLQuery = SQLQuery & "vefProdPct6, "
    SQLQuery = SQLQuery & "vefProdPct7, "
    SQLQuery = SQLQuery & "vefProdPct8, "
    SQLQuery = SQLQuery & "vefState, "
    SQLQuery = SQLQuery & "vefmnfGroup, "
    SQLQuery = SQLQuery & "vefmnfGroup2, "
    SQLQuery = SQLQuery & "vefmnfGroup3, "
    SQLQuery = SQLQuery & "vefmnfGroup4, "
    SQLQuery = SQLQuery & "vefmnfGroup5, "
    SQLQuery = SQLQuery & "vefmnfGroup6, "
    SQLQuery = SQLQuery & "vefmnfGroup7, "
    SQLQuery = SQLQuery & "vefmnfGroup8, "
    SQLQuery = SQLQuery & "vefSort, "
    SQLQuery = SQLQuery & "vefdnfCode, "
    SQLQuery = SQLQuery & "vefmnfDemo, "
    SQLQuery = SQLQuery & "vefmnfSSCode1, "
    SQLQuery = SQLQuery & "vefmnfSSCode2, "
    SQLQuery = SQLQuery & "vefmnfSSCode3, "
    SQLQuery = SQLQuery & "vefmnfSSCode4, "
    SQLQuery = SQLQuery & "vefmnfSSCode5, "
    SQLQuery = SQLQuery & "vefmnfSSCode6, "
    SQLQuery = SQLQuery & "vefmnfSSCode7, "
    SQLQuery = SQLQuery & "vefmnfSSCode8, "
    SQLQuery = SQLQuery & "vefExportRAB, "
    SQLQuery = SQLQuery & "vefVsfCode, "
    SQLQuery = SQLQuery & "vefRateAud, "
    SQLQuery = SQLQuery & "vefCPPCPM, "
    SQLQuery = SQLQuery & "vefYearAvails, "
    SQLQuery = SQLQuery & "vefPctSellout, "
    SQLQuery = SQLQuery & "vefMnfVehGp2, "
    SQLQuery = SQLQuery & "vefMnfVehGp3Mkt, "
    SQLQuery = SQLQuery & "vefMnfVehGp4Fmt, "
    SQLQuery = SQLQuery & "vefMnfVehGp5Rsch, "
    SQLQuery = SQLQuery & "vefMnfVehGp6Sub, "
    SQLQuery = SQLQuery & "vefMnfVehGp7, "
    SQLQuery = SQLQuery & "vefSSMnfCode, "
    SQLQuery = SQLQuery & "vefStdPrice, "
    SQLQuery = SQLQuery & "vefStdInvTime, "
    SQLQuery = SQLQuery & "vefStdAlter, "
    SQLQuery = SQLQuery & "vefStdIndex, "
    SQLQuery = SQLQuery & "vefUnused, "
    SQLQuery = SQLQuery & "vefRemoteID, "
    SQLQuery = SQLQuery & "vefAutoCode, "
    SQLQuery = SQLQuery & "vefExtUpdateRvf1, "
    SQLQuery = SQLQuery & "vefExtUpdateRvf2, "
    SQLQuery = SQLQuery & "vefExtUpdateRvf3, "
    SQLQuery = SQLQuery & "vefExtUpdateRvf4, "
    SQLQuery = SQLQuery & "vefExtUpdateRvf5, "
    SQLQuery = SQLQuery & "vefExtUpdateRvf6, "
    SQLQuery = SQLQuery & "vefExtUpdateRvf7, "
    SQLQuery = SQLQuery & "vefExtUpdateRvf8, "
    SQLQuery = SQLQuery & "vefStdSelCriteria, "
    SQLQuery = SQLQuery & "vefStdOverrideFlag, "
    SQLQuery = SQLQuery & "vefContact "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & tlVef.iCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sName) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sAddr1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sAddr2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sAddr3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sPhone) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sFax) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sUnused1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sDialPos) & "', "
    SQLQuery = SQLQuery & tlVef.lPvfCode & ", "
    SQLQuery = SQLQuery & tlVef.iReallDnfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sUpdateRvf1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sUpdateRvf2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sUpdateRvf3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sUpdateRvf4) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sUpdateRvf5) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sUpdateRvf6) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sUpdateRvf7) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sUpdateRvf8) & "', "
    SQLQuery = SQLQuery & tlVef.iCombineVefCode & ", "
    SQLQuery = SQLQuery & tlVef.iMnfHubCode & ", "
    SQLQuery = SQLQuery & tlVef.iTrfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sType) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sCodeStn) & "', "
    SQLQuery = SQLQuery & tlVef.iVefCode & ", "
    SQLQuery = SQLQuery & tlVef.iUnused2 & ", "
    SQLQuery = SQLQuery & tlVef.iProdPct & ", "
    SQLQuery = SQLQuery & tlVef.iProdPct2 & ", "
    SQLQuery = SQLQuery & tlVef.iProdPct3 & ", "
    SQLQuery = SQLQuery & tlVef.iProdPct4 & ", "
    SQLQuery = SQLQuery & tlVef.iProdPct5 & ", "
    SQLQuery = SQLQuery & tlVef.iProdPct6 & ", "
    SQLQuery = SQLQuery & tlVef.iProdPct7 & ", "
    SQLQuery = SQLQuery & tlVef.iProdPct8 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sState) & "', "
    SQLQuery = SQLQuery & tlVef.imnfGroup & ", "
    SQLQuery = SQLQuery & tlVef.imnfGroup2 & ", "
    SQLQuery = SQLQuery & tlVef.imnfGroup3 & ", "
    SQLQuery = SQLQuery & tlVef.imnfGroup4 & ", "
    SQLQuery = SQLQuery & tlVef.imnfGroup5 & ", "
    SQLQuery = SQLQuery & tlVef.imnfGroup6 & ", "
    SQLQuery = SQLQuery & tlVef.imnfGroup7 & ", "
    SQLQuery = SQLQuery & tlVef.imnfGroup8 & ", "
    SQLQuery = SQLQuery & tlVef.iSort & ", "
    SQLQuery = SQLQuery & tlVef.idnfCode & ", "
    SQLQuery = SQLQuery & tlVef.imnfDemo & ", "
    SQLQuery = SQLQuery & tlVef.imnfSSCode1 & ", "
    SQLQuery = SQLQuery & tlVef.imnfSSCode2 & ", "
    SQLQuery = SQLQuery & tlVef.imnfSSCode3 & ", "
    SQLQuery = SQLQuery & tlVef.imnfSSCode4 & ", "
    SQLQuery = SQLQuery & tlVef.imnfSSCode5 & ", "
    SQLQuery = SQLQuery & tlVef.imnfSSCode6 & ", "
    SQLQuery = SQLQuery & tlVef.imnfSSCode7 & ", "
    SQLQuery = SQLQuery & tlVef.imnfSSCode8 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sExportRAB) & "', "
    SQLQuery = SQLQuery & tlVef.lVsfCode & ", "
    SQLQuery = SQLQuery & tlVef.lRateAud & ", "
    SQLQuery = SQLQuery & tlVef.lCPPCPM & ", "
    SQLQuery = SQLQuery & tlVef.lYearAvails & ", "
    SQLQuery = SQLQuery & tlVef.iPctSellout & ", "
    SQLQuery = SQLQuery & tlVef.iMnfVehGp2 & ", "
    SQLQuery = SQLQuery & tlVef.iMnfVehGp3Mkt & ", "
    SQLQuery = SQLQuery & tlVef.iMnfVehGp4Fmt & ", "
    SQLQuery = SQLQuery & tlVef.iMnfVehGp5Rsch & ", "
    SQLQuery = SQLQuery & tlVef.iMnfVehGp6Sub & ", "
    SQLQuery = SQLQuery & tlVef.iMnfVehGp7 & ", "
    SQLQuery = SQLQuery & tlVef.iSSMnfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sStdPrice) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sStdInvTime) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sStdAlter) & "', "
    SQLQuery = SQLQuery & tlVef.iStdIndex & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sStdAlterName) & "', "
    SQLQuery = SQLQuery & tlVef.iRemoteID & ", "
    SQLQuery = SQLQuery & tlVef.iAutoCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf4) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf5) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf6) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf7) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sExtUpdateRvf8) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sStdSelCriteria) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sStdOverrideFlag) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVef.sContact) & "' "
    SQLQuery = SQLQuery & ")"

    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddVehicle"
        mAddVehicle = -1
        Exit Function
    End If
    
    'Reload array and find vehicle, then update
    '11/26/17
    gFileChgdUpdate "vef.btr", True
    ilRet = gPopVehicles()
    ilVefCode = -1
    If slVehicleStationCode <> "" Then
        For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
            If StrComp(Trim$(tgVehicleInfo(ilVef).sCodeStn), slVehicleStationCode, vbTextCompare) = 0 Then
                ilVefCode = tgVehicleInfo(ilVef).iCode
                Exit For
            End If
        Next ilVef
    Else
        For ilVef = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
            If StrComp(Trim$(tgVehicleInfo(ilVef).sVehicle), slVehicleName, vbTextCompare) = 0 Then
                ilVefCode = tgVehicleInfo(ilVef).iCode
                Exit For
            End If
        Next ilVef
    End If
    If ilVefCode = -1 Then
        mAddVehicle = -1
        Exit Function
    End If
    SQLQuery = "UPDATE VEF_Vehicles SET "
    SQLQuery = SQLQuery & "vefAutoCode = " & ilVefCode
    SQLQuery = SQLQuery & " WHERE vefCode = " & ilVefCode
    
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddVehicle"
        mAddVehicle = -1
        Exit Function
    End If
    
    ilRet = mAddVPF(ilVefCode)
    '11/26/17
    gFileChgdUpdate "vpf.btr", True
    ilRet = gPopVehicleOptions()
    
    mAddVehicle = ilVefCode
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddVehicle"
    mAddVehicle = -1
    Exit Function
End Function

Private Function mAddTeam(slName As String, slAbbreviation As String) As Integer
    Dim tlMnf As MNF
    Dim ilMnfCode As Integer
    Dim ilRet As Integer
    Dim ilMnf As Integer

    On Error GoTo ErrHand
    
    tlMnf.iCode = 0
    tlMnf.sType = "Z"
    tlMnf.sName = slName
    tlMnf.sRPU = "" 'Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(15)
    tlMnf.sUnitType = slAbbreviation
    tlMnf.sSSComm = ""  'Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    tlMnf.iMerge = 0
    tlMnf.iGroupNo = 0
    tlMnf.sCodeStn = ""
    tlMnf.iRemoteID = 0
    tlMnf.iAutoCode = 0
    tlMnf.sSyncDate = Format$(gNow(), sgShowDateForm)
    tlMnf.sSyncTime = Format$(gNow(), sgShowTimeWSecForm)
    tlMnf.sUnitsPer = ""
    tlMnf.lCost = 0
    tlMnf.sUnused = ""
    
    SQLQuery = "Insert Into MNF_Multi_Names ( "
    SQLQuery = SQLQuery & "mnfCode, "
    SQLQuery = SQLQuery & "mnfType, "
    SQLQuery = SQLQuery & "mnfName, "
    SQLQuery = SQLQuery & "mnfRPU, "
    SQLQuery = SQLQuery & "mnfUnitType, "
    SQLQuery = SQLQuery & "mnfSSComm, "
    SQLQuery = SQLQuery & "mnfMerge, "
    SQLQuery = SQLQuery & "mnfGroupNo, "
    SQLQuery = SQLQuery & "mnfCodeStn, "
    SQLQuery = SQLQuery & "mnfRemoteID, "
    SQLQuery = SQLQuery & "mnfAutoCode, "
    SQLQuery = SQLQuery & "mnfSyncDate, "
    SQLQuery = SQLQuery & "mnfSyncTime, "
    SQLQuery = SQLQuery & "mnfUnitsPer, "
    SQLQuery = SQLQuery & "mnfCost, "
    SQLQuery = SQLQuery & "mnfUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & tlMnf.iCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlMnf.sType) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlMnf.sName) & "', "
    SQLQuery = SQLQuery & "'" & tlMnf.sRPU & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlMnf.sUnitType) & "', "
    SQLQuery = SQLQuery & "'" & tlMnf.sSSComm & "', "
    SQLQuery = SQLQuery & tlMnf.iMerge & ", "
    SQLQuery = SQLQuery & tlMnf.iGroupNo & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlMnf.sCodeStn) & "', "
    SQLQuery = SQLQuery & tlMnf.iRemoteID & ", "
    SQLQuery = SQLQuery & tlMnf.iAutoCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlMnf.sSyncDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlMnf.sSyncTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlMnf.sUnitsPer) & "', "
    SQLQuery = SQLQuery & tlMnf.lCost & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlMnf.sUnused) & "' "
    SQLQuery = SQLQuery & ") "
    
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddTeam"
        mAddTeam = -1
        Exit Function
    End If
    
    'Reload array and find vehicle, then update
    ilRet = gPopTeams()
    ilMnfCode = -1
    For ilMnf = LBound(tgTeamInfo) To UBound(tgTeamInfo) - 1 Step 1
        If StrComp(Trim$(tgTeamInfo(ilMnf).sName), slName, vbTextCompare) = 0 Then
            ilMnfCode = tgTeamInfo(ilMnf).iCode
            Exit For
        End If
    Next ilMnf
    If ilMnfCode = -1 Then
        mAddTeam = -1
        Exit Function
    End If
    SQLQuery = "UPDATE MNF_Multi_Names SET "
    SQLQuery = SQLQuery & "mnfAutoCode = " & ilMnfCode
    SQLQuery = SQLQuery & " WHERE mnfCode = " & ilMnfCode
    
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddTeam"
        mAddTeam = -1
        Exit Function
    End If
    
    mAddTeam = ilMnfCode
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddTeam"
    mAddTeam = -1
    Exit Function
End Function
Private Function mAddLang(slName As String, slEnglish As String) As Integer
    Dim tlMnf As MNF
    Dim ilMnfCode As Integer
    Dim ilRet As Integer
    Dim ilMnf As Integer

    On Error GoTo ErrHand
    
    tlMnf.iCode = 0
    tlMnf.sType = "Z"
    tlMnf.sName = slName
    tlMnf.sRPU = "" 'Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(15)
    tlMnf.sUnitType = slEnglish
    tlMnf.sSSComm = ""  'Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    tlMnf.iMerge = 0
    tlMnf.iGroupNo = 0
    tlMnf.sCodeStn = ""
    tlMnf.iRemoteID = 0
    tlMnf.iAutoCode = 0
    tlMnf.sSyncDate = Format$(gNow(), sgShowDateForm)
    tlMnf.sSyncTime = Format$(gNow(), sgShowTimeWSecForm)
    tlMnf.sUnitsPer = ""
    tlMnf.lCost = 0
    tlMnf.sUnused = ""
    
    SQLQuery = "Insert Into MNF_Multi_Names ( "
    SQLQuery = SQLQuery & "mnfCode, "
    SQLQuery = SQLQuery & "mnfType, "
    SQLQuery = SQLQuery & "mnfName, "
    SQLQuery = SQLQuery & "mnfRPU, "
    SQLQuery = SQLQuery & "mnfUnitType, "
    SQLQuery = SQLQuery & "mnfSSComm, "
    SQLQuery = SQLQuery & "mnfMerge, "
    SQLQuery = SQLQuery & "mnfGroupNo, "
    SQLQuery = SQLQuery & "mnfCodeStn, "
    SQLQuery = SQLQuery & "mnfRemoteID, "
    SQLQuery = SQLQuery & "mnfAutoCode, "
    SQLQuery = SQLQuery & "mnfSyncDate, "
    SQLQuery = SQLQuery & "mnfSyncTime, "
    SQLQuery = SQLQuery & "mnfUnitsPer, "
    SQLQuery = SQLQuery & "mnfCost, "
    SQLQuery = SQLQuery & "mnfUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & tlMnf.iCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlMnf.sType) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlMnf.sName) & "', "
    SQLQuery = SQLQuery & "'" & tlMnf.sRPU & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlMnf.sUnitType) & "', "
    SQLQuery = SQLQuery & "'" & tlMnf.sSSComm & "', "
    SQLQuery = SQLQuery & tlMnf.iMerge & ", "
    SQLQuery = SQLQuery & tlMnf.iGroupNo & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlMnf.sCodeStn) & "', "
    SQLQuery = SQLQuery & tlMnf.iRemoteID & ", "
    SQLQuery = SQLQuery & tlMnf.iAutoCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlMnf.sSyncDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlMnf.sSyncTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlMnf.sUnitsPer) & "', "
    SQLQuery = SQLQuery & tlMnf.lCost & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlMnf.sUnused) & "' "
    SQLQuery = SQLQuery & ") "
    
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddLang"
        mAddLang = -1
        Exit Function
    End If
    
    'Reload array and find vehicle, then update
    ilRet = gPopLangs()
    ilMnfCode = -1
    For ilMnf = LBound(tgLangInfo) To UBound(tgLangInfo) - 1 Step 1
        If StrComp(Trim$(tgLangInfo(ilMnf).sName), slName, vbTextCompare) = 0 Then
            ilMnfCode = tgLangInfo(ilMnf).iCode
            Exit For
        End If
    Next ilMnf
    If ilMnfCode = -1 Then
        mAddLang = -1
        Exit Function
    End If
    SQLQuery = "UPDATE MNF_Multi_Names SET "
    SQLQuery = SQLQuery & "mnfAutoCode = " & ilMnfCode
    SQLQuery = SQLQuery & " WHERE mnfCode = " & ilMnfCode
    
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddLang"
        mAddLang = -1
        Exit Function
    End If
    
    mAddLang = ilMnfCode
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "AffErorLog.txt", "frmImportCSISpot-mAddLang"
    mAddLang = -1
    Exit Function
End Function

Private Function mAddGameHeader(ilVefCode As Integer, slSeason As String, ilNoGames As Integer) As Long
    Dim tlGhf As GHF
    Dim ilRet As Integer

    On Error GoTo ErrHand
    
    tlGhf.lCode = 0
    tlGhf.iVefCode = ilVefCode
    tlGhf.sSeasonName = slSeason
    tlGhf.sSeasonStartDate = "1/1/1970"
    tlGhf.sSeasonEndDate = "12/31/2069"
    tlGhf.iNoGames = ilNoGames
    tlGhf.sUnused = ""
    
    
    SQLQuery = "Insert Into GHF_Game_Header ( "
    SQLQuery = SQLQuery & "ghfCode, "
    SQLQuery = SQLQuery & "ghfVefCode, "
    SQLQuery = SQLQuery & "ghfSeasonName, "
    SQLQuery = SQLQuery & "ghfSeasonStartDate, "
    SQLQuery = SQLQuery & "ghfSeasonEndDate, "
    SQLQuery = SQLQuery & "ghfNoGames, "
    SQLQuery = SQLQuery & "ghfUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & tlGhf.lCode & ", "
    SQLQuery = SQLQuery & tlGhf.iVefCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlGhf.sSeasonName) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlGhf.sSeasonStartDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlGhf.sSeasonEndDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlGhf.iNoGames & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlGhf.sUnused) & "' "
    SQLQuery = SQLQuery & ") "
    
    
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddGameHeader"
        mAddGameHeader = -1
        Exit Function
    End If
    
    SQLQuery = "SELECT * FROM GHF_Game_Header WHERE ghfVefCode = " & ilVefCode
    SQLQuery = SQLQuery & " AND ghfSeasonName = '" & slSeason & "'"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF Then
        mAddGameHeader = -1
    Else
        mAddGameHeader = rst!ghfCode
    End If
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mAddGameHeader"
    mAddGameHeader = -1
    Exit Function
End Function

Private Function mUpdateGameHeader(llGhfCode As Long, ilNoGames As Integer) As Integer

    On Error GoTo ErrHand
    
    SQLQuery = "UPDATE GHF_Game_Header SET "
    SQLQuery = SQLQuery & "ghfNoGames = " & ilNoGames
    SQLQuery = SQLQuery & " WHERE ghfCode = " & llGhfCode
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mUpdateGameHeader"
        mUpdateGameHeader = False
        Exit Function
    End If
    
    mUpdateGameHeader = True
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mUpdateGameHeader"
    mUpdateGameHeader = False
    Exit Function
End Function

Private Function mAddGame(llGhfCode As Long, ilVefCode As Integer, ilGameNo As Integer, ilVisitMnfCode As Integer, ilHomeMnfCode As Integer, slAirDate As String, slAirTime As String, slGameStatus As String, ilLangMnfCode As Integer, slFeedSource As String) As Long
    Dim tlGsf As GSF
    Dim ilRet As Integer

    On Error GoTo ErrHand
    
    tlGsf.lCode = 0
    tlGsf.lGhfCode = llGhfCode
    tlGsf.iVefCode = ilVefCode
    tlGsf.iGameNo = ilGameNo
    tlGsf.sFeedSource = slFeedSource
    tlGsf.iLangMnfCode = ilLangMnfCode
    tlGsf.iVisitMnfCode = ilVisitMnfCode
    tlGsf.iHomeMnfCode = ilHomeMnfCode
    tlGsf.lLvfCode = 0
    tlGsf.sAirDate = Format$(slAirDate, sgShowDateForm)
    tlGsf.sAirTime = Format$(slAirTime, sgShowTimeWSecForm)
    tlGsf.iAirVefCode = 0
    tlGsf.sGameStatus = slGameStatus
    tlGsf.sLiveLogMerge = ""
    tlGsf.sXDSProgCodeID = ""
    tlGsf.sBus = ""
    tlGsf.iSubtotal1MnfCode = 0
    tlGsf.iSubtotal2MnfCode = 0
    tlGsf.sUnused = ""
    
    
    SQLQuery = "Insert Into GSF_Game_Schd ( "
    SQLQuery = SQLQuery & "gsfCode, "
    SQLQuery = SQLQuery & "gsfGhfCode, "
    SQLQuery = SQLQuery & "gsfVefCode, "
    SQLQuery = SQLQuery & "gsfGameNo, "
    SQLQuery = SQLQuery & "gsfFeedSource, "
    SQLQuery = SQLQuery & "gsfLangMnfCode, "
    SQLQuery = SQLQuery & "gsfVisitMnfCode, "
    SQLQuery = SQLQuery & "gsfHomeMnfCode, "
    SQLQuery = SQLQuery & "gsfLvfCode, "
    SQLQuery = SQLQuery & "gsfAirDate, "
    SQLQuery = SQLQuery & "gsfAirTime, "
    SQLQuery = SQLQuery & "gsfAirVefCode, "
    SQLQuery = SQLQuery & "gsfGameStatus, "
    SQLQuery = SQLQuery & "gsfLiveLogMerge, "
    SQLQuery = SQLQuery & "gsfXDSProgCodeID, "
    SQLQuery = SQLQuery & "gsfBus, "
    SQLQuery = SQLQuery & "gsfSubtotal1MnfCode, "
    SQLQuery = SQLQuery & "gsfSubtotal2MnfCode, "
    SQLQuery = SQLQuery & "gsfUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & tlGsf.lCode & ", "
    SQLQuery = SQLQuery & tlGsf.lGhfCode & ", "
    SQLQuery = SQLQuery & tlGsf.iVefCode & ", "
    SQLQuery = SQLQuery & tlGsf.iGameNo & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlGsf.sFeedSource) & "', "
    SQLQuery = SQLQuery & tlGsf.iLangMnfCode & ", "
    SQLQuery = SQLQuery & tlGsf.iVisitMnfCode & ", "
    SQLQuery = SQLQuery & tlGsf.iHomeMnfCode & ", "
    SQLQuery = SQLQuery & tlGsf.lLvfCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlGsf.sAirDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlGsf.sAirTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & tlGsf.iAirVefCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlGsf.sGameStatus) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlGsf.sLiveLogMerge) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlGsf.sXDSProgCodeID) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlGsf.sBus) & "', "
    SQLQuery = SQLQuery & tlGsf.iSubtotal1MnfCode & ", "
    SQLQuery = SQLQuery & tlGsf.iSubtotal2MnfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlGsf.sUnused) & "' "
    SQLQuery = SQLQuery & ") "
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddGame"
        mAddGame = -1
        Exit Function
    End If
    
    
    SQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfGhfCode = " & llGhfCode & " AND gsfVefCode = " & ilVefCode & " AND gsfGameNo = " & ilGameNo & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF Then
        mAddGame = -1
    Else
        mAddGame = rst!gsfCode
    End If
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mAddGame"
    mAddGame = -1
    Exit Function
End Function

Private Function mUpdateGame(llGsfCode As Long, ilVisitMnfCode As Integer, ilHomeMnfCode As Integer, slAirDate As String, slAirTime As String, slGameStatus As String) As Integer

    On Error GoTo ErrHand
    
    SQLQuery = "UPDATE GSF_Game_Schd SET "
    SQLQuery = SQLQuery & "gsfVisitMnfCode = " & ilVisitMnfCode & ", "
    SQLQuery = SQLQuery & "gsfHomeMnfCode = " & ilHomeMnfCode & ", "
    SQLQuery = SQLQuery & "gsfAirDate = '" & Format$(slAirDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "gsfAirTime = '" & Format$(slAirTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "gsfGameStatus = '" & gFixQuote(slGameStatus) & "'"
    SQLQuery = SQLQuery & " WHERE gsfCode = " & llGsfCode
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mUpdateGame"
        mUpdateGame = False
        Exit Function
    End If
    
    mUpdateGame = True
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mUpdateGame"
    mUpdateGame = False
    Exit Function
End Function

Private Function mAddSpot(ilAdfCode As Integer, ilVefCode As Integer, ilAnfcode As Integer, llGsfCode As Long, slGameDate As String, slFields() As String) As Long
    Dim tlLST As LST
    Dim ilRet As Integer
    Dim llCode As Long

    On Error GoTo ErrHand
    

    tlLST.lCode = 0
    tlLST.iType = 0
    tlLST.lSdfCode = 0
    tlLST.lCntrNo = 0
    tlLST.iAdfCode = ilAdfCode
    tlLST.iAgfCode = 0
    'tlLST.sProd = gFixQuote(slFields(7))
    tlLST.sProd = gFixQuote(slFields(6))
    tlLST.iLineNo = 0
    tlLST.iLnVefCode = 0
    tlLST.sStartDate = Format$("1/1/1970", sgShowDateForm)
    tlLST.sEndDate = Format$("1/1/1970", sgShowDateForm)
    tlLST.iMon = 0
    tlLST.iTue = 0
    tlLST.iWed = 0
    tlLST.iThu = 0
    tlLST.iFri = 0
    tlLST.iSat = 0
    tlLST.iSun = 0
    tlLST.iSpotsWk = 0
    tlLST.iPriceType = 1
    tlLST.lPrice = 0
    tlLST.iSpotType = 5
    tlLST.iLogVefCode = ilVefCode
    tlLST.sLogDate = Format$(slGameDate, sgShowDateForm)
    'tlLST.sLogTime = Format$(gConvertTime(slFields(2)), sgShowTimeWSecForm)
    tlLST.sLogTime = Format$(gConvertTime(slFields(1)), sgShowTimeWSecForm)
    tlLST.sDemo = ""
    tlLST.lAud = 0
    'tlLST.sISCI = gFixQuote(slFields(10))
    tlLST.sISCI = gFixQuote(slFields(9))
    tlLST.iWkNo = 0
    'tlLST.iBreakNo = slFields(4)
    tlLST.iBreakNo = slFields(3)
    'tlLST.iPositionNo = slFields(3)
    tlLST.iPositionNo = slFields(2)
    tlLST.iSeqNo = 0
    'tlLST.sZone = slFields(13)
    tlLST.sZone = slFields(12)
    'tlLST.sCart = gFixQuote(slFields(9))
    tlLST.sCart = gFixQuote(slFields(8))
    tlLST.lCpfCode = 0
    tlLST.lCrfCsfCode = 0
    tlLST.iStatus = 0
    'tlLST.iLen = Val(slFields(8))
    tlLST.iLen = Val(slFields(7))
    tlLST.iUnits = 0
    tlLST.lCifCode = 0
    tlLST.iAnfCode = ilAnfcode
    tlLST.lEvtIDCefCode = 0
    tlLST.sSplitNetwork = "N"
    tlLST.lRafCode = 0
    tlLST.lFsfCode = 0
    tlLST.lgsfCode = llGsfCode
    tlLST.sImportedSpot = "Y"
    tlLST.lBkoutLstCode = 0
    tlLST.sLnStartTime = "12am"
    tlLST.sLnEndTime = "12am"
    tlLST.sUnused = ""
    
    SQLQuery = "Insert Into lst ( "
    SQLQuery = SQLQuery & "lstCode, "
    SQLQuery = SQLQuery & "lstType, "
    SQLQuery = SQLQuery & "lstSdfCode, "
    SQLQuery = SQLQuery & "lstCntrNo, "
    SQLQuery = SQLQuery & "lstAdfCode, "
    SQLQuery = SQLQuery & "lstAgfCode, "
    SQLQuery = SQLQuery & "lstProd, "
    SQLQuery = SQLQuery & "lstLineNo, "
    SQLQuery = SQLQuery & "lstLnVefCode, "
    SQLQuery = SQLQuery & "lstStartDate, "
    SQLQuery = SQLQuery & "lstEndDate, "
    SQLQuery = SQLQuery & "lstMon, "
    SQLQuery = SQLQuery & "lstTue, "
    SQLQuery = SQLQuery & "lstWed, "
    SQLQuery = SQLQuery & "lstThu, "
    SQLQuery = SQLQuery & "lstFri, "
    SQLQuery = SQLQuery & "lstSat, "
    SQLQuery = SQLQuery & "lstSun, "
    SQLQuery = SQLQuery & "lstSpotsWk, "
    SQLQuery = SQLQuery & "lstPriceType, "
    SQLQuery = SQLQuery & "lstPrice, "
    SQLQuery = SQLQuery & "lstSpotType, "
    SQLQuery = SQLQuery & "lstLogVefCode, "
    SQLQuery = SQLQuery & "lstLogDate, "
    SQLQuery = SQLQuery & "lstLogTime, "
    SQLQuery = SQLQuery & "lstDemo, "
    SQLQuery = SQLQuery & "lstAud, "
    SQLQuery = SQLQuery & "lstISCI, "
    SQLQuery = SQLQuery & "lstWkNo, "
    SQLQuery = SQLQuery & "lstBreakNo, "
    SQLQuery = SQLQuery & "lstPositionNo, "
    SQLQuery = SQLQuery & "lstSeqNo, "
    SQLQuery = SQLQuery & "lstZone, "
    SQLQuery = SQLQuery & "lstCart, "
    SQLQuery = SQLQuery & "lstCpfCode, "
    SQLQuery = SQLQuery & "lstCrfCsfCode, "
    SQLQuery = SQLQuery & "lstStatus, "
    SQLQuery = SQLQuery & "lstLen, "
    SQLQuery = SQLQuery & "lstUnits, "
    SQLQuery = SQLQuery & "lstCifCode, "
    SQLQuery = SQLQuery & "lstAnfCode, "
    SQLQuery = SQLQuery & "lstEvtIDCefCode, "
    SQLQuery = SQLQuery & "lstSplitNetwork, "
    SQLQuery = SQLQuery & "lstRafCode, "
    SQLQuery = SQLQuery & "lstFsfCode, "
    SQLQuery = SQLQuery & "lstGsfCode, "
    SQLQuery = SQLQuery & "lstImportedSpot, "
    SQLQuery = SQLQuery & "lstBkoutLstCode, "
    SQLQuery = SQLQuery & "lstLnStartTime, "
    SQLQuery = SQLQuery & "lstLnEndTime, "
    SQLQuery = SQLQuery & "lstUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & "Replace" & ", "
    SQLQuery = SQLQuery & tlLST.iType & ", "
    SQLQuery = SQLQuery & tlLST.lSdfCode & ", "
    SQLQuery = SQLQuery & tlLST.lCntrNo & ", "
    SQLQuery = SQLQuery & tlLST.iAdfCode & ", "
    SQLQuery = SQLQuery & tlLST.iAgfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sProd) & "', "
    SQLQuery = SQLQuery & tlLST.iLineNo & ", "
    SQLQuery = SQLQuery & tlLST.iLnVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlLST.sStartDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlLST.sEndDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlLST.iMon & ", "
    SQLQuery = SQLQuery & tlLST.iTue & ", "
    SQLQuery = SQLQuery & tlLST.iWed & ", "
    SQLQuery = SQLQuery & tlLST.iThu & ", "
    SQLQuery = SQLQuery & tlLST.iFri & ", "
    SQLQuery = SQLQuery & tlLST.iSat & ", "
    SQLQuery = SQLQuery & tlLST.iSun & ", "
    SQLQuery = SQLQuery & tlLST.iSpotsWk & ", "
    SQLQuery = SQLQuery & tlLST.iPriceType & ", "
    SQLQuery = SQLQuery & tlLST.lPrice & ", "
    SQLQuery = SQLQuery & tlLST.iSpotType & ", "
    SQLQuery = SQLQuery & tlLST.iLogVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlLST.sLogDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlLST.sLogTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sDemo) & "', "
    SQLQuery = SQLQuery & tlLST.lAud & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sISCI) & "', "
    SQLQuery = SQLQuery & tlLST.iWkNo & ", "
    SQLQuery = SQLQuery & tlLST.iBreakNo & ", "
    SQLQuery = SQLQuery & tlLST.iPositionNo & ", "
    SQLQuery = SQLQuery & tlLST.iSeqNo & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sZone) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sCart) & "', "
    SQLQuery = SQLQuery & tlLST.lCpfCode & ", "
    SQLQuery = SQLQuery & tlLST.lCrfCsfCode & ", "
    SQLQuery = SQLQuery & tlLST.iStatus & ", "
    SQLQuery = SQLQuery & tlLST.iLen & ", "
    SQLQuery = SQLQuery & tlLST.iUnits & ", "
    SQLQuery = SQLQuery & tlLST.lCifCode & ", "
    SQLQuery = SQLQuery & tlLST.iAnfCode & ", "
    SQLQuery = SQLQuery & tlLST.lEvtIDCefCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sSplitNetwork) & "', "
    SQLQuery = SQLQuery & tlLST.lRafCode & ", "
    SQLQuery = SQLQuery & tlLST.lFsfCode & ", "
    SQLQuery = SQLQuery & tlLST.lgsfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sImportedSpot) & "', "
    SQLQuery = SQLQuery & tlLST.lBkoutLstCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlLST.sLnStartTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlLST.sLnEndTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlLST.sUnused) & "' "
    SQLQuery = SQLQuery & ")"

    'cnn.Execute SQLQuery, rdExecDirect
    llCode = gInsertAndReturnCode(SQLQuery, "lst", "lstCode", "Replace")
    If llCode <= 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddSpot"
        mAddSpot = -1
        Exit Function
    End If

    On Error GoTo 0
    mAddSpot = llCode
    Exit Function
ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mAddSpot"
    mAddSpot = -1
    Exit Function
End Function

Private Function mAddAdvt(slAdvtName As String, slAdvtAbbr As String) As Integer
    Dim tlAdf As ADF
    Dim ilRet As Integer
    Dim ilAdf As Integer
    Dim ilAdfCode As Integer

    On Error GoTo ErrHand
    
    tlAdf.iCode = 0
    tlAdf.sName = slAdvtName
    tlAdf.sAbbr = slAdvtAbbr
    tlAdf.sProd = ""
    tlAdf.iSlfCode = 0
    tlAdf.iAgfCode = 0
    tlAdf.sBuyer = ""
    tlAdf.sCodeRep = ""
    tlAdf.sCodeAgy = ""
    tlAdf.sCodeStn = ""
    tlAdf.imnfComp2 = 0
    tlAdf.imnfExcl1 = 0
    tlAdf.imnfExcl2 = 0
    tlAdf.sCppCpm = ""
    tlAdf.sDemo1 = ""
    tlAdf.sDemo2 = ""
    tlAdf.sDemo3 = ""
    tlAdf.sDemo4 = ""
    tlAdf.imnfDemo1 = 0
    tlAdf.imnfDemo2 = 0
    tlAdf.imnfDemo3 = 0
    tlAdf.imnfDemo4 = 0
    tlAdf.lTarget1 = 0
    tlAdf.lTarget2 = 0
    tlAdf.lTarget3 = 0
    tlAdf.lTarget4 = 0
    tlAdf.sCreditRestr = "N"
    tlAdf.lCreditLimit = 0
    tlAdf.sPaymRating = "1"
    tlAdf.sISCI = ""
    tlAdf.imnfSort = 0
    tlAdf.sBilAgyDir = ""
    tlAdf.sCntrAddr1 = ""
    tlAdf.sCntrAddr2 = ""
    tlAdf.sCntrAddr3 = ""
    tlAdf.sBillAddr1 = ""
    tlAdf.sBillAddr2 = ""
    tlAdf.sBillAddr3 = ""
    tlAdf.iarfLkCode = 0
    tlAdf.sPhone = ""
    tlAdf.sFax = ""
    tlAdf.iarfContrCode = 0
    tlAdf.iarfInvCode = 0
    tlAdf.sCntrPrtSz = ""
    tlAdf.iTrfCode = 0
    tlAdf.sCrdApp = "A"
    tlAdf.sCrdRtg = ""
    tlAdf.ipnfBuyer = 0
    tlAdf.ipnfPay = 0
    tlAdf.iPct90 = 0
    tlAdf.sCurrAR = ""  'Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    tlAdf.sUnbilled = ""    'Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    tlAdf.sHiCredit = ""    'Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    tlAdf.sTotalGross = ""  'Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
    tlAdf.sDateEntrd = Format$(gNow(), sgShowDateForm)
    tlAdf.iNSFChks = 0
    tlAdf.sDateLstInv = Format$("1/1/1970", sgShowDateForm)
    tlAdf.sDateLstPaym = Format$("1/1/1970", sgShowDateForm)
    tlAdf.iAvgToPay = 0
    tlAdf.iLstToPay = 0
    tlAdf.iNoInvPd = 0
    tlAdf.sNewBus = ""
    tlAdf.sEndDate = Format$("1/1/1970", sgShowDateForm)
    tlAdf.iMerge = 0
    tlAdf.iurfCode = 0
    tlAdf.sState = "A"
    tlAdf.sCrdAppDate = Format$("1/1/1970", sgShowDateForm)
    tlAdf.sCrdAppTime = Format$("1/1/1970", sgShowTimeWSecForm)
    tlAdf.sPkInvShow = ""
    tlAdf.lGuar = 0
    tlAdf.sBkoutPoolStatus = "N"
    tlAdf.sUnused2 = ""
    tlAdf.sRateOnInv = ""
    tlAdf.iMnfBus = 0
    tlAdf.iUnused1 = 0
    tlAdf.sAllowRepMG = ""
    tlAdf.sBonusOnInv = ""
    tlAdf.sRepInvGen = ""
    tlAdf.iMnfInvTerms = 0
    tlAdf.sPolitical = "N"
    tlAdf.sAddrID = ""
    
    
    SQLQuery = "Insert Into ADF_Advertisers ( "
    SQLQuery = SQLQuery & "adfCode, "
    SQLQuery = SQLQuery & "adfName, "
    SQLQuery = SQLQuery & "adfAbbr, "
    SQLQuery = SQLQuery & "adfProd, "
    SQLQuery = SQLQuery & "adfslfCode, "
    SQLQuery = SQLQuery & "adfagfCode, "
    SQLQuery = SQLQuery & "adfBuyer, "
    SQLQuery = SQLQuery & "adfCodeRep, "
    SQLQuery = SQLQuery & "adfCodeAgy, "
    SQLQuery = SQLQuery & "adfCodeStn, "
    SQLQuery = SQLQuery & "adfmnfComp1, "
    SQLQuery = SQLQuery & "adfmnfComp2, "
    SQLQuery = SQLQuery & "adfmnfExcl1, "
    SQLQuery = SQLQuery & "adfmnfExcl2, "
    SQLQuery = SQLQuery & "adfCppCpm, "
    SQLQuery = SQLQuery & "adfDemo1, "
    SQLQuery = SQLQuery & "adfDemo2, "
    SQLQuery = SQLQuery & "adfDemo3, "
    SQLQuery = SQLQuery & "adfDemo4, "
    SQLQuery = SQLQuery & "adfmnfDemo1, "
    SQLQuery = SQLQuery & "adfmnfDemo2, "
    SQLQuery = SQLQuery & "adfmnfDemo3, "
    SQLQuery = SQLQuery & "adfmnfDemo4, "
    SQLQuery = SQLQuery & "adfTarget1, "
    SQLQuery = SQLQuery & "adfTarget2, "
    SQLQuery = SQLQuery & "adfTarget3, "
    SQLQuery = SQLQuery & "adfTarget4, "
    SQLQuery = SQLQuery & "adfCreditRestr, "
    SQLQuery = SQLQuery & "adfCreditLimit, "
    SQLQuery = SQLQuery & "adfPaymRating, "
    SQLQuery = SQLQuery & "adfISCI, "
    SQLQuery = SQLQuery & "adfmnfSort, "
    SQLQuery = SQLQuery & "adfBilAgyDir, "
    SQLQuery = SQLQuery & "adfCntrAddr1, "
    SQLQuery = SQLQuery & "adfCntrAddr2, "
    SQLQuery = SQLQuery & "adfCntrAddr3, "
    SQLQuery = SQLQuery & "adfBillAddr1, "
    SQLQuery = SQLQuery & "adfBillAddr2, "
    SQLQuery = SQLQuery & "adfBillAddr3, "
    SQLQuery = SQLQuery & "adfarfLkCode, "
    SQLQuery = SQLQuery & "adfPhone, "
    SQLQuery = SQLQuery & "adfFax, "
    SQLQuery = SQLQuery & "adfarfContrCode, "
    SQLQuery = SQLQuery & "adfarfInvCode, "
    SQLQuery = SQLQuery & "adfCntrPrtSz, "
    SQLQuery = SQLQuery & "adfTrfCode, "
    SQLQuery = SQLQuery & "adfCrdApp, "
    SQLQuery = SQLQuery & "adfCrdRtg, "
    SQLQuery = SQLQuery & "adfpnfBuyer, "
    SQLQuery = SQLQuery & "adfpnfPay, "
    SQLQuery = SQLQuery & "adfPct90, "
    SQLQuery = SQLQuery & "adfCurrAR, "
    SQLQuery = SQLQuery & "adfUnbilled, "
    SQLQuery = SQLQuery & "adfHiCredit, "
    SQLQuery = SQLQuery & "adfTotalGross, "
    SQLQuery = SQLQuery & "adfDateEntrd, "
    SQLQuery = SQLQuery & "adfNSFChks, "
    SQLQuery = SQLQuery & "adfDateLstInv, "
    SQLQuery = SQLQuery & "adfDateLstPaym, "
    SQLQuery = SQLQuery & "adfAvgToPay, "
    SQLQuery = SQLQuery & "adfLstToPay, "
    SQLQuery = SQLQuery & "adfNoInvPd, "
    SQLQuery = SQLQuery & "adfNewBus, "
    SQLQuery = SQLQuery & "adfEndDate, "
    SQLQuery = SQLQuery & "adfMerge, "
    SQLQuery = SQLQuery & "adfurfCode, "
    SQLQuery = SQLQuery & "adfState, "
    SQLQuery = SQLQuery & "adfCrdAppDate, "
    SQLQuery = SQLQuery & "adfCrdAppTime, "
    SQLQuery = SQLQuery & "adfPkInvShow, "
    SQLQuery = SQLQuery & "adfGuar, "
    SQLQuery = SQLQuery & "adfBkoutPoolStatus"
    SQLQuery = SQLQuery & "adfUnused2, "
    SQLQuery = SQLQuery & "adfRateOnInv, "
    SQLQuery = SQLQuery & "adfMnfBus, "
    SQLQuery = SQLQuery & "adfUnused1, "
    SQLQuery = SQLQuery & "adfAllowRepMG, "
    SQLQuery = SQLQuery & "adfBonusOnInv, "
    SQLQuery = SQLQuery & "adfRepInvGen, "
    SQLQuery = SQLQuery & "adfMnfInvTerms, "
    SQLQuery = SQLQuery & "adfPolitical, "
    SQLQuery = SQLQuery & "adfAddrID "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & tlAdf.iCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sName) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sAbbr) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sProd) & "', "
    SQLQuery = SQLQuery & tlAdf.iSlfCode & ", "
    SQLQuery = SQLQuery & tlAdf.iAgfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBuyer) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCodeRep) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCodeAgy) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCodeStn) & "', "
    SQLQuery = SQLQuery & tlAdf.imnfComp1 & ", "
    SQLQuery = SQLQuery & tlAdf.imnfComp2 & ", "
    SQLQuery = SQLQuery & tlAdf.imnfExcl1 & ", "
    SQLQuery = SQLQuery & tlAdf.imnfExcl2 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCppCpm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sDemo1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sDemo2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sDemo3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sDemo4) & "', "
    SQLQuery = SQLQuery & tlAdf.imnfDemo1 & ", "
    SQLQuery = SQLQuery & tlAdf.imnfDemo2 & ", "
    SQLQuery = SQLQuery & tlAdf.imnfDemo3 & ", "
    SQLQuery = SQLQuery & tlAdf.imnfDemo4 & ", "
    SQLQuery = SQLQuery & tlAdf.lTarget1 & ", "
    SQLQuery = SQLQuery & tlAdf.lTarget2 & ", "
    SQLQuery = SQLQuery & tlAdf.lTarget3 & ", "
    SQLQuery = SQLQuery & tlAdf.lTarget4 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCreditRestr) & "', "
    SQLQuery = SQLQuery & tlAdf.lCreditLimit & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sPaymRating) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sISCI) & "', "
    SQLQuery = SQLQuery & tlAdf.imnfSort & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBilAgyDir) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCntrAddr1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCntrAddr2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCntrAddr3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBillAddr1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBillAddr2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBillAddr3) & "', "
    SQLQuery = SQLQuery & tlAdf.iarfLkCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sPhone) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sFax) & "', "
    SQLQuery = SQLQuery & tlAdf.iarfContrCode & ", "
    SQLQuery = SQLQuery & tlAdf.iarfInvCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCntrPrtSz) & "', "
    SQLQuery = SQLQuery & tlAdf.iTrfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCrdApp) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sCrdRtg) & "', "
    SQLQuery = SQLQuery & tlAdf.ipnfBuyer & ", "
    SQLQuery = SQLQuery & tlAdf.ipnfPay & ", "
    SQLQuery = SQLQuery & tlAdf.iPct90 & ", "
    SQLQuery = SQLQuery & "'" & tlAdf.sCurrAR & "', "
    SQLQuery = SQLQuery & "'" & tlAdf.sUnbilled & "', "
    SQLQuery = SQLQuery & "'" & tlAdf.sHiCredit & "', "
    SQLQuery = SQLQuery & "'" & tlAdf.sTotalGross & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlAdf.sDateEntrd, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlAdf.iNSFChks & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlAdf.sDateLstInv, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlAdf.sDateLstPaym, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlAdf.iAvgToPay & ", "
    SQLQuery = SQLQuery & tlAdf.iLstToPay & ", "
    SQLQuery = SQLQuery & tlAdf.iNoInvPd & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sNewBus) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlAdf.sEndDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlAdf.iMerge & ", "
    SQLQuery = SQLQuery & tlAdf.iurfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sState) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlAdf.sCrdAppDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlAdf.sCrdAppTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sPkInvShow) & "', "
    SQLQuery = SQLQuery & tlAdf.lGuar & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBkoutPoolStatus) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sUnused2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sRateOnInv) & "', "
    SQLQuery = SQLQuery & tlAdf.iMnfBus & ", "
    SQLQuery = SQLQuery & tlAdf.iUnused1 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sAllowRepMG) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sBonusOnInv) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sRepInvGen) & "', "
    SQLQuery = SQLQuery & tlAdf.iMnfInvTerms & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sPolitical) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAdf.sAddrID) & "' "
    SQLQuery = SQLQuery & ")"
    
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddAdvt"
        mAddAdvt = -1
        Exit Function
    End If
    
    ilRet = gPopAdvertisers()
    
    ilAdfCode = -1
    For ilAdf = LBound(tgAdvtInfo) To UBound(tgAdvtInfo) - 1 Step 1
        If StrComp(Trim$(tgAdvtInfo(ilAdf).sAdvtName), slAdvtName, vbTextCompare) = 0 Then
            ilAdfCode = tgAdvtInfo(ilAdf).iCode
            Exit For
        End If
    Next ilAdf
    If ilAdfCode = -1 Then
        mAddAdvt = -1
        Exit Function
    End If
    
    
    mAddAdvt = ilAdfCode
    
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mAddAdvt"
    mAddAdvt = -1
    Exit Function
End Function


Private Sub mClearPrevImport(ilVefCode As Integer, slClearDate As String, llGsfCode As Long)
    Dim ilCheck As Integer
    Dim ilDateClearedPrev As Integer
    Dim ilGameClearedPrev As Integer
    Dim llClearDate As Long
    Dim llLst As Long
    Dim slLstDate As String
    Dim slMonDate As String
    Dim slSunDate As String
    ReDim llLsfCode(0 To 0) As Long
    
    On Error GoTo ErrHand
    
    ilDateClearedPrev = False
    ilGameClearedPrev = False
    llClearDate = DateValue(slClearDate)
    slMonDate = gObtainPrevMonday(slClearDate)
    slSunDate = gObtainNextSunday(slMonDate)
    For ilCheck = 0 To UBound(tmClearImportInfo) - 1 Step 1
        If (tmClearImportInfo(ilCheck).iVefCode = ilVefCode) Then
            If tmClearImportInfo(ilCheck).lClearDate = llClearDate Then
                ilDateClearedPrev = True
            End If
            If tmClearImportInfo(ilCheck).lgsfCode = llGsfCode Then
                ilGameClearedPrev = True
            End If
            If ilDateClearedPrev And ilGameClearedPrev Then
                Exit Sub
            End If
        End If
    Next ilCheck
    
    If Not ilDateClearedPrev Then
        SQLQuery = "SELECT * FROM att WHERE (attVefCode = " & ilVefCode
        SQLQuery = SQLQuery & " AND " & "(attOnAir <= '" & Format$(gAdjYear(slClearDate), sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " AND " & "(attOffAir >= '" & Format$(gAdjYear(slClearDate), sgSQLDateForm) & "') AND (attDropDate >= '" & Format$(gAdjYear(slClearDate), sgSQLDateForm) & "')" & ")"
        Set att_rst = gSQLSelectCall(SQLQuery)
        Do While Not att_rst.EOF
            
            SQLQuery = "DELETE FROM Ast WHERE (astAtfCode = " & att_rst!attCode
            SQLQuery = SQLQuery & " AND astFeedDate = '" & Format$(slClearDate, sgSQLDateForm) & "')"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand1:
                gHandleError "CSIImportLog.Txt", "ImportCSISpot-mClearPrevImport"
                Exit Sub
            End If
            
            SQLQuery = "DELETE FROM Aet WHERE (aetAtfCode = " & att_rst!attCode
            SQLQuery = SQLQuery & " AND aetFeedDate = '" & Format$(slClearDate, sgSQLDateForm) & "')"
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand1:
                gHandleError "CSIImportLog.Txt", "ImportCSISpot-mClearPrevImport"
                Exit Sub
            End If
            
            'Doug-Remove spots from web for this att and date
            
            att_rst.MoveNext
        Loop
        
        SQLQuery = "DELETE FROM lst WHERE (lstLogVefCode = " & ilVefCode & " AND (lstLogDate = '" & Format$(slClearDate, sgSQLDateForm) & "')" & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand1:
            gHandleError "CSIImportLog.Txt", "ImportCSISpot-mClearPrevImport"
            Exit Sub
        End If
    End If
    
    If Not ilGameClearedPrev Then
        'Handle case where game moved to a different day
        If llGsfCode > 0 Then
            SQLQuery = "SELECT * FROM lst WHERE (lstLogVefCode = " & ilVefCode & " AND lstGsfCode = " & llGsfCode & ")"
            Set lst_rst = gSQLSelectCall(SQLQuery)
            If Not lst_rst.EOF Then
                Do While Not lst_rst.EOF
                    If UBound(llLsfCode) = 0 Then
                        slLstDate = Format$(lst_rst!lstLogDate, sgShowDateForm)
                    End If
                    llLsfCode(UBound(llLsfCode)) = lst_rst!lstCode
                    ReDim Preserve llLsfCode(0 To UBound(llLsfCode) + 1) As Long
                    lst_rst.MoveNext
                Loop
                SQLQuery = "SELECT * FROM att WHERE (attVefCode = " & ilVefCode
                SQLQuery = SQLQuery & " AND " & "(attOnAir <= '" & Format$(gAdjYear(slClearDate), sgSQLDateForm) & "')"
                SQLQuery = SQLQuery & " AND " & "(attOffAir >= '" & Format$(gAdjYear(slClearDate), sgSQLDateForm) & "') AND (attDropDate >= '" & Format$(gAdjYear(slClearDate), sgSQLDateForm) & "')" & ")"
                Set att_rst = gSQLSelectCall(SQLQuery)
                Do While Not att_rst.EOF
                    For llLst = 0 To UBound(llLsfCode) - 1 Step 1
                        SQLQuery = "DELETE FROM Ast WHERE (astAtfCode = " & att_rst!attCode
                        SQLQuery = SQLQuery & " AND (astFeedDate = '" & Format$(slLstDate, sgSQLDateForm) & "')"
                        SQLQuery = SQLQuery & " AND astlsfCode = " & llLsfCode(llLst) & ")"
                        'cnn.Execute SQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand1:
                            gHandleError "CSIImportLog.Txt", "ImportCSISpot-mClearPrevImport"
                            Exit Sub
                        End If
                        
                        SQLQuery = "DELETE FROM Aet WHERE (aetAtfCode = " & att_rst!attCode
                        SQLQuery = SQLQuery & " AND (aetFeedDate = '" & Format$(slLstDate, sgSQLDateForm) & "')"
                        SQLQuery = SQLQuery & " AND aetVefCode = " & ilVefCode & ")"
                        'cnn.Execute SQLQuery            ', rdExecDirect
                        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                            '6/10/16: Replaced GoSub
                            'GoSub ErrHand1:
                            gHandleError "CSIImportLog.Txt", "ImportCSISpot-mClearPrevImport"
                            Exit Sub
                        End If
                        'Doug- Remove spots from web for this att and date and game
                    Next llLst
                    att_rst.MoveNext
                Loop
                SQLQuery = "DELETE FROM lst WHERE (lstLogVefCode = " & ilVefCode & " AND lstGsfCode = " & llGsfCode & ")"
                'cnn.Execute SQLQuery, rdExecDirect
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand1:
                    gHandleError "CSIImportLog.Txt", "ImportCSISpot-mClearPrevImport"
                    Exit Sub
                End If
            End If
        End If
    End If
    tmClearImportInfo(UBound(tmClearImportInfo)).iVefCode = ilVefCode
    tmClearImportInfo(UBound(tmClearImportInfo)).lClearDate = llClearDate
    tmClearImportInfo(UBound(tmClearImportInfo)).lgsfCode = llGsfCode
    ReDim Preserve tmClearImportInfo(0 To UBound(tmClearImportInfo) + 1) As CLEASRIMPORTINFO
    'Check if CPTT should be removed.  If no lst exist, then remove CPTT
    'CPTT will be recreated if required in mCheckCPTT
    SQLQuery = "SELECT * FROM lst WHERE (lstLogVefCode = " & ilVefCode & " AND (lstLogDate >= '" & Format$(slMonDate, sgSQLDateForm) & "')" & " AND (lstLogDate <= '" & Format$(slSunDate, sgSQLDateForm) & "')" & ")"
    Set lst_rst = gSQLSelectCall(SQLQuery)
    If lst_rst.EOF Then
        SQLQuery = "DELETE FROM CPTT WHERE (cpttVefCode = " & ilVefCode & " AND (cpttStartDate = '" & Format$(slMonDate, sgSQLDateForm) & "')" & ")"
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand1:
            gHandleError "CSIImportLog.Txt", "ImportCSISpot-mClearPrevImport"
            Exit Sub
        End If
    End If
    On Error GoTo 0
    Exit Sub
ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mClearPrevImport"
    Exit Sub
End Sub

Private Function mCheckCptt(ilVefCode As Integer, slClearDate As String) As Integer
    Dim slMonDate As String
    Dim slSunDate As String
    Dim ilRet As Integer
    Dim llAttCode As Long
    Dim ilShttCode As Integer
    Dim llCpttCode As Long
    Dim slServiceAgreement As String
    
    On Error GoTo ErrHand
    
    slMonDate = gObtainPrevMonday(slClearDate)
    slSunDate = gObtainNextSunday(slMonDate)
    'Determine if any LST exist.  If so, then create cptt for each agreement
    SQLQuery = "SELECT * FROM lst WHERE (lstLogVefCode = " & ilVefCode
    SQLQuery = SQLQuery & " AND " & "(lstLogDate >= '" & Format$(gAdjYear(slMonDate), sgSQLDateForm) & "')"
    SQLQuery = SQLQuery & " AND " & "(lstLogDate <= '" & Format$(gAdjYear(slSunDate), sgSQLDateForm) & "')" & ")"
    Set lst_rst = gSQLSelectCall(SQLQuery)
    If Not lst_rst.EOF Then
        SQLQuery = "SELECT * FROM att WHERE (attVefCode = " & ilVefCode
        SQLQuery = SQLQuery & " AND " & "(attOnAir <= '" & Format$(gAdjYear(slMonDate), sgSQLDateForm) & "')"
        SQLQuery = SQLQuery & " AND " & "(attOffAir >= '" & Format$(gAdjYear(slSunDate), sgSQLDateForm) & "') AND (attDropDate >= '" & Format$(gAdjYear(slSunDate), sgSQLDateForm) & "')" & ")"
        Set att_rst = gSQLSelectCall(SQLQuery)
        Do While Not att_rst.EOF
            llAttCode = att_rst!attCode
            ilShttCode = att_rst!attshfCode
            slServiceAgreement = att_rst!attServiceAgreement
            SQLQuery = "SELECT * FROM CPTT WHERE (cpttVefCode = " & ilVefCode & " AND cpttAtfCode = " & llAttCode & " AND (cpttStartDate = '" & Format$(slMonDate, sgSQLDateForm) & "')" & ")"
            Set rst = gSQLSelectCall(SQLQuery)
            If rst.EOF Then
                llCpttCode = mAddCPTT(llAttCode, ilShttCode, ilVefCode, slMonDate, slServiceAgreement)
            End If
            att_rst.MoveNext
        Loop
    End If
    mCheckCptt = True
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mCheckCptt"
    Resume Next
ErrHand1:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mCheckCptt"
    Return
End Function



Private Function mAddCPTT(llAttCode As Long, ilShttCode As Integer, ilVefCode As Integer, slMonDate As String, slServiceAgreement As String) As Long
    Dim tlCPTT As CPTT
    Dim ilRet As Integer

    On Error GoTo ErrHand
    tlCPTT.lCode = 0
    tlCPTT.lAtfCode = llAttCode
    tlCPTT.iShfCode = ilShttCode
    tlCPTT.iVefCode = ilVefCode
    tlCPTT.sCreateDate = Format$(gNow(), sgShowDateForm)
    tlCPTT.sStartDate = Format$(slMonDate, sgShowDateForm)
    'tlCPTT.iCycle = 1
    tlCPTT.sReturnDate = Format$("1/1/1970", sgShowDateForm)
    'tlCPTT.sAirTime = Format$("12:00AM", sgShowTimeWSecForm)
    If slServiceAgreement = "Y" Then
        tlCPTT.iStatus = 1
    Else
        tlCPTT.iStatus = 0
    End If
    tlCPTT.iUsfCode = igUstCode
    'tlCPTT.iPrintStatus = 0
    tlCPTT.iNoSpotsGen = 0
    tlCPTT.iNoSpotsAired = 0
    tlCPTT.iPostingStatus = 0
    tlCPTT.sAstStatus = "N"
    tlCPTT.sUnused = ""
    
    
    SQLQuery = "Insert Into cptt ( "
    SQLQuery = SQLQuery & "cpttCode, "
    SQLQuery = SQLQuery & "cpttAtfCode, "
    SQLQuery = SQLQuery & "cpttShfCode, "
    SQLQuery = SQLQuery & "cpttVefCode, "
    SQLQuery = SQLQuery & "cpttCreateDate, "
    SQLQuery = SQLQuery & "cpttStartDate, "
    SQLQuery = SQLQuery & "cpttReturnDate, "
    SQLQuery = SQLQuery & "cpttStatus, "
    SQLQuery = SQLQuery & "cpttUsfCode, "
    SQLQuery = SQLQuery & "cpttNoSpotsGen, "
    SQLQuery = SQLQuery & "cpttNoSpotsAired, "
    SQLQuery = SQLQuery & "cpttPostingStatus, "
    SQLQuery = SQLQuery & "cpttAstStatus, "
    SQLQuery = SQLQuery & "cpttUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & tlCPTT.lCode & ", "
    SQLQuery = SQLQuery & tlCPTT.lAtfCode & ", "
    SQLQuery = SQLQuery & tlCPTT.iShfCode & ", "
    SQLQuery = SQLQuery & tlCPTT.iVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlCPTT.sCreateDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlCPTT.sStartDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlCPTT.sReturnDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlCPTT.iStatus & ", "
    SQLQuery = SQLQuery & tlCPTT.iUsfCode & ", "
    SQLQuery = SQLQuery & tlCPTT.iNoSpotsGen & ", "
    SQLQuery = SQLQuery & tlCPTT.iNoSpotsAired & ", "
    SQLQuery = SQLQuery & tlCPTT.iPostingStatus & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlCPTT.sAstStatus) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlCPTT.sUnused) & "' "
    SQLQuery = SQLQuery & ")"
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddCPTT"
        mAddCPTT = -1
        Exit Function
    End If
    gFileChgdUpdate "cptt.mkd", True
    SQLQuery = "SELECT * FROM CPTT WHERE (cpttVefCode = " & ilVefCode & " AND cpttAtfCode = " & llAttCode & " AND (cpttStartDate = '" & Format$(slMonDate, sgSQLDateForm) & "')" & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF Then
        mAddCPTT = -1
    Else
        mAddCPTT = rst!cpttCode
    End If
    
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mAddCPTT"
    mAddCPTT = -1
    Exit Function
End Function

Private Function mAddAvailName(slAvailName As String) As Integer
    Dim tlAnf As ANF
    Dim ilAnfcode As Integer
    Dim ilAnf As Integer
    Dim ilRet As Integer

    On Error GoTo ErrHand
    
    tlAnf.iCode = 0
    tlAnf.sName = slAvailName
    tlAnf.sSustain = "N"
    tlAnf.sState = "A"
    tlAnf.sSponsorship = "N"
    tlAnf.iMerge = 0
    tlAnf.iRemoteID = 0
    tlAnf.iAutoCode = 0
    tlAnf.sBookLocalFeed = "F"
    tlAnf.sRptDefault = "Y"
    tlAnf.iSortCode = 0
    tlAnf.sTrafToAff = "Y"
    tlAnf.sISCIExport = "Y"
    tlAnf.sAudioExport = "Y"
    tlAnf.sAutomationExport = "Y"
    tlAnf.sUnused = ""
    
    
    SQLQuery = "Insert Into ANF_Avail_Names ( "
    SQLQuery = SQLQuery & "anfCode, "
    SQLQuery = SQLQuery & "anfName, "
    SQLQuery = SQLQuery & "anfSustain, "
    SQLQuery = SQLQuery & "anfState, "
    SQLQuery = SQLQuery & "anfSponsorship, "
    SQLQuery = SQLQuery & "anfMerge, "
    SQLQuery = SQLQuery & "anfRemoteID, "
    SQLQuery = SQLQuery & "anfAutoCode, "
    SQLQuery = SQLQuery & "anfBookLocalFeed, "
    SQLQuery = SQLQuery & "anfRptDefault, "
    SQLQuery = SQLQuery & "anfSortCode, "
    SQLQuery = SQLQuery & "anfTrafToAff, "
    SQLQuery = SQLQuery & "anfISCIExport, "
    SQLQuery = SQLQuery & "anfAudioExport, "
    SQLQuery = SQLQuery & "anfAutomationExport, "
    SQLQuery = SQLQuery & "anfUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & tlAnf.iCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAnf.sName) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAnf.sSustain) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAnf.sState) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAnf.sSponsorship) & "', "
    SQLQuery = SQLQuery & tlAnf.iMerge & ", "
    SQLQuery = SQLQuery & tlAnf.iRemoteID & ", "
    SQLQuery = SQLQuery & tlAnf.iAutoCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAnf.sBookLocalFeed) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAnf.sRptDefault) & "', "
    SQLQuery = SQLQuery & tlAnf.iSortCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAnf.sTrafToAff) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAnf.sISCIExport) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAnf.sAudioExport) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAnf.sAutomationExport) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlAnf.sUnused) & "' "
    SQLQuery = SQLQuery & ") "
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddAvailName"
        mAddAvailName = -1
        Exit Function
    End If
    
    
    'Reload array and find avail name, then update
    ilRet = gPopAvailNames()
    ilAnfcode = -1
    For ilAnf = LBound(tgAvailNamesInfo) To UBound(tgAvailNamesInfo) - 1 Step 1
        If StrComp(Trim$(tgAvailNamesInfo(ilAnf).sName), slAvailName, vbTextCompare) = 0 Then
            ilAnfcode = tgAvailNamesInfo(ilAnf).iCode
            Exit For
        End If
    Next ilAnf
    If ilAnfcode = -1 Then
        mAddAvailName = -1
        Exit Function
    End If
    SQLQuery = "UPDATE ANF_Avail_Names SET "
    SQLQuery = SQLQuery & "anfAutoCode = " & ilAnfcode
    SQLQuery = SQLQuery & " WHERE anfCode = " & ilAnfcode
    
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddAvailName"
        mAddAvailName = -1
        Exit Function
    End If
    
    mAddAvailName = ilAnfcode
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mAddAvailName"
    mAddAvailName = -1
    Exit Function
End Function

Private Function mAddVPF(ilVefCode As Integer)
    Dim tlVPF As VPF
    Dim ilRet As Integer

    On Error GoTo ErrHand

    tlVPF.ivefKCode = ilVefCode
    tlVPF.sGTime = Format$("12:00AM", sgShowTimeWSecForm)
    tlVPF.sGMedium = "N"
    tlVPF.iurfGCode = 0
    tlVPF.sAdvtSep = "B"
    tlVPF.sCPLogo = ""
    tlVPF.lLgHd1CefCode = 0
    tlVPF.lLgNmCefCode = 0
    tlVPF.sUnsoldBlank = "Y"
    tlVPF.sUsingFeatures1 = Chr(0)  'this might not work
    tlVPF.iSAGroupNo = 0
    'tlVPF.sGPriceStat = "N"
    tlVPF.sOwnership = "A"
    tlVPF.sGGridRes = "F"
    tlVPF.sGScript = "N"
    tlVPF.iGLocalAdj1 = 0
    tlVPF.iGLocalAdj2 = 0
    tlVPF.iGLocalAdj3 = 0
    tlVPF.iGLocalAdj4 = 0
    tlVPF.iGLocalAdj5 = 0
    tlVPF.iGFeedAdj1 = 0
    tlVPF.iGFeedAdj2 = 0
    tlVPF.iGFeedAdj3 = 0
    tlVPF.iGFeedAdj4 = 0
    tlVPF.iGFeedAdj5 = 0
    tlVPF.sGZone1 = ""
    tlVPF.sGZone2 = ""
    tlVPF.sGZone3 = ""
    tlVPF.sGZone4 = ""
    tlVPF.sGZone5 = ""
    tlVPF.iGV2Z1 = 0
    tlVPF.iGV2Z2 = 0
    tlVPF.iGV2Z3 = 0
    tlVPF.iGV2Z4 = 0
    tlVPF.iGV2Z5 = 0
    tlVPF.iGV3Z1 = 0
    tlVPF.iGV3Z2 = 0
    tlVPF.iGV3Z3 = 0
    tlVPF.iGV3Z4 = 0
    tlVPF.iGV3Z5 = 0
    tlVPF.iGV4Z1 = 0
    tlVPF.iGV4Z2 = 0
    tlVPF.iGV4Z3 = 0
    tlVPF.iGV4Z4 = 0
    tlVPF.iGV4Z5 = 0
    tlVPF.sGCSVer1 = ""
    tlVPF.sGCSVer2 = ""
    tlVPF.sGCSVer3 = ""
    tlVPF.sGCSVer4 = ""
    tlVPF.sGCSVer5 = ""
    tlVPF.iGmnfNCode1 = 0
    tlVPF.iGmnfNCode2 = 0
    tlVPF.iGmnfNCode3 = 0
    tlVPF.iGmnfNCode4 = 0
    tlVPF.iGmnfNCode5 = 0
    tlVPF.sGBus1 = ""
    tlVPF.sGBus2 = ""
    tlVPF.sGBus3 = ""
    tlVPF.sGBus4 = ""
    tlVPF.sGBus5 = ""
    tlVPF.sGSked1 = ""
    tlVPF.sGSked2 = ""
    tlVPF.sGSked3 = ""
    tlVPF.sGSked4 = ""
    tlVPF.sGSked5 = ""
    tlVPF.sSVarComm = "N"
    tlVPF.sSCompType = "B"
    tlVPF.sSCompLen = Format$("12:00AM", sgShowTimeWSecForm)
    tlVPF.iSBBLen = 5
    tlVPF.sSSellout = "B"
    tlVPF.sSOverBook = "N"
    tlVPF.sSForceMG = "A"
    tlVPF.sEmbeddedOrROS = "R"
    tlVPF.sWegenerExport = "N"
    tlVPF.sOLAExport = "N"
    tlVPF.sAvailNameOnWeb = "N"
    'tlVPF.sSPTA = "A"
    tlVPF.sUsingFeatures2 = Chr(0)  'this might not work
    tlVPF.sSAvailOrder = 7
    tlVPF.iSLen1 = 30
    tlVPF.iSLen2 = 60
    tlVPF.iSLen3 = 0
    tlVPF.iSLen4 = 0
    tlVPF.iSLen5 = 0
    tlVPF.iSLen6 = 0
    tlVPF.iSLen7 = 0
    tlVPF.iSLen8 = 0
    tlVPF.iSLen9 = 0
    tlVPF.iSLen10 = 0
    tlVPF.iSLenGroup1 = 0
    tlVPF.iSLenGroup2 = 0
    tlVPF.iSLenGroup3 = 0
    tlVPF.iSLenGroup4 = 0
    tlVPF.iSLenGroup5 = 0
    tlVPF.iSLenGroup6 = 0
    tlVPF.iSLenGroup7 = 0
    tlVPF.iSLenGroup8 = 0
    tlVPF.iSLenGroup9 = 0
    tlVPF.iSLenGroup10 = 0
    tlVPF.sSCommCalc = "B"
    tlVPF.lMPSA60 = 0
    tlVPF.lMPSA30 = 0
    tlVPF.lMPSA10 = 0
    tlVPF.lMPromo60 = 0
    tlVPF.lMPromo30 = 0
    tlVPF.lMPromo10 = 0
    tlVPF.iMMFPSA1 = 0
    tlVPF.iMMFPSA2 = 0
    tlVPF.iMMFPSA3 = 0
    tlVPF.iMMFPSA4 = 0
    tlVPF.iMMFPSA5 = 0
    tlVPF.iMMFPSA6 = 0
    tlVPF.iMMFPSA7 = 0
    tlVPF.iMMFPSA8 = 0
    tlVPF.iMMFPSA9 = 0
    tlVPF.iMMFPSA10 = 0
    tlVPF.iMMFPSA11 = 0
    tlVPF.iMMFPSA12 = 0
    tlVPF.iMMFPSA13 = 0
    tlVPF.iMMFPSA14 = 0
    tlVPF.iMMFPSA15 = 0
    tlVPF.iMMFPSA16 = 0
    tlVPF.iMMFPSA17 = 0
    tlVPF.iMMFPSA18 = 0
    tlVPF.iMMFPSA19 = 0
    tlVPF.iMMFPSA20 = 0
    tlVPF.iMMFPSA21 = 0
    tlVPF.iMMFPSA22 = 0
    tlVPF.iMMFPSA23 = 0
    tlVPF.iMMFPSA24 = 0
    tlVPF.iMSaPSA1 = 0
    tlVPF.iMSaPSA2 = 0
    tlVPF.iMSaPSA3 = 0
    tlVPF.iMSaPSA4 = 0
    tlVPF.iMSaPSA5 = 0
    tlVPF.iMSaPSA6 = 0
    tlVPF.iMSaPSA7 = 0
    tlVPF.iMSaPSA8 = 0
    tlVPF.iMSaPSA9 = 0
    tlVPF.iMSaPSA10 = 0
    tlVPF.iMSaPSA11 = 0
    tlVPF.iMSaPSA12 = 0
    tlVPF.iMSaPSA13 = 0
    tlVPF.iMSaPSA14 = 0
    tlVPF.iMSaPSA15 = 0
    tlVPF.iMSaPSA16 = 0
    tlVPF.iMSaPSA17 = 0
    tlVPF.iMSaPSA18 = 0
    tlVPF.iMSaPSA19 = 0
    tlVPF.iMSaPSA20 = 0
    tlVPF.iMSaPSA21 = 0
    tlVPF.iMSaPSA22 = 0
    tlVPF.iMSaPSA23 = 0
    tlVPF.iMSaPSA24 = 0
    tlVPF.iMSuPSA1 = 0
    tlVPF.iMSuPSA2 = 0
    tlVPF.iMSuPSA3 = 0
    tlVPF.iMSuPSA4 = 0
    tlVPF.iMSuPSA5 = 0
    tlVPF.iMSuPSA6 = 0
    tlVPF.iMSuPSA7 = 0
    tlVPF.iMSuPSA8 = 0
    tlVPF.iMSuPSA9 = 0
    tlVPF.iMSuPSA10 = 0
    tlVPF.iMSuPSA11 = 0
    tlVPF.iMSuPSA12 = 0
    tlVPF.iMSuPSA13 = 0
    tlVPF.iMSuPSA14 = 0
    tlVPF.iMSuPSA15 = 0
    tlVPF.iMSuPSA16 = 0
    tlVPF.iMSuPSA17 = 0
    tlVPF.iMSuPSA18 = 0
    tlVPF.iMSuPSA19 = 0
    tlVPF.iMSuPSA20 = 0
    tlVPF.iMSuPSA21 = 0
    tlVPF.iMSuPSA22 = 0
    tlVPF.iMSuPSA23 = 0
    tlVPF.iMSuPSA24 = 0
    tlVPF.iMMFPr1 = 0
    tlVPF.iMMFPr2 = 0
    tlVPF.iMMFPr3 = 0
    tlVPF.iMMFPr4 = 0
    tlVPF.iMMFPr5 = 0
    tlVPF.iMMFPr6 = 0
    tlVPF.iMMFPr7 = 0
    tlVPF.iMMFPr8 = 0
    tlVPF.iMMFPr9 = 0
    tlVPF.iMMFPr10 = 0
    tlVPF.iMMFPr11 = 0
    tlVPF.iMMFPr12 = 0
    tlVPF.iMMFPr13 = 0
    tlVPF.iMMFPr14 = 0
    tlVPF.iMMFPr15 = 0
    tlVPF.iMMFPr16 = 0
    tlVPF.iMMFPr17 = 0
    tlVPF.iMMFPr18 = 0
    tlVPF.iMMFPr19 = 0
    tlVPF.iMMFPr20 = 0
    tlVPF.iMMFPr21 = 0
    tlVPF.iMMFPr22 = 0
    tlVPF.iMMFPr23 = 0
    tlVPF.iMMFPr24 = 0
    tlVPF.iMSaPr1 = 0
    tlVPF.iMSaPr2 = 0
    tlVPF.iMSaPr3 = 0
    tlVPF.iMSaPr4 = 0
    tlVPF.iMSaPr5 = 0
    tlVPF.iMSaPr6 = 0
    tlVPF.iMSaPr7 = 0
    tlVPF.iMSaPr8 = 0
    tlVPF.iMSaPr9 = 0
    tlVPF.iMSaPr10 = 0
    tlVPF.iMSaPr11 = 0
    tlVPF.iMSaPr12 = 0
    tlVPF.iMSaPr13 = 0
    tlVPF.iMSaPr14 = 0
    tlVPF.iMSaPr15 = 0
    tlVPF.iMSaPr16 = 0
    tlVPF.iMSaPr17 = 0
    tlVPF.iMSaPr18 = 0
    tlVPF.iMSaPr19 = 0
    tlVPF.iMSaPr20 = 0
    tlVPF.iMSaPr21 = 0
    tlVPF.iMSaPr22 = 0
    tlVPF.iMSaPr23 = 0
    tlVPF.iMSaPr24 = 0
    tlVPF.iMSuPr1 = 0
    tlVPF.iMSuPr2 = 0
    tlVPF.iMSuPr3 = 0
    tlVPF.iMSuPr4 = 0
    tlVPF.iMSuPr5 = 0
    tlVPF.iMSuPr6 = 0
    tlVPF.iMSuPr7 = 0
    tlVPF.iMSuPr8 = 0
    tlVPF.iMSuPr9 = 0
    tlVPF.iMSuPr10 = 0
    tlVPF.iMSuPr11 = 0
    tlVPF.iMSuPr12 = 0
    tlVPF.iMSuPr13 = 0
    tlVPF.iMSuPr14 = 0
    tlVPF.iMSuPr15 = 0
    tlVPF.iMSuPr16 = 0
    tlVPF.iMSuPr17 = 0
    tlVPF.iMSuPr18 = 0
    tlVPF.iMSuPr19 = 0
    tlVPF.iMSuPr20 = 0
    tlVPF.iMSuPr21 = 0
    tlVPF.iMSuPr22 = 0
    tlVPF.iMSuPr23 = 0
    tlVPF.iMSuPr24 = 0
    tlVPF.sLLD = Format$("1/1/1970", sgShowDateForm)
    tlVPF.sLPD = Format$("1/1/1970", sgShowDateForm)
    tlVPF.slTimeZone = "E"
    tlVPF.sLDaylight = "N"
    tlVPF.sLTiming = "N"
    tlVPF.sLAvailLen = "Y"
    tlVPF.iSDLen = 30
    tlVPF.iFTPArfCode = 0
    tlVPF.sLShowCut = "N"
    tlVPF.sLTimeFormat = "A"
    tlVPF.slZone = "A"
    tlVPF.sCPTitle = ""
    tlVPF.sPrtCPStation = "N"
    tlVPF.iRnfPlayCode = 0
    tlVPF.sLastCP = Format$("1/1/1970", sgShowDateForm)
    tlVPF.sStnFdCart = "N"
    tlVPF.sStnFdXRef = "Y"
    tlVPF.sGenLog = "N"
    tlVPF.sCopyOnAir = "N"
    tlVPF.sBillSA = "S"
    tlVPF.sExpVehNo = ""
    tlVPF.sExpBkCpyCart = "N"
    tlVPF.sExpHiCmmlChg = "N"
    tlVPF.lAPenny = 0
    tlVPF.iGV1Z1 = 0
    tlVPF.iGV1Z2 = 0
    tlVPF.iGV1Z3 = 0
    tlVPF.iGV1Z4 = 0
    tlVPF.iGV1Z5 = 0
    tlVPF.sFedZ1 = 0
    tlVPF.sFedZ2 = 0
    tlVPF.sFedZ3 = 0
    tlVPF.sFedZ4 = 0
    tlVPF.sFedZ5 = 0
    tlVPF.sGGroupNo = ""
    tlVPF.sLLastDateCpyAsgn = Format$("1/1/1970", sgShowDateForm)
    tlVPF.iESTEndTime1 = 0
    tlVPF.iESTEndTime2 = 0
    tlVPF.iESTEndTime3 = 0
    tlVPF.iESTEndTime4 = 0
    tlVPF.iESTEndTime5 = 0
    tlVPF.iCSTEndTime1 = 0
    tlVPF.iCSTEndTime2 = 0
    tlVPF.iCSTEndTime3 = 0
    tlVPF.iCSTEndTime4 = 0
    tlVPF.iCSTEndTime5 = 0
    tlVPF.iMSTEndTime1 = 0
    tlVPF.iMSTEndTime2 = 0
    tlVPF.iMSTEndTime3 = 0
    tlVPF.iMSTEndTime4 = 0
    tlVPF.iMSTEndTime5 = 0
    tlVPF.iPSTEndTime1 = 0
    tlVPF.iPSTEndTime2 = 0
    tlVPF.iPSTEndTime3 = 0
    tlVPF.iPSTEndTime4 = 0
    tlVPF.iPSTEndTime5 = 0
    tlVPF.sMapZone1 = ""
    tlVPF.sMapZone2 = ""
    tlVPF.sMapZone3 = ""
    tlVPF.sMapZone4 = ""
    tlVPF.sMapProgCode1 = ""
    tlVPF.sMapProgCode2 = ""
    tlVPF.sMapProgCode3 = ""
    tlVPF.sMapProgCode4 = ""
    tlVPF.iMapDPNo1 = 0
    tlVPF.iMapDPNo2 = 0
    tlVPF.iMapDPNo3 = 0
    tlVPF.iMapDPNo4 = 0
    tlVPF.sExpHiClear = "N"
    tlVPF.sExpHiDallas = "N"
    tlVPF.sExpHiPhoenix = "N"
    tlVPF.sExpHiNY = "N"
    tlVPF.sBulkXFer = "N"
    tlVPF.sClearAsSell = "N"
    tlVPF.sClearChgTime = "N"
    tlVPF.sMoveLLD = "Y"
    tlVPF.irnfLogCode = 0
    tlVPF.irnfCertCode = 0
    tlVPF.iLNoDaysCycle = 1
    tlVPF.iLLeadTime = 8
    tlVPF.sShowTime = "S"
    tlVPF.sEDICallLetter = ""
    tlVPF.sAccruedRevenue = ""
    tlVPF.sAccruedTrade = ""
    tlVPF.sBilledRevenue = ""
    tlVPF.sBilledTrade = ""
    tlVPF.sLCmmlSmmyAvNm = ""
    tlVPF.lEDASWindow = 400
    tlVPF.sKCGenRot = "Y"
    tlVPF.sExportSQL = "N"
    tlVPF.sAllowSplitCopy = "N"
    tlVPF.sUnunsed1 = ""
    tlVPF.sLastLog = Format$("1/1/1970", sgShowDateForm)
    tlVPF.iRnfSvLogCode = 0
    tlVPF.iRnfSvCertCode = 0
    tlVPF.iRnfSvPlayCode = 0
    tlVPF.lLgFt1CefCode = 0
    tlVPF.lLgFt2CefCode = 0
    tlVPF.sStnFdCode = ""
    tlVPF.iProducerArfCode = 0
    tlVPF.iProgProvArfCode = 0
    tlVPF.iCommProvArfCode = 0
    tlVPF.sEmbeddedComm = "N"
    tlVPF.sARBCode = ""
    tlVPF.lEMailCefCode = 0
    tlVPF.sShowRateOnInsert = "N"
    tlVPF.iAutoExptArfCode = 0
    tlVPF.iAutoImptArfCode = 0
    tlVPF.sWebLogSummary = "N"
    tlVPF.sWebLogFeedTime = "N"
    tlVPF.sRadarCode = ""
    tlVPF.sEDIBand = ""
    tlVPF.iInterfaceID = 0
    
    
    SQLQuery = "Insert Into VPF_Vehicle_Options ( "
    SQLQuery = SQLQuery & "vpfvefKCode, "
    SQLQuery = SQLQuery & "vpfGTime, "
    SQLQuery = SQLQuery & "vpfGMedium, "
    SQLQuery = SQLQuery & "vpfurfGCode, "
    SQLQuery = SQLQuery & "vpfAdvtSep, "
    SQLQuery = SQLQuery & "vpfCPLogo, "
    SQLQuery = SQLQuery & "vpfLgHd1CefCode, "
    SQLQuery = SQLQuery & "vpfLgNmCefCode, "
    SQLQuery = SQLQuery & "vpfUnsoldBlank, "
    SQLQuery = SQLQuery & "vpfUsingFeatures1, "
    SQLQuery = SQLQuery & "vpfSAGroupNo, "
    'SQLQuery = SQLQuery & "vpfGPriceStat, "
    SQLQuery = SQLQuery & "vpfOwnership, "
    SQLQuery = SQLQuery & "vpfGGridRes, "
    SQLQuery = SQLQuery & "vpfGScript, "
    SQLQuery = SQLQuery & "vpfGLocalAdj1, "
    SQLQuery = SQLQuery & "vpfGLocalAdj2, "
    SQLQuery = SQLQuery & "vpfGLocalAdj3, "
    SQLQuery = SQLQuery & "vpfGLocalAdj4, "
    SQLQuery = SQLQuery & "vpfGLocalAdj5, "
    SQLQuery = SQLQuery & "vpfGFeedAdj1, "
    SQLQuery = SQLQuery & "vpfGFeedAdj2, "
    SQLQuery = SQLQuery & "vpfGFeedAdj3, "
    SQLQuery = SQLQuery & "vpfGFeedAdj4, "
    SQLQuery = SQLQuery & "vpfGFeedAdj5, "
    SQLQuery = SQLQuery & "vpfGZone1, "
    SQLQuery = SQLQuery & "vpfGZone2, "
    SQLQuery = SQLQuery & "vpfGZone3, "
    SQLQuery = SQLQuery & "vpfGZone4, "
    SQLQuery = SQLQuery & "vpfGZone5, "
    SQLQuery = SQLQuery & "vpfGV2Z1, "
    SQLQuery = SQLQuery & "vpfGV2Z2, "
    SQLQuery = SQLQuery & "vpfGV2Z3, "
    SQLQuery = SQLQuery & "vpfGV2Z4, "
    SQLQuery = SQLQuery & "vpfGV2Z5, "
    SQLQuery = SQLQuery & "vpfGV3Z1, "
    SQLQuery = SQLQuery & "vpfGV3Z2, "
    SQLQuery = SQLQuery & "vpfGV3Z3, "
    SQLQuery = SQLQuery & "vpfGV3Z4, "
    SQLQuery = SQLQuery & "vpfGV3Z5, "
    SQLQuery = SQLQuery & "vpfGV4Z1, "
    SQLQuery = SQLQuery & "vpfGV4Z2, "
    SQLQuery = SQLQuery & "vpfGV4Z3, "
    SQLQuery = SQLQuery & "vpfGV4Z4, "
    SQLQuery = SQLQuery & "vpfGV4Z5, "
    SQLQuery = SQLQuery & "vpfGCSVer1, "
    SQLQuery = SQLQuery & "vpfGCSVer2, "
    SQLQuery = SQLQuery & "vpfGCSVer3, "
    SQLQuery = SQLQuery & "vpfGCSVer4, "
    SQLQuery = SQLQuery & "vpfGCSVer5, "
    SQLQuery = SQLQuery & "vpfGmnfNCode1, "
    SQLQuery = SQLQuery & "vpfGmnfNCode2, "
    SQLQuery = SQLQuery & "vpfGmnfNCode3, "
    SQLQuery = SQLQuery & "vpfGmnfNCode4, "
    SQLQuery = SQLQuery & "vpfGmnfNCode5, "
    SQLQuery = SQLQuery & "vpfGBus1, "
    SQLQuery = SQLQuery & "vpfGBus2, "
    SQLQuery = SQLQuery & "vpfGBus3, "
    SQLQuery = SQLQuery & "vpfGBus4, "
    SQLQuery = SQLQuery & "vpfGBus5, "
    SQLQuery = SQLQuery & "vpfGSked1, "
    SQLQuery = SQLQuery & "vpfGSked2, "
    SQLQuery = SQLQuery & "vpfGSked3, "
    SQLQuery = SQLQuery & "vpfGSked4, "
    SQLQuery = SQLQuery & "vpfGSked5, "
    SQLQuery = SQLQuery & "vpfSVarComm, "
    SQLQuery = SQLQuery & "vpfSCompType, "
    SQLQuery = SQLQuery & "vpfSCompLen, "
    SQLQuery = SQLQuery & "vpfSBBLen, "
    SQLQuery = SQLQuery & "vpfSSellout, "
    SQLQuery = SQLQuery & "vpfSOverBook, "
    SQLQuery = SQLQuery & "vpfSForceMG, "
    SQLQuery = SQLQuery & "vpfEmbeddedOrROS, "
    SQLQuery = SQLQuery & "vpfWegenerExport, "
    SQLQuery = SQLQuery & "vpfOLAExport, "
    SQLQuery = SQLQuery & "vpfAvailNameOnWeb, "
    'SQLQuery = SQLQuery & "vpfSPTA, "
    SQLQuery = SQLQuery & "vpfUsingFeatures2, "
    SQLQuery = SQLQuery & "vpfSAvailOrder, "
    SQLQuery = SQLQuery & "vpfSLen1, "
    SQLQuery = SQLQuery & "vpfSLen2, "
    SQLQuery = SQLQuery & "vpfSLen3, "
    SQLQuery = SQLQuery & "vpfSLen4, "
    SQLQuery = SQLQuery & "vpfSLen5, "
    SQLQuery = SQLQuery & "vpfSLen6, "
    SQLQuery = SQLQuery & "vpfSLen7, "
    SQLQuery = SQLQuery & "vpfSLen8, "
    SQLQuery = SQLQuery & "vpfSLen9, "
    SQLQuery = SQLQuery & "vpfSLen10, "
    SQLQuery = SQLQuery & "vpfSLenGroup1, "
    SQLQuery = SQLQuery & "vpfSLenGroup2, "
    SQLQuery = SQLQuery & "vpfSLenGroup3, "
    SQLQuery = SQLQuery & "vpfSLenGroup4, "
    SQLQuery = SQLQuery & "vpfSLenGroup5, "
    SQLQuery = SQLQuery & "vpfSLenGroup6, "
    SQLQuery = SQLQuery & "vpfSLenGroup7, "
    SQLQuery = SQLQuery & "vpfSLenGroup8, "
    SQLQuery = SQLQuery & "vpfSLenGroup9, "
    SQLQuery = SQLQuery & "vpfSLenGroup10, "
    SQLQuery = SQLQuery & "vpfSCommCalc, "
    SQLQuery = SQLQuery & "vpfMPSA60, "
    SQLQuery = SQLQuery & "vpfMPSA30, "
    SQLQuery = SQLQuery & "vpfMPSA10, "
    SQLQuery = SQLQuery & "vpfMPromo60, "
    SQLQuery = SQLQuery & "vpfMPromo30, "
    SQLQuery = SQLQuery & "vpfMPromo10, "
    SQLQuery = SQLQuery & "vpfMMFPSA1, "
    SQLQuery = SQLQuery & "vpfMMFPSA2, "
    SQLQuery = SQLQuery & "vpfMMFPSA3, "
    SQLQuery = SQLQuery & "vpfMMFPSA4, "
    SQLQuery = SQLQuery & "vpfMMFPSA5, "
    SQLQuery = SQLQuery & "vpfMMFPSA6, "
    SQLQuery = SQLQuery & "vpfMMFPSA7, "
    SQLQuery = SQLQuery & "vpfMMFPSA8, "
    SQLQuery = SQLQuery & "vpfMMFPSA9, "
    SQLQuery = SQLQuery & "vpfMMFPSA10, "
    SQLQuery = SQLQuery & "vpfMMFPSA11, "
    SQLQuery = SQLQuery & "vpfMMFPSA12, "
    SQLQuery = SQLQuery & "vpfMMFPSA13, "
    SQLQuery = SQLQuery & "vpfMMFPSA14, "
    SQLQuery = SQLQuery & "vpfMMFPSA15, "
    SQLQuery = SQLQuery & "vpfMMFPSA16, "
    SQLQuery = SQLQuery & "vpfMMFPSA17, "
    SQLQuery = SQLQuery & "vpfMMFPSA18, "
    SQLQuery = SQLQuery & "vpfMMFPSA19, "
    SQLQuery = SQLQuery & "vpfMMFPSA20, "
    SQLQuery = SQLQuery & "vpfMMFPSA21, "
    SQLQuery = SQLQuery & "vpfMMFPSA22, "
    SQLQuery = SQLQuery & "vpfMMFPSA23, "
    SQLQuery = SQLQuery & "vpfMMFPSA24, "
    SQLQuery = SQLQuery & "vpfMSaPSA1, "
    SQLQuery = SQLQuery & "vpfMSaPSA2, "
    SQLQuery = SQLQuery & "vpfMSaPSA3, "
    SQLQuery = SQLQuery & "vpfMSaPSA4, "
    SQLQuery = SQLQuery & "vpfMSaPSA5, "
    SQLQuery = SQLQuery & "vpfMSaPSA6, "
    SQLQuery = SQLQuery & "vpfMSaPSA7, "
    SQLQuery = SQLQuery & "vpfMSaPSA8, "
    SQLQuery = SQLQuery & "vpfMSaPSA9, "
    SQLQuery = SQLQuery & "vpfMSaPSA10, "
    SQLQuery = SQLQuery & "vpfMSaPSA11, "
    SQLQuery = SQLQuery & "vpfMSaPSA12, "
    SQLQuery = SQLQuery & "vpfMSaPSA13, "
    SQLQuery = SQLQuery & "vpfMSaPSA14, "
    SQLQuery = SQLQuery & "vpfMSaPSA15, "
    SQLQuery = SQLQuery & "vpfMSaPSA16, "
    SQLQuery = SQLQuery & "vpfMSaPSA17, "
    SQLQuery = SQLQuery & "vpfMSaPSA18, "
    SQLQuery = SQLQuery & "vpfMSaPSA19, "
    SQLQuery = SQLQuery & "vpfMSaPSA20, "
    SQLQuery = SQLQuery & "vpfMSaPSA21, "
    SQLQuery = SQLQuery & "vpfMSaPSA22, "
    SQLQuery = SQLQuery & "vpfMSaPSA23, "
    SQLQuery = SQLQuery & "vpfMSaPSA24, "
    SQLQuery = SQLQuery & "vpfMSuPSA1, "
    SQLQuery = SQLQuery & "vpfMSuPSA2, "
    SQLQuery = SQLQuery & "vpfMSuPSA3, "
    SQLQuery = SQLQuery & "vpfMSuPSA4, "
    SQLQuery = SQLQuery & "vpfMSuPSA5, "
    SQLQuery = SQLQuery & "vpfMSuPSA6, "
    SQLQuery = SQLQuery & "vpfMSuPSA7, "
    SQLQuery = SQLQuery & "vpfMSuPSA8, "
    SQLQuery = SQLQuery & "vpfMSuPSA9, "
    SQLQuery = SQLQuery & "vpfMSuPSA10, "
    SQLQuery = SQLQuery & "vpfMSuPSA11, "
    SQLQuery = SQLQuery & "vpfMSuPSA12, "
    SQLQuery = SQLQuery & "vpfMSuPSA13, "
    SQLQuery = SQLQuery & "vpfMSuPSA14, "
    SQLQuery = SQLQuery & "vpfMSuPSA15, "
    SQLQuery = SQLQuery & "vpfMSuPSA16, "
    SQLQuery = SQLQuery & "vpfMSuPSA17, "
    SQLQuery = SQLQuery & "vpfMSuPSA18, "
    SQLQuery = SQLQuery & "vpfMSuPSA19, "
    SQLQuery = SQLQuery & "vpfMSuPSA20, "
    SQLQuery = SQLQuery & "vpfMSuPSA21, "
    SQLQuery = SQLQuery & "vpfMSuPSA22, "
    SQLQuery = SQLQuery & "vpfMSuPSA23, "
    SQLQuery = SQLQuery & "vpfMSuPSA24, "
    SQLQuery = SQLQuery & "vpfMMFPr1, "
    SQLQuery = SQLQuery & "vpfMMFPr2, "
    SQLQuery = SQLQuery & "vpfMMFPr3, "
    SQLQuery = SQLQuery & "vpfMMFPr4, "
    SQLQuery = SQLQuery & "vpfMMFPr5, "
    SQLQuery = SQLQuery & "vpfMMFPr6, "
    SQLQuery = SQLQuery & "vpfMMFPr7, "
    SQLQuery = SQLQuery & "vpfMMFPr8, "
    SQLQuery = SQLQuery & "vpfMMFPr9, "
    SQLQuery = SQLQuery & "vpfMMFPr10, "
    SQLQuery = SQLQuery & "vpfMMFPr11, "
    SQLQuery = SQLQuery & "vpfMMFPr12, "
    SQLQuery = SQLQuery & "vpfMMFPr13, "
    SQLQuery = SQLQuery & "vpfMMFPr14, "
    SQLQuery = SQLQuery & "vpfMMFPr15, "
    SQLQuery = SQLQuery & "vpfMMFPr16, "
    SQLQuery = SQLQuery & "vpfMMFPr17, "
    SQLQuery = SQLQuery & "vpfMMFPr18, "
    SQLQuery = SQLQuery & "vpfMMFPr19, "
    SQLQuery = SQLQuery & "vpfMMFPr20, "
    SQLQuery = SQLQuery & "vpfMMFPr21, "
    SQLQuery = SQLQuery & "vpfMMFPr22, "
    SQLQuery = SQLQuery & "vpfMMFPr23, "
    SQLQuery = SQLQuery & "vpfMMFPr24, "
    SQLQuery = SQLQuery & "vpfMSaPr1, "
    SQLQuery = SQLQuery & "vpfMSaPr2, "
    SQLQuery = SQLQuery & "vpfMSaPr3, "
    SQLQuery = SQLQuery & "vpfMSaPr4, "
    SQLQuery = SQLQuery & "vpfMSaPr5, "
    SQLQuery = SQLQuery & "vpfMSaPr6, "
    SQLQuery = SQLQuery & "vpfMSaPr7, "
    SQLQuery = SQLQuery & "vpfMSaPr8, "
    SQLQuery = SQLQuery & "vpfMSaPr9, "
    SQLQuery = SQLQuery & "vpfMSaPr10, "
    SQLQuery = SQLQuery & "vpfMSaPr11, "
    SQLQuery = SQLQuery & "vpfMSaPr12, "
    SQLQuery = SQLQuery & "vpfMSaPr13, "
    SQLQuery = SQLQuery & "vpfMSaPr14, "
    SQLQuery = SQLQuery & "vpfMSaPr15, "
    SQLQuery = SQLQuery & "vpfMSaPr16, "
    SQLQuery = SQLQuery & "vpfMSaPr17, "
    SQLQuery = SQLQuery & "vpfMSaPr18, "
    SQLQuery = SQLQuery & "vpfMSaPr19, "
    SQLQuery = SQLQuery & "vpfMSaPr20, "
    SQLQuery = SQLQuery & "vpfMSaPr21, "
    SQLQuery = SQLQuery & "vpfMSaPr22, "
    SQLQuery = SQLQuery & "vpfMSaPr23, "
    SQLQuery = SQLQuery & "vpfMSaPr24, "
    SQLQuery = SQLQuery & "vpfMSuPr1, "
    SQLQuery = SQLQuery & "vpfMSuPr2, "
    SQLQuery = SQLQuery & "vpfMSuPr3, "
    SQLQuery = SQLQuery & "vpfMSuPr4, "
    SQLQuery = SQLQuery & "vpfMSuPr5, "
    SQLQuery = SQLQuery & "vpfMSuPr6, "
    SQLQuery = SQLQuery & "vpfMSuPr7, "
    SQLQuery = SQLQuery & "vpfMSuPr8, "
    SQLQuery = SQLQuery & "vpfMSuPr9, "
    SQLQuery = SQLQuery & "vpfMSuPr10, "
    SQLQuery = SQLQuery & "vpfMSuPr11, "
    SQLQuery = SQLQuery & "vpfMSuPr12, "
    SQLQuery = SQLQuery & "vpfMSuPr13, "
    SQLQuery = SQLQuery & "vpfMSuPr14, "
    SQLQuery = SQLQuery & "vpfMSuPr15, "
    SQLQuery = SQLQuery & "vpfMSuPr16, "
    SQLQuery = SQLQuery & "vpfMSuPr17, "
    SQLQuery = SQLQuery & "vpfMSuPr18, "
    SQLQuery = SQLQuery & "vpfMSuPr19, "
    SQLQuery = SQLQuery & "vpfMSuPr20, "
    SQLQuery = SQLQuery & "vpfMSuPr21, "
    SQLQuery = SQLQuery & "vpfMSuPr22, "
    SQLQuery = SQLQuery & "vpfMSuPr23, "
    SQLQuery = SQLQuery & "vpfMSuPr24, "
    SQLQuery = SQLQuery & "vpfLLD, "
    SQLQuery = SQLQuery & "vpfLPD, "
    SQLQuery = SQLQuery & "vpfLTimeZone, "
    SQLQuery = SQLQuery & "vpfLDaylight, "
    SQLQuery = SQLQuery & "vpfLTiming, "
    SQLQuery = SQLQuery & "vpfLAvailLen, "
    SQLQuery = SQLQuery & "vpfSDLen, "
    SQLQuery = SQLQuery & "vpfFTPArfCode, "
    SQLQuery = SQLQuery & "vpfLShowCut, "
    SQLQuery = SQLQuery & "vpfLTimeFormat, "
    SQLQuery = SQLQuery & "vpfLZone, "
    SQLQuery = SQLQuery & "vpfCPTitle, "
    SQLQuery = SQLQuery & "vpfPrtCPStation, "
    SQLQuery = SQLQuery & "vpfrnfPlayCode, "
    SQLQuery = SQLQuery & "vpfLastCP, "
    SQLQuery = SQLQuery & "vpfStnFdCart, "
    SQLQuery = SQLQuery & "vpfStnFdXRef, "
    SQLQuery = SQLQuery & "vpfGenLog, "
    SQLQuery = SQLQuery & "vpfCopyOnAir, "
    SQLQuery = SQLQuery & "vpfBillSA, "
    SQLQuery = SQLQuery & "vpfExpVehNo, "
    SQLQuery = SQLQuery & "vpfExpBkCpyCart, "
    SQLQuery = SQLQuery & "vpfExpHiCmmlChg, "
    SQLQuery = SQLQuery & "vpfAPenny, "
    SQLQuery = SQLQuery & "vpfGV1Z1, "
    SQLQuery = SQLQuery & "vpfGV1Z2, "
    SQLQuery = SQLQuery & "vpfGV1Z3, "
    SQLQuery = SQLQuery & "vpfGV1Z4, "
    SQLQuery = SQLQuery & "vpfGV1Z5, "
    SQLQuery = SQLQuery & "vpfFedZ1, "
    SQLQuery = SQLQuery & "vpfFedZ2, "
    SQLQuery = SQLQuery & "vpfFedZ3, "
    SQLQuery = SQLQuery & "vpfFedZ4, "
    SQLQuery = SQLQuery & "vpfFedZ5, "
    SQLQuery = SQLQuery & "vpgGGroupNo, "
    SQLQuery = SQLQuery & "vpfLLastDateCpyAsgn, "
    SQLQuery = SQLQuery & "vpfESTEndTime1, "
    SQLQuery = SQLQuery & "vpfESTEndTime2, "
    SQLQuery = SQLQuery & "vpfESTEndTime3, "
    SQLQuery = SQLQuery & "vpfESTEndTime4, "
    SQLQuery = SQLQuery & "vpfESTEndTime5, "
    SQLQuery = SQLQuery & "vpfCSTEndTime1, "
    SQLQuery = SQLQuery & "vpfCSTEndTime2, "
    SQLQuery = SQLQuery & "vpfCSTEndTime3, "
    SQLQuery = SQLQuery & "vpfCSTEndTime4, "
    SQLQuery = SQLQuery & "vpfCSTEndTime5, "
    SQLQuery = SQLQuery & "vpfMSTEndTime1, "
    SQLQuery = SQLQuery & "vpfMSTEndTime2, "
    SQLQuery = SQLQuery & "vpfMSTEndTime3, "
    SQLQuery = SQLQuery & "vpfMSTEndTime4, "
    SQLQuery = SQLQuery & "vpfMSTEndTime5, "
    SQLQuery = SQLQuery & "vpfPSTEndTime1, "
    SQLQuery = SQLQuery & "vpfPSTEndTime2, "
    SQLQuery = SQLQuery & "vpfPSTEndTime3, "
    SQLQuery = SQLQuery & "vpfPSTEndTime4, "
    SQLQuery = SQLQuery & "vpfPSTEndTime5, "
    SQLQuery = SQLQuery & "vpfMapZone1, "
    SQLQuery = SQLQuery & "vpfMapZone2, "
    SQLQuery = SQLQuery & "vpfMapZone3, "
    SQLQuery = SQLQuery & "vpfMapZone4, "
    SQLQuery = SQLQuery & "vpfMapProgCode1, "
    SQLQuery = SQLQuery & "vpfMapProgCode2, "
    SQLQuery = SQLQuery & "vpfMapProgCode3, "
    SQLQuery = SQLQuery & "vpfMapProgCode4, "
    SQLQuery = SQLQuery & "vpfMapDPNo1, "
    SQLQuery = SQLQuery & "vpfMapDPNo2, "
    SQLQuery = SQLQuery & "vpfMapDPNo3, "
    SQLQuery = SQLQuery & "vpfMapDPNo4, "
    SQLQuery = SQLQuery & "vpfExpHiClear, "
    SQLQuery = SQLQuery & "vpfExpHiDallas, "
    SQLQuery = SQLQuery & "vpfExpHiPhoenix, "
    SQLQuery = SQLQuery & "vpfExpHiNY, "
    SQLQuery = SQLQuery & "vpfBulkXFer, "
    SQLQuery = SQLQuery & "vpfClearAsSell, "
    SQLQuery = SQLQuery & "vpfClearChgTime, "
    SQLQuery = SQLQuery & "vpfMoveLLD, "
    SQLQuery = SQLQuery & "vpfrnfLogCode, "
    SQLQuery = SQLQuery & "vpfrnfCertCode, "
    SQLQuery = SQLQuery & "vpfLNoDaysCycle, "
    SQLQuery = SQLQuery & "vpfLLeadTime, "
    SQLQuery = SQLQuery & "vpfShowTime, "
    SQLQuery = SQLQuery & "vpfEDICallLetter, "
    SQLQuery = SQLQuery & "vpfAccruedRevenue, "
    SQLQuery = SQLQuery & "vpfAccruedTrade, "
    SQLQuery = SQLQuery & "vpfBilledRevenue, "
    SQLQuery = SQLQuery & "vpfBilledTrade, "
    SQLQuery = SQLQuery & "vpfLCmmlSmmyAvNm, "
    SQLQuery = SQLQuery & "vpfEDASWindow, "
    SQLQuery = SQLQuery & "vpfKCGenRot, "
    SQLQuery = SQLQuery & "vpfExportSQL, "
    SQLQuery = SQLQuery & "vpfAllowSplitCopy, "
    SQLQuery = SQLQuery & "vpfUnunsed1, "
    SQLQuery = SQLQuery & "vpfLastLog, "
    SQLQuery = SQLQuery & "vpfRnfSvLogCode, "
    SQLQuery = SQLQuery & "vpfRnfSvCertCode, "
    SQLQuery = SQLQuery & "vpfRnfSvPlayCode, "
    SQLQuery = SQLQuery & "vpfLgFt1CefCode, "
    SQLQuery = SQLQuery & "vpfLgFt2CefCode, "
    SQLQuery = SQLQuery & "vpfStnFdCode, "
    SQLQuery = SQLQuery & "vpfProducerArfCode, "
    SQLQuery = SQLQuery & "vpfProgProvArfCode, "
    SQLQuery = SQLQuery & "vpfCommProvArfCode, "
    SQLQuery = SQLQuery & "vpfEmbeddedComm, "
    SQLQuery = SQLQuery & "vpfARBCode, "
    SQLQuery = SQLQuery & "vpfEMailCefCode, "
    SQLQuery = SQLQuery & "vpfShowRateOnInsert, "
    SQLQuery = SQLQuery & "vpfAutoExptArfCode, "
    SQLQuery = SQLQuery & "vpfAutoImptArfCode, "
    SQLQuery = SQLQuery & "vpfWebLogSummary, "
    SQLQuery = SQLQuery & "vpfWebLogFeedTime, "
    SQLQuery = SQLQuery & "vpfInterfaceID "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & tlVPF.ivefKCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlVPF.sGTime, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGMedium) & "', "
    SQLQuery = SQLQuery & tlVPF.iurfGCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sAdvtSep) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sCPLogo) & "', "
    SQLQuery = SQLQuery & tlVPF.lLgHd1CefCode & ", "
    SQLQuery = SQLQuery & tlVPF.lLgNmCefCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sUnsoldBlank) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sUsingFeatures1) & "', "
    SQLQuery = SQLQuery & tlVPF.iSAGroupNo & ", "
    'SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGPriceStat) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sOwnership) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGGridRes) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGScript) & "', "
    SQLQuery = SQLQuery & tlVPF.iGLocalAdj1 & ", "
    SQLQuery = SQLQuery & tlVPF.iGLocalAdj2 & ", "
    SQLQuery = SQLQuery & tlVPF.iGLocalAdj3 & ", "
    SQLQuery = SQLQuery & tlVPF.iGLocalAdj4 & ", "
    SQLQuery = SQLQuery & tlVPF.iGLocalAdj5 & ", "
    SQLQuery = SQLQuery & tlVPF.iGFeedAdj1 & ", "
    SQLQuery = SQLQuery & tlVPF.iGFeedAdj2 & ", "
    SQLQuery = SQLQuery & tlVPF.iGFeedAdj3 & ", "
    SQLQuery = SQLQuery & tlVPF.iGFeedAdj4 & ", "
    SQLQuery = SQLQuery & tlVPF.iGFeedAdj5 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGZone1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGZone2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGZone3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGZone4) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGZone5) & "', "
    SQLQuery = SQLQuery & tlVPF.iGV2Z1 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV2Z2 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV2Z3 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV2Z4 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV2Z5 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV3Z1 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV3Z2 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV3Z3 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV3Z4 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV3Z5 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV4Z1 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV4Z2 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV4Z3 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV4Z4 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV4Z5 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGCSVer1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGCSVer2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGCSVer3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGCSVer4) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGCSVer5) & "', "
    SQLQuery = SQLQuery & tlVPF.iGmnfNCode1 & ", "
    SQLQuery = SQLQuery & tlVPF.iGmnfNCode2 & ", "
    SQLQuery = SQLQuery & tlVPF.iGmnfNCode3 & ", "
    SQLQuery = SQLQuery & tlVPF.iGmnfNCode4 & ", "
    SQLQuery = SQLQuery & tlVPF.iGmnfNCode5 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGBus1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGBus2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGBus3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGBus4) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGBus5) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGSked1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGSked2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGSked3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGSked4) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGSked5) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sSVarComm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sSCompType) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlVPF.sSCompLen, sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & tlVPF.iSBBLen & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sSSellout) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sSOverBook) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sSForceMG) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sEmbeddedOrROS) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sWegenerExport) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sOLAExport) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sAvailNameOnWeb) & "', "
    'SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sSPTA) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sUsingFeatures2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sSAvailOrder) & "', "
    SQLQuery = SQLQuery & tlVPF.iSLen1 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLen2 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLen3 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLen4 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLen5 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLen6 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLen7 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLen8 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLen9 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLen10 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLenGroup1 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLenGroup2 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLenGroup3 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLenGroup4 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLenGroup5 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLenGroup6 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLenGroup7 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLenGroup8 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLenGroup9 & ", "
    SQLQuery = SQLQuery & tlVPF.iSLenGroup10 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sSCommCalc) & "', "
    SQLQuery = SQLQuery & tlVPF.lMPSA60 & ", "
    SQLQuery = SQLQuery & tlVPF.lMPSA30 & ", "
    SQLQuery = SQLQuery & tlVPF.lMPSA10 & ", "
    SQLQuery = SQLQuery & tlVPF.lMPromo60 & ", "
    SQLQuery = SQLQuery & tlVPF.lMPromo30 & ", "
    SQLQuery = SQLQuery & tlVPF.lMPromo10 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA1 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA2 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA3 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA4 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA5 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA6 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA7 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA8 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA9 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA10 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA11 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA12 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA13 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA14 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA15 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA16 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA17 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA18 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA19 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA20 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA21 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA22 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA23 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPSA24 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA1 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA2 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA3 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA4 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA5 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA6 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA7 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA8 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA9 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA10 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA11 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA12 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA13 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA14 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA15 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA16 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA17 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA18 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA19 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA20 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA21 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA22 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA23 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPSA24 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA1 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA2 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA3 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA4 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA5 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA6 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA7 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA8 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA9 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA10 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA11 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA12 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA13 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA14 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA15 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA16 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA17 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA18 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA19 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA20 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA21 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA22 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA23 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPSA24 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr1 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr2 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr3 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr4 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr5 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr6 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr7 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr8 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr9 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr10 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr11 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr12 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr13 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr14 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr15 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr16 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr17 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr18 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr19 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr20 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr21 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr22 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr23 & ", "
    SQLQuery = SQLQuery & tlVPF.iMMFPr24 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr1 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr2 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr3 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr4 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr5 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr6 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr7 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr8 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr9 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr10 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr11 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr12 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr13 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr14 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr15 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr16 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr17 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr18 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr19 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr20 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr21 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr22 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr23 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSaPr24 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr1 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr2 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr3 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr4 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr5 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr6 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr7 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr8 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr9 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr10 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr11 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr12 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr13 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr14 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr15 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr16 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr17 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr18 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr19 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr20 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr21 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr22 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr23 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSuPr24 & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlVPF.sLLD, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlVPF.sLPD, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.slTimeZone) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sLDaylight) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sLTiming) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sLAvailLen) & "', "
    SQLQuery = SQLQuery & tlVPF.iSDLen & ", "
    SQLQuery = SQLQuery & tlVPF.iFTPArfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sLShowCut) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sLTimeFormat) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.slZone) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sCPTitle) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sPrtCPStation) & "', "
    SQLQuery = SQLQuery & tlVPF.iRnfPlayCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(tlVPF.sLastCP, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sStnFdCart) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sStnFdXRef) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGenLog) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sCopyOnAir) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sBillSA) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sExpVehNo) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sExpBkCpyCart) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sExpHiCmmlChg) & "', "
    SQLQuery = SQLQuery & tlVPF.lAPenny & ", "
    SQLQuery = SQLQuery & tlVPF.iGV1Z1 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV1Z2 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV1Z3 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV1Z4 & ", "
    SQLQuery = SQLQuery & tlVPF.iGV1Z5 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sFedZ1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sFedZ2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sFedZ3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sFedZ4) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sFedZ5) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sGGroupNo) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlVPF.sLLastDateCpyAsgn, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlVPF.iESTEndTime1 & ", "
    SQLQuery = SQLQuery & tlVPF.iESTEndTime2 & ", "
    SQLQuery = SQLQuery & tlVPF.iESTEndTime3 & ", "
    SQLQuery = SQLQuery & tlVPF.iESTEndTime4 & ", "
    SQLQuery = SQLQuery & tlVPF.iESTEndTime5 & ", "
    SQLQuery = SQLQuery & tlVPF.iCSTEndTime1 & ", "
    SQLQuery = SQLQuery & tlVPF.iCSTEndTime2 & ", "
    SQLQuery = SQLQuery & tlVPF.iCSTEndTime3 & ", "
    SQLQuery = SQLQuery & tlVPF.iCSTEndTime4 & ", "
    SQLQuery = SQLQuery & tlVPF.iCSTEndTime5 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSTEndTime1 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSTEndTime2 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSTEndTime3 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSTEndTime4 & ", "
    SQLQuery = SQLQuery & tlVPF.iMSTEndTime5 & ", "
    SQLQuery = SQLQuery & tlVPF.iPSTEndTime1 & ", "
    SQLQuery = SQLQuery & tlVPF.iPSTEndTime2 & ", "
    SQLQuery = SQLQuery & tlVPF.iPSTEndTime3 & ", "
    SQLQuery = SQLQuery & tlVPF.iPSTEndTime4 & ", "
    SQLQuery = SQLQuery & tlVPF.iPSTEndTime5 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sMapZone1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sMapZone2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sMapZone3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sMapZone4) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sMapProgCode1) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sMapProgCode2) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sMapProgCode3) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sMapProgCode4) & "', "
    SQLQuery = SQLQuery & tlVPF.iMapDPNo1 & ", "
    SQLQuery = SQLQuery & tlVPF.iMapDPNo2 & ", "
    SQLQuery = SQLQuery & tlVPF.iMapDPNo3 & ", "
    SQLQuery = SQLQuery & tlVPF.iMapDPNo4 & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sExpHiClear) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sExpHiDallas) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sExpHiPhoenix) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sExpHiNY) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sBulkXFer) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sClearAsSell) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sClearChgTime) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sMoveLLD) & "', "
    SQLQuery = SQLQuery & tlVPF.irnfLogCode & ", "
    SQLQuery = SQLQuery & tlVPF.irnfCertCode & ", "
    SQLQuery = SQLQuery & tlVPF.iLNoDaysCycle & ", "
    SQLQuery = SQLQuery & tlVPF.iLLeadTime & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sShowTime) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sEDICallLetter) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sAccruedRevenue) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sAccruedTrade) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sBilledRevenue) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sBilledTrade) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sLCmmlSmmyAvNm) & "', "
    SQLQuery = SQLQuery & tlVPF.lEDASWindow & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sKCGenRot) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sExportSQL) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sAllowSplitCopy) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sUnunsed1) & "', "
    SQLQuery = SQLQuery & "'" & Format$(tlVPF.sLastLog, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & tlVPF.iRnfSvLogCode & ", "
    SQLQuery = SQLQuery & tlVPF.iRnfSvCertCode & ", "
    SQLQuery = SQLQuery & tlVPF.iRnfSvPlayCode & ", "
    SQLQuery = SQLQuery & tlVPF.lLgFt1CefCode & ", "
    SQLQuery = SQLQuery & tlVPF.lLgFt2CefCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sStnFdCode) & "', "
    SQLQuery = SQLQuery & tlVPF.iProducerArfCode & ", "
    SQLQuery = SQLQuery & tlVPF.iProgProvArfCode & ", "
    SQLQuery = SQLQuery & tlVPF.iCommProvArfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sEmbeddedComm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sARBCode) & "', "
    SQLQuery = SQLQuery & tlVPF.lEMailCefCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sShowRateOnInsert) & "', "
    SQLQuery = SQLQuery & tlVPF.iAutoExptArfCode & ", "
    SQLQuery = SQLQuery & tlVPF.iAutoImptArfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sWebLogSummary) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sWebLogFeedTime) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sRadarCode) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(tlVPF.sEDIBand) & "', "
    SQLQuery = SQLQuery & tlVPF.iInterfaceID
    SQLQuery = SQLQuery & ") "
    
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddVPF"
        mAddVPF = False
        Exit Function
    End If
    
    SQLQuery = "SELECT * FROM VPF_Vehicle_Options WHERE (vpfvefKCode = " & ilVefCode & ")"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF Then
        mAddVPF = False
    Else
        mAddVPF = True
    End If
    
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mAddVPF"
    mAddVPF = False
    Exit Function
End Function

Private Function mUpdateLLD(ilVefCode As Integer, slLLDDate As String) As Integer
    Dim slLLD As String

    On Error GoTo ErrHand
    
    If (ilVefCode = imPrevLLDVefCode) And (gDateValue(slLLDDate) = lmPrevLLDDate) Then
        mUpdateLLD = True
        Exit Function
    End If
    SQLQuery = "SELECT vpfLLD"
    SQLQuery = SQLQuery + " FROM VPF_Vehicle_Options"
    SQLQuery = SQLQuery + " WHERE (vpfvefKCode =" & ilVefCode & ")"
    
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        If IsNull(rst!vpfLLD) Then
            slLLD = "1/1/1970"
        Else
            If Not gIsDate(rst!vpfLLD) Then
                slLLD = "1/1/1970"
            Else
                'set sLLD to last log date
                slLLD = Format$(rst!vpfLLD, sgShowDateForm)
            End If
        End If
    Else
        mUpdateLLD = False
        Exit Function
    End If
    If DateValue(slLLDDate) > DateValue(slLLD) Then
        SQLQuery = "UPDATE VPF_Vehicle_Options SET "
        SQLQuery = SQLQuery & "vpfLLD = '" & Format$(slLLDDate, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " WHERE vpfvefKCode =" & ilVefCode
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand1:
            gHandleError "CSIImportLog.Txt", "ImportCSISpot-mUpdateLLD"
            mUpdateLLD = False
            Exit Function
        End If
        imPrevLLDVefCode = ilVefCode
        lmPrevLLDDate = gDateValue(slLLDDate)
        '11/26/17
        gFileChgdUpdate "vpf.btr", True
    End If
    mUpdateLLD = True
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mUpdateLLD"
    mUpdateLLD = False
    Exit Function
End Function

Private Function mUpdateTeam(ilMnfCode As Integer, slAbbreviation As String) As Integer
    On Error GoTo ErrHand

    SQLQuery = "UPDATE MNF_Multi_Names SET "
    SQLQuery = SQLQuery & "mnfUnitType = '" & slAbbreviation & "'"
    SQLQuery = SQLQuery & " WHERE mnfCode = " & ilMnfCode
    
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mUpdateTeam"
        mUpdateTeam = False
        Exit Function
    End If
    
    mUpdateTeam = True
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mUpdateTeam"
    mUpdateTeam = False
    Exit Function
End Function

Private Function mUpdateLang(ilMnfCode As Integer, slEnglish As String) As Integer
    On Error GoTo ErrHand

    SQLQuery = "UPDATE MNF_Multi_Names SET "
    SQLQuery = SQLQuery & "mnfUnitType = '" & slEnglish & "'"
    SQLQuery = SQLQuery & " WHERE mnfCode = " & ilMnfCode
    
    'cnn.Execute SQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mUpdateLang"
        mUpdateLang = False
        Exit Function
    End If
    
    mUpdateLang = True
    On Error GoTo 0
    Exit Function

ErrHand:
    gHandleError "CSIImportLog.Txt", "frmImportCSISpot-mUpdateLang"
    mUpdateLang = False
    Exit Function
End Function

Private Sub txtFile_Change()
    lacResult.Caption = ""
    cmdImport.Caption = "&Import"
    cmdCancel.Caption = "&Cancel"
    cmdImport.Enabled = True
End Sub

Private Function mGetFileNames() As Integer
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    Dim slFromFile As String
    Dim slDrivePath As String
    
    ReDim smFileNames(0 To 0) As String
    
    slFromFile = txtFile.Text
    If Trim$(slFromFile) = "" Then
        mGetFileNames = False
        Exit Function
    End If
    mGetFileNames = True
    ilPos1 = InStr(1, slFromFile, " ", vbTextCompare)
    If ilPos1 <= 0 Then
        smFileNames(0) = slFromFile
        ReDim Preserve smFileNames(0 To 1) As String
        Exit Function
    End If
    slDrivePath = Trim$(Left$(slFromFile, ilPos1))
    Do
        ilPos2 = InStr(ilPos1 + 1, slFromFile, " ", vbTextCompare)
        If ilPos2 <= 0 Then
            smFileNames(UBound(smFileNames)) = slDrivePath & Trim$(Mid$(slFromFile, ilPos1 + 1))
            ReDim Preserve smFileNames(0 To UBound(smFileNames) + 1) As String
            Exit Function
        Else
            smFileNames(UBound(smFileNames)) = slDrivePath & Trim$(Mid$(slFromFile, ilPos1 + 1, ilPos2 - ilPos1 - 1))
            ReDim Preserve smFileNames(0 To UBound(smFileNames) + 1) As String
            ilPos1 = ilPos2
        End If
    Loop
    Exit Function
End Function

Private Function mAddCopy(llInLstCode As Long, slType As String, ilAdfCode As Integer, slCart As String, slISCI As String, slProduct As String, slCreativeTitle As String, ilShttCode As Integer, slXDSCue As String) As Integer
    Dim llLstCode As Long
    On Error GoTo ErrHand

    llLstCode = llInLstCode
    
    SQLQuery = "Insert Into irt ( "
    SQLQuery = SQLQuery & "irtCode, "
    SQLQuery = SQLQuery & "irtLstCode, "
    SQLQuery = SQLQuery & "irtShttCode, "
    SQLQuery = SQLQuery & "irtType, "
    SQLQuery = SQLQuery & "irtAdfCode, "
    SQLQuery = SQLQuery & "irtCart, "
    SQLQuery = SQLQuery & "irtISCI, "
    SQLQuery = SQLQuery & "irtProduct, "
    SQLQuery = SQLQuery & "irtCreativeTitle, "
    SQLQuery = SQLQuery & "irtXDSCue, "
    SQLQuery = SQLQuery & "irtUnused "
    SQLQuery = SQLQuery & ") "
    SQLQuery = SQLQuery & "Values ( "
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & llLstCode & ", "
    SQLQuery = SQLQuery & ilShttCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slType) & "', "
    SQLQuery = SQLQuery & ilAdfCode & ", "
    SQLQuery = SQLQuery & "'" & gFixQuote(slCart) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(slISCI) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(slProduct) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(slCreativeTitle) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(slXDSCue) & "', "
    SQLQuery = SQLQuery & "'" & "" & "' "
    SQLQuery = SQLQuery & ") "
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand1:
        gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddCopy"
        mAddCopy = False
        Exit Function
    End If
    mAddCopy = True
    Exit Function
ErrHand:
    gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddCopy"
    'Resume Next
    mAddCopy = False
'ErrHand1:
'    gHandleError "CSIImportLog.Txt", "ImportCSISpot-mAddCopy"
'    'Return
'    mAddCopy = False
End Function

Private Function mUpdateCptt(ilVefCode As Integer, slGameDate As String)
    Dim slMoDate As String
    
    On Error GoTo ErrHand
    slMoDate = gObtainPrevMonday(slGameDate)
    If (ilVefCode <> imPrevCPTTVefCode) Or (gDateValue(slMoDate) <> lmPrevCPTTDate) Then
        SQLQuery = "UPDATE cptt SET "
        SQLQuery = SQLQuery & "cpttAstStatus = 'N'"
        SQLQuery = SQLQuery & " WHERE cpttvefCode = " & ilVefCode
        SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format(slMoDate, sgSQLDateForm) & "'"
        
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/10/16: Replaced GoSub
            'GoSub ErrHand1:
            gHandleError "CSIImportLog.Txt", "ImportCSISpot-mUpdateCPTT"
            mUpdateCptt = False
            Exit Function
        End If
        imPrevCPTTVefCode = ilVefCode
        lmPrevCPTTDate = gDateValue(slMoDate)
    End If
    mUpdateCptt = True
    Exit Function
ErrHand:
    gHandleError "CSIImportLog.Txt", "ImportCSISpot-mUpdateCPTT"
    'Resume Next
    mUpdateCptt = False
'ErrHand1:
'    gHandleError "CSIImportLog.Txt", "ImportCSISpot-mUpdateCPTT"
'    'Return
'    mUpdateCptt = False
End Function
