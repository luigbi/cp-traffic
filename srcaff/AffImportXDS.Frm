VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmImportXDS 
   Caption         =   "Import XDS"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   6930
   Icon            =   "AffImportXDS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6930
   Begin VB.TextBox txtNumDays 
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Text            =   "7"
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox pbcTextWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      ScaleHeight     =   225
      ScaleWidth      =   1005
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   1035
   End
   Begin V81Affiliate.CSI_Calendar csiStartDate 
      Height          =   285
      Left            =   1545
      TabIndex        =   6
      Top             =   405
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
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
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   0
   End
   Begin VB.ListBox lbcMsg 
      Height          =   2205
      ItemData        =   "AffImportXDS.frx":08CA
      Left            =   240
      List            =   "AffImportXDS.frx":08CC
      TabIndex        =   2
      Top             =   1485
      Width           =   6300
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   3855
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4515
      FormDesignWidth =   6930
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3810
      TabIndex        =   1
      Top             =   3885
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Import"
      Height          =   375
      Left            =   1815
      TabIndex        =   0
      Top             =   3885
      Width           =   1335
   End
   Begin VB.Frame frcDates 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   90
      TabIndex        =   3
      Top             =   390
      Width           =   6585
      Begin VB.Label lacEndDate 
         Caption         =   "Number of days"
         Height          =   225
         Left            =   3435
         TabIndex        =   5
         Top             =   45
         Width           =   1290
      End
      Begin VB.Label lacStartDate 
         Caption         =   "Import: Start Date"
         Height          =   225
         Left            =   0
         TabIndex        =   4
         Top             =   60
         Width           =   1380
      End
   End
   Begin VB.Label lacTitle1 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   285
      TabIndex        =   9
      Top             =   1080
      Width           =   6255
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu MnuDebug 
         Caption         =   "Create Debug File"
      End
      Begin VB.Menu mnuRemote 
         Caption         =   "Don't Download"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Don't Import"
      End
      Begin VB.Menu mnuSingleFile 
         Caption         =   "Download each day as one file"
      End
   End
End
Attribute VB_Name = "frmImportXDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmImportXDS - enters ImportXDS Station information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

'  Mode           UnitID                     TransmissionID           SiteID
'  Cue-By AST     AstCode                    FeedDate: yyyymmdd       SiteID: StationID or Agreement ReceiverID
'  Cue-Not Ast    yyyymmdd+hb+VefCode        FeedDate: yyyymmdd       SiteID: StationID or Agreement ReceiverID
'  Cue-Not Ast    yymmdd+hbp+VefCode         FeedDate: yyyymmdd       SiteID: StationID or Agreement ReceiverID
'
'  ISCI           yyyymmdd+SeqNo+vefCode     FeedDate: yyyymmdd       SiteID: StationID or Agreement ReceiverID

'<TrackEventsResult>
'  <Events>
'    <Event UnitID="159357">
'    <SiteID>753951</SiteID>
'    <ProgramCode>P109</ProgramCode>
'    <Cue>SPH</Cue>
'    <National>0</National>
'    <ISCI>SPORTS_AUTH-RET</ISCI>
'    <Status>insert</Status>
'    <AirTime>10/30/2014 11:22:03 AM</AirTime>
'    <ReportTime>10/30/2014 11:22:04 AM</ReportTime>
'    <SerialNumber>9</SerialNumber>
'    <Index>11</Index>
'    <Watermark />
'    </Event>
'…
'  </Events>
'</TrackEventsResult>

Private imFieldChgd As Integer
Private imExporting As Integer
Private imTerminate As Integer
Private smHeadEndZone As String
Private imHeadEndAdjZone As Integer
Private smLastImportDate As String
Private oMyFileObj As FileSystemObject

Private myImport As CXDS

'replaces logfile
Private smPathForgLogMsg As String
Private myLogInfo As CLogger
Private Const FILELOG As String = "XDigitalImport"
Private Const FILEDEBUG As String = "XDigitalImportDebug"

Private hmAst As Integer
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private myEnt As CENThelper
Private smUnitDate As String
'Private xht_rst As ADODB.Recordset
Private AgreementInfo_rst As ADODB.Recordset
Private ProgramInfo_rst As ADODB.Recordset
Private lmMaxWidth As Long
Private Const MESSAGEBLACK As Long = 0
Private Const MESSAGERED As Long = 255
Private Const MESSAGEGREEN As Long = 39680
Private Const COMPLETEDFOLDER As String = "Completed\"
Private Const COMPLETEDDAYS As Integer = 365
Private Const ISCILENGTH As Integer = 32
Private Const PROGLENGTH As Integer = 8
Private Const ERRORSQL As Integer = 9101
Private Const NODATE As String = "1970-01-01"
Private Const FORMNAME As String = "FrmImportXDS"

'Private Sub cmcBrowse_Click()
'    Dim slCurDir As String
'
'    slCurDir = CurDir
'    ' Set CancelError is True
'    CommonDialog1.CancelError = True
'    On Error GoTo ErrHandler
'    ' Set flags
'    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNFileMustExist
'    ' Set filters
'    'CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
'    '"(*.txt)|*.txt|CSV Files (*.csv)|*.csv"
'    CommonDialog1.Filter = "All Files (*.*)|*.*|XML Files (*.xml)|*.xml|Text Files (*.txt)|*.txt"
'    ' Specify default filter
'    CommonDialog1.FilterIndex = 2
'    ' Display the Open dialog box
'    CommonDialog1.ShowOpen
'    ' Display name of selected file
'
'    edcFile.Text = Trim$(CommonDialog1.fileName)
'    ChDir slCurDir
'    Exit Sub
'ErrHandler:
'    'User pressed the Cancel button
'    Exit Sub
'End Sub

'Private Type IMPORTINFO
'    lAttCode As Long
'    sUnitID As String * 20
'    lAirDate As Long
'    lAirTime As Long
'    sStatus As String * 1
'End Type
'Private tmImportInfo() As IMPORTINFO

    


Private Sub cmdCancel_Click()
    If imExporting Then
        imTerminate = True
        Exit Sub
    End If
    Unload frmImportXDS
End Sub

Private Sub cmdOk_Click()
    Dim slDate As String
    Dim slFile As String
    Dim blRet As Boolean
    Dim fs As New FileSystemObject
    
   ' If rbcSource(0).Value Then
        slDate = csiStartDate.Text
        If slDate = "" Then
            gMsgBox "Import Start Date must be specified.", vbOKOnly
            csiStartDate.SetFocus
            Exit Sub
        End If
        If gIsDate(slDate) = False Then
            Beep
            gMsgBox "Please enter a valid Import Start Date (m/d/yy).", vbCritical
            csiStartDate.SetFocus
            Exit Sub
        End If
        If Not IsNumeric(txtNumDays) Then
            Beep
            gMsgBox "Number of days must be specified.", vbOKOnly
            'txtNumberDays.SetFocus
            Exit Sub
        End If
        If Val(txtNumDays.Text) <= 0 Then
            Beep
            gMsgBox "Number of days must be specified.", vbOKOnly
            'txtNumberDays.SetFocus
            Exit Sub
        End If

'        slDate = csiEndDate.Text
'        If slDate = "" Then
'            gMsgBox "Import End Date must be specified.", vbOKOnly
'            csiEndDate.SetFocus
'            Exit Sub
'        End If
'        If gIsDate(slDate) = False Then
'            Beep
'            gMsgBox "Please enter a valid Import End Date (m/d/yy).", vbCritical
'            csiEndDate.SetFocus
'            Exit Sub
'        End If
'    Else
'        slFile = edcFile.Text
'        If Not fs.FileExists(slFile) Then
'            Beep
'            gMsgBox "Please enter a valid File to Process.", vbCritical
'            Exit Sub
'        End If
'    End If
    imTerminate = False
    imExporting = True
    bgTaskBlocked = False
    sgTaskBlockedName = "X-Digital Import"
    blRet = mImport()
    If bgTaskBlocked Then
         mSetResults "Some spots were blocked during Import.", MESSAGERED
         gMsgBox "Some spots were blocked during the Import." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
    End If
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    imExporting = False
    mUpdateSafImportDate
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / 3
    Me.Height = (Screen.Height) / 3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmCPDat
    Erase tmAstInfo
    AgreementInfo_rst.Close
    ProgramInfo_rst.Close
    Set oMyFileObj = Nothing
    Set myEnt = Nothing
    Set myLogInfo = Nothing
    Set myImport = Nothing
    Set frmImportXDS = Nothing
End Sub

Private Sub mnuDebug_Click()
    mnuDebug.Checked = Not mnuDebug.Checked
    myImport.LogPath = myImport.CreateLogName(sgMsgDirectory & FILEDEBUG)
End Sub
Private Sub mnuImport_Click()
    mnuImport.Checked = Not mnuImport.Checked
End Sub

Private Sub mnuRemote_Click()
    mnuRemote.Checked = Not mnuRemote.Checked
End Sub

Private Sub mnuSingleFile_Click()
    mnuSingleFile.Checked = Not mnuSingleFile.Checked
End Sub
Private Sub mInit()
    If (StrComp(sgUserName, "Guide", 1) = 0) And Not bgLimitedGuide Then
        mnuTools.Visible = True
    Else
        mnuTools.Visible = False
    End If
    txtNumDays.Text = 7
    lmMaxWidth = lbcMsg.Width
    imExporting = False
    imTerminate = False
    frmImportXDS.Caption = "Import XDS - " & sgClientName
    csiStartDate.Top = frcDates.Top
    csiStartDate.Left = frcDates.Left + lacStartDate.Left + lacStartDate.Width
'    csiEndDate.Top = frcDates.Top
'    csiEndDate.Left = frcDates.Left + lacEndDate.Left + lacEndDate.Width
    Set oMyFileObj = New FileSystemObject
    Set myLogInfo = New CLogger
    myLogInfo.LogPath = myLogInfo.CreateLogName(sgMsgDirectory & FILELOG)
    smPathForgLogMsg = FILELOG & "Log_" & Format(gNow(), "mm-dd-yy") & ".txt"
    Set myImport = New CXDS
    'xml.ini sections: [X-Digital] is checked first. For dual provider, use [BackOffice]
    If Not myImport.GetIniValues() Then
        mSetResults "Issue with xml.ini.  This form cannot be acitivated.", MESSAGERED
        myLogInfo.WriteWarning myImport.ErrorMessage
        cmdOK.Enabled = False
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    mSetSafInfo
    If Len(smLastImportDate) > 0 Then
        csiStartDate.Text = smLastImportDate
    Else
        csiStartDate.Text = gObtainPrevMonday(gNow)
    End If
    gPopShttInfo
    Screen.MousePointer = vbDefault
End Sub

Private Function mImport() As Boolean
    Dim blRet As Boolean
    Dim llDate As Long
    Dim slUnitDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llCpttStartDate As Long
    Dim slFile As String
    Dim blSpotsAired As Boolean
    Dim shtt_rst As ADODB.Recordset
    Dim slSite As String
    Dim blDebug As Boolean
    Dim blOneFilePerDay As Boolean
    Dim myFile As file
    Dim slImportPath As String
    Dim ilReturn As Integer
    Dim slFileName As String
    Dim ilSpotCount As Integer
    Dim slPreviousSite As String
    Dim slPreviousProg As String
    Dim blNotAllImported As Boolean
    Dim slStation As String
    Dim slVehicle As String
    Dim blMarkAsUnread As Boolean
    Dim blAtLeastOne As Boolean
    
    Screen.MousePointer = vbHourglass
    imExporting = True
    blRet = True
    lbcMsg.Clear
    lbcMsg.ForeColor = MESSAGEBLACK
On Error Resume Next
    myLogInfo.CleanThisFolder = messages
    myLogInfo.CleanFolder
On Error GoTo 0
    blDebug = mnuDebug.Checked
    blOneFilePerDay = mnuSingleFile.Checked
    slImportPath = mXDSFolder()
    gOpenMKDFile hmAst, "Ast.Mkd"
    If mnuImport.Checked Then
        mSetResults "   User chose to not import files", MESSAGEGREEN
    End If
    mSetResults "Contacting XDS Server", MESSAGEBLACK
    If mnuRemote.Checked Then
        mSetResults "   User chose to not retrieve from XDS", MESSAGEBLACK
    Else
        If blDebug Then
            myImport.LogStart
        End If
        myImport.ImportPath = slImportPath
        myLogInfo.WriteFacts "Importing for " & csiStartDate.Text & " for " & txtNumDays.Text & " days", True
        llStartDate = gDateValue(csiStartDate.Text)
        llEndDate = DateAdd("d", txtNumDays.Text - 1, llStartDate)
        With myImport
            If Not blOneFilePerDay Then
                ''Obtain unique SiteID's
                'SELECT distinct attShfCode,attXDReceiverID, shttStationID FROM att Left Outer Join shtt On attShfCode = shttCode where  (attOnAir <= '2015-11-16') AND (attOffAir >= '2015-11-16') AND (attDropDate >= '2015-11-16')
                SQLQuery = "SELECT Distinct attshfcode, attXDReceiverID, shttStationID,shttcallletters"
                SQLQuery = SQLQuery + " FROM att Left Outer Join shtt On attShfCode = shttCode "
                SQLQuery = SQLQuery & "WHERE (attOnAir <= '" & Format(llStartDate, sgSQLDateForm) & "')"
                SQLQuery = SQLQuery & " AND (attOffAir >= '" & Format(llStartDate, sgSQLDateForm) & "')"
                SQLQuery = SQLQuery & " AND (attDropDate >= '" & Format(llStartDate, sgSQLDateForm) & "') ORDER BY shttcallletters"
                Set shtt_rst = gSQLSelectCall(SQLQuery)
                    Do While Not shtt_rst.EOF
                        slSite = shtt_rst!attXDReceiverId
                        If slSite = "0" Then
                            slSite = shtt_rst!shttStationId
                            If Len(slSite) = 0 Then
                                slSite = "0"
                            End If
                        End If
                        'dan test alter site here
                      '  slSite = "1101354"
                        If slSite <> "0" Then
                            For llDate = llStartDate To llEndDate Step 1
                                .Clear False
                                slUnitDate = Format(llDate, "mm/dd/yyyy")
                                .startDate = slUnitDate & " 00:00:00"
                                .endDate = slUnitDate & " 23:59:59"
                                .SiteID = slSite
                                slFile = "XDSImport_" & Trim$(shtt_rst!shttCallLetters) & "_" & slSite & "_" & Format(llDate, "mm-dd-yyyy")
                                slFile = gFileNameFilter(slFile) & ".xml"
                                mSetResults slFile, MESSAGEBLACK
                                If Not .ImportTimes(slFile) Then
                                    blRet = False
                                    mSetResults .ErrorMessage, MESSAGERED
                                End If
                            Next llDate
                        End If
                        shtt_rst.MoveNext
                    Loop
            Else
                For llDate = llStartDate To llEndDate Step 1
                    .Clear False
                    slUnitDate = Format(llDate, "mm/dd/yyyy")
                    .startDate = slUnitDate & " 00:00:00"
                    .endDate = slUnitDate & " 23:59:59"
                    '.SiteID = ""
                    slFile = "XDSImport_" & Format(llDate, "mm-dd-yyyy") & ".xml"
                    mSetResults slFile, MESSAGEBLACK
                    If Not .ImportTimes(slFile) Then
                        blRet = False
                        mSetResults .ErrorMessage, MESSAGERED
                    End If
                Next llDate
            End If
        End With
    End If
    DoEvents
    If imTerminate Then
        mSetResults "** User Terminated **", MESSAGERED
        imExporting = False
        GoTo Cleanup
    End If
    If Not mnuImport.Checked Then
        Set AgreementInfo_rst = mInitAgreementInfo()
        Set myEnt = New CENThelper
        With myEnt
            .User = igUstCode
            .TypeEnt = Importposted3rdparty
            .ThirdParty = XDS_Break
            .ErrorLog = smPathForgLogMsg
        End With
        For Each myFile In oMyFileObj.GetFolder(slImportPath).Files
            smUnitDate = ""
            slFileName = ""
            blMarkAsUnread = False
            Set ProgramInfo_rst = mInitProgramInfo()
            myEnt.fileName = myFile.Name
            ilReturn = mProcessXml(slImportPath, myFile.Name)
            If ilReturn > 0 Then
                slPreviousSite = ""
                slPreviousProg = ""
                With ProgramInfo_rst
                    If Not .BOF Then
                        .MoveFirst
                    End If
                    Do While Not .EOF
                        If slPreviousSite <> .Fields("SiteID") Then
                            slPreviousSite = .Fields("SiteID")
                            slPreviousProg = ""
                        End If
                        If slPreviousProg <> .Fields("ProgCode") Then
                            slPreviousProg = .Fields("ProgCode")
                            'set all posting times for this agreement
                            ilSpotCount = mProcessSpots(slPreviousSite, slPreviousProg, blNotAllImported, slStation, slVehicle)
                            If ilSpotCount > 0 And Not blNotAllImported Then
                                blAtLeastOne = True
                                mSetResults "Finished posting for " & slStation & " / " & slVehicle & ". Spots: " & ilSpotCount, MESSAGEBLACK
                            ElseIf ilSpotCount = 0 Then
                                blMarkAsUnread = True
                                mSetResults "Could not post for " & slStation & " / " & slVehicle, MESSAGERED
                            ElseIf ilSpotCount < 0 Then
                                blMarkAsUnread = True
                                mSetResults "Problem reading data for " & slStation & " / " & slVehicle, MESSAGERED
                            Else
                                blMarkAsUnread = True
                                mSetResults "Not all spots could be matched for " & slStation & " / " & slVehicle, MESSAGERED
                            End If
                        End If
                        .MoveNext
                    Loop
                End With
            ElseIf ilReturn = 0 Then
                If InStr(1, myFile.Name, "NoProcess_") = 1 Then
                    slFileName = myFile.Name
                Else
                    slFileName = "NoProcess_" & myFile.Name
                End If
                mSetResults myFile.Name & " had no posting information.", MESSAGEBLACK
            Else
                blMarkAsUnread = True
            End If
            If imTerminate Then
                mSetResults "** User Terminated **", MESSAGERED
                imExporting = False
                GoTo Cleanup
            End If
            If blMarkAsUnread Then
                If InStr(1, myFile.Name, "Unread_") = 1 Then
                    slFileName = myFile.Name
                Else
                    slFileName = "Unread_" & myFile.Name
                    If oMyFileObj.FILEEXISTS(slImportPath & COMPLETEDFOLDER & myFile.Name) Then
                        oMyFileObj.DeleteFile (slImportPath & COMPLETEDFOLDER & myFile.Name)
                    End If
                End If
            End If
            If Len(slFileName) = 0 Then
                slFileName = myFile.Name
            End If
            ''        'move file whether read or not
            If oMyFileObj.FILEEXISTS(slImportPath & COMPLETEDFOLDER & slFileName) Then
                oMyFileObj.DeleteFile (slImportPath & COMPLETEDFOLDER & slFileName)
            End If
            myFile.Move (slImportPath & COMPLETEDFOLDER & slFileName)
        Next myFile
On Error Resume Next
         AgreementInfo_rst.Filter = adFilterNone
        If Not (AgreementInfo_rst.EOF And AgreementInfo_rst.BOF) Then
            AgreementInfo_rst.MoveFirst
        End If
        'one record for each agreement/week
        Do While Not AgreementInfo_rst.EOF
            llCpttStartDate = AgreementInfo_rst!MoDate
            blSpotsAired = AgreementInfo_rst!SpotAired
            blRet = mUpdateCptt(blSpotsAired, AgreementInfo_rst!vefCode, AgreementInfo_rst!attCode, Format(llCpttStartDate, "m/d/yy"))
            AgreementInfo_rst.MoveNext
        Loop
    End If
    blRet = True
Cleanup:
    On Error Resume Next
    If Not blAtLeastOne Then
        mSetResults "**No files to read--posting ended**", MESSAGEBLACK
    End If
    If lbcMsg.ForeColor = MESSAGERED Then
        mSetResults "Import finished with issues. See today's " & FILELOG & " log file for more details.", MESSAGERED
        myLogInfo.WriteFacts "Import finished with issues.", True
    Else
        mSetResults "Import Complete", MESSAGEGREEN
        myLogInfo.WriteFacts "Import Complete", True
    End If
    If blDebug Then
        myImport.LogEnd
    End If
    gCloseMKDFile hmAst, "Ast.Mkd"
    mCloseAgreementInfo
    mCloseProgramInfo
    If Not shtt_rst Is Nothing Then
        If (shtt_rst.State And adStateOpen) <> 0 Then
            shtt_rst.Close
        End If
        Set shtt_rst = Nothing
    End If
    cmdCancel.Caption = "&Done"
    Screen.MousePointer = vbDefault
    Set myEnt = Nothing
    Set myFile = Nothing
    imExporting = False
    mImport = blRet
End Function
'
'Private Function mParseString(slUnitDate As String, slXML As String, llDate As Long) As Boolean
'    Dim llAstCode As Long
'    Dim slUnitID As String
'    Dim slSiteId As String
'    Dim slProgramCode As String
'    Dim slCue As String
'    Dim slNational As String
'    Dim slISCI As String
'    Dim slStatus As String
'    Dim slAirDateTime As String
'    Dim slAirDate As String
'    Dim slAirTime As String
'    Dim llAirTime As Long
'    Dim ilAdjTime As Integer
'    Dim ilObjSub As Integer
'    Dim llUpper As Long
'    Dim ilPos As Integer
'    Dim slForm As String
'    Dim blRet As Boolean
'    Dim objXML As New MSXML2.DOMDocument
'    Dim objElem As MSXML2.IXMLDOMElement
'    Dim objSub As MSXML2.IXMLDOMElement
'    Dim objChild As MSXML2.IXMLDOMElement
'    Dim ast_rst As ADODB.Recordset
'
'    mParseString = True
'    If Len(slXML) = 0 Then
'        Exit Function
'    End If
'    If Not objXML.loadXML(slXML) Then
'        mParseString = False
'        Exit Function
'    End If
'    Set objElem = objXML.selectSingleNode("//Events")
'    For Each objSub In objElem.childNodes
'        If objSub.Attributes.LENGTH > 0 Then
'            For ilObjSub = 0 To objSub.Attributes.LENGTH - 1
'                If objSub.Attributes(ilObjSub).nodeName = "UnitID" Then
'                    slUnitID = objSub.Attributes(ilObjSub).nodeValue
'                    If objSub.Attributes(ilObjSub).hasChildNodes = True Then
'                        For Each objChild In objSub.childNodes
'                            Select Case objChild.baseName
'                                Case "SiteID"       'StationID is shttStationID or is attXDReceiverID
'                                    slSiteId = objSub.selectSingleNode(objChild.baseName).Text
'                                Case "ProgramCode"
'                                    slProgramCode = objSub.selectSingleNode(objChild.baseName).Text
'                                Case "Cue"
'                                    slCue = objSub.selectSingleNode(objChild.baseName).Text
'                                Case "National"
'                                    slNational = objSub.selectSingleNode(objChild.baseName).Text
'                                Case "ISCI"     'Might not exist if national spot
'                                    slISCI = objSub.selectSingleNode(objChild.baseName).Text
'                                Case "Status"   'Live if National ran; Insert if Region ran; Skip is region missed
'                                    slStatus = Left$(objSub.selectSingleNode(objChild.baseName).Text, 1)
'                                Case "AirTime"  'identifies the exact time (relative to NOC timezone) that the event executed on the receiver
'                                    slAirDateTime = objSub.selectSingleNode(objChild.baseName).Text
'                                    ilPos = InStr(1, slAirDateTime, " ", vbTextCompare)
'                                    If ilPos > 0 Then
'                                        slAirDate = Left$(slAirDateTime, ilPos - 1)
'                                        slAirTime = Mid$(slAirDateTime, ilPos + 1)
'                                    End If
'                                Case "ReportTime"   'indicates the time the event was reported to the NOC
'                                Case "SerialNumber"
'                                Case "Index"
'                                Case "Watermark"
'                            End Select
'                        Next
'                        slForm = "AST"
'                        If InStr(1, slUnitID, slUnitDate, vbTextCompare) = 1 Then
'                            If Len(slUnitID) = 17 Then
'                                slForm = "ISCI"    'yyyymmddssssvvvvv
'                            Else
'                                slForm = "HB"       'yyyymmddhhbbbvvvvv
'                            End If
'                        ElseIf InStr(1, Mid$(slUnitID, 2, 6), slUnitDate, vbTextCompare) = 1 Then
'                            slForm = "HBP"   'HBP
'                        End If
'                        If slForm = "AST" Then
'                            SQLQuery = "SELECT * FROM ast"
'                            SQLQuery = SQLQuery & " WHERE astCode = " & slUnitID
'                            Set ast_rst = gSQLSelectCall(SQLQuery)
'                            If Not ast_rst.EOF Then
'                                ilAdjTime = mZoneAdjustments(ast_rst!astShfCode, ast_rst!astVefCode)
'                                llAirTime = gTimeToLong(slAirTime, False) + 3600 * (ilAdjTime + imHeadEndAdjZone)
'                                If llAirTime < 0 Then
'                                    llAirTime = llAirTime + 86400
'                                    slAirTime = gLongToTime(llAirTime)
'                                    slAirDate = DateAdd("d", -1, slAirDate)
'                                Else
'                                    slAirTime = gLongToTime(llAirTime)
'                                End If
'                                'Update ast
'                                SQLQuery = "UPDATE ast SET "
'                                SQLQuery = SQLQuery + "astCPStatus = 1" & ", " 'Received
'                                If slStatus <> "S" Then
'                                    If ast_rst!astStatus <= 1 Or ast_rst!astStatus = 9 Or ast_rst!astStatus = 10 Then
'                                        SQLQuery = SQLQuery & "astStatus = " & ast_rst!astStatus
'                                    Else
'                                        SQLQuery = SQLQuery & "astStatus = 1"
'                                    End If
'                                    SQLQuery = SQLQuery + ", "
'                                    SQLQuery = SQLQuery & "astAirDate = '" & Format$(slAirDate, sgSQLDateForm) & "', "
'                                    SQLQuery = SQLQuery & "astAirTime = '" & Format$(slAirTime, sgSQLTimeForm) & "'"
'                                Else
'                                    SQLQuery = SQLQuery + "astStatus = 4"   'Not Aired
'                                End If
'                                SQLQuery = SQLQuery + " WHERE (astCode = " & slUnitID & ")"
'                                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                                    'Error
'                                    mSetResults "Unable to update Affiliate Spot.", MESSAGERED
'                                    If slStatus = "S" Then
'                                        myLogInfo.WriteFacts "Unable to update Affiliate Spot. Import Spot Info:" & " Spot ID " & slUnitID & " Air Date " & slAirDate & " Air Time " & slAirTime & " ISCI " & slISCI & " Status " & "Missed", False
'                                    Else
'                                        myLogInfo.WriteFacts "Unable to update Affiliate Spot. Import Spot Info:" & " Spot ID " & slUnitID & " Air Date " & slAirDate & " Air Time " & slAirTime & " ISCI " & slISCI & " Status " & "Aired", False
'                                    End If
'                                Else
'                                    If slStatus <> "S" Then
'                                        blRet = mAddAgreementInfo(ast_rst!astAtfCode, ast_rst!astVefCode, llDate, True)
'                                    Else
'                                        blRet = mAddAgreementInfo(ast_rst!astAtfCode, ast_rst!astVefCode, llDate, False)
'                                    End If
'                                End If
'                            Else
'                                mSetResults "Matching Affiliate Spot not Found.", MESSAGERED
'                                If slStatus = "S" Then
'                                    myLogInfo.WriteFacts "Matching Affiliate Spot not Found. Import Spot Info:" & " Air Date " & slAirDate & " Air Time " & slAirTime & " ISCI " & slISCI & " Status " & "Missed", False
'                                Else
'                                    myLogInfo.WriteFacts "Matching Affiliate Spot not Found. Import Spot Info:" & " Air Date " & slAirDate & " Air Time " & slAirTime & " ISCI " & slISCI & " Status " & "Aired", False
'                                End If
'                            End If
'                        Else
'                            'Code later: Unit ID by: HB or HBP or ISCI
'                            mSetResults "Matching Affiliate Spot not Found.", MESSAGERED
'                            If slStatus = "S" Then
'                                myLogInfo.WriteFacts "Matching Affiliate Spot not Found. Import Spot Info:" & " Form " & slForm & " Unit ID " & slUnitID & " Site ID " & slSiteId & " Air Date " & slAirDate & " Air Time " & slAirTime & " ISCI " & slISCI & " Status " & "Missed", False
'                            Else
'                                myLogInfo.WriteFacts "Matching Affiliate Spot not Found. Import Spot Info:" & " Form " & slForm & " Unit ID " & slUnitID & " Site ID " & slSiteId & " Air Date " & slAirDate & " Air Time " & slAirTime & " ISCI " & slISCI & " Status " & "Aired", False
'                            End If
'                        End If  'recordset empty
'                    End If 'by astcode
'                End If '"unitID" found
'            Next ilObjSub
'        End If
'    Next
'    If Not ast_rst Is Nothing Then
'        If (ast_rst.State And adStateOpen) <> 0 Then
'            ast_rst.Close
'        End If
'        Set ast_rst = Nothing
'    End If
'
'End Function
Private Function mInitAgreementInfo() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "attCode", adInteger
        .Append "vefCode", adInteger
        .Append "MoDate", adInteger
        .Append "SpotAired", adBoolean
    End With
    rst.Open
    rst!attCode.Properties("optimize") = True
    Set mInitAgreementInfo = rst
End Function
Private Function mInitProgramInfo() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "astCode", adInteger
        .Append "ProgCode", adVarChar, PROGLENGTH
        .Append "SiteID", adInteger
        .Append "ISCI", adVarChar, ISCILENGTH
        .Append "Status", adVarChar, 1
        .Append "AirDate", adDBDate
        .Append "AirTime", adDBTime
        .Append "Found", adBoolean
    End With
    rst.Open
    rst.Sort = "SiteID,ProgCode"
    Set mInitProgramInfo = rst
End Function
Private Sub mCloseProgramInfo()
    On Error Resume Next
    If Not ProgramInfo_rst Is Nothing Then
        If (ProgramInfo_rst.State And adStateOpen) <> 0 Then
            ProgramInfo_rst.Close
        End If
        Set ProgramInfo_rst = Nothing
    End If
End Sub
Private Sub mCloseAgreementInfo()
    On Error Resume Next
    If Not AgreementInfo_rst Is Nothing Then
        If (AgreementInfo_rst.State And adStateOpen) <> 0 Then
            AgreementInfo_rst.Close
        End If
        Set AgreementInfo_rst = Nothing
    End If
End Sub

Private Function mAddAgreementInfo(llAttCode As Long, ilVefCode As Integer, blSpotAired As Boolean) As Integer
    Dim llMoDate As Long
    llMoDate = gDateValue(smUnitDate)
    Do While gWeekDayLong(llMoDate) <> 0
        llMoDate = llMoDate - 1
    Loop
    mAddAgreementInfo = False
    AgreementInfo_rst.Filter = "attCode = " & llAttCode & " And vefCode = " & ilVefCode & " And MoDate = " & llMoDate
    If AgreementInfo_rst.EOF Then
        AgreementInfo_rst.AddNew Array("attCode", "vefCode", "MoDate", "SpotAired"), Array(llAttCode, ilVefCode, llMoDate, blSpotAired)
        mAddAgreementInfo = True
    Else
        If blSpotAired Then
            AgreementInfo_rst!SpotAired = True
        End If
    End If
End Function
Private Function mUpdateCptt(blSpotsAired As Boolean, ilVefCode As Integer, llAtfCode As Long, slMondayFeedDate As String) As Boolean
    'Created by D.S. June 2007  Modified Dan M 11/02/10 V81 new values in cptt added 2/25/2011
    'Set the CPTT value, but only for days between monday and 'last date'
    Dim slSuDate As String
    Dim ilStatus As Integer
    Dim llVeh As Long
    Dim ilAst As Integer
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilPledgeCompliantCount As Integer
    Dim ilAgyCompliantCount As Integer
    Dim blRet As Boolean
    
    On Error GoTo ErrHand
    blRet = True
    slSuDate = DateAdd("d", 6, slMondayFeedDate)
    'Set any Not Aired to received as they are not exported
    For ilStatus = 0 To UBound(tgStatusTypes) Step 1
        If (tgStatusTypes(ilStatus).iPledged = 2) Then
            SQLQuery = "UPDATE ast SET "
            SQLQuery = SQLQuery & "astCPStatus = " & "1"    'Received
            SQLQuery = SQLQuery & " WHERE (astAtfCode = " & llAtfCode
            SQLQuery = SQLQuery & " AND astCPStatus = 0"
            SQLQuery = SQLQuery & " AND astStatus = " & tgStatusTypes(ilStatus).iStatus
            SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')" & ")"
            cnn.BeginTrans
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.Txt", "frmImportXDS-mUpdateCPTT"
                cnn.RollbackTrans
                mUpdateCptt = False
                Exit Function
            End If
            cnn.CommitTrans
        End If
    Next ilStatus
    '8/20/15: Until it is determined how we know if all or region selected, bypass this feature
    'Might use eht-> ect to see FieldName of rbcXSpots and field value of -1 (True) or 0 (False)
    '             -> evt to determine vehicle
    'Issue is know way to know how they ran it for the date being set
    ''ast's not found  are marked as not aired 4
    'SQLQuery = "UPDATE ast SET "
    'SQLQuery = SQLQuery & "astCPStatus = 1, astStatus = 4"    'Received
    'SQLQuery = SQLQuery & " WHERE (astAtfCode = " & llAtfCode
    'SQLQuery = SQLQuery & " AND astCPStatus = 0"
    'SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slLastDate, sgSQLDateForm) & "')" & ")"
    'If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
    '    GoSub ErrHand:
    'End If
    
    'Determine if CPTTStatus should to set to 0=Partial or 1=Completed
    SQLQuery = "Select astCode FROM ast WHERE astCPStatus = 0"
    SQLQuery = SQLQuery & " AND astAtfCode = " & llAtfCode
    SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
    Set rst = gSQLSelectCall(SQLQuery)
    If rst.EOF Then
        If Not blSpotsAired Then
            SQLQuery = "Select astCode FROM ast WHERE astStatus <= 1 or astStatus = 9 Or astStatus = 10"
            SQLQuery = SQLQuery & " AND astAtfCode = " & llAtfCode
            SQLQuery = SQLQuery & " AND (astFeedDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                '6/11/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.Txt", "frmImportXDS-mUpdateCPTT"
                mUpdateCptt = False
                Exit Function
            End If
        End If
        'Set CPTT as complete
        SQLQuery = "UPDATE cptt SET "
        llVeh = gBinarySearchVef(CLng(ilVefCode))
        If llVeh <> -1 Then
            If (tgVehicleInfo(llVeh).sVehType = "G") And (DateValue(slSuDate) > DateValue(Format$(gNow(), "m/d/yy"))) Then
                SQLQuery = SQLQuery & "cpttStatus = 0" & ", " 'Partial
                SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & "cpttPostingStatus = 1" 'Partial
            Else
                If Not blSpotsAired Then
                    'Test for complete week to handle case where only partial days proocessed
                    
                    SQLQuery = SQLQuery & "cpttStatus = 2" & ", " 'Complete
                Else
                    SQLQuery = SQLQuery & "cpttStatus = 1" & ", " 'Complete
                End If
                SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                SQLQuery = SQLQuery & "cpttPostingStatus = 2"  'Complete
            End If
        Else
            If Not blSpotsAired Then
                SQLQuery = SQLQuery & "cpttStatus = 2" & ", " 'Complete
            Else
                SQLQuery = SQLQuery & "cpttStatus = 1" & ", " 'Complete
            End If
            SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
            SQLQuery = SQLQuery & "cpttPostingStatus = 2"  'Complete
        End If
        SQLQuery = SQLQuery & " WHERE cpttAtfCode = " & llAtfCode
        SQLQuery = SQLQuery & " AND (cpttStartDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "frmImportXDS-mUpdateCPTT"
            mUpdateCptt = False
            Exit Function
        End If
    Else
        'Set CPTT as partial
        SQLQuery = "UPDATE cptt SET "
        SQLQuery = SQLQuery & "cpttStatus = 0" & ", " 'Partial
        SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "cpttPostingStatus = 1" 'Partial
        SQLQuery = SQLQuery & " WHERE cpttAtfCode = " & llAtfCode
        SQLQuery = SQLQuery & " AND (cpttStartDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/11/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.Txt", "frmImportXDS-mUpdateCPTT"
            mUpdateCptt = False
            Exit Function
        End If
    End If
    'Dan M V81 has new fields: how many spots to be aired? how many aired?  how many compliant? first step, get changes above into tmastInfo
    ilSchdCount = 0
    ilAiredCount = 0
    ilPledgeCompliantCount = 0
    ilAgyCompliantCount = 0
    gClearASTInfo False
    mLoadCpPosting llAtfCode, slMondayFeedDate
    gGetAstInfo hmAst, tmCPDat(), tmAstInfo(), -1, False, False, True
    For ilAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
        gIncSpotCounts tmAstInfo(ilAst), ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount
    Next ilAst
    SQLQuery = "Update cptt Set "
    SQLQuery = SQLQuery & "cpttNoSpotsGen = " & ilSchdCount & ", "
    SQLQuery = SQLQuery & "cpttNoSpotsAired = " & ilAiredCount & ", "
    SQLQuery = SQLQuery & "cpttNoCompliant = " & ilPledgeCompliantCount & ", "
    SQLQuery = SQLQuery & " cpttReturnDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "cpttAgyCompliant = " & ilAgyCompliantCount & " "
    SQLQuery = SQLQuery & " WHERE cpttAtfCode = " & llAtfCode
    SQLQuery = SQLQuery & " AND (cpttStartDate >= '" & Format$(slMondayFeedDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.Txt", "frmImportXDS-mUpdateCPTT"
        mUpdateCptt = False
        Exit Function
    End If
    gFileChgdUpdate "cptt.mkd", True
    mUpdateCptt = blRet
    Exit Function
ErrHand:
    'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmImportXDS-mUpdateCptt"
    mUpdateCptt = False
End Function

Private Sub mLoadCpPosting(llAtt As Long, slFeedDate As String)
    Dim cprst As ADODB.Recordset
    Dim SQLQuery As String
    Dim ilVpf As Integer
    
    SQLQuery = "SELECT cpttCode,cpttStatus,cpttPostingStatus,cpttAstStatus,attTimeType,shttTimeZone,ShttackDaylight as Daylight, shttTztCode as TimeZone, cpttshfcode, cpttvefcode"
    SQLQuery = SQLQuery & " FROM cptt,att,shtt WHERE (shttCode = cpttShfCode AND attCode = cpttAtfCode "
    SQLQuery = SQLQuery & " AND cpttAtfCode = " & llAtt
    SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(slFeedDate, sgSQLDateForm) & "')"
    Set cprst = gSQLSelectCall(SQLQuery)
    If Not cprst.EOF Then
        ReDim tgCPPosting(0 To 1) As CPPOSTING
        tgCPPosting(0).lCpttCode = cprst!cpttCode
        tgCPPosting(0).iStatus = cprst!cpttStatus
        tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
        tgCPPosting(0).lAttCode = llAtt
        tgCPPosting(0).iAttTimeType = cprst!attTimeType
        tgCPPosting(0).iVefCode = cprst!cpttvefcode
        tgCPPosting(0).iShttCode = cprst!cpttshfcode
        tgCPPosting(0).sZone = cprst!shttTimeZone
        tgCPPosting(0).sDate = Format$(slFeedDate, sgShowDateForm)
        tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
    End If
    cprst.Close
    Set cprst = Nothing
End Sub

Private Sub mSetSafInfo()
    Dim slSQLQuery As String
    Dim saf_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    smHeadEndZone = "E"
    imHeadEndAdjZone = 0
    smLastImportDate = ""
    slSQLQuery = "Select safXDSHeadEndZone, safXDSLastImptDate From SAF_Schd_Attributes WHERE safVefCode = 0"
    Set saf_rst = gSQLSelectCall(slSQLQuery)
    If Not saf_rst.EOF Then
        smHeadEndZone = saf_rst!safXDSHeadEndZone
        'Adjust time as times are relative to location of head end
        If smHeadEndZone = "C" Then
            imHeadEndAdjZone = 1
        ElseIf smHeadEndZone = "M" Then
            imHeadEndAdjZone = 2
        ElseIf smHeadEndZone = "P" Then
            imHeadEndAdjZone = 3
        End If
        If Not IsNull(saf_rst!safxdslastimptdate) Then
            If IsDate(saf_rst!safxdslastimptdate) Then
                smLastImportDate = DateAdd("d", 1, saf_rst!safxdslastimptdate)
            End If
        End If
    End If
    saf_rst.Close
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmImportXDS-mSetSafInfo"
    On Error Resume Next
    saf_rst.Close
    Exit Sub
    
End Sub

Private Function mZoneAdjustments(ilShttCode As Integer, ilVefCode As Integer) As Integer
    Dim llShttRet As Long
    Dim slZone As String
    Dim ilAdjTime As Integer
    
    llShttRet = gBinarySearchShtt(ilShttCode)
    'If Not rst_Shtt.EOF Then
    If llShttRet <> -1 Then
        slZone = Trim$(tgShttInfo1(llShttRet).shttTimeZone)
    Else
        slZone = ""
    End If
    Select Case Left(slZone, 1)
        Case "E"
            ilAdjTime = 0
        Case "C"
            ilAdjTime = -1
        Case "M"
            ilAdjTime = -2
        Case "P"
            ilAdjTime = -3
        Case Else
            ilAdjTime = 0
    End Select
    'mZoneAdjustments = ilAdjTime + mGetZoneAdj(ilVefCode, slZone)
    'adjust based on vehicle time zone table only
    mZoneAdjustments = mGetZoneAdj(ilVefCode, slZone)
End Function
Private Function mGetZoneAdj(ilVefCode As Integer, slZone As String) As Integer
    Dim ilLocalAdj As Integer
    Dim ilNumberAsterisk As Integer
    Dim ilVef As Integer
    Dim ilZone As Integer
    
    ilLocalAdj = 0
    ilVef = gBinarySearchVef(CLng(ilVefCode))
    If Len(Trim$(slZone)) <> 0 Then
        'Get zone
        If ilVef <> -1 Then
            For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                If Trim$(tgVehicleInfo(ilVef).sZone(ilZone)) = Trim$(slZone) Then
                    If (tgVehicleInfo(ilVef).sFed(ilZone) <> "*") And (Trim$(tgVehicleInfo(ilVef).sFed(ilZone)) <> "") And (tgVehicleInfo(ilVef).iBaseZone(ilZone) <> -1) Then
                        ilLocalAdj = tgVehicleInfo(ilVef).iLocalAdj(ilZone)
                    End If
                    Exit For
                End If
            Next ilZone
        End If
    End If
    mGetZoneAdj = ilLocalAdj
End Function
Private Sub mSetResults(slMsg As String, llFGC As Long)
        
    If Len(slMsg) < 1000 Then
        gAddMsgToListBox frmImportXDS, lmMaxWidth, slMsg, lbcMsg
    Else
        gAddMsgToListBox frmImportXDS, lmMaxWidth, "Message too long for listbox--see log for more information.", lbcMsg
    End If
    If lbcMsg.ForeColor <> MESSAGERED Then
        lbcMsg.ForeColor = llFGC
    End If
    DoEvents
    myLogInfo.WriteFacts slMsg
End Sub
Private Function mXDSFolder() As String
    Dim slFolderPath As String
    Dim slNestedPath As String
    Dim myFile As file
    Dim dlDeleteDate As Date
    Dim slFolderName As String
    Dim blFirstTime As Boolean
    
    blFirstTime = False
    slFolderName = mSafeFileName(sgClientName) & "-XDS"
    dlDeleteDate = gNow()
    slFolderPath = oMyFileObj.BuildPath(sgImportDirectory, slFolderName & "\")
    If Not oMyFileObj.FolderExists(slFolderPath) Then
        oMyFileObj.CreateFolder slFolderPath
    End If
    slNestedPath = slFolderPath & COMPLETEDFOLDER
    If Not oMyFileObj.FolderExists(slNestedPath) Then
        oMyFileObj.CreateFolder (slNestedPath)
    End If
    For Each myFile In oMyFileObj.GetFolder(slNestedPath).Files
        If InStr(1, myFile.Name, "UNREAD", vbTextCompare) > 0 Then
        Else
            If DateDiff("d", myFile.DateCreated, dlDeleteDate) > COMPLETEDDAYS Then
                myFile.Delete
            End If
        End If
    Next myFile
    mXDSFolder = slFolderPath
End Function
Private Function mProcessXml(slPath As String, slName As String) As Integer
    ' 1 ok, 0 no spots, -1 an issue
    Dim slUnitID As String
    Dim slSiteId As String
    Dim slProgramCode As String
    Dim slCue As String
    Dim slNational As String
    Dim slISCI As String
    Dim slStatus As String
    Dim slAirDateTime As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim ilObjSub As Integer
    Dim ilPos As Integer
    Dim ilRet  As Integer
    '7878
'    Dim objXML As New MSXML2.DOMDocument
'    Dim objElem As MSXML2.IXMLDOMElement
'    Dim objSub As MSXML2.IXMLDOMElement
'    Dim objChild As MSXML2.IXMLDOMElement
    Dim tlTxtStream As TextStream
    Dim slXML As String
    Dim llAstCode As Long
    Dim blIsAst As Boolean
    Dim blWriteAstMessage As Boolean
    Dim llDateValue As Long
    
'    blWriteAstMessage = True
'    ilRet = 0
'    llDateValue = 0
'On Error GoTo ERRORBOX
'    ilPos = InStrRev(slName, "_", , vbTextCompare)
'    If ilPos > 0 Then
'        smUnitDate = Mid$(slName, ilPos + 1)
'        ilPos = InStrRev(smUnitDate, ".")
'        If ilPos > 0 Then
'            smUnitDate = Mid$(smUnitDate, 1, ilPos - 1)
'        End If
'        llDateValue = gDateValue(smUnitDate)
'    End If
'    If llDateValue = 0 Then
'        mSetResults "Could not extract the Feed Date from the File Name " & slName, MESSAGERED
'        mProcessXml = -1
'        Exit Function
'    End If
'    Set tlTxtStream = oMyFileObj.OpenTextFile(slPath & slName, ForReading, False)
'    Do While tlTxtStream.AtEndOfStream <> True
'        slXML = tlTxtStream.ReadLine
'        If Len(slXML) > 0 Then
'            If Not objXML.loadXML(slXML) Then
'                ilRet = False
'                mSetResults "", MESSAGERED
'                Exit Function
'            End If
'            Set objElem = objXML.selectSingleNode("//Events")
'            For Each objSub In objElem.childNodes
'                slUnitID = ""
'                slSiteId = ""
'                slProgramCode = ""
'                slISCI = ""
'                slStatus = ""
'                slAirDateTime = ""
'                slAirDate = ""
'                slAirTime = ""
'                If objSub.Attributes.LENGTH > 0 Then
'                    For ilObjSub = 0 To objSub.Attributes.LENGTH - 1
'                        If objSub.Attributes(ilObjSub).nodeName = "UnitID" Then
'                            slUnitID = objSub.Attributes(ilObjSub).nodeValue
'                            If objSub.Attributes(ilObjSub).hasChildNodes = True Then
'                                For Each objChild In objSub.childNodes
'                                    Select Case objChild.baseName
'                                        Case "SiteID"       'StationID is shttStationID or is attXDReceiverID
'                                            slSiteId = objSub.selectSingleNode(objChild.baseName).Text
'                                        Case "ProgramCode"
'                                           slProgramCode = objSub.selectSingleNode(objChild.baseName).Text
'                                            If Len(slProgramCode) > PROGLENGTH Then
'                                                slProgramCode = Left(slProgramCode, PROGLENGTH)
'                                            End If
'                                        Case "Cue"
'                                          '  slCue = objSub.selectSingleNode(objChild.baseName).Text
'                                        Case "National"
'                                          '  slNational = objSub.selectSingleNode(objChild.baseName).Text
'                                        Case "ISCI"     'Might not exist if national spot
'                                            slISCI = objSub.selectSingleNode(objChild.baseName).Text
'                                            If Len(slISCI) > ISCILENGTH Then
'                                                slISCI = Left(slISCI, ISCILENGTH)
'                                            End If
'                                        Case "Status"   'Live if National ran; Insert if Region ran; Skip is region missed
'                                            slStatus = Left$(objSub.selectSingleNode(objChild.baseName).Text, 1)
'                                        Case "AirTime"  'identifies the exact time (relative to NOC timezone) that the event executed on the receiver
'                                            slAirDateTime = objSub.selectSingleNode(objChild.baseName).Text
'                                            ilPos = InStr(1, slAirDateTime, " ", vbTextCompare)
'                                            If ilPos > 0 Then
'                                                slAirDate = Left$(slAirDateTime, ilPos - 1)
'                                                slAirTime = Mid$(slAirDateTime, ilPos + 1)
'                                            End If
'                                        Case "ReportTime"   'indicates the time the event was reported to the NOC
'                                        Case "SerialNumber"
'                                        Case "Index"
'                                        Case "Watermark"
'                                    End Select
'                                Next
'                               ' slUnitID = "15"
'                               'make sure astcodes and not date unit ids.  If wrong, write out message once for this file, then continue to try to parse-maybe others are ok?
'                                blIsAst = True
'                                If IsNumeric(slUnitID) Then
'                            On Error GoTo ERRORAST
'                                    llAstCode = slUnitID
'                            On Error GoTo ERRORBOX
'                                    If blIsAst Then
'                                        If llAstCode > 0 And Len(slProgramCode) > 0 And Len(slSiteId) > 0 And Len(slISCI) > 0 And Len(slStatus) > 0 And Len(slAirDate) > 0 And Len(slAirTime) > 0 Then
'                                        Else
'                                            blIsAst = False
'                                        End If
'                                    End If
'                                Else
'                                    blIsAst = False
'                                End If
'                                If blIsAst Then
'                                    'ok, only if not previous error
'                                    If ilRet = 0 Then
'                                        ilRet = 1
'                                    End If
'                                    ProgramInfo_rst.AddNew Array("astCode", "ProgCode", "SiteID", "ISCI", "Status", "AirDate", "AirTime", "Found"), Array(llAstCode, slProgramCode, slSiteId, slISCI, slStatus, slAirDate, Format(slAirTime, "hh:nn:ss"), False)
'                                ElseIf blWriteAstMessage Then
'                                    blWriteAstMessage = False
'                                    ilRet = -1
'                                    mSetResults "Issue in mProcessXml-" & slName, MESSAGERED
'                                    myLogInfo.WriteWarning Err.Description
'                                End If
'                            End If
'                        End If 'UnitID
'                    Next ilObjSub
'                End If
'            Next objSub
'        End If
'    Loop
'    mProcessXml = ilRet
    Exit Function
ERRORAST:
    blIsAst = False
    Resume Next
ERRORBOX:
    mSetResults "Problem in mProcessXml-" & slName, MESSAGERED
    myLogInfo.WriteWarning Err.Description, False
    mProcessXml = -1
End Function
Private Function mProcessSpots(slSiteId As String, slProgramCode As String, blNotAllProcessed As Boolean, slStation As String, slVehicle As String) As Integer
    'O # of spots processed
    '0 blNotAllProcessed.  Were there spots in the import file that couldn't be fit with database spots?  Notice and mark file as such.
    'O station and vehicle for information only
    Dim ilRet As Integer
    Dim llAtt As Long
    Dim ilvehicle As Integer
    Dim ilStation As Integer
    Dim slMondayFeedDate As String
    Dim rsClone As ADODB.Recordset
    Dim ilAdjTime As Integer
    Dim slNoAstExists As String
    
    ilRet = 0
    blNotAllProcessed = False
On Error GoTo ErrHand
    Set rsClone = ProgramInfo_rst.Clone
    'just this agreement
    rsClone.Filter = "SiteID = '" & slSiteId & "' AND progCode = '" & slProgramCode & "'"
    If Not rsClone.EOF Then
        llAtt = mGetCodesFromFirstAst(rsClone, ilvehicle, ilStation, slMondayFeedDate, slStation, slVehicle)
        If llAtt = 0 Or ilvehicle = 0 Or ilStation = 0 Then
            ilRet = -1
            GoTo Cleanup
        End If
        mSetResults "preparing spots to be posted against.", MESSAGEBLACK
        With myEnt
            .Vehicle = ilvehicle
            .Station = ilStation
            .Agreement = llAtt
            .ProcessStart
        End With
        mPrepAst llAtt, slMondayFeedDate
        ilAdjTime = mZoneAdjustments(ilStation, ilvehicle)
        mImportByAstCode blNotAllProcessed, llAtt, ilvehicle, ilStation, rsClone, ilAdjTime

'        If Not mUpdateCptt(blNoSpotAired, ilVehicle, llAtt, slMondayFeedDate) Then
'            mSetResults "Error with make goods in mUpdateCptt", MESSAGERED
'            ilRet = 0
'            GoTo CLEANUP
'        End If
        mInsertToWebLog slVehicle, slStation, llAtt, smUnitDate
        rsClone.Filter = "SiteID = '" & slSiteId & "' AND progCode = '" & slProgramCode & "' AND Found = true"
        ilRet = rsClone.RecordCount
        rsClone.Filter = "SiteID = '" & slSiteId & "' AND progCode = '" & slProgramCode & "' AND Found = false"
        Do While Not rsClone.EOF
            '7458 unmatched...use aired date
            If Not myEnt.Add(Format(rsClone!airDate, sgSQLDateForm), 0, SentOrReceived) Then
                myLogInfo.WriteWarning "using air date: " & myEnt.ErrorMessage
            End If
            '6714
            If rsClone!astCode > 0 Then
                blNotAllProcessed = True
                slNoAstExists = slVehicle & "," & slStation & ", attCode: "
                slNoAstExists = slNoAstExists & llAtt & ", ISCI: "
                slNoAstExists = slNoAstExists & Trim$(rsClone!ISCI) & ", astCode:"
                slNoAstExists = slNoAstExists & rsClone!astCode & ", Air date: "
                slNoAstExists = slNoAstExists & Format(rsClone!airDate, sgShowDateForm) & ", Air time: "
                slNoAstExists = slNoAstExists & Format(rsClone!airTime, sgSQLTimeForm) & ", Status: "
                slNoAstExists = slNoAstExists & rsClone!Status
                mSetResults "Unable to process: " & slNoAstExists & " AST missing", MESSAGERED
            Else
                myLogInfo.WriteWarning "Unable to process a spot because ast code was 0"
            End If
            rsClone.MoveNext
        Loop
    End If
Cleanup:
    If Not rsClone Is Nothing Then
        If (rsClone.State And adStateOpen) <> 0 Then
            rsClone.Close
        End If
        Set rsClone = Nothing
    End If
    mProcessSpots = ilRet
    Exit Function
ErrHand:
    ilRet = -1
    myLogInfo.WriteError "mProcessSpots-" & Err.Description, False, False
    GoTo Cleanup
End Function
Private Function mGetCodesFromFirstAst(rsClone As ADODB.Recordset, ilvehicle As Integer, ilStation As Integer, slFeedDate As String, slStation As String, slVehicle As String) As Long
'return  AttCode  O-ilVehicle and ilStation and slFeedDate, and slStation, and slVehicle
    Dim rs As ADODB.Recordset
    Dim Sql As String
    Dim blAstFound As Boolean
    Dim c As Integer
    
    ilvehicle = 0
    ilStation = 0
    slFeedDate = ""
On Error GoTo ErrHandle
    Do While Not rsClone.EOF
        Sql = "Select astAtfCode,astvefCode,astshfcode,astFeedDate,shttcallletters,vefname from ast inner join shtt on astshfcode = shttcode inner join VEF_Vehicles on astvefcode = vefcode  where astcode = " & rsClone!astCode
        Set rs = gSQLSelectCall(Sql)
        If Not (rs.EOF And rs.BOF) Then
            mGetCodesFromFirstAst = rs!astAtfCode
            ilvehicle = rs!astVefCode
            ilStation = rs!astShfCode
            slFeedDate = gAdjYear(gObtainPrevMonday(rs!astFeedDate))
            slVehicle = Trim(rs!vefName)
            slStation = Trim(rs!shttCallLetters)
            GoTo Cleanup
        End If
        rsClone.MoveNext
    Loop
Cleanup:
    If Not rs Is Nothing Then
        If (rs.State And adStateOpen) <> 0 Then
            rs.Close
        End If
        Set rs = Nothing
    End If
    rsClone.MoveFirst
    Exit Function
ErrHandle:
    mGetCodesFromFirstAst = 0
    mSetResults "issue in mGetCodesFromFirstAst:" & Err.Description, MESSAGERED
    GoTo Cleanup

End Function

Private Function mImportByAstCode(blNoSpotAired As Boolean, llAtfCode As Long, ilVefCode As Integer, ilShfCode As Integer, rsClone As ADODB.Recordset, ilAdjTime As Integer) As Boolean
'    update the AST if it exists and mark it as found in the array. Return false only on error OR code matches but vehicle and station don't.
'   O-blNoSpotAired
'    6/7/2011 added multicasting
    Dim llAstCode As Long
    Dim llSpotLoop As Long
    Dim blFound As Boolean
    Dim tlMyAst As AST
    Dim slStatus As String
    Dim slDate As String
    Dim slTime As String
    Dim slBaseFilter As String
    Dim llAirTime As Long
    Dim blThisSpotMissed As Boolean
    Dim llNewCpf As Long
    Dim ilAddToStatus As Integer

On Error GoTo ErrHand
    mImportByAstCode = True
    blNoSpotAired = True
    slBaseFilter = rsClone.Filter
    For llSpotLoop = 0 To UBound(tmAstInfo) - 1
        If Not myEnt.Add(tmAstInfo(llSpotLoop).sFeedDate, tmAstInfo(llSpotLoop).lgsfCode, Asts) Then
            myLogInfo.WriteWarning myEnt.ErrorMessage
        End If
        DoEvents
        llAstCode = tmAstInfo(llSpotLoop).lCode
        blFound = False
        rsClone.Filter = slBaseFilter & " AND astcode = " & llAstCode
        If Not rsClone.EOF Then
            'Dan M note:  Use attcode,station, and vehicle to test if same spot.
            'because the import doesn't pass most of this info, ilshfcode, and ilvefcode are set by looking at the first spot.
            ' All we are doing is making sure all further spots match this one.
            If tmAstInfo(llSpotLoop).lAttCode = llAtfCode Then
                blFound = True
                If Not myEnt.Add(tmAstInfo(llSpotLoop).sFeedDate, tmAstInfo(llSpotLoop).lgsfCode, Ingested) Then
                    myLogInfo.WriteWarning myEnt.ErrorMessage
                End If
                '7639
                slStatus = rsClone!Status
                If slStatus <> "S" Then
                    llNewCpf = mAdjustISCIAsNeeded(llSpotLoop, rsClone!ISCI)
                Else
                    llNewCpf = 0
                End If
                With tlMyAst
                    'update ast
                    .lCode = llAstCode
                    .iStatus = tmAstInfo(llSpotLoop).iStatus
                    .iVefCode = ilVefCode
                    .iCPStatus = 1
                    .lSdfCode = tmAstInfo(llSpotLoop).lSdfCode
                    .lLsfCode = tmAstInfo(llSpotLoop).lLstCode
                    gPackDate tmAstInfo(llSpotLoop).sFeedDate, .iFeedDate(0), .iFeedDate(1)
                    gPackTime tmAstInfo(llSpotLoop).sFeedTime, .iFeedTime(0), .iFeedTime(1)
                    .iAdfCode = tmAstInfo(llSpotLoop).iAdfCode
                    .lDatCode = tmAstInfo(llSpotLoop).lDatCode
                    '7639
                    If llNewCpf > 0 Then
                        .lCpfCode = llNewCpf
                        ilAddToStatus = ASTEXTENDED_ISCICHGD
                    Else
                        .lCpfCode = tmAstInfo(llSpotLoop).lCpfCode
                        ilAddToStatus = 0
                    End If
                    .lCpfCode = tmAstInfo(llSpotLoop).lCpfCode
                    ilAddToStatus = 0
                    .lRsfCode = tmAstInfo(llSpotLoop).lRRsfCode
                    .sStationCompliant = ""
                    .sAgencyCompliant = ""
                    .sAffidavitSource = ""
                    .iUstCode = igUstCode
                End With
                slDate = Format(rsClone!airDate, sgSQLDateForm)
                slTime = Format(rsClone!airTime, sgSQLTimeForm)
                llAirTime = gTimeToLong(slTime, False) + 3600 * (ilAdjTime + imHeadEndAdjZone)
                If llAirTime < 0 Then
                    llAirTime = llAirTime + 86400
                    slTime = gLongToTime(llAirTime)
                    slDate = DateAdd("d", -1, slDate)
                Else
                    slTime = gLongToTime(llAirTime)
                End If
                blThisSpotMissed = Not mUpdateAst(slStatus, Format(slDate, sgSQLDateForm), Format(slTime, sgSQLTimeForm), tlMyAst, ilAddToStatus)
                If Not blThisSpotMissed Then
                    blNoSpotAired = False
                    rsClone!found = True
                    If slStatus <> "S" Then
                        mAddAgreementInfo llAtfCode, ilVefCode, True
                    Else
                        mAddAgreementInfo llAtfCode, ilVefCode, False
                    End If
                End If
            End If  'astcodes match
        End If
    Next llSpotLoop
    Exit Function
ErrHand:
    mImportByAstCode = False
    'Throw to mProcessSpots to stop all processing.
    Err.Raise ERRORSQL, "mImportByAstCode", Err.Description
End Function


Private Sub mPrepAst(llAtt As Long, slMondayFeedDate As String)
    mLoadCpPosting llAtt, slMondayFeedDate
    DoEvents
    igTimes = 1 'By Week
    gGetAstInfo hmAst, tmCPDat(), tmAstInfo(), -1, True, False, True
End Sub
Private Function mSafeFileName(slOldName As String) As String
    Dim slTempName As String
    slTempName = Replace(slOldName, "?", "-")
    slTempName = Replace(slTempName, "/", "-")
    slTempName = Replace(slTempName, "\", "-")
    slTempName = Replace(slTempName, "%", "-")
    slTempName = Replace(slTempName, "*", "-")
    slTempName = Replace(slTempName, ":", "-")
    slTempName = Replace(slTempName, "|", "-")
    slTempName = Replace(slTempName, """", "-")
    slTempName = Replace(slTempName, ".", "-")
    slTempName = Replace(slTempName, "<", "-")
    slTempName = Replace(slTempName, ">", "-")
    mSafeFileName = slTempName
End Function
Private Function mUpdateAst(slStatus As String, slDate As String, slTime As String, myTlAst As AST, ilAddToStatus As Integer) As Boolean
    ' O- did spots air?
    Dim llAstCode As Long
    Dim ilAstStatus As Integer
    
    llAstCode = myTlAst.lCode
    If slStatus <> "S" Then
        'C - Program and Commercial aired.
        '   XDS returns = "L"(national) or "I"(regional)  Status = 0  Screen = 1-Aired Live
        'N - Neither the spot nor the program aired.
        '   XDS returns = "S"  Status = 4   Screen = 5-Not Aired Other
        'unused:
        'D - Program and Commercial were both delayed.
        '   XDS does not return  Status = 9   Screen = 10-Delay Cmml/Prg
        'S - Program did not air, but spot aired, either live or delayed.
        '   XDS does not return  Status = 10  Screen = 11-Air Cmml Only
        'P - Program aired spot did not.
        '   XDS does not return  Status = 4   Screen = 5-Not Aired Other
        'K - Delay B'cast
        '   XDS does not return  Status = 1   Screen = 2-Delay B'cast
        mUpdateAst = True
        'update date/time aired
        If myTlAst.iStatus <= 1 Or myTlAst.iStatus = 9 Or myTlAst.iStatus = 10 Then
            ilAstStatus = myTlAst.iStatus
        Else
            ilAstStatus = 1
        End If
        ilAstStatus = ilAstStatus + ilAddToStatus
        SQLQuery = "UPDATE ast SET astCPStatus = 1, astStatus = " & ilAstStatus & ", "
        SQLQuery = SQLQuery & "astAirDate = '" & slDate & "', "
        SQLQuery = SQLQuery & "astAirTime = '" & slTime & "'"
        SQLQuery = SQLQuery & ", astcpfCode = " & myTlAst.lCpfCode
        SQLQuery = SQLQuery & " WHERE (astCode = " & llAstCode & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            mUpdateAst = False
            Err.Raise ERRORSQL, "mUpdateAst", "Problem in mUpdateAst"
        End If
    Else
         'update status as not aired
        ilAstStatus = 4
        SQLQuery = "UPDATE ast SET astCPStatus = 1, astStatus = " & ilAstStatus
        SQLQuery = SQLQuery & " WHERE (astCode = " & llAstCode & ")"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            mUpdateAst = False
            Err.Raise ERRORSQL, "mUpdateAst", "Problem in mUpdateAst"
        End If
    End If
End Function
Private Function mAdjustISCIAsNeeded(llSpotLoop As Long, slImportISCI As String) As Long
    'return CpfCode if import isci is different; otherwise 0
    'add or update cpf with new isci; then create alt for old isci
    Dim slAstISCI As String
    Dim llRet As Long
    
    llRet = 0
    'Dan M 7/29/15 don't do 7639 until later: reports need to be fixed. To restore: lose goto line below
    GoTo Cleanup
On Error GoTo ERRORBOX
    If llSpotLoop <= UBound(tmAstInfo) Then
        If tmAstInfo(llSpotLoop).iRegionType = 0 Then
            slAstISCI = tmAstInfo(llSpotLoop).sISCI
        Else
            slAstISCI = tmAstInfo(llSpotLoop).sRISCI
        End If
        If slAstISCI <> slImportISCI Then
            SQLQuery = "select cpfCode from cpf_Copy_Prodct_ISCI where cpfisci = '" & slImportISCI & "'"
            Set rst = gSQLSelectCall(SQLQuery)
            If Not rst.EOF Then
                llRet = rst!cpfCode
            Else
                SQLQuery = "INSERT into cpf_Copy_Prodct_ISCI (cpfCode,cpfName,cpfIsci,cpfCreative,cpfRotEndDate,cpfsifCode) VALUES (Replace,'','" & slImportISCI & "','','" & NODATE & "',0)"
                llRet = gInsertAndReturnCode(SQLQuery, "cpf_Copy_Prodct_ISCI", "cpfCode", "Replace")
            End If
            'now create alt
            If llRet <> 0 Then
                mAddAltForIsci tmAstInfo(llSpotLoop).lCode, tmAstInfo(llSpotLoop).iAdfCode, tmAstInfo(llSpotLoop).lCpfCode
            Else
                mSetResults "warning!  Issue in mAdjustISCIAsNeeded. Please see log", MESSAGERED
                myLogInfo.WriteWarning "sql call invalid:" & SQLQuery, False
            End If

        End If
    Else
        mSetResults "warning!  Issue in mAdjustISCIAsNeeded. Please see log", MESSAGERED
        myLogInfo.WriteWarning "index passed: " & llSpotLoop & " is not valid.  Max: " & UBound(tmAstInfo), False
    End If
Cleanup:
    mAdjustISCIAsNeeded = llRet
    Exit Function
ERRORBOX:
    mSetResults "warning!  Issue in mAdjustISCIAsNeeded. Please see log", MESSAGERED
    gHandleError smPathForgLogMsg, "frmImportXDS-mAdjustISCIAsNeeded"
    llRet = 0
    GoTo Cleanup
End Function
Private Sub mInsertToWebLog(slVehicleName As String, slStation As String, llAttCode As Long, slDate As String)
    Dim SQLQuery As String
    Dim slCurrent As String
    Dim slSignature As String
    
    slCurrent = gNow()
    slSignature = "XDS"
    SQLQuery = "Insert Into WebL (weblType, weblattCode, weblCallLetters, weblVehicleName, weblUserName, weblPostDay, weblDate, weblTime) " ' weblIP, weblCPUName,
    SQLQuery = SQLQuery & "Values (4," & llAttCode & ",'" & gFixQuote(slStation) & "','" & gFixQuote(slVehicleName) & "','" & gFixQuote(slSignature) & "', '" & Format(slDate, sgSQLDateForm) & "',"
    SQLQuery = SQLQuery & "'" & Format$(slCurrent, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(slCurrent, sgSQLTimeForm) & "'"
    SQLQuery = SQLQuery & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        GoTo ERRSQL
    End If
    Exit Sub
ERRSQL:
    'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError smPathForgLogMsg, "frmImportXDS-mInsertToWebLog"
    mSetResults gMsg, MESSAGERED
End Sub
Private Sub mAddAltForIsci(llAstCode As Long, ilAdfCode As Integer, llCpfCode As Long)
    'note that the tmastInfo has never been updated with any new isci data
    SQLQuery = "insert into alt (altAstCode,altMissedDate,altAdfCode,altMgDate,altCpfCode) values (" & llAstCode & ",'" & NODATE & "'," & ilAdfCode & ",'" & NODATE & "'," & llCpfCode & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/11/16: Replaced GoSub
        'GoSub errbox:
        mSetResults "warning!  Issue in mAddAltForIsci. Please see log", MESSAGERED
        gHandleError "AffErrorLog.Txt", "frmImportXDS-mAddAltForISCI"
        Exit Sub
    End If
    Exit Sub
End Sub
Private Sub txtNumDays_Change()
    gCtrlGotFocus ActiveControl
End Sub
Private Sub mUpdateSafImportDate()
    Dim SQLQuery As String
    Dim slCurrent As String
    Dim blContinue As Boolean
    
    blContinue = False
    slCurrent = DateAdd("d", txtNumDays - 1, csiStartDate.Text)
    slCurrent = Format(slCurrent, sgSQLDateForm)
    If IsDate(slCurrent) And IsDate(smLastImportDate) Then
        ' only update if the date they imported is later than what is in database
        If DateDiff("d", smLastImportDate, slCurrent) > 0 Then
            blContinue = True
        End If
    ElseIf IsDate(slCurrent) Then
        blContinue = True
    End If
    If blContinue Then
        SQLQuery = "update  SAF_Schd_Attributes set safXDSLastImptDate = '" & slCurrent & "' WHERE safVefCode = 0"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            GoTo ERRSQL
        End If
        smLastImportDate = slCurrent
    End If
    Exit Sub
ERRSQL:
    gHandleError smPathForgLogMsg, "frmImportXDS-mUpdateSafImportDate"
    mSetResults gMsg, MESSAGERED
End Sub
