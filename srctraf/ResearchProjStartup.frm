VERSION 5.00
Begin VB.Form ResearchProjStartup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Research"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   630
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmcMain 
      Interval        =   50
      Left            =   45
      Top             =   120
   End
End
Attribute VB_Name = "ResearchProjStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim omResearchChoice As ResearchChoice
Private Enum ResearchChoice
    ResearchX = 0
    ByCust = 1
End Enum
Dim imTerminate As Integer  'True = terminating task, False= OK
Private Sub Form_Load()
    mInit
    If imTerminate Then
        Unload Me
    End If
End Sub
Private Sub mInit()
    tmcMain.Enabled = False
    Screen.MousePointer = vbHourglass
    Me.Visible = False
    Dim blContinue As Boolean
    
    mParseCmmdLine
    Screen.MousePointer = vbDefault
    If Not igExitTraffic Then
        mGoToFormChoice
    End If
'    Do While Not blContinue
'        DoEvents
'        If igExitTraffic Then
'            tmcMain.Enabled = True
'            blContinue = True
'            Exit Do
'        End If
'        Sleep 500
'    Loop
    tmcMain.Enabled = True
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
    Dim slStartIn As String
    Dim blIsCsi As Boolean
    Dim blIsFromTraffic As Boolean

    sgCommandStr = Command$
    slStartIn = CurDir$
    igShowVersionNo = 0
    slCommand = sgCommandStr
    lgCurrHRes = GetDeviceCaps(Traffic!pbcList.hdc, HORZRES)
    lgCurrVRes = GetDeviceCaps(Traffic!pbcList.hdc, VERTRES)
    lgCurrBPP = GetDeviceCaps(Traffic!pbcList.hdc, BITSPIXEL)
    '10365 we need to connect to database now, so need 'test' or 'prod' now.
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        igTestSystem = False
    Else
        igTestSystem = True
    End If
    mTestPervasive
    lgUlfCode = 0
    blIsFromTraffic = False
    'not coming from traffic
    If InStr(1, sgCommandStr, "^", vbTextCompare) <= 0 Then
        Signon.Show vbModal
        If igExitTraffic Then
            imTerminate = True
            Exit Sub
        End If
        '10365 already did this
'        If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
'            igTestSystem = False
'        Else
'            igTestSystem = True
'        End If
        slStr = sgUserName
        sgCallAppName = "Traffic"
    Else
        blIsFromTraffic = True
        gParseItem slCommand, 1, "\", slStr    'Get application name
        gParseItem slStr, 1, "^", sgCallAppName    'Get application name
        '10365 already did this.  No longer testing what gets sent from traffic.
'        gParseItem slStr, 2, "^", slStr    'Get Test or prod
'        If slStr = "Prod" Then
'            igTestSystem = False
'        Else
'            igTestSystem = True
'        End If
        gParseItem slCommand, 2, "\", slStr    'user
        sgUrfStamp = "~" 'Clear time stamp incase same name
        sgUserName = Trim$(slStr)
    End If
    If StrComp(sgUserName, "CSI", vbTextCompare) = 0 Then
        sgSpecialPassword = mDetermineCsiLogin()
    End If
    gUrfRead ResearchProjStartup, sgUserName, True, tgUrf(), False  'Obtain user records
    If Len(sgSpecialPassword) > 0 Then
        'sets the igWinStatus that we need.  Reruns gInitSuperUser (from gUrfRead) but that doesn't do anything because tlUrf.sName is blank from gUrfRead
        gExpandGuideAsUser tgUrf(0)
        sgUserName = "Guide"
        tgUrf(0).sName = sgUserName
    Else
       ' gUrfRead ResearchProjStartup, sgUserName, True, tgUrf(), False
        If Not blIsFromTraffic Then
            'is this user allowed to access traffic?
            If igWinStatus(RESEARCHLIST) < 1 Then
                gMsgBox "This user does not have access to Research screens", vbExclamation, "Access Denied"
                imTerminate = True
                igExitTraffic = True
                Exit Sub
            End If
        End If
    End If
    mGetUlfCode
    gParseItem slCommand, 3, "\", slStr    'get which report
    gParseItem slStr, 1, "/", slStr
    If Trim$(slStr) = "RschByCust" Then
        omResearchChoice = ByCust
    Else
        omResearchChoice = ResearchX
    End If
    DoEvents
    mInitStdAloneResearch
    mCheckForDate
End Sub
Private Sub mTestPervasive()
    Dim ilRet As Integer
    Dim ilRecLen As Integer
    Dim hlSpf As Integer
    Dim tlSpf As SPF

    gInitGlobalVar
    hlSpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlSpf, "", sgDBPath & "Spf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    ilRecLen = Len(tlSpf)
    ilRet = btrGetFirst(hlSpf, tlSpf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSpf
        'btrStopAppl
        hgDB = CBtrvMngrInit(0, sgMDBPath, sgSDBPath, sgTDBPath, igRetrievalDB, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
        Do While csiHandleValue(0, 3) = 0
            '7/6/11
            Sleep 1000
        Loop
        Exit Sub
    End If
    btrDestroy hlSpf
End Sub
Private Sub mCheckForDate()
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slDate As String
    Dim slSetDate As String
    Dim ilRet As Integer

    ilPos = InStr(1, sgCommandStr, "/D:", 1)
    If ilPos > 0 Then
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            slDate = Trim$(Mid$(sgCommandStr, ilPos + 3))
        Else
            slDate = Trim$(Mid$(sgCommandStr, ilPos + 3, ilSpace - ilPos - 3))
        End If
        If gValidDate(slDate) Then
            slDate = gAdjYear(slDate)
            slSetDate = slDate
        End If
    End If
    If Trim$(slSetDate) = "" Then
        If (InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) > 0) Or (InStr(1, tgSpf.sGClient, "XYZ Network", vbTextCompare) > 0) Then
            slSetDate = "12/15/1999"
            slDate = slSetDate
        End If
    End If
    If Trim$(slSetDate) <> "" Then
        ilRet = gCsiSetName(slDate)
    End If
End Sub
Private Sub mGetUlfCode()
    Dim ilPos As Integer
    Dim ilSpace As Integer

    ilPos = InStr(1, sgCommandStr, "/ULF:", 1)
    If ilPos > 0 Then
        ilSpace = InStr(ilPos, sgCommandStr, " ")
        If ilSpace = 0 Then
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5)))
        Else
            lgUlfCode = Val(Trim$(Mid$(sgCommandStr, ilPos + 5, ilSpace - ilPos - 3)))
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub tmcMain_Timer()
    If igExitTraffic Then
        tmcMain.Enabled = False
        Unload Me
        Exit Sub
    End If
End Sub
Private Sub mInitStdAloneResearch()
    'set sgusername before calling?
    
    Dim ilRet As Integer
    Dim slVehType As String
    Dim hlUrf As Integer        'User Option file handle
    Dim tlUrf As URF
    Dim ilRecLen As Integer

    If igStdAloneMode Then
        ilRet = csiSetAlloc("NAMES", 0, 2)
    End If
    sgSystemDate = gAdjYear(Format$(gNow(), "m/d/yy"))    'Used to reset date when exiting traffic
    igResetSystemDate = False
    ReDim tgJobHelp(0 To 0) As HLF
    ReDim tgListHelp(0 To 0) As HLF
    gSpfRead
    sgCPName = gGetCSIName("CPNAME")
    sgSUName = gGetCSIName("SUNAME")
    If (sgCPName = "") Or (sgSUName = "") Then
        hlUrf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet = BTRV_ERR_NONE Then
            ilRecLen = Len(tlUrf)  'btrRecordLength(hlUrf)  'Get and save record length
            ilRet = btrGetFirst(hlUrf, tlUrf, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                gUrfDecrypt tlUrf
                If tlUrf.iCode = 1 Then
                    sgCPName = Trim$(tlUrf.sName)
                End If
                ilRet = btrGetNext(hlUrf, tlUrf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    gUrfDecrypt tlUrf
                    If tlUrf.iCode = 2 Then
                        sgSUName = Trim$(tlUrf.sName)
                    End If
                End If
            End If
        End If
        ilRet = btrClose(hlUrf)
        btrDestroy hlUrf
    End If
    igUpdateAllowed = True
End Sub
Private Sub mGoToFormChoice()
    
    If omResearchChoice = ResearchX Then
        mPrepResearch
        Research.Show
    Else
        mPrepByCust
        RschByCust.Show
    End If
End Sub
Private Sub mPrepResearch()
    gObtainSAF
    gVpfRead
End Sub
Private Sub mPrepByCust()
    fgFlexGridRowH = 225
    fgPanelAdj = 90
    gObtainVef
End Sub
Private Function mDetermineCsiLogin() As String
    Dim slDate As String
    Dim slMonth As String
    Dim slYear As String
    Dim llValue As Long
    Dim ilValue As Integer
    Dim slRet As String
    
    slDate = Format$(Now(), "m/d/yy")
    slMonth = Month(slDate)
    slYear = Year(slDate)
    llValue = Val(slMonth) * Val(slYear)
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    llValue = ilValue
    ilValue = Int(10000 * Rnd(-llValue) + 1)
    slRet = Trim$(Str$(ilValue))
    Do While Len(slRet) < 4
        slRet = "0" & slRet
    Loop
    mDetermineCsiLogin = slRet
End Function
