Attribute VB_Name = "ResearchProject"
Option Explicit
Option Compare Text

Dim tmAdf As ADF
Dim tmAgf As AGF
Dim tmAnf As ANF
Dim tmAuf As AUF
Dim tmArf As ARF
Dim tmBof As BOF
Dim tmBsf As BSF
Dim tmBvf As BVF
Dim tmCaf As CAF
Dim tmCcf As CCF
Dim tmCdf As CDF
Dim tmCef As CEF
Dim tmCff As CFF
Dim tmCgf As CGF
Dim tmChf As CHF
Dim tmCif As CIF
Dim tmClf As CLF
Dim tmCmf As CMF
Dim tmCnf As CNF
Dim tmCof As COF
Dim tmCpf As CPF
Dim tmCrf As CRF
Dim tmCsf As CSF
'Dim tmCtf As CTF
Dim tmCxf As CXF
Dim tmCyf As CYF
'Dim tmDaf As DAF
Dim tmDef As DEF
Dim tmDlf As DLF
Dim tmDnf As DNF
Dim tmDrf As DRF
Dim tmDsf As DSF
'Dim tmElf As ELF
Dim tmEnf As ENF
Dim tmEtf As ETF
'Dim tmFsf As FSF
'Dim tmFxf As FXF
'Dim tmGmf As GMF
Dim tmGhf As GHF
Dim tmGsf As GSF
Dim tmLcf As LCF
Dim tmLef As LEF
'Dim tmLgf As LGF
'Dim tmLhf As LHF   'in XTRAFILE.BAS
Dim tmLtf As LTF
Dim tmLvf As LVF
Dim tmMcf As MCF
Dim tmMnf As MNF
Dim tmPhf As RVF    'PHF
Dim tmPjf As PJF
Dim tmPrf As PRF
Dim tmPvf As PVF
Dim tmRcf As RCF
Dim tmRdf As RDF
'Dim tmRgf As RGF
Dim tmRif As RIF
Dim tmRlf As RLF
'Dim tmRpf As RPF
Dim tmRvf As RVF
Dim tmSbf As SBF
Dim tmSdf As SDF
'Dim tmSff As SFF
'Dim tmShf As SHF
Dim tmSif As SIF
Dim tmSlf As SLF
Dim tmSmf As SMF
Dim tmSof As SOF
Dim tmSpf As SPF
'Dim tmSsf As SSF
Dim tmSwf As SWF
'Dim tmSvf As SVF
Dim tmTzf As TZF
Dim tmUrf As URF
Dim tmVcf As VCF
Dim tmVef As VEF
Dim tmVlf As VLF
Dim tmVpf As VPF
Dim tmVsf As VSF
'*******************************************************
'*                                                     *
'*      Procedure Name:gSetMenuState                   *
'*                                                     *
'*             Created:5/8/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Turn File and Window menu on or *
'*                     false as required               *
'*                     This sub should not be used.    *
'*                                                     *
'*******************************************************
'Sub gSetMenuState (ilMenuOn As Integer)
'
'   gSetMenuState ilOn
'   Where:
'       ilOn (I)- Indicates if menu should to turned on or off
'
    'If ilMenuOn Then
    '    igNoMenuDisabled = igNoMenuDisabled - 1
    '    If igNoMenuDisabled <= 0 Then
    '        Traffic!mnuFile.Enabled = True
    '        Traffic!mnuWnd.Enabled = True
    '    End If
    'Else
    '    igNoMenuDisabled = igNoMenuDisabled + 1
    '    If igNoMenuDisabled = 1 Then
    '        Traffic!mnuFile.Enabled = False
    '        Traffic!mnuWnd.Enabled = False
    '    End If
    'End If
'End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gCenterForm                     *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Center form within Traffic Form *
'*                                                     *
'*******************************************************
'Sub gCenterForm(FrmName As Form)
''
''   gCenterForm FrmName
''   Where:
''       FrmName (I)- Name of modeless form to be centered within Traffic form
''
'    Dim flLeft As Single
'    Dim flTop As Single
'    gSetPictureBoxFontSize FrmName, 9
'    flLeft = Traffic.Left + (Traffic.Width - Traffic.ScaleWidth) / 2 + (Traffic.ScaleWidth - FrmName.Width) / 2
''    flTop = Traffic.Top + (Traffic.Height - Traffic.ScaleHeight) + (Traffic.ScaleHeight - FrmName.Height) / 2
'    flTop = Traffic.Top + (Traffic.Height - 8 * Traffic.cmcTask(0).Height / 3 - FrmName.Height) / 2 - 60
'    FrmName.Move flLeft, flTop
'End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gCenterModalForm                *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Center modal form within        *
'*                     Traffic Form                    *
'*                                                     *
'*******************************************************
Sub gCenterModalForm(FrmName As Form)
'
'   gCenterModalForm FrmName
'   Where:
'       FrmName (I)- Name of modal form to be centered within Traffic form
'
'10066
'    Dim flLeft As Single
'    Dim flTop As Single
'    flLeft = Traffic.Left + (Traffic.Width - Traffic.ScaleWidth) / 2 + (Traffic.ScaleWidth - FrmName.Width) / 2
'    flTop = Traffic.Top + (Traffic.Height - FrmName.Height + 2 * Traffic.cmcTask(0).Height - 60) / 2 + Traffic.cmcTask(0).Height
'    FrmName.Move flLeft, flTop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gShowBranner                    *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Show branner in main title bar  *
'*                                                     *
'*******************************************************
Sub gShowBranner(ilUpdateAllowed As Integer)
'
'   gShowBranner
'   Where:
'
    Dim sAllowed As String
    Dim slName As String
    Dim slDateTime As String
    'Dim slDateTime As String

    'slDateTime = " " & Format$(gNow(), "ddd, m/d/yy h:mm AM/PM")
    'Remove spf test on 12/11/02- Jim
    'If tgSpf.sGTBar = "C" Then
    '    If ilUpdateAllowed Then
    '        sAllowed = sAllowed & ", Input OK"
    '    Else
    '        sAllowed = sAllowed & ", View Only"
    '    End If
    '    Traffic.Caption = Trim$(tgSpf.sGClient) & " on " & sgBrannerMsg & sAllowed '& " ** " & slDateTime & " **"
    'Else
        If ilUpdateAllowed Then
            sAllowed = sAllowed & ", Input OK"
        Else
            sAllowed = sAllowed & ", View Only"
        End If
        If Trim$(tgUrf(0).sRept) <> "" Then
            slName = Trim$(tgUrf(0).sRept)
        Else
            slName = sgUserName
        End If
        'removed for 10066
'        'Traffic.Caption = slName & " on " & sgBrannerMsg & sAllowed '& " ** " & slDateTime & " **"
'        sgDateBrannerMsg = slName & " on " & sgBrannerMsg & " for " & Trim$(tgSpf.sGClient) & sAllowed
'        slDateTime = " " & Format$(gNow(), "ddd, m/d/yy h:mm AM/PM")
'        Traffic.mnuDate.Caption = slDateTime & " " & sgDateBrannerMsg '"                                                   "
'        Traffic.Caption = "CSI Traffic" 'slName & " on " & sgBrannerMsg & " for " & Trim$(tgSpf.sGClient) & sAllowed '& " ** " & slDateTime & " **"
    'End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name: gShowHelpMess                  *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Show help message.              *
'*                     Use gHlfRead to load tlHlf.     *
'*                                                     *
'*******************************************************
Sub gShowHelpMess(tlHlf() As HLF, ilMsgNo As Integer)
'
'   gShowHelpMess tlHlf(), ilMsgNo
'   Where:
'   tlHlf (I)- Records of messages
'   ilMsgNo (I)- message number
'
    If ilMsgNo = -1 Then
        ''Traffic!plcHelp.Caption = " "
        'Traffic!plcHelp.CurrentX = 0
        'Traffic!plcHelp.CurrentY = 0
        'Traffic!plcHelp.Print "                          "
        'Traffic!plcHelp.Cls
        Exit Sub
    End If
    On Error GoTo gShowHelpMessErr
    If (ilMsgNo - 1 < LBound(tlHlf)) Or (ilMsgNo - 1 > UBound(tlHlf)) Then
        ''Traffic!plcHelp.Caption = " "
        'Traffic!plcHelp.CurrentX = 0
        'Traffic!plcHelp.CurrentY = 0
        'Traffic!plcHelp.Print "                          "
        'Traffic!plcHelp.Cls
    Else
        'Traffic!plcHelp.Caption = " " & tlHlf(ilMsgNo - 1).sMessage
        'Traffic!plcHelp.Cls
        'Traffic!plcHelp.CurrentX = 0
        'Traffic!plcHelp.CurrentY = 0
        'Traffic!plcHelp.Print " " & tlHlf(ilMsgNo - 1).sMessage
    End If
    Exit Sub
gShowHelpMessErr:
    ''Traffic!plcHelp.Caption = " "
    'Traffic!plcHelp.CurrentX = 0
    'Traffic!plcHelp.CurrentY = 0
    'Traffic!plcHelp.Print "                          "
    'Traffic!plcHelp.Cls
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gShowPartBranner                *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Show branner in main title bar  *
'*                                                     *
'*******************************************************
Sub gShowPartBranner()
'
'   gShowBranner
'   Where:
'
    Dim slName As String
    Dim slDateTime As String
    Dim ilErr As Integer
    ''Dim slDateTime As String

    ''slDateTime = " " & Format$(gNow(), "ddd, m/d/yy h:mm AM/PM")
    ilErr = False
    'If tgSpf.sGTBar = "C" Then
    '    Traffic.Caption = Trim$(tgSpf.sGClient) & " on " & sgBrannerMsg '& " ** " & slDateTime & " **"
    'Else
        On Error GoTo ShowErr
        If Not ilErr Then
            If Trim$(tgUrf(0).sRept) <> "" Then
                slName = Trim$(tgUrf(0).sRept)
            Else
                slName = sgUserName
            End If
     '       Traffic.Caption = slName & " on " & sgBrannerMsg '& " ** " & slDateTime & " **"
            sgDateBrannerMsg = slName & " on " & sgBrannerMsg & " for " & Trim$(tgSpf.sGClient)
            slDateTime = " " & Format$(gNow(), "ddd, m/d/yy h:mm AM/PM")
            'removed 10066
'            Traffic.mnuDate.Caption = slDateTime & " " & sgDateBrannerMsg '"                                                   "
'            Traffic.Caption = "CSI Traffic" 'slName & " on " & sgBrannerMsg & " for " & Trim$(tgSpf.sGClient)
        End If
        On Error GoTo 0
    'End If
    Exit Sub
ShowErr:
    ilErr = True
    Resume Next
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gTestRecLengths                 *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if record lengths match    *
'*                                                     *
'*******************************************************
Sub gTestRecLengths()
    Dim ilSize As Integer
    Dim ilValue As Integer

    ilSize = mGetRecLength("Adf.Btr")
    If ilSize <> Len(tmAdf) Then
        If ilSize > 0 Then
            MsgBox "Adf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmAdf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Adf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Agf.Btr")
    If ilSize <> Len(tmAgf) Then
        If ilSize > 0 Then
            MsgBox "Agf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmAgf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Agf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Anf.Btr")
    If ilSize <> Len(tmAnf) Then
        If ilSize > 0 Then
            MsgBox "Anf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmAnf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Anf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Arf.Btr")
    If ilSize <> Len(tmArf) Then
        If ilSize > 0 Then
            MsgBox "Arf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmArf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Arf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Auf.Btr")
    If ilSize <> Len(tmAuf) Then
        If ilSize > 0 Then
            MsgBox "Auf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmAuf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Auf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Bof.Btr")
        If ilSize <> Len(tmBof) Then
            If ilSize > 0 Then
                MsgBox "Bof size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmBof)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Bof error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Bsf.Btr")
        If ilSize <> Len(tmBsf) Then
            If ilSize > 0 Then
                MsgBox "Bsf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmBsf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Bsf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Bvf.Btr")
        If ilSize <> Len(tmBvf) Then
            If ilSize > 0 Then
                MsgBox "Bvf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmBvf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Bvf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Caf.Btr")
        If ilSize <> Len(tmCaf) Then
            If ilSize > 0 Then
                MsgBox "Caf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCaf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Caf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Ccf.Btr")
        If ilSize <> Len(tmCcf) Then
            If ilSize > 0 Then
                MsgBox "Ccf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCcf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Ccf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Cdf.Btr")
        If ilSize <> Len(tmCdf) Then    '- Len(tmCdf.iStrLen) - Len(tmCdf.sComment) Then
            If ilSize > 0 Then
                'MsgBox "Cdf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmCdf) - Len(tmCdf.iStrLen) - Len(tmCdf.sComment)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
                MsgBox "Cdf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCdf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Cdf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Cef.Btr")
        If ilSize <> Len(tmCef) Then    '- Len(tmCef.iStrLen) - Len(tmCef.sComment) Then
            If ilSize > 0 Then
                'MsgBox "Cef size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmCef) - Len(tmCef.iStrLen) - Len(tmCef.sComment)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
                MsgBox "Cef size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCef)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Cef error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilSize = mGetRecLength("Cff.Btr")
    If ilSize <> Len(tmCff) Then
        If ilSize > 0 Then
            MsgBox "Cff size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCff)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Cff error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Chf.Btr")
    If ilSize <> Len(tmChf) Then
        If ilSize > 0 Then
            MsgBox "Chf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmChf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Chf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Cif.Btr")
        If ilSize <> Len(tmCif) Then
            If ilSize > 0 Then
                MsgBox "Cif size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCif)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Cif error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilSize = mGetRecLength("Clf.Btr")
    If ilSize <> Len(tmClf) Then
        If ilSize > 0 Then
            MsgBox "Clf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmClf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Clf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Cmf.Btr")
    If ilSize <> Len(tmCmf) Then    '- Len(tmCmf.iStrLen) - Len(tmCmf.sComment) Then
        If ilSize > 0 Then
            'MsgBox "Cmf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmCmf) - Len(tmCmf.iStrLen) - Len(tmCmf.sComment)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            MsgBox "Cmf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCmf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Cmf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Cnf.Btr")
        If ilSize <> Len(tmCnf) Then
            If ilSize > 0 Then
                MsgBox "Cnf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCnf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Cnf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilSize = mGetRecLength("Cof.Btr")
    If ilSize <> Len(tmCof) Then
        If ilSize > 0 Then
            MsgBox "Cof size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCof)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Cof error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Cpf.Btr")
        If ilSize <> Len(tmCpf) Then
            If ilSize > 0 Then
                MsgBox "Cpf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCpf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Cpf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Crf.Btr")
        If ilSize <> Len(tmCrf) Then
            If ilSize > 0 Then
                MsgBox "Crf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCrf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Crf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Csf.Btr")
        If ilSize <> Len(tmCsf) Then    '- Len(tmCsf.iStrLen) - Len(tmCsf.sComment) Then
            If ilSize > 0 Then
                'MsgBox "Csf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmCsf) - Len(tmCsf.iStrLen) - Len(tmCsf.sComment)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
                MsgBox "Csf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCsf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Csf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    'ilSize = mGetRecLength("Ctf.Btr")
    'If ilSize <> Len(tmCtf) Then
    '    If ilSize > 0 Then
    '        MsgBox "Ctf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCtf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
    '    Else
    '        MsgBox "Ctf error: " & Str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "File Error"
    '    End If
    'End If
    ilValue = Asc(tgSpf.sSportInfo)
    If (ilValue And USINGSPORTS) = USINGSPORTS Then
        ilSize = mGetRecLength("Cgf.Btr")
        If ilSize <> Len(tmCgf) Then
            If ilSize > 0 Then
                MsgBox "Cgf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCgf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Cgf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilSize = mGetRecLength("Cxf.Btr")
    If ilSize <> Len(tmCxf) Then    '- Len(tmCxf.iStrLen) - Len(tmCxf.sComment) Then
        If ilSize > 0 Then
            'MsgBox "Cxf size error: Btrieve Size" & str$(ilSize) & " Internal size" & str$(Len(tmCxf) - Len(tmCxf.iStrLen) - Len(tmCxf.sComment)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
            MsgBox "Cxf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCxf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Cxf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Cyf.Btr")
        If ilSize <> Len(tmCyf) Then
            If ilSize > 0 Then
                MsgBox "Cyf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmCyf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Cyf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    'ilSize = mGetRecLength("Daf.Btr")
    'If ilSize <> Len(tmDaf) Then
    '    If ilSize > 0 Then
    '        MsgBox "Daf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmDaf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
    '    Else
    '        MsgBox "Daf error: " & Str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "File Error"
    '    End If
    'End If
    ilSize = mGetRecLength("Def.Btr")
    If ilSize <> Len(tmDef) Then
        If ilSize > 0 Then
            MsgBox "Def size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmDef)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Def error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Dlf.Btr")
        If ilSize <> Len(tmDlf) Then
            If ilSize > 0 Then
                MsgBox "Dlf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmDlf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Dlf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilSize = mGetRecLength("Dnf.Btr")
    If ilSize <> Len(tmDnf) Then
        If ilSize > 0 Then
            MsgBox "Dnf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmDnf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Dnf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Drf.Btr")
    If ilSize <> Len(tmDrf) Then
        If ilSize > 0 Then
            MsgBox "Drf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmDrf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Drf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Dsf.Btr")
    If ilSize <> Len(tmDsf) Then
        If ilSize > 0 Then
            MsgBox "Dsf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmDsf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Dsf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    'If mGetRecLength("Elf.Btr") <> Len(tmElf) Then
    '    MsgBox "Elf size error: Btrieve Size" & Str$(mGetRecLength("Elf.Btr")) & " Internal size" & Str$(Len(tmElf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
    'End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Enf.Btr")
        If ilSize <> Len(tmEnf) Then
            If ilSize > 0 Then
                MsgBox "Enf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmEnf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Enf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Etf.Btr")
        If ilSize <> Len(tmEtf) Then
            If ilSize > 0 Then
                MsgBox "Etf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmEtf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Etf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    'If mGetRecLength("Fsf.Btr") <> Len(tmFsf) Then
    '    MsgBox "Fsf size error: Btrieve Size" & Str$(mGetRecLength("Fsf.Btr")) & " Internal size" & Str$(Len(tmFsf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
    'End If
    'If mGetRecLength("Fxf.Btr") <> Len(tmFxf) Then
    '    MsgBox "Fxf size error: Btrieve Size" & Str$(mGetRecLength("Fxf.Btr")) & " Internal size" & Str$(Len(tmFxf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
    'End If
    'If mGetRecLength("Gmf.Btr") <> Len(tmGmf) Then
    '    MsgBox "Gmf size error: Btrieve Size" & Str$(mGetRecLength("Gmf.Btr")) & " Internal size" & Str$(Len(tmGmf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
    'End If
    ilValue = Asc(tgSpf.sSportInfo)
    If (ilValue And USINGSPORTS) = USINGSPORTS Then
        ilSize = mGetRecLength("Ghf.Btr")
        If ilSize <> Len(tmGhf) Then
            If ilSize > 0 Then
                MsgBox "Ghf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmGhf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Ghf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilValue = Asc(tgSpf.sSportInfo)
    If (ilValue And USINGSPORTS) = USINGSPORTS Then
        ilSize = mGetRecLength("Gsf.Btr")
        If ilSize <> Len(tmGsf) Then
            If ilSize > 0 Then
                MsgBox "Gsf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmGsf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Gsf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If

    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Lcf.Btr")
        If ilSize <> Len(tmLcf) Then
            If ilSize > 0 Then
                MsgBox "Lcf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmLcf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Lcf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Lef.Btr")
        If ilSize <> Len(tmLef) Then
            If ilSize > 0 Then
                MsgBox "Lef size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmLef)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Lef error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    'If mGetRecLength("Lgf.Btr") <> Len(tmLgf) Then
    '    MsgBox "Lgf size error: Btrieve Size" & Str$(mGetRecLength("Lgf.Btr")) & " Internal size" & Str$(Len(tmLgf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
    'End If
    'If mGetRecLength("Lhf.Btr") <> Len(tmLhf) Then
    '    MsgBox "Lhf size error: Btrieve Size" & Str$(mGetRecLength("Lhf.Btr")) & " Internal size" & Str$(Len(tmLhf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
    'End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Ltf.Btr")
        If ilSize <> Len(tmLtf) Then
            If ilSize > 0 Then
                MsgBox "Ltf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmLtf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Ltf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Lvf.Btr")
        If ilSize <> Len(tmLvf) Then
            If ilSize > 0 Then
                MsgBox "Lvf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmLvf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Lvf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Mcf.Btr")
        If ilSize <> Len(tmMcf) Then
            If ilSize > 0 Then
                MsgBox "Mcf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmMcf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Mcf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilSize = mGetRecLength("Mnf.Btr")
    If ilSize <> Len(tmMnf) Then
        If ilSize > 0 Then
            MsgBox "Mnf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmMnf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Mnf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Phf.Btr")
        If ilSize <> Len(tmPhf) Then
            If ilSize > 0 Then
                MsgBox "Phf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmPhf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Phf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilSize = mGetRecLength("Pjf.Btr")
    If ilSize <> Len(tmPjf) Then
        If ilSize > 0 Then
            MsgBox "Pjf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmPjf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Pjf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Prf.Btr")
    If ilSize <> Len(tmPrf) Then
        If ilSize > 0 Then
            MsgBox "Prf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmPrf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Prf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Pvf.Btr")
    If ilSize <> Len(tmPvf) Then
        If ilSize > 0 Then
            MsgBox "Pvf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmPvf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Pvf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Rcf.Btr")
    If ilSize <> Len(tmRcf) Then
        If ilSize > 0 Then
            MsgBox "Rcf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmRcf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Rcf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Rdf.Btr")
    If ilSize <> Len(tmRdf) Then
        If ilSize > 0 Then
            MsgBox "Rdf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmRdf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Rdf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    'ilSize = mGetRecLength("Rgf.Btr")
    'If ilSize <> Len(tmRgf) Then
    '    If ilSize > 0 Then
    '        MsgBox "Rgf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmRgf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
    '    Else
    '        MsgBox "Rgf error: " & Str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "File Error"
    '    End If
    'End If
    ilSize = mGetRecLength("Rif.Btr")
    If ilSize <> Len(tmRif) Then
        If ilSize > 0 Then
            MsgBox "Rif size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmRif)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Rif error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Rlf.Btr")
    If ilSize <> Len(tmRlf) Then
        If ilSize > 0 Then
            MsgBox "Rlf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmRlf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Rlf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    'ilSize = mGetRecLength("Rpf.Btr")
    'If ilSize <> Len(tmRpf) Then
    '    If ilSize > 0 Then
    '        MsgBox "Rpf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmRpf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
    '    Else
    '        MsgBox "Rpf error: " & Str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "File Error"
    '    End If
    'End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Rvf.Btr")
        If ilSize <> Len(tmRvf) Then
            If ilSize > 0 Then
                MsgBox "Rvf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmRvf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Rvf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilSize = mGetRecLength("Sbf.Btr")
    If ilSize <> Len(tmSbf) Then
        If ilSize > 0 Then
            MsgBox "Sbf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmSbf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Sbf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Sdf.Btr")
        If ilSize <> Len(tmSdf) Then
            If ilSize > 0 Then
                MsgBox "Sdf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmSdf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Sdf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    'If mGetRecLength("Sff.Btr") <> Len(tmSff) Then
    '    MsgBox "Sff size error: Btrieve Size" & Str$(mGetRecLength("Sff.Btr")) & " Internal size" & Str$(Len(tmSff)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
    'End If
    ilSize = mGetRecLength("Sif.Btr")
    If ilSize <> Len(tmSif) Then
        If ilSize > 0 Then
            MsgBox "Sif size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmSif)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Sif error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        End If
    End If
    ilSize = mGetRecLength("Slf.Btr")
    If ilSize <> Len(tmSlf) Then
        If ilSize > 0 Then
            MsgBox "Slf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmSlf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Slf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Smf.Btr")
        If ilSize <> Len(tmSmf) Then
            If ilSize > 0 Then
                MsgBox "Smf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmSmf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Smf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilSize = mGetRecLength("Sof.Btr")
    If ilSize <> Len(tmSof) Then
        If ilSize > 0 Then
            MsgBox "Sof size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmSof)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Sof error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Spf.Btr")
    If ilSize <> Len(tmSpf) Then
        If ilSize > 0 Then
            MsgBox "Spf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmSpf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Spf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        'ilSize = mGetRecLength("Ssf.Btr")
        'If ilSize <> Len(tmSsf) - UBound(tmSsf.tPAS) * Len(tmSsf.tPas(ADJSSFPASBZ + 1)) Then
        '    If ilSize > 0 Then
        '        MsgBox "Ssf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmSsf) - UBound(tmSsf.tPAS) * Len(tmSsf.tPas(ADJSSFPASBZ + 1))), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
        '    Else
        '        MsgBox "Ssf error: " & Str$(-ilSize), vbOkOnly + vbCritical + vbApplicationModal, "File Error"
        '    End If
        'End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Swf.Btr")
        If ilSize <> Len(tmSwf) Then
            If ilSize > 0 Then
                MsgBox "Swf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmSwf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Swf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    'If mGetRecLength("Svf.Btr") <> Len(tmSvf) Then
    '    MsgBox "Svf size error: Btrieve Size" & Str$(mGetRecLength("Svf.Btr")) & " Internal size" & Str$(Len(tmSvf)), vbOkOnly + vbCritical + vbApplicationModal, "Size Error"
    'End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Tzf.Btr")
        If ilSize <> Len(tmTzf) Then
            If ilSize > 0 Then
                MsgBox "Tzf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmTzf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Tzf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilSize = mGetRecLength("Urf.Btr")
    If ilSize <> Len(tmUrf) Then
        If ilSize > 0 Then
            MsgBox "Urf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmUrf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Urf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Vcf.Btr")
        If ilSize <> Len(tmVcf) Then
            If ilSize > 0 Then
                MsgBox "Vcf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmVcf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Vcf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilSize = mGetRecLength("Vef.Btr")
    If ilSize <> Len(tmVef) Then
        If ilSize > 0 Then
            MsgBox "Vef size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmVef)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Vef error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    If tgUrf(0).iRemoteUserID <= 0 Then
        ilSize = mGetRecLength("Vlf.Btr")
        If ilSize <> Len(tmVlf) Then
            If ilSize > 0 Then
                MsgBox "Vlf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmVlf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
            Else
                MsgBox "Vlf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
            End If
        End If
    End If
    ilSize = mGetRecLength("Vpf.Btr")
    If ilSize <> Len(tmVpf) Then
        If ilSize > 0 Then
            MsgBox "Vpf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmVpf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Vpf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
    ilSize = mGetRecLength("Vsf.Btr")
    If ilSize <> Len(tmVsf) Then
        If ilSize > 0 Then
            MsgBox "Vsf size error: Btrieve Size" & Str$(ilSize) & " Internal size" & Str$(Len(tmVsf)), vbOKOnly + vbCritical + vbApplicationModal, "Size Error"
        Else
            MsgBox "Vsf error: " & Str$(-ilSize), vbOKOnly + vbCritical + vbApplicationModal, "File Error"
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gGetRecLength                   *
'*                                                     *
'*             Created:10/09/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the record length from   *
'*                     the database                    *
'*                                                     *
'*******************************************************
Private Function mGetRecLength(slFileName As String) As Integer
'
'   ilRecLen = mGetRecLength(slName)
'   Where:
'       slName (I)- Name of the file
'       ilRecLen (O)- record length within the file
'
    Dim hlFile As Integer
    Dim ilRet As Integer
    hlFile = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mGetRecLength = -ilRet
        ilRet = btrClose(hlFile)
        btrDestroy hlFile
        Exit Function
    End If
    mGetRecLength = btrRecordLength(hlFile)  'Get and save record length
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
End Function
'9459 I need a new form, but am reusing 'GenMsg' added radio buttons
Public Function gShowGenMsgWithRadioButtons(slMessage As String, ilDefaultRadioButton As Integer, ParamArray slRadioButtons() As Variant) As Integer
    'I:slButtons: each radio button's name. Must be less then 5
    'O: selected button's index  -1 means there was an issue
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim blIsEdit As Boolean
    Dim ilMaxButtons As Integer
    
    ilRet = -1
    ilMaxButtons = 4
    bgUseRadioButtons = True
    sgGenMsg = slMessage
    For ilIndex = 0 To ilMaxButtons - 1
        If UBound(slRadioButtons) >= ilIndex Then
            sgCMCTitle(ilIndex) = slRadioButtons(ilIndex)
        Else
            sgCMCTitle(ilIndex) = ""
        End If
    Next ilIndex
    igDefCMC = ilDefaultRadioButton
    GenMsg.Show vbModal
    ilRet = igAnsCMC
    For ilIndex = 0 To 3
        sgCMCTitle(ilIndex) = ""
    Next ilIndex
    bgUseRadioButtons = False
    gShowGenMsgWithRadioButtons = ilRet
End Function






