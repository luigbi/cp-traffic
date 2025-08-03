Attribute VB_Name = "ContractModules"

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ContractModules.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the extra XML routine definitions for contracts
Option Explicit
Dim tmChfSrchKey As LONGKEY0  'CHF key record image
Dim tmEffSrchKey As LONGKEY0
Dim tmAdfSrchKey As INTKEY0 'ADF key record image
Dim tmAgfSrchKey As INTKEY0 'AGF key record image
Dim tmVsfSrchKey As LONGKEY0



Public Function mSetBuyerEmail(ilPnfCode As Integer, hlPnf As Integer) As Boolean
    Dim ilRet As Integer
    Dim tlPnf As PNF
    Dim ilRecLen As Integer
    Dim llCefCode As Long
    
    llCefCode = ogEmailer.SetEmailAddress(ogEmailer.ToAddress)
    If llCefCode > 0 Then
        ilRecLen = Len(tlPnf)
        ilRet = btrGetEqual(hlPnf, tlPnf, ilRecLen, ilPnfCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            tlPnf.lEMailCefCode = llCefCode
            ilRet = btrUpdate(hlPnf, tlPnf, ilRecLen)
            If ilRet = BTRV_ERR_NONE Then
                mSetBuyerEmail = True
            Else
                mSetBuyerEmail = False
            End If
        Else
            mSetBuyerEmail = False
        End If
    End If

End Function


Public Function mEDSSetNeg2(ilVefCode As Integer, slLineType As String, bmAnyVehicleOnInsertion As Boolean) As Boolean
    Dim ilVff As Integer
    
    If (Asc(tgSaf(0).sFeatures2) And EMAILDISTRIBUTION) <> EMAILDISTRIBUTION Then
        mEDSSetNeg2 = True
        Exit Function
    End If
    If Not bmAnyVehicleOnInsertion Then
        mEDSSetNeg2 = True
        Exit Function
    End If
    If ilVefCode <= 0 Then
        mEDSSetNeg2 = False
        Exit Function
    End If
        
    If (slLineType = "A") Or (slLineType = "O") Or (slLineType = "E") Then
        mEDSSetNeg2 = True
        Exit Function
    End If
    
    ilVff = gBinarySearchVff(ilVefCode)
    If ilVff = -1 Then
        mEDSSetNeg2 = True
        Exit Function
    End If
    If tgVff(ilVff).sOnInsertions = "N" Then
        mEDSSetNeg2 = True
        Exit Function
    Else
        mEDSSetNeg2 = False
        Exit Function
    End If
End Function

Public Function mEDSSetNeg1(iEDSSentExtRevNo As Integer, ilVefCode As Integer, slLineType As String, bmAnyVehicleOnInsertion As Boolean) As Boolean
    Dim ilVff As Integer
    
    If iEDSSentExtRevNo <> -2 Then
        mEDSSetNeg1 = False
        Exit Function
    End If
    If (Asc(tgSaf(0).sFeatures2) And EMAILDISTRIBUTION) <> EMAILDISTRIBUTION Then
        mEDSSetNeg1 = False
        Exit Function
    End If
    If Not bmAnyVehicleOnInsertion Then
        mEDSSetNeg1 = False
        Exit Function
    End If
    If ilVefCode <= 0 Then
        mEDSSetNeg1 = True
        Exit Function
    End If
        
    If (slLineType = "A") Or (slLineType = "O") Or (slLineType = "E") Then
        mEDSSetNeg1 = False
        Exit Function
    End If
    
    ilVff = gBinarySearchVff(ilVefCode)
    If ilVff = -1 Then
        mEDSSetNeg1 = False
        Exit Function
    End If
    If tgVff(ilVff).sOnInsertions = "N" Then
        mEDSSetNeg1 = False
        Exit Function
    Else
        mEDSSetNeg1 = True
        Exit Function
    End If
End Function

Public Function mGetStatusForEmail() As String
    '6552 status not as simple as it appears
    'returns "Proposal to " and then "hold" or "order" , or just "Proposal"
    Dim slStatus As String
    Dim slStatusInitial As String
    Dim slDefaultStatus As String
    
    slDefaultStatus = "Proposal"
    slStatusInitial = tgChfCntr.sStatus
    'Dick says:  H is hold, O is order.  Otherwise proposal
    Select Case slStatusInitial
        Case "H", "G"
            slStatus = "Proposal to Hold"
        Case "N", "O"
            slStatus = "Proposal to Order"
        Case Else
            slStatus = slDefaultStatus
    End Select
    mGetStatusForEmail = slStatus
End Function

Public Function mEmailProposal(slContract As String, slAdv As String, hmPnf As Integer, tmAgf As AGF) As Boolean

'tmagf.ibuyerpnf, pnfcode, pnfname, pnfemailcefcode,  cefcode, cefComment
    Dim slToAddress As String
    Dim slBuyerName As String
    Dim ilRet As VbMsgBoxResult
    Dim slString As String
    '6552
    Dim slStatus As String
    Set ogEmailer = New CEmail
    '5719 'don't need slbuyername anymore.
    If tgChfCntr.iPnfBuyer > 0 Then
        slToAddress = mGetBuyerEmail(tgChfCntr.iPnfBuyer, slBuyerName, hmPnf)
    Else
        slToAddress = mGetBuyerEmail(tgChfCntr.iPnfBuyer, slBuyerName, hmPnf)
    End If
'    slToAddress = mGetBuyerEmail(tmAgf.iPnfBuyer, slBuyerName)
    With ogEmailer
        .ToName = slBuyerName
        .ToAddress = slToAddress
        .Message = " "
        If Len(slContract) > 0 Then
            slStatus = mGetStatusForEmail()
            'slString = "Proposal " & slContract
            '6552
            slString = slStatus & " " & slContract
            If Len(slAdv) > 0 Then
                slString = slString & " for " & slAdv
            End If
        End If
        .Subject = slString
    End With
    'dan M 10/25/11 moved from emailGeneric
    With EmailProposal
        '6189
        If tmAgf.iCode = 0 Then
            .AgencyCode = 0
            'email will change to X later
            .AgencyOutput = "Z"
        Else
            .AgencyCode = tmAgf.iCode
            .AgencyOutput = tmAgf.sCntrExptForm
        End If
'        .AgencyCode = tmAgf.iCode
'        .AgencyOutput = tmAgf.sCntrExptForm
        .Show vbModal
    End With
    mEmailProposal = ogEmailer.Success
    If StrComp(slToAddress, ogEmailer.ToAddress, vbTextCompare) <> 0 Then
        'note: IF site not set, to addresses are cleared, so will not fall into here.
        '6236 block if not agency
        If tmAgf.iPnfBuyer <> 0 Then
            If StrComp(slBuyerName, ogEmailer.ToName, vbTextCompare) = 0 Then
                DoEvents
                ilRet = MsgBox("The email was changed for the buyer.  Do you wish to save this email?", vbYesNo, "Save Buyer's email?")
                If ilRet = vbYes Then
                    DoEvents
                    If mSetBuyerEmail(tmAgf.iPnfBuyer, hmPnf) Then
                         MsgBox "Address saved.", vbOKOnly, "Buyer email address"
                     Else
                         MsgBox "Address was not saved.", vbInformation + vbOKOnly, "Buyer email address"
                    End If
                End If  'yes, please change
            End If  'name same?
        End If  'not an agency
    End If
    Set ogEmailer = Nothing
End Function

Public Function mHdCommentBranch(ilIndex As Integer, imInCommentBranch As Integer, imDoubleClickName As Integer, imBSMode As Integer, imComboBoxIndex As Integer, imUpdateAllowed As Integer, smComment() As String) As Integer
    If imInCommentBranch Then
        imInCommentBranch = False
        imDoubleClickName = False
        mHdCommentBranch = False
        Exit Function
    End If
    gMatchLookAhead Contract.edcHdDropDown, Contract.lbcComment(0), imBSMode, imComboBoxIndex
    If (Contract.lbcComment(0).ListIndex <= 0) Then
        Contract.edcComment.Text = ""
        smComment(ilIndex) = ""
        imDoubleClickName = False
        mHdCommentBranch = False
        Exit Function
    End If
    If (Not imDoubleClickName) And (smComment(ilIndex) <> "") Then
        imDoubleClickName = False
        mHdCommentBranch = False
        Exit Function
    End If
    imDoubleClickName = False
    Contract.mBPlatePop
    imInCommentBranch = True
    Contract.plcComment.Visible = True
    '2/16/09:  aLLOW COMMENTS
    'If (Not imUpdateAllowed) Or (igRepNetChgAllowed) Then
    If (Not imUpdateAllowed) Then
        Contract.plcComment.Enabled = False
    Else
        Contract.plcComment.Enabled = True
        Contract.edcComment.SetFocus
    End If
    mHdCommentBranch = True
End Function


Public Function mRevSetBranch(ilIndex As Integer, imDoubleClickName As Integer, imBSMode As Integer, imUpdateAllowed As Integer, smRevSetCodeTag() As String, imRSBoxNo As Integer) As Integer
'
'   ilRet = mRevSetBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    'Dim ilNoForms As Integer
    'Dim ilUsage As Integer
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    'Dim ilFound As Integer
    If (igRepNetChgAllowed) Then
        imDoubleClickName = False
        mRevSetBranch = False
        Exit Function
    End If
    ilRet = gOptionalLookAhead(Contract.edcShareDropDown, Contract.lbcRevSet(ilIndex), imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (Contract.edcShareDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mRevSetBranch = False
        Exit Function
    End If
    If igWinStatus(REVENUESETSLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mRevSetBranch = True
        Contract.mRSSetFocus imRSBoxNo
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(REVENUESETSLIST)) Then
    '    imDoubleClickName = False
    '    mRevSetBranch = True
    '    mRSEnableBox imRSBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourglass  'Wait
    sgMnfCallType = "R"
    igMNmCallSource = CALLSOURCECONTRACT
    If Contract.edcShareDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'Traffic!edcLinkSrceHelpMsg.Text = ""
    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
    End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Traffic.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Traffic.Enabled = True
    'Traffic!edcLinkSrceHelpMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mRevSetBranch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        Contract.lbcRevSet(ilIndex).Clear
        'Traffic!lbcRevSetCode(ilIndex).Tag = ""
        smRevSetCodeTag(ilIndex) = ""
        sgRevSetStamp = ""
        mRevSetPop 1
        If imTerminate Then
            mRevSetBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, Contract.lbcRevSet(ilIndex)
        If gLastFound(Contract.lbcRevSet(ilIndex)) > 0 Then
            imChgMode = True
            Contract.lbcRevSet(ilIndex).ListIndex = gLastFound(Contract.lbcRevSet(ilIndex))
            Contract.edcShareDropDown.Text = Contract.lbcRevSet(ilIndex).List(Contract.lbcRevSet(ilIndex).ListIndex)
            imChgMode = False
            mRevSetBranch = False
            Contract.mRSSetChg imRSBoxNo
        Else
            imChgMode = True
            Contract.lbcRevSet(ilIndex).ListIndex = 1
            Contract.edcShareDropDown.Text = Contract.lbcRevSet(ilIndex).List(1)
            imChgMode = False
            Contract.mRSSetChg imRSBoxNo
            Contract.edcShareDropDown.SetFocus
            sgMNmName = ""
            Exit Function
        End If
        sgMNmName = ""
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        Contract.mRSEnableBox imRSBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        Contract.mRSEnableBox imRSBoxNo
        Exit Function
    End If
    Exit Function
End Function

Public Function mShareBudgetBranch(ilIndex As Integer, imDoubleClickName As Integer, imBSMode As Integer, imSBBoxNo As Integer, imUpdateAllowed As Integer, smShareBudgetCodeTag() As String) As Integer
'
'   ilRet = mShareBudgetBranch()
'   Where:
'       ilRet (O)- True = Stay on control
'                  False = go to next control
'
    Dim ilRet As Integer
    Dim slStr As String
    'Dim ilNoForms As Integer
    'Dim ilUsage As Integer
    Dim ilParse As Integer
    Dim ilUpdateAllowed As Integer
    'Dim ilFound As Integer
    If (igRepNetChgAllowed) Then
        imDoubleClickName = False
        mShareBudgetBranch = False
        Exit Function
    End If
    ilRet = gOptionalLookAhead(Contract.edcShareDropDown, Contract.lbcShareBudget(ilIndex), imBSMode, slStr)
    If ((ilRet = 0) And (Not imDoubleClickName)) Or (Contract.edcShareDropDown.Text = "[None]") Then
        imDoubleClickName = False
        mShareBudgetBranch = False
        Exit Function
    End If
    If igWinStatus(COMPETITORSLIST) <> 2 Then
        Beep
        imDoubleClickName = False
        mShareBudgetBranch = True
        Contract.mSBSetFocus imSBBoxNo
        Exit Function
    End If
    'Unload IconTraf
    'If Not gWinRoom(igNoLJWinRes(REVENUESETSLIST)) Then
    '    imDoubleClickName = False
    '    mShareBudgetBranch = True
    '    mSBEnableBox imSBBoxNo
    '    Exit Function
    'End If
    'Screen.MousePointer = vbHourglass  'Wait
    sgMnfCallType = "O"
    igMNmCallSource = CALLSOURCECONTRACT
    If Contract.edcShareDropDown.Text = "[New]" Then
        sgMNmName = ""
    Else
        sgMNmName = slStr
    End If
    ilUpdateAllowed = imUpdateAllowed
    'igChildDone = False 'edcLinkDestDoneMsg.Text = ""
    'Traffic!edcLinkSrceHelpMsg.Text = ""
    If igTestSystem Then
        slStr = "Traffic^Test\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
    Else
        slStr = "Traffic^Prod\" & sgUserName & "\" & sgMnfCallType & "\" & Trim$(str$(igMNmCallSource)) & "\" & sgMNmName
    End If
    'lgShellRet = Shell(sgExePath & "MultiNm.Exe " & slStr, 1)
    'Traffic.Enabled = False
    'Screen.MousePointer = vbDefault  'Wait
    'Do While Not igChildDone
    '    DoEvents
    'Loop
    sgCommandStr = slStr
    MultiNm.Show vbModal
    slStr = sgDoneMsg
    ilParse = gParseItem(slStr, 1, "\", sgMNmName)
    igMNmCallSource = Val(sgMNmName)
    ilParse = gParseItem(slStr, 2, "\", sgMNmName)
    'Traffic.Enabled = True
    'Traffic!edcLinkSrceHelpMsg.Text = "Ok"
    'For ilLoop = 0 To 10
    '    DoEvents
    'Next ilLoop
    'Screen.MousePointer = vbDefault    'Default
    imDoubleClickName = False
    mShareBudgetBranch = True
    imUpdateAllowed = ilUpdateAllowed
    gShowBranner imUpdateAllowed
    If igMNmCallSource = CALLDONE Then  'Done
        igMNmCallSource = CALLNONE
'        gSetMenuState True
        sgShareBudgetStamp = ""
        Contract.lbcShareBudget(ilIndex).Clear
        'Traffic!lbcShareBudgetCode(ilIndex).Tag = ""
        If ilIndex = 0 Then
            smShareBudgetCodeTag(0) = ""
        Else
            smShareBudgetCodeTag(1) = ""
        End If
        Contract.mShareBudgetPop
        If imTerminate Then
            mShareBudgetBranch = False
            Exit Function
        End If
        gFindMatch sgMNmName, 1, Contract.lbcShareBudget(ilIndex)
        If gLastFound(Contract.lbcShareBudget(ilIndex)) > 0 Then
            imChgMode = True
            Contract.lbcShareBudget(ilIndex).ListIndex = gLastFound(Contract.lbcShareBudget(ilIndex))
            Contract.edcShareDropDown.Text = Contract.lbcShareBudget(ilIndex).List(Contract.lbcShareBudget(ilIndex).ListIndex)
            imChgMode = False
            mShareBudgetBranch = False
            Contract.mSBSetChg imSBBoxNo
        Else
            imChgMode = True
            Contract.lbcShareBudget(ilIndex).ListIndex = 1
            Contract.edcShareDropDown.Text = Contract.lbcShareBudget(ilIndex).List(1)
            imChgMode = False
            Contract.mSBSetChg imSBBoxNo
            Contract.edcShareDropDown.SetFocus
            sgMNmName = ""
            Exit Function
        End If
        sgMNmName = ""
    End If
    If igMNmCallSource = CALLCANCELLED Then  'Cancelled
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        Contract.mSBEnableBox imSBBoxNo
        Exit Function
    End If
    If igMNmCallSource = CALLTERMINATED Then
'        gSetMenuState True
        igMNmCallSource = CALLNONE
        sgMNmName = ""
        Contract.mSBEnableBox imSBBoxNo
        Exit Function
    End If
    Exit Function
End Function
'12/5/14: Disallow Installment if Billing Cycle other then Standadrd selected
Public Function mCheckInstallment() As Integer
    '0=Not using Installemnt; 1=Installemnt defined but wrong billing cycle; 2=Installment defined and is Ok
    Dim ilLoop As Integer
    Dim ilClf As Integer
    Dim ilVef As Integer
    Dim slStr As String
    
    mCheckInstallment = 0
    If (Asc(tgSpf.sUsingFeatures6) And INSTALLMENT) <> INSTALLMENT Then
        Exit Function
    End If
    If Contract.lbcBillCycle.ListIndex < 0 Then
        slStr = ""
    Else
        slStr = Contract.lbcBillCycle.List(Contract.lbcBillCycle.ListIndex)
    End If
    
    'If Contract.lbcBillCycle.ListIndex <= 0 Then
    If (slStr = "") Or (Left(slStr, 1) = "S") Then
        For ilLoop = LBound(tgFBSbf) To UBound(tgFBSbf) - 1 Step 1
            If (tgFBSbf(ilLoop).iStatus >= 0) And (tgFBSbf(ilLoop).iStatus <= 1) Then
                mCheckInstallment = 2
                Exit For
            End If
        Next ilLoop
        Exit Function
    End If
    For ilLoop = LBound(tgFBSbf) To UBound(tgFBSbf) - 1 Step 1
        If (tgFBSbf(ilLoop).iStatus >= 0) And (tgFBSbf(ilLoop).iStatus <= 1) Then
            Screen.MousePointer = vbDefault    'Default
            MsgBox "Installment disallowed when Standard not selected as the Billing Cycle", vbOKOnly + vbExclamation, "Warning"
            mCheckInstallment = 1
            Exit Function
        End If
    Next ilLoop
End Function

'12/16/14: Add Hide/Show Hidden Lines
Public Function mRemapRowNo(ilPaintRow As Integer) As Integer
    Dim ilRow As Integer
    'add test if using feature
    If UBound(tgMapLinePaint) <= 0 Then
        mRemapRowNo = ilPaintRow
        Exit Function
    End If
    For ilRow = 0 To UBound(tgMapLinePaint) - 1 Step 1
        If (tgMapLinePaint(ilRow).lPaintLnRowNo = ilPaintRow) Then
            mRemapRowNo = tgMapLinePaint(ilRow).lMapLnRowNo
            Exit Function
        End If
    Next ilRow
    mRemapRowNo = UBound(tgClfCntr) + 1
End Function

Public Sub mClearAvailAud(ilLnRowNo As Integer, ilClearAvail As Integer)
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    If ilLnRowNo <= 0 Then
        Exit Sub
    End If
    tgClfCntr(ilLnRowNo - 1).iResearch = False
    tgClfCntr(ilLnRowNo - 1).lAvgAud = 0
    tgClfCntr(ilLnRowNo - 1).iAvgRating = 0
    If (Contract.plcDemoBar.Visible) Then
        ilIndex = imLnSave(10, ilLnRowNo)
        Do While ilIndex > 0
            For ilLoop = LBound(tmWkInfo(ilIndex).lSpotCount) To UBound(tmWkInfo(ilIndex).lSpotCount) Step 1
                If ilClearAvail Then
                    tmWkInfo(ilIndex).iAvail(ilLoop) = -29999   'Not defined
                End If
                tmWkInfo(ilIndex).lRCPrice(ilLoop) = -1
                tmWkInfo(ilIndex).lAvgAud(ilLoop) = 0
                tmWkInfo(ilIndex).lGRP(ilLoop) = 0
                tmWkInfo(ilIndex).lGrImp(ilLoop) = 0
                tmWkInfo(ilIndex).lPopEst(ilLoop) = 0
            Next ilLoop
            ilIndex = tmWkInfo(ilIndex).iNextIndex
        Loop
        If imDBTotals < 5 Then
            Contract.pbcDemoBar(0).Cls
            Contract.pbcDemoBar(1).Cls
            Contract.pbcDemoBar_Paint 0
            Contract.pbcDemoBar_Paint 1
        Else
            Contract.pbcDemoBar(1).Cls
            Contract.pbcDemoBar(2).Cls
            Contract.pbcDemoBar(3).Cls
            Contract.mSummary
            Contract.pbcDemoBar_Paint 1
            If imDBTotals = 5 Then
                Contract.pbcDemoBar_Paint 2
            Else
                Contract.pbcDemoBar_Paint 3
            End If
        End If
        Screen.MousePointer = vbDefault  'Wait
    End If
End Sub

Public Sub mCalculateTotalSpots()
    'Dan 5/3/2011 added because spot count not calculating properly if not scheduled.
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llMoStartDate As Long
    Dim ilSpots As Integer
    Dim ilDay As Integer
    Dim llDate As Long
    
    'Compute Total spots
    For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        tgClfCntr(ilClf).lCost = 0
        tgClfCntr(ilClf).lTSpots = 0
        ilCff = tgClfCntr(ilClf).iFirstCff
        Do While ilCff <> -1
            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                llStartDate = tgCffCntr(ilCff).lStartDate
                llEndDate = tgCffCntr(ilCff).lEndDate
                llMoStartDate = llStartDate
                Do While gWeekDayLong(llMoStartDate) <> 0
                    llMoStartDate = llMoStartDate - 1
                Loop
                For llDate = llMoStartDate To llEndDate Step 7
                    If tgCffCntr(ilCff).CffRec.sDyWk = "D" Then
                        ilSpots = 0
                        For ilDay = 0 To 6 Step 1
                            If (llDate + ilDay >= llStartDate) And (llDate + ilDay <= llEndDate) Then
                                ilSpots = ilSpots + tgCffCntr(ilCff).CffRec.iDay(ilDay)
                            End If
                        Next ilDay
                    Else
                        ilSpots = tgCffCntr(ilCff).CffRec.iSpotsWk + tgCffCntr(ilCff).CffRec.iXSpotsWk
                    End If
                    tgClfCntr(ilClf).lTSpots = tgClfCntr(ilClf).lTSpots + ilSpots
                Next llDate
            End If
            ilCff = tgCffCntr(ilCff).iNextCff
        Loop
    Next ilClf
End Sub
'12/16/14: Add Hide/Show Hidden Lines
Public Function mFindPaintLnRowNo(ilLnRowNo As Integer) As Integer
    Dim ilRow As Integer

    'add test if using feature
    If Not bgHideHiddenLines Then
        mFindPaintLnRowNo = ilLnRowNo
        Exit Function
    End If
    For ilRow = 0 To UBound(tgMapLinePaint) - 1 Step 1
        If (tgMapLinePaint(ilRow).lMapLnRowNo = ilLnRowNo) Then
            mFindPaintLnRowNo = tgMapLinePaint(ilRow).lPaintLnRowNo
            Exit Function
        End If
    Next ilRow
    'mFindPaintLnRowNo = UBound(tgClfCntr) + 1
    mFindPaintLnRowNo = UBound(tgMapLinePaint) + 1
End Function
'12/16/14: Add Hide/Show Hidden Lines
Public Sub mBuildMapLinePaint()
    Dim ilRow As Integer
    Dim ilPaintRowNo As Integer
    Dim slLineType As String
    
    ReDim tgMapLinePaint(0 To 0) As MAPLINEPAINT
    If Not bgHideHiddenLines Then
        Exit Sub
    End If
    'Map air time
    ilPaintRowNo = 1
    For ilRow = LBound(tgClfCntr) + 1 To UBound(tgClfCntr) + 1 Step 1
        If (tgClfCntr(ilRow - 1).ClfRec.iLine > 0) And (tgClfCntr(ilRow - 1).bShowLine) Then
            slLineType = mGetLineType(ilRow)
            tgMapLinePaint(UBound(tgMapLinePaint)).lMapLnRowNo = ilRow
            tgMapLinePaint(UBound(tgMapLinePaint)).lPaintLnRowNo = ilPaintRowNo
            tgMapLinePaint(UBound(tgMapLinePaint)).sLnType = slLineType
            ReDim Preserve tgMapLinePaint(0 To UBound(tgMapLinePaint) + 1) As MAPLINEPAINT
            ilPaintRowNo = ilPaintRowNo + 1
        End If
    Next ilRow
'    'Map Podcast Spots
'    ilPaintRowNo = 1
'    For ilRow = LBound(tgClfCntr) + 1 To UBound(tgClfCntr) + 1 Step 1
'        If (tgClfCntr(ilRow - 1).ClfRec.iLine > 0) And (tgClfCntr(ilRow - 1).bShowLine) Then
'            slLineType = mGetLineType(ilRow)
'            tgMapLinePaint(UBound(tgMapLinePaint)).lMapLnRowNo = ilRow
'            tgMapLinePaint(UBound(tgMapLinePaint)).lPaintLnRowNo = ilPaintRowNo
'            tgMapLinePaint(UBound(tgMapLinePaint)).sLnType = slLineType
'            ReDim Preserve tgMapLinePaint(0 To UBound(tgMapLinePaint) + 1) As MAPLINEPAINT
'            ilPaintRowNo = ilPaintRowNo + 1
'        End If
'    Next ilRow
End Sub

Public Sub mSetHideParameters()
    Dim ilPkg As Integer
    Dim ilPkgClf As Integer
    Dim ilIndex As Integer
    Dim ilHdClf As Integer
    Dim ilClf As Integer
    Dim slLineType As String
    
    For ilPkg = LBound(tgPackLn) To UBound(tgPackLn) - 1 Step 1
        If (tgPackLn(ilPkg).sType = "O") Or (tgPackLn(ilPkg).sType = "A") Or (tgPackLn(ilPkg).sType = "E") Then
            For ilPkgClf = 0 To UBound(tgClfCntr) Step 1
                If tgPackLn(ilPkg).iLineNo = tgClfCntr(ilPkgClf).ClfRec.iLine Then
                    tgClfCntr(ilPkgClf).bShowLine = True
                    ilIndex = tgPackLn(ilPkg).iIndex
                    Do While ilIndex > 0
                        For ilHdClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                            If tgClfCntr(ilHdClf).ClfRec.iLine = tgPackLn(ilIndex).iLineNo Then
                                tgClfCntr(ilHdClf).bShowLine = tgClfCntr(ilPkgClf).bPkExpanded
                                tgClfCntr(ilHdClf).bPkExpanded = tgClfCntr(ilPkgClf).bPkExpanded
                                Exit For
                            End If
                        Next ilHdClf
                        ilIndex = tgPackLn(ilIndex).iIndex
                    Loop
                    Exit For
                End If
            Next ilPkgClf
        End If
    Next ilPkg
    For ilClf = 0 To UBound(tgClfCntr) - 1 Step 1
        slLineType = mGetLineType(ilClf + 1)
        If (slLineType = "") Or (slLineType = "S") Then
            tgClfCntr(ilClf).bShowLine = True
            tgClfCntr(ilClf).bPkExpanded = True
        End If
    Next ilClf
End Sub
'12/16/14: Add Hide/Show Hidden Lines
Public Sub mGetHideHiddenLinesFlag()
    Dim hlSaf As Integer
    Dim tlSaf As SAF            'Schedule Attributes record image
    Dim tlSafSrchKey1 As SAFKEY1    'Vef key record image
    Dim ilSafRecLen As Integer
    Dim ilRet As Integer
    
    On Error Resume Next
    bgHideHiddenLines = False
    hlSaf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        ilSafRecLen = Len(tlSaf)  'btrRecordLength(hlSlf)  'Get and save record length
        tlSafSrchKey1.iVefCode = 0
        ilRet = btrGetEqual(hlSaf, tlSaf, ilSafRecLen, tlSafSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If (Asc(tlSaf.sFeatures2) And HIDEHIDDENLINES) = HIDEHIDDENLINES Then
                bgHideHiddenLines = True
            End If
        End If
    End If
    btrDestroy hlSaf
    ilRet = btrClose(hlSaf)
End Sub

Public Function mRegionAllowed(imLnRowNo As Integer) As Integer
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilTerminate As Integer
    Dim ilVef As Integer

    mRegionAllowed = False
    If imLnRowNo <= 0 Then
        Exit Function
    End If
    If (Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS Then  'Allocation
        mCVehOptIndex imLnRowNo, ilVefCode, ilVpfIndex, ilTerminate
        ilVef = gBinarySearchVef(ilVefCode)
        If ilVef <> -1 Then
            ''7/3/05: Jim- disallow split networks with selling vehicles
            'If (tgMVef(ilVef).sType <> "R") And (tgMVef(ilVef).sType <> "S") Then
            '    mRegionAllowed = True
            'End If
            '6/21/07: Jim- new rules-> Allow split networks for Package and selling.  Test Conventional, Airing and Game vehicles.
            '5/11/11: Allow selling to be set to No
            'If (tgMVef(ilVef).sType = "P") Or (tgMVef(ilVef).sType = "S") Then
            If (tgMVef(ilVef).sType = "P") Then
                mRegionAllowed = True
            'ElseIf (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "A") Or (tgMVef(ilVef).sType = "G") Then
            ElseIf (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "A") Or (tgMVef(ilVef).sType = "G") Or (tgMVef(ilVef).sType = "S") Then
                If ilVpfIndex <> -1 Then
                    If tgVpf(ilVpfIndex).sAllowSplitCopy = "Y" Then
                        mRegionAllowed = True
                    End If
                End If
            End If
        End If
    End If
End Function
Public Sub mSetBackgroundColor(ilStep1Index As Integer)
    'If Contract.grdCPMBuy.Visible Then
        Contract.pbcLineArrow.BackColor = vbButtonFace
        Contract.pbcLineArrow.Picture = IconTraf!imcArrow.Picture
    'Else
        If sgCntrScreen = "Orders" Then
            '8033
            Contract.imcShare(1).Visible = False
            'View- Show in yellow
            If ilStep1Index >= 4 Then
                '8033
                If mIsStationXmlAllowedForSite() Then
                    Contract.imcShare(1).Visible = True
                End If
                Contract.pbcHeader.Picture = Contract.pbcHdImage(1).Picture
                Contract.plcScreen.BackColor = LIGHTYELLOW
                Contract.BackColor = LIGHTYELLOW
                'DoEvents
                Contract.pbcLineArrow.BackColor = LIGHTYELLOW
                Contract.pbcLineArrow.Picture = IconTraf!imcArrow.Picture
                'imcKey.Picture = IconTraf!imcKeyYellow.Picture
            Else
                Contract.pbcHeader.Picture = Contract.pbcHdImage(0).Picture
                Contract.plcScreen.BackColor = LIGHTGREEN
                Contract.BackColor = LIGHTGREEN
                'DoEvents
                Contract.pbcLineArrow.BackColor = LIGHTGREEN
                Contract.pbcLineArrow.Picture = IconTraf!imcArrow.Picture
                'imcKey.Picture = IconTraf!imcKeyGreen.Picture
            End If
            'plcScreen_Paint
            Contract.plcScreen.Caption = sgCntrScreen
            lmAirTimeArrowBackColor = Contract.pbcLineArrow.BackColor
        Else
            'View- Show in yellow
            If ilStep1Index >= 6 Then
                Contract.pbcHeader.Picture = Contract.pbcHdImage(1).Picture
                Contract.plcScreen.BackColor = LIGHTYELLOW
                Contract.BackColor = LIGHTYELLOW
                'DoEvents
                Contract.pbcLineArrow.BackColor = LIGHTYELLOW
                Contract.pbcLineArrow.Picture = IconTraf!imcArrow.Picture
                'imcKey.Picture = IconTraf!imcKeyYellow.Picture
            ElseIf (ilStep1Index <= 3) Then     'Cyan
                Contract.pbcHeader.Picture = Contract.pbcHdImage(0).Picture
                Contract.plcScreen.BackColor = CYAN
                Contract.BackColor = CYAN
                'DoEvents
                Contract.pbcLineArrow.BackColor = CYAN
                Contract.pbcLineArrow.Picture = IconTraf!imcArrow.Picture
                'imcKey.Picture = IconTraf!imcKeyCyan.Picture
            Else
                Contract.pbcHeader.Picture = Contract.pbcHdImage(0).Picture
                Contract.plcScreen.BackColor = LIGHTGREEN
                Contract.BackColor = LIGHTGREEN
                'DoEvents
                Contract.pbcLineArrow.BackColor = LIGHTGREEN
                Contract.pbcLineArrow.Picture = IconTraf!imcArrow.Picture
                'imcKey.Picture = IconTraf!imcKeyGreen.Picture
            End If
            'plcScreen_Paint
            Contract.plcScreen.Caption = sgCntrScreen
            lmAirTimeArrowBackColor = Contract.pbcLineArrow.BackColor
        End If
        Contract.lacCode.BackColor = Contract.BackColor
        '4/28/07:  Adjust for varibale screen size
        Contract.pbcHeader.Picture = LoadPicture()
        If ilStep1Index = 10 Then   'Force View mode
            If sgCntrScreen = "Orders" Then
                imStep1Index = 4
            Else
                imStep1Index = 6
            End If
            gPaintArea Contract.pbcSelect, tmStepCtrls(1).fBoxX, tmStepCtrls(1).fBoxY, tmStepCtrls(1).fBoxW - 15, tmStepCtrls(1).fBoxH - 15, WHITE
            tmStepCtrls(1).sShow = "View"
            Contract.pbcSelect_Paint
        End If
    'End If
End Sub

Public Sub mGetOwnership(ilLnRowNo As Integer)
    Dim slLineType As String
    Dim ilVpf As Integer
    
    slLineType = mGetLineType(ilLnRowNo)
    tgClfCntr(ilLnRowNo - 1).sOwnership = "A"
    If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
        ilVpf = gBinarySearchVpf(tgClfCntr(ilLnRowNo - 1).ClfRec.iVefCode)
        If ilVpf <> -1 Then
            tgClfCntr(ilLnRowNo - 1).sOwnership = tgVpf(ilVpf).sOwnership
        End If
    End If
    tgClfCntr(ilLnRowNo - 1).iAcqCostSpotRate = mGetAcqCostSpotRateFlag(ilLnRowNo)
End Sub

Public Sub mAdvtPop()
'
'   mAdvtPop
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Dim slResvName As String
    Dim ilResvIndex As Integer
    ilIndex = Contract.lbcAdvt.ListIndex
    If ilIndex > 0 Then
        slName = Contract.lbcAdvt.List(ilIndex)
    End If
    ilResvIndex = Contract.lbcResvAdvt.ListIndex
    If ilResvIndex > 0 Then
        slName = Contract.lbcResvAdvt.List(ilResvIndex)
    End If
    If imStep1Index = 0 Then
        igPopAdfAgfDormant = False
    ElseIf imStep1Index = -1 Then
        If sgCntrScreen = "Orders" Then
            igPopAdfAgfDormant = True
        Else
            igPopAdfAgfDormant = False
        End If
    Else
        igPopAdfAgfDormant = True
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    ilRet = gPopAdvtBox(Contract, Contract.lbcAdvt, tmCntAdvertiser(), smCntAdvertiserTag) 'Traffic!lbcAdvertiser)
    igPopAdfAgfDormant = True
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAdvtPopErr
        gCPErrorMsg ilRet, "mAdvtPop (gPopAdvtBox)", Contract
        On Error GoTo 0
        Contract.lbcResvAdvt.Clear
        For ilLoop = 0 To Contract.lbcAdvt.ListCount - 1 Step 1
            Contract.lbcResvAdvt.AddItem Contract.lbcAdvt.List(ilLoop)
        Next ilLoop
        Contract.lbcAdvt.AddItem "[New]", 0  'Force as first item on list
        Contract.lbcResvAdvt.AddItem "[None]", 0
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, Contract.lbcAdvt
            If gLastFound(Contract.lbcAdvt) > 0 Then
                Contract.lbcAdvt.ListIndex = gLastFound(Contract.lbcAdvt)
            Else
                Contract.lbcAdvt.ListIndex = -1
            End If
        Else
            Contract.lbcAdvt.ListIndex = ilIndex
        End If
        If ilResvIndex > 0 Then
            gFindMatch slResvName, 1, Contract.lbcResvAdvt
            If gLastFound(Contract.lbcResvAdvt) > 0 Then
                Contract.lbcResvAdvt.ListIndex = gLastFound(Contract.lbcResvAdvt)
            Else
                Contract.lbcResvAdvt.ListIndex = -1
            End If
        Else
            Contract.lbcResvAdvt.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mAdvtPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Public Sub mAgencyPop()
'
'   mAgencyPop
'   Where:
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = Contract.lbcAgency.ListIndex
    If ilIndex > 1 Then
        slName = Contract.lbcAgency.List(ilIndex)
    End If
    If imStep1Index = 0 Then
        igPopAdfAgfDormant = False
    ElseIf imStep1Index = -1 Then
        If sgCntrScreen = "Orders" Then
            igPopAdfAgfDormant = True
        Else
            igPopAdfAgfDormant = False
        End If
    Else
        igPopAdfAgfDormant = True
    End If
    'Repopulate if required- if agency changed by another user while in this screen
    ilRet = gPopAgyBox(Contract, Contract.lbcAgency, tmCntAgency(), smCntAgencyTag)    'lbcAgency)
    igPopAdfAgfDormant = True
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mAgencyPopErr
        gCPErrorMsg ilRet, "mAgencyPop (gIMoveListBox)", Contract
        On Error GoTo 0
        Contract.lbcAgency.AddItem "[Direct Advertiser]", 0
        Contract.lbcAgency.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slName, 2, Contract.lbcAgency
            If gLastFound(Contract.lbcAgency) > 1 Then
                    Contract.lbcAgency.ListIndex = gLastFound(Contract.lbcAgency)
            Else
                Contract.lbcAgency.ListIndex = -1
            End If
        Else
            Contract.lbcAgency.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mAgencyPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Public Sub mAvailPop()


'This module is not used any longer- it was used to help build the daypart name

'
'   mAvailPop
'   Where:
'
    Dim ilRet As Integer 'btrieve status
    ReDim ilFilter(0) As Integer
    ReDim slFilter(0) As String
    ReDim ilOffSet(0) As Integer
    Dim ilLoop As Integer
    Dim ilLp As Integer
    Dim slStr As String
'    ilFilter(0) = NOFILTER
'    slFilter(0) = ""
'    ilOffset(0) = 0
    ilFilter(0) = CHARFILTERNOT
    slFilter(0) = "F"
    ilOffSet(0) = gFieldOffset("Anf", "AnfBookLocalFeed") '2
    'ilRet = gIMoveListBox(Contract, lbcAvail, lbcAvailCode, "Anf.btr", gFieldOffset("Anf", "AnfName"), 20, ilFilter(), slFilter(), ilOffset())
    ilRet = gIMoveListBox(Contract, Contract.lbcAvail, tgAvailCode(), sgAvailCodeTag, "Anf.btr", gFieldOffset("Anf", "AnfName"), 20, ilFilter(), slFilter(), ilOffSet())
    If ilRet <> CP_MSG_NOPOPREQ Then
        'Remove "Post Log" avail name
        For ilLoop = 0 To Contract.lbcAvail.ListCount - 1 Step 1
            slStr = Trim$(Contract.lbcAvail.List(ilLoop))
            If StrComp(slStr, "Post Log", 1) = 0 Then
                Contract.lbcAvail.RemoveItem ilLoop
                For ilLp = ilLoop To UBound(tgAvailCode) - 1 Step 1
                    tgAvailCode(ilLp) = tgAvailCode(ilLp + 1)
                Next ilLp
                ReDim Preserve tgAvailCode(LBound(tgAvailCode) To UBound(tgAvailCode) - 1) As SORTCODE
                Exit For
            End If
        Next ilLoop
        On Error GoTo mAvailPopErr
        gCPErrorMsg ilRet, "mAvailPop (gIMoveListBox)", Contract
        On Error GoTo 0
    End If
    Exit Sub
mAvailPopErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Public Sub mBuyerPop(ilRetrieve As Integer, ilAdvtAgyCode As Integer)
'
'   mBuyerPop
'   Where:
'       ilAdvtCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    ilIndex = Contract.lbcBuyer.ListIndex
    If ilIndex > 0 Then
        slName = Contract.lbcBuyer.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopPersonnelBox(Contract, ilRetrieve, ilAdvtAgyCode, "B", True, 0, lbcBuyer, Traffic!lbcBuyerCode)
    ilRet = gPopPersonnelBox(Contract, ilRetrieve, ilAdvtAgyCode, "B", True, 0, Contract.lbcBuyer, tmCntBuyerCode(), smCntBuyerCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mBuyerPopErr
        gCPErrorMsg ilRet, "mBuyerPop (gPopPersonnelBox)", Contract
        On Error GoTo 0
        Contract.lbcBuyer.AddItem "[None]", 0  'Force as first item on list
        Contract.lbcBuyer.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, Contract.lbcBuyer
            If gLastFound(Contract.lbcBuyer) > 0 Then
                Contract.lbcBuyer.ListIndex = gLastFound(Contract.lbcBuyer)
            Else
                Contract.lbcBuyer.ListIndex = -1
            End If
        Else
            Contract.lbcBuyer.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mBuyerPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

Public Sub mDemoPop()
'
'   mDemoPop
'   Where:
'
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slDemo As String      'Demo name, saved to determine if changed
    Dim ilIndex As Integer      'Demo name, saved to determine if changed
    Dim ilDemo As Integer
    'Repopulate if required- if sales source changed by another user while in this screen
    For ilLoop = 0 To 3 Step 1
        ilIndex = Contract.lbcDemo(ilLoop).ListIndex
        If ilIndex > 1 Then
            slDemo = Contract.lbcDemo(ilLoop).List(ilIndex)
        End If
        If ilLoop = 0 Then
            'ilRet = gPopMnfPlusFieldsBox(Contract, lbcDemo(ilLoop), Traffic!lbcDemoCode, "D")
            ilRet = gPopMnfPlusFieldsBox(Contract, Contract.lbcDemo(ilLoop), tgDemoCode(), sgDemoCodeTag, "D")
        Else
            For ilDemo = 1 To Contract.lbcDemo(0).ListCount - 1 Step 1
                Contract.lbcDemo(ilLoop).AddItem Contract.lbcDemo(0).List(ilDemo)
            Next ilDemo
            ilRet = CP_MSG_POPREQ
        End If
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mDemoPopErr
            gCPErrorMsg ilRet, "mDemoPop (gPopMnfPlusFieldsBox)", Contract
            On Error GoTo 0
            Contract.lbcDemo(ilLoop).AddItem "[None]", 0
            imChgMode = True
            If ilIndex >= 1 Then
                gFindMatch slDemo, 1, Contract.lbcDemo(ilLoop)
                If gLastFound(Contract.lbcDemo(ilLoop)) > 0 Then
                    Contract.lbcDemo(ilLoop).ListIndex = gLastFound(Contract.lbcDemo(ilLoop))
                Else
                    Contract.lbcDemo(ilLoop).ListIndex = -1
                End If
            Else
                Contract.lbcDemo(ilLoop).ListIndex = ilIndex
            End If
            imChgMode = False
        End If
    Next ilLoop
    Exit Sub
mDemoPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

Public Sub mMnfPop(ilType As Integer)
'
'   mMnfPop
'      ilType (I): 1 = Business Categories
'                  2 = Product Protection
'                  3 = Exclusion
'                  4 = Potential
'                  5 = Segment
'   Where:
'
    ReDim ilFilter(0 To 0) As Integer
    ReDim slFilter(0 To 0) As String
    ReDim ilOffSet(0 To 0) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slBusCat As String      'Potential name, saved to determine if changed
    Dim slPot As String
    Dim slSeg As String
    Dim ilIndex As Integer      'Potential name, saved to determine if changed
    ReDim slComp(0 To 1) As String      'Competitive name, saved to determine if changed
    ReDim ilComp(0 To 1) As Integer      'Competitive name, saved to determine if changed
    ReDim slExcl(0 To 1) As String      'Exclusion name, saved to determine if changed
    ReDim ilExcl(0 To 1) As Integer      'Exclusion name, saved to determine if changed
    'Repopulate if required- if sales source changed by another user while in this screen
    If ilType = 1 Then
        ilFilter(0) = CHARFILTER
        slFilter(0) = "B"
        ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
        ilIndex = Contract.lbcBusCat.ListIndex
        If ilIndex > 1 Then
            slBusCat = Contract.lbcBusCat.List(ilIndex)
        End If
        'ilRet = gIMoveListBox(Contract, lbcBusCat, Traffic!lbcBusCatCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
        ilRet = gIMoveListBox(Contract, Contract.lbcBusCat, tgBusCatCode(), sgBusCatCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mMnfPopErr
            gCPErrorMsg ilRet, "mMnfPop (gIMoveListBox: Business Categories)", Contract
            On Error GoTo 0
            Contract.lbcBusCat.AddItem "[None]", 0
            Contract.lbcBusCat.AddItem "[New]", 0  'Force as first item on list
            imChgMode = True
            If ilIndex > 1 Then
                gFindMatch slBusCat, 2, Contract.lbcBusCat
                If gLastFound(Contract.lbcBusCat) > 1 Then
                    Contract.lbcBusCat.ListIndex = gLastFound(Contract.lbcBusCat)
                Else
                    Contract.lbcBusCat.ListIndex = -1
                End If
            Else
                Contract.lbcBusCat.ListIndex = ilIndex
            End If
            imChgMode = False
        End If
    ElseIf ilType = 2 Then
        'Repopulate if required- if sales source changed by another user while in this screen
        ilFilter(0) = CHARFILTER
        slFilter(0) = "C"
        ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
        ilComp(0) = Contract.lbcComp(0).ListIndex
        ilComp(1) = Contract.lbcComp(1).ListIndex
        If ilComp(0) > 1 Then
            slComp(0) = Contract.lbcComp(0).List(ilComp(0))
        End If
        If ilComp(1) > 1 Then
            slComp(1) = Contract.lbcComp(1).List(ilComp(1))
        End If
        If Contract.lbcComp(0).ListCount <> Contract.lbcComp(1).ListCount Then
            Contract.lbcComp(0).Clear
        End If
        'ilRet = gIMoveListBox(Contract, lbcComp(0), Traffic!lbcCompCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
        ilRet = gIMoveListBox(Contract, Contract.lbcComp(0), tgCompCode(), sgCompCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mMnfPopErr
            gCPErrorMsg ilRet, "mMnfPop (gIMoveListBox: Product Protection)", Contract
            On Error GoTo 0
            Contract.lbcComp(0).AddItem "[None]", 0
            Contract.lbcComp(0).AddItem "[New]", 0  'Force as first item on list
            Contract.lbcComp(1).Clear
            For ilLoop = Contract.lbcComp(0).ListCount - 1 To 0 Step -1
                Contract.lbcComp(1).AddItem Contract.lbcComp(0).List(ilLoop), 0
            Next ilLoop
            imChgMode = True
            If ilComp(0) > 1 Then
                gFindMatch slComp(0), 2, Contract.lbcComp(0)
                If gLastFound(Contract.lbcComp(0)) > 1 Then
                    Contract.lbcComp(0).ListIndex = gLastFound(Contract.lbcComp(0))
                Else
                    Contract.lbcComp(0).ListIndex = -1
                End If
            Else
                Contract.lbcComp(0).ListIndex = ilComp(0)
            End If
            If ilComp(1) > 1 Then
                gFindMatch slComp(1), 2, Contract.lbcComp(1)
                If gLastFound(Contract.lbcComp(1)) > 1 Then
                    Contract.lbcComp(1).ListIndex = gLastFound(Contract.lbcComp(1))
                Else
                    Contract.lbcComp(1).ListIndex = -1
                End If
            Else
                Contract.lbcComp(1).ListIndex = ilComp(1)
            End If
            imChgMode = False
        End If
    ElseIf ilType = 3 Then
        ilFilter(0) = CHARFILTER
        slFilter(0) = "X"
        ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
        ilExcl(0) = Contract.lbcExcl(0).ListIndex
        ilExcl(1) = Contract.lbcExcl(1).ListIndex
        If ilExcl(0) > 1 Then
            slExcl(0) = Contract.lbcExcl(0).List(ilExcl(0))
        End If
        If ilExcl(1) > 1 Then
            slExcl(1) = Contract.lbcExcl(1).List(ilExcl(1))
        End If
        If Contract.lbcExcl(0).ListCount <> Contract.lbcExcl(1).ListCount Then
            Contract.lbcExcl(0).Clear
        End If
        'ilRet = gIMoveListBox(Contract, lbcExcl(0), Traffic!lbcExclCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
        ilRet = gIMoveListBox(Contract, Contract.lbcExcl(0), tgExclCode(), sgExclCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mMnfPopErr
            gCPErrorMsg ilRet, "mMnfPop (gIMoveListBox: Exclusions)", Contract
            On Error GoTo 0
            Contract.lbcExcl(0).AddItem "[None]", 0
            Contract.lbcExcl(0).AddItem "[New]", 0  'Force as first item on list
            Contract.lbcExcl(1).Clear
            For ilLoop = Contract.lbcExcl(0).ListCount - 1 To 0 Step -1
                Contract.lbcExcl(1).AddItem Contract.lbcExcl(0).List(ilLoop), 0
            Next ilLoop
            imChgMode = True
            If ilExcl(0) > 1 Then
                gFindMatch slExcl(0), 2, Contract.lbcExcl(0)
                If gLastFound(Contract.lbcExcl(0)) > 1 Then
                    Contract.lbcExcl(0).ListIndex = gLastFound(Contract.lbcExcl(0))
                Else
                    Contract.lbcExcl(0).ListIndex = -1
                End If
            Else
                Contract.lbcExcl(0).ListIndex = ilExcl(0)
            End If
            If ilExcl(1) > 1 Then
                gFindMatch slExcl(1), 2, Contract.lbcExcl(1)
                If gLastFound(Contract.lbcExcl(1)) > 1 Then
                    Contract.lbcExcl(1).ListIndex = gLastFound(Contract.lbcExcl(1))
                Else
                    Contract.lbcExcl(1).ListIndex = -1
                End If
            Else
                Contract.lbcExcl(1).ListIndex = ilExcl(1)
            End If
            imChgMode = False
        End If
    ElseIf ilType = 4 Then
        ilFilter(0) = CHARFILTER
        slFilter(0) = "P"
        ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
        ilIndex = Contract.lbcPot.ListIndex
        If ilIndex > 1 Then
            slPot = Contract.lbcPot.List(ilIndex)
        End If
        'ilRet = gIMoveListBox(Contract, lbcPot, Traffic!lbcPotCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
        ilRet = gIMoveListBox(Contract, Contract.lbcPot, tgPotCode(), sgPotCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mMnfPopErr
            gCPErrorMsg ilRet, "mMnfPop (gIMoveListBox: Potential)", Contract
            On Error GoTo 0
            Contract.lbcPot.AddItem "[None]", 0
            Contract.lbcPot.AddItem "[New]", 0  'Force as first item on list
            imChgMode = True
            If ilIndex > 1 Then
                gFindMatch slPot, 2, Contract.lbcPot
                If gLastFound(Contract.lbcPot) > 1 Then
                    Contract.lbcPot.ListIndex = gLastFound(Contract.lbcPot)
                Else
                    Contract.lbcPot.ListIndex = -1
                End If
            Else
                Contract.lbcPot.ListIndex = ilIndex
            End If
            imChgMode = False
        End If
    ElseIf ilType = 5 Then
        If tgSpf.sCUseSegments <> "Y" Then  'Using Segments
            Contract.lbcSeg.Clear
            ReDim tgSegCode(0 To 0) As SORTCODE
            Exit Sub
        End If
        ilFilter(0) = CHARFILTER
        slFilter(0) = "K"
        ilOffSet(0) = gFieldOffset("Mnf", "MnfType") '2
        ilIndex = Contract.lbcSeg.ListIndex
        If ilIndex > 1 Then
            slSeg = Contract.lbcSeg.List(ilIndex)
        End If
        'ilRet = gIMoveListBox(Contract, lbcSeg, Traffic!lbcSegCode, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffset())
        ilRet = gIMoveListBox(Contract, Contract.lbcSeg, tgSegCode(), sgSegCodeTag, "Mnf.Btr", gFieldOffset("Mnf", "MnfName"), 20, ilFilter(), slFilter(), ilOffSet())
        If ilRet <> CP_MSG_NOPOPREQ Then
            On Error GoTo mMnfPopErr
            gCPErrorMsg ilRet, "mMnfPop (gIMoveListBox: Segment)", Contract
            On Error GoTo 0
            Contract.lbcSeg.AddItem "[New]", 0  'Force as first item on list
            imChgMode = True
            If ilIndex > 0 Then
                gFindMatch slSeg, 1, Contract.lbcSeg
                If gLastFound(Contract.lbcSeg) > 1 Then
                    Contract.lbcSeg.ListIndex = gLastFound(Contract.lbcSeg)
                Else
                    Contract.lbcSeg.ListIndex = -1
                End If
            Else
                Contract.lbcSeg.ListIndex = ilIndex
            End If
            imChgMode = False
        End If
    End If
    Exit Sub
mMnfPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

Public Sub mProdPop(ilAdvtCode As Integer, tmProdCode() As SORTCODE, smProdCodeTag As String)
'
'   mProdPop
'   Where:
'       ilAdvtCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Screen.MousePointer = vbHourglass  'Wait
    ilIndex = Contract.lbcProd.ListIndex
    If ilIndex > 0 Then
        slName = Contract.lbcProd.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopAdvtProdBox(Contract, ilAdvtCode, lbcProd, Traffic!lbcProdCode)
    ilRet = gPopAdvtProdBox(Contract, ilAdvtCode, Contract.lbcProd, tmProdCode(), smProdCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mProdPopErr
        gCPErrorMsg ilRet, "mProdPop (gIMoveListBox)", Contract
        On Error GoTo 0
        Contract.lbcProd.AddItem "[None]", 0  'Force as first item on list
'            lbcProd.AddItem "[New]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex > 0 Then
            gFindMatch slName, 1, Contract.lbcProd
            If gLastFound(Contract.lbcProd) > 0 Then
                Contract.lbcProd.ListIndex = gLastFound(Contract.lbcProd)
            Else
                Contract.lbcProd.ListIndex = -1
            End If
        Else
            Contract.lbcProd.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mProdPopErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
End Sub

Public Sub mRevSetPop(RSREV1INDEX As Integer)
'
'   mRevSetPop
'   Where:
'
    'ReDim ilFilter(0 To 1) As Integer
    'ReDim slFilter(0 To 1) As String
    'ReDim ilOffset(0 To 1) As Integer
    Dim ilLoop As Integer
    'Dim ilCount As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    'Dim slRev As String      'Exclusion name, saved to determine if changed
    'Dim ilRev As Integer      'Exclusion name, saved to determine if changed
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim slStr As String
    Dim slShow As String
    'Dim slNameCode As String
    'Dim slCode As String
    ilRet = gObtainMnfForType("R", sgRevSetStamp, tgRevSet())
    For ilLoop = 0 To 4 Step 1
        imChgMode = True
        Select Case ilLoop
            Case 0
                mPopRevSet ilLoop + 1, Contract.lbcRevSet(ilLoop), tmRevSetCode0(), smRevSetCodeTag(ilLoop), RSREV1INDEX
            Case 1
                mPopRevSet ilLoop + 1, Contract.lbcRevSet(ilLoop), tmRevSetCode1(), smRevSetCodeTag(ilLoop), RSREV1INDEX
            Case 2
                mPopRevSet ilLoop + 1, Contract.lbcRevSet(ilLoop), tmRevSetCode2(), smRevSetCodeTag(ilLoop), RSREV1INDEX
            Case 3
                mPopRevSet ilLoop + 1, Contract.lbcRevSet(ilLoop), tmRevSetCode3(), smRevSetCodeTag(ilLoop), RSREV1INDEX
            Case 4
                mPopRevSet ilLoop + 1, Contract.lbcRevSet(ilLoop), tmRevSetCode4(), smRevSetCodeTag(ilLoop), RSREV1INDEX
        End Select
        imChgMode = False
        smRevSetName(ilLoop) = ""
        For ilIndex = 2 To Contract.lbcRevSet(ilLoop).ListCount - 1 Step 1
            If smRevSetName(ilLoop) = "" Then
                smRevSetName(ilLoop) = Trim$(Contract.lbcRevSet(ilLoop).List(ilIndex))
            Else
                smRevSetName(ilLoop) = smRevSetName(ilLoop) & "/" & Trim$(Contract.lbcRevSet(ilLoop).List(ilIndex))
            End If
        Next ilIndex
    Next ilLoop

    slShow = tmRSCtrls(1).sShow
    llColor = Contract.pbcShareBudget.ForeColor
    slFontName = Contract.pbcShareBudget.FontName
    flFontSize = Contract.pbcShareBudget.FontSize
    Contract.pbcShareBudget.ForeColor = BLUE
    Contract.pbcShareBudget.FontBold = False
    Contract.pbcShareBudget.FontSize = 7
    Contract.pbcShareBudget.FontName = "Arial"
    Contract.pbcShareBudget.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    For ilLoop = 0 To 4 Step 1
        slStr = smRevSetName(ilLoop)
        gSetShow Contract.pbcShareBudget, slStr, tmRSCtrls(1)
        smRevSetName(ilLoop) = tmRSCtrls(1).sShow
    Next ilLoop
    Contract.pbcShareBudget.FontSize = flFontSize
    Contract.pbcShareBudget.FontName = slFontName
    Contract.pbcShareBudget.FontSize = flFontSize
    Contract.pbcShareBudget.ForeColor = llColor
    Contract.pbcShareBudget.FontBold = True
    tmRSCtrls(1).sShow = slShow
    Exit Sub

    On Error GoTo 0
    imTerminate = True
End Sub

Public Sub mShtTitlePop(ilAdvtCode As Integer)
'
'   mShtTitlePop
'   Where:
'       ilAdvtCode (I)- Adsvertiser code value
'
    Dim ilRet As Integer
    'Dim slNameCode As String
    'Dim slCode As String
    'Dim ilCode As Integer
    'Dim ilFound As Integer
    Dim slName As String
    Dim ilIndex As Integer
    Screen.MousePointer = vbHourglass  'Wait
    ilIndex = Contract.lbcShtTitle.ListIndex
    If ilIndex >= 0 Then
        slName = Contract.lbcShtTitle.List(ilIndex)
    End If
    'Repopulate if required- if sales source changed by another user while in this screen
    'ilRet = gPopShortTitleBox(Contract, ilAdvtCode, lbcShtTitle, lbcShtTitleCode)
    ilRet = gPopShortTitleBox(Contract, ilAdvtCode, Contract.lbcShtTitle, tgShtTitleCode(), sgShtTitleCodeTag)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mShtTitlePopErr
        gCPErrorMsg ilRet, "mShtTitlePop (gIMoveListBox)", Contract
        On Error GoTo 0
        'lbcShtTitle.AddItem "[None]", 0  'Force as first item on list
        imChgMode = True
        If ilIndex >= 0 Then
            gFindMatch slName, 0, Contract.lbcShtTitle
            If gLastFound(Contract.lbcShtTitle) >= 0 Then
                Contract.lbcShtTitle.ListIndex = gLastFound(Contract.lbcShtTitle)
            Else
                Contract.lbcShtTitle.ListIndex = -1
            End If
        Else
            Contract.lbcShtTitle.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
mShtTitlePopErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    imTerminate = True
End Sub

Public Sub mSocEcoPop()
'
'   mSocEcoPop
'   Where:
'
    Dim ilRet As Integer
    'Dim ilFound As Integer
    Dim slSocEco As String      'Demo name, saved to determine if changed
    Dim ilIndex As Integer      'Demo name, saved to determine if changed
    'Repopulate if required- if sales source changed by another user while in this screen
    ilIndex = Contract.lbcDBSocEco.ListIndex
    If ilIndex > 1 Then
        slSocEco = Contract.lbcDBSocEco.List(ilIndex)
    End If
    'ilRet = gPopMnfPlusFieldsBox(Contract, Contract.lbcDBSocEco, Traffic!lbcSocEcoCode, "FNG")
    ilRet = gPopMnfPlusFieldsBox(Contract, Contract.lbcDBSocEco, tgSocEcoCode(), sgSocEcoCodeTag, "FNG")
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mSocEcoPopErr
        gCPErrorMsg ilRet, "mSocEcoPop (gPopMnfPlusFieldsBox)", Contract
        On Error GoTo 0
        Contract.lbcDBSocEco.AddItem "[None]", 0
        imChgMode = True
        If ilIndex > 1 Then
            gFindMatch slSocEco, 1, Contract.lbcDBSocEco
            If gLastFound(Contract.lbcDBSocEco) > 0 Then
                Contract.lbcDBSocEco.ListIndex = gLastFound(Contract.lbcDBSocEco)
            Else
                Contract.lbcDBSocEco.ListIndex = -1
            End If
        Else
            Contract.lbcDBSocEco.ListIndex = ilIndex
        End If
        imChgMode = False
    End If
    Exit Sub
mSocEcoPopErr:
    On Error GoTo 0
    imTerminate = True
End Sub

Public Sub mSetAcqOrSpotPrice(slSetAcqOrSpot As String, ilLnRowNo As Integer, slPrice As String, ilPRICEINDEX As Integer)
    Dim slLineType As String
    Dim ilCff As Integer
    Dim slPriceType As String
    Dim llPrice As Long
    Dim slStr As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    
    If imDiffMode Then
        Exit Sub
    End If
    If ilLnRowNo <= 0 Then
        Exit Sub
    End If
    If ilLnRowNo >= UBound(tgClfCntr) + 1 Then
        Exit Sub
    End If
    If (Asc(tgSpf.sUsingFeatures2) And BARTER) <> BARTER Then
        Exit Sub
    End If
    If tgClfCntr(ilLnRowNo - 1).iAcqCostSpotRate <> 2 Then
        Exit Sub
    End If
    'Check if CBS line
    If tgClfCntr(ilLnRowNo - 1).iCancel Then
        Exit Sub
    End If
    gUnpackDateLong tgClfCntr(ilLnRowNo - 1).ClfRec.iStartDate(0), tgClfCntr(ilLnRowNo - 1).ClfRec.iStartDate(1), llStartDate
    gUnpackDateLong tgClfCntr(ilLnRowNo - 1).ClfRec.iEndDate(0), tgClfCntr(ilLnRowNo - 1).ClfRec.iEndDate(1), llEndDate
    If llEndDate < llStartDate Then
        Exit Sub
    End If
    If slSetAcqOrSpot = "A" Then
        ilCff = tgClfCntr(ilLnRowNo - 1).iFirstCff
        If ilCff <> -1 Then
            slPriceType = tgCffCntr(ilCff).CffRec.sPriceType
            llPrice = tgCffCntr(ilCff).CffRec.lActPrice
        Else
            If Asc(Left(slPrice, 1)) >= 0 And Asc(Left(slPrice, 1)) + 9 Then
                slPriceType = "T"
                llPrice = gStrDecToLong(slPrice, 2)
            Else
                slPriceType = "X"
                llPrice = 0
            End If
        End If
        If slPriceType = "T" Then  'Price type (0=dollars)
            smLnSave(24, ilLnRowNo) = gLongToStrDec(llPrice, 2)
            tgClfCntr(ilLnRowNo - 1).ClfRec.lAcquisitionCost = gStrDecToLong(slPrice, 2)
        Else
            smLnSave(24, ilLnRowNo) = ".00"
            tgClfCntr(ilLnRowNo - 1).ClfRec.lAcquisitionCost = 0
        End If
    Else
        'Set Flight price
        ilCff = tgClfCntr(ilLnRowNo - 1).iFirstCff
        Do While ilCff <> -1
            If (tgCffCntr(ilCff).iStatus = 0) Or (tgCffCntr(ilCff).iStatus = 1) Then
                tgCffCntr(ilCff).CffRec.lActPrice = gStrDecToLong(slPrice, 2)
                tgCffCntr(ilCff).CffRec.sPriceType = "T"
                'set show and recompute package price
                slStr = slPrice
                gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                'gSetShow Contract.pbcLine, slStr, tmLnCtrls(ilPRICEINDEX)
                mSetPriceShow slStr, ilPRICEINDEX
                smLnShow(ilPRICEINDEX, imLnRowNo) = tmLnCtrls(ilPRICEINDEX).sShow
            End If
            ilCff = tgCffCntr(ilCff).iNextCff
        Loop
        
    End If
End Sub

Private Function mGetAcqCostSpotRateFlag(ilLnRowNo As Integer) As Integer
    Dim slLineType As String
    
    mGetAcqCostSpotRateFlag = 0
    If ilLnRowNo <= 0 Then
        Exit Function
    End If
    'If ilLnRowNo >= UBound(tgClfCntr) + 1 Then
    '    Exit Function
    'End If
    If (Asc(tgSpf.sUsingFeatures2) And BARTER) <> BARTER Then
        Exit Function
    End If
    slLineType = mGetLineType(ilLnRowNo)
    If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
      
        Select Case tgClfCntr(ilLnRowNo - 1).sOwnership
            Case "A"    'Owned-Network
            Case "B"    'Owned-Station
                mGetAcqCostSpotRateFlag = 1
            Case "C"    'Unowned-Network
            Case "D"    'Unowned-Station
                mGetAcqCostSpotRateFlag = 2
        End Select
    Else
        mGetAcqCostSpotRateFlag = 1
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveRecToCtrl                  *
'*                                                     *
'*             Created:7/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Move record values to controls *
'*                      on the screen                  *
'*                                                     *
'*******************************************************
Public Sub mMoveHdRecToCtrl(DATEINDEX As Integer, tmChf As CHF, tmEff As EFF, tmAdf As ADF, tmAgf As AGF, tmVsf As VSF, tmRcf As RCF, imTerminate As Integer)
'
'   mMoveRecToCtrl
'   Where:
'
    Dim ilLoop As Integer
    Dim slRecCode As String
    Dim slNameCode As String  'Name and code
    Dim ilRet As Integer    'Return call status
    Dim slCode As String    'Sales source code number
    Dim slName As String
    Dim slStr As String
    Dim ilOutLp As Integer
    Dim ilRevSet As Integer
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim ilRcf As Integer
    Dim ilCount As Integer
    Dim ilVsf As Integer
    Dim ilVef As Integer
    Dim blFound As Boolean
    Dim ilPass As Integer
    
    '11/2/19: Part of CFlight
    Contract.lbcDWP.AddItem "Daily"
    Contract.lbcDWP.AddItem "Weekly"
    'Contract.lbcDW.Clear
    'Contract.lbcDW.AddItem "Daily"
    'Contract.lbcDW.AddItem "Weekly"
    Select Case tgChfCntr.sStatus
        Case "W"    'Working
            If tgChfCntr.iCntRevNo > 0 Then
                slStr = "Rev Working"
            Else
                slStr = "Working Proposal"
            End If
        Case "D"
            slStr = "Rejected"
        Case "C"
            If tgChfCntr.iCntRevNo > 0 Then
                slStr = "Rev Completed"
            Else
                slStr = "Completed Proposal"
            End If
        Case "I"
            If tgChfCntr.iCntRevNo > 0 Then
                slStr = "Rev Unapproved"
            Else
                slStr = "Unapproved Proposal"
            End If
        Case "G"
            slStr = "Approved Hold"
        Case "H"
            slStr = "Hold"
        Case "N"
            slStr = "Approved Order"
        Case "O"
            slStr = "Order"
    End Select
    gFindMatch slStr, 0, Contract.lbcStatus
    If gLastFound(Contract.lbcStatus) < 0 Then
        Contract.lbcStatus.ListIndex = -1
    Else
        Contract.lbcStatus.ListIndex = gLastFound(Contract.lbcStatus)
    End If
    'Potential
    Contract.lbcPot.ListIndex = 1
    smPot = ""
    slRecCode = Trim$(str$(tgChfCntr.iMnfPotnType))
    For ilLoop = 0 To UBound(tgPotCode) - 1 Step 1  'Traffic!Contract.lbcPotCode.ListCount - 1 Step 1
        slNameCode = tgPotCode(ilLoop).sKey   'Traffic!Contract.lbcPotCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
        On Error GoTo 0
        If slRecCode = slCode Then
            Contract.lbcPot.ListIndex = ilLoop + 2
            smPot = Contract.lbcPot.List(ilLoop + 2)
            Exit For
        End If
    Next ilLoop
    Select Case tgChfCntr.sType
        Case "C"    'Standard
            slStr = "Standard"
        Case "V"  'Reservation
            slStr = "Reservation"
        Case "T"  'Remnant
            slStr = "Remnant"
        Case "R"  'DR
            slStr = "Direct Response"
        Case "Q"    'PI
            slStr = "Per Inquiry"
        Case "S"  'PSA
            slStr = "PSA"
        Case "M"
            slStr = "Promo"
        Case Else
            slStr = ""
    End Select
    'imTestTypeIndexSetting = False
    gFindMatch slStr, 0, Contract.lbcType
    If gLastFound(Contract.lbcType) < 0 Then
        Contract.lbcType.ListIndex = -1
    Else
        Contract.lbcType.ListIndex = gLastFound(Contract.lbcType)
    End If
    'imCntrOrigTypeIndex = Contract.lbcType.ListIndex
    'imTestTypeIndexSetting = True
    If (tgChfCntr.lCntrNo <= 0) Or (imStep1Index <= 1) Then
        Contract.edcCntrNo.Text = ""
        smPropNo = ""
        lmOrigCntrNo = -1   'tgChfCntr.lCntrNo
    Else
        Contract.edcCntrNo.Text = Trim$(str$(tgChfCntr.lCntrNo))
        smPropNo = Trim$(str$(tgChfCntr.lCntrNo))
        lmOrigCntrNo = tgChfCntr.lCntrNo
    End If
    'If tgChfCntr.lAgyEstNo > 0 Then
        'Contract.edcEstNo.Text = Trim$(Str$(tgChfCntr.lAgyEstNo))
        smEstNo = Trim$(tgChfCntr.sAgyEstNo) & Trim$(tgChfCntr.sTitle)
    'Else
        'Contract.edcEstNo.Text = ""
        'smEstNo = ""
    'End If
    Contract.lbcResvAdvt.ListIndex = 0
    smResvAdvt = ""
    smInitResvNo = ""
    Contract.lbcResv.ListIndex = 0   '[None]
    If tgChfCntr.lChfCode > 0 Then
        tmChfSrchKey.lCode = tgChfCntr.lChfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        slRecCode = Trim$(str$(tmChf.iAdfCode))
        tmAdfSrchKey.iCode = tmChf.iAdfCode
        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmAdf.iCode = 0
            tmAdf.sName = ""
            tmAdf.iAgfCode = 0
            tmAdf.iSlfCode = 0
        End If
        For ilLoop = 0 To UBound(tmCntAdvertiser) - 1 Step 1 'Traffic!Contract.lbcAdvertiser.ListCount - 1 Step 1
            slNameCode = tmCntAdvertiser(ilLoop).sKey   'Traffic!Contract.lbcAdvertiser.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
            On Error GoTo 0
            If slRecCode = slCode Then
                Contract.lbcResvAdvt.ListIndex = ilLoop + 1
                smResvAdvt = Contract.lbcResvAdvt.List(ilLoop + 1)
                Exit For
            End If
        Next ilLoop
        Contract.mResvPop
        For ilLoop = 0 To UBound(tgResvCode) - 1 Step 1 'Traffic!Contract.lbcResvCode.ListCount - 1 Step 1
            slNameCode = tgResvCode(ilLoop).sKey    'Traffic!Contract.lbcResvCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
            On Error GoTo 0
            If tgChfCntr.lChfCode = Val(slCode) Then
                Contract.lbcResv.ListIndex = ilLoop + 1
                smInitResvNo = Contract.lbcResv.List(ilLoop + 1)
                Exit For
            End If
        Next ilLoop
        'tmChfSrchKey.lCode = tgChfCntr.lCode
        'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, ilLockBias, SETFORREADONLY)
    End If
    smEDIClientCode = ""
    smEDIProductCode = ""
    If tgChfCntr.lEffCode > 0 Then
        tmEffSrchKey.lCode = tgChfCntr.lEffCode
        ilRet = btrGetEqual(hmEff, tmEff, imEffRecLen, tmEffSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            smEDIClientCode = Trim$(tmEff.sString(0))
            smEDIProductCode = Trim$(tmEff.sString(1))
        End If
    End If
    'smInvProjCode = ""
    'smInvSMSNo = ""
    'If tgChfCntr.lEffCode > 0 Then
    '    tmEffSrchKey.lCode = tgChfCntr.lEffCode
    '    ilRet = btrGetEqual(hmEff, tmEff, imEffRecLen, tmEffSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '    If ilRet = BTRV_ERR_NONE Then
    '        smInvProjCode = Trim$(tmEff.sString(0))
    '        smInvSMSNo = Trim$(tmEff.sString(1))
    '    End If
    'End If
    gUnpackDate tgChfCntr.iDtNeed(0), tgChfCntr.iDtNeed(1), smInitDtNeed
    smDtNeeded = smInitDtNeed
    imFromExtRevNo = tgChfCntr.iExtRevNo
    imToExtRevNo = tgChfCntr.iExtRevNo + 1
    'If ((tgChfCntr.sStatus = "H") Or (tgChfCntr.sStatus = "O")) And (((sgCntrScreen = "Orders") And (imStep1Index >= 4)) Or ((sgCntrScreen <> "Orders") And (imStep1Index >= 6))) Then
    If (((sgCntrScreen = "Orders") And (imStep1Index >= 4)) Or ((sgCntrScreen <> "Orders") And (imStep1Index >= 6))) Then
        If tgChfCntr.iCntRevNo <> 0 Then
            Contract.lbcRevNo.Clear
            Contract.lbcRevNo.AddItem Trim$(str$(tgChfCntr.iCntRevNo)) & "-" & Trim$(str$(tgChfCntr.iExtRevNo))
            Contract.edcRevNo.Text = Trim$(str$(tgChfCntr.iCntRevNo)) & "-" & Trim$(str$(tgChfCntr.iExtRevNo))
        Else
            Contract.lbcRevNo.Clear
            Contract.lbcRevNo.AddItem "0-0"
            Contract.edcRevNo.Text = "0-0"
        End If
    ElseIf ((sgCntrScreen = "Orders") And (imStep1Index < 3)) Or ((sgCntrScreen <> "Orders") And (imStep1Index = 4)) Then
        Contract.lbcRevNo.Clear
        Contract.edcRevNo.Text = Trim$(str$(tgChfCntr.iExtRevNo))
    Else
        If tgChfCntr.iCntRevNo <> 0 Then
            Contract.lbcRevNo.Clear
            Contract.lbcRevNo.AddItem Trim$(str$(tgChfCntr.iCntRevNo)) & "-" & Trim$(str$(tgChfCntr.iExtRevNo))
            Contract.edcRevNo.Text = Trim$(str$(tgChfCntr.iCntRevNo)) & "-" & Trim$(str$(tgChfCntr.iExtRevNo))
        Else
            Contract.lbcRevNo.Clear
            Contract.lbcRevNo.AddItem "0-0"
            Contract.edcRevNo.Text = "0-0"
        End If
    End If
    If tgChfCntr.iPropVer <> 0 Then
        Contract.edcPropVer.Text = Trim$(str$(tgChfCntr.iPropVer))
    Else
        Contract.edcPropVer.Text = ""
    End If
    If sgCntrScreen = "Orders" Then
        gUnpackDate tgChfCntr.iOHDDate(0), tgChfCntr.iOHDDate(1), slStr
        If slStr = "" Then
            gUnpackDate tgChfCntr.iPropDate(0), tgChfCntr.iPropDate(1), slStr
        End If
        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(DATEINDEX)
    Else
        gUnpackDate tgChfCntr.iPropDate(0), tgChfCntr.iPropDate(1), slStr
        If slStr = "" Then
            gUnpackDate tgChfCntr.iOHDDate(0), tgChfCntr.iOHDDate(1), slStr
        End If
        gSetShow Contract.pbcHeader, slStr, tmHdCtrls(DATEINDEX)
    End If
    'look up advertiser name from code number
    Contract.lbcAdvt.ListIndex = 0
    smAdvertiser = ""
    slRecCode = Trim$(str$(tgChfCntr.iAdfCode))
    tmAdfSrchKey.iCode = tgChfCntr.iAdfCode
    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        tmAdf.iCode = 0
        tmAdf.sName = ""
        tmAdf.iAgfCode = 0
        tmAdf.iSlfCode = 0
    End If
    '10/2/14: Re=populate list if name not found
    blFound = False
    For ilPass = 0 To 1 Step 1
        For ilLoop = 0 To UBound(tmCntAdvertiser) - 1 Step 1 'Traffic!Contract.lbcAdvertiser.ListCount - 1 Step 1
            slNameCode = tmCntAdvertiser(ilLoop).sKey   'Traffic!Contract.lbcAdvertiser.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
            On Error GoTo 0
            If slRecCode = slCode Then
                blFound = True
                Contract.lbcAdvt.ListIndex = ilLoop + 1
                smAdvertiser = Contract.lbcAdvt.List(ilLoop + 1)
                Exit For
            End If
        Next ilLoop
        If Not blFound Then
            smCntAdvertiserTag = ""
            mAdvtPop
        Else
            Exit For
        End If
    Next ilPass
    smProduct = Trim$(tgChfCntr.sProduct)
    If smProduct = "" Then
        smProduct = "[None]"
    End If
    mProdPop tmAdf.iCode, tmProdCode(), smProdCodeTag
    Screen.MousePointer = vbHourglass  'Reset as mProdPop removes HourGlass
    smInitShtTitle = ""
    smShtTitle = ""
    If tgSpf.sUseProdSptScr = "P" Then
        mShtTitlePop tmAdf.iCode
        Screen.MousePointer = vbHourglass  'Wait
        ReDim tgCShtTitle(0 To 0) As CSHTTITLE
        If tgChfCntr.lSifCode > 0 Then
            slRecCode = Trim$(str$(tgChfCntr.lSifCode))
            For ilLoop = 0 To UBound(tgShtTitleCode) - 1 Step 1 'Contract.lbcShtTitleCode.ListCount - 1 Step 1
                slNameCode = tgShtTitleCode(ilLoop).sKey   'Contract.lbcShtTitleCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mMoveRecToCtrlErr
                gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
                On Error GoTo 0
                If slRecCode = slCode Then
                    Contract.lbcShtTitle.ListIndex = ilLoop ' + 1
                    smShtTitle = Contract.lbcShtTitle.List(ilLoop) ' + 1)
                    smInitShtTitle = smShtTitle
                    Exit For
                End If
            Next ilLoop
            'Initialize Short Title by Vehicle
            If tgChfCntr.lVefCode < 0 Then
                tmVsfSrchKey.lCode = -tgChfCntr.lVefCode
                ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                        If (tmVsf.iFSCode(ilVsf) > 0) And (tmVsf.lFSComm(ilVsf) > 0) Then
                            tgCShtTitle(UBound(tgCShtTitle)).sKey = ""    'Key will be build in CShTitle mVehPop
                            tgCShtTitle(UBound(tgCShtTitle)).iVefCode = tmVsf.iFSCode(ilVsf)
                            tgCShtTitle(UBound(tgCShtTitle)).lSifCode = tmVsf.lFSComm(ilVsf)
                            tgCShtTitle(UBound(tgCShtTitle)).sShortTitle = ""
                            slRecCode = Trim$(str$(tmVsf.lFSComm(ilVsf)))
                            For ilLoop = 0 To UBound(tgShtTitleCode) - 1 Step 1 'Contract.lbcShtTitleCode.ListCount - 1 Step 1
                                slNameCode = tgShtTitleCode(ilLoop).sKey   'Contract.lbcShtTitleCode.List(ilLoop)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                On Error GoTo mMoveRecToCtrlErr
                                gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
                                On Error GoTo 0
                                If slRecCode = slCode Then
                                    tgCShtTitle(UBound(tgCShtTitle)).sShortTitle = Contract.lbcShtTitle.List(ilLoop) ' + 1)
                                    Exit For
                                End If
                            Next ilLoop
                            ReDim Preserve tgCShtTitle(0 To UBound(tgCShtTitle) + 1) As CSHTTITLE
                        End If
                    Next ilVsf
                End If
            End If
        End If
    End If
    'look up agency name from code number
    Contract.lbcAgency.ListIndex = 1
    smAgency = ""
    slRecCode = Trim$(str$(tgChfCntr.iAgfCode))
    If tgChfCntr.iAgfCode = 0 Then  'Direct
        Contract.lbcAgency.ListIndex = 1
        smAgency = Contract.lbcAgency.List(1)
        tmAgf.iCode = 0
        tmAgf.sName = ""
        tmAgf.iSlfCode = 0
    Else
        tmAgfSrchKey.iCode = tgChfCntr.iAgfCode
        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            tmAgf.iCode = 0
            tmAgf.sName = ""
            tmAgf.iSlfCode = 0
        End If
        '10/2/14: Re-populate list if name not found
        blFound = False
        For ilPass = 0 To 1 Step 1
            For ilLoop = 0 To UBound(tmCntAgency) - 1 Step 1 'Traffic!Contract.lbcAgency.ListCount - 1 Step 1
                slNameCode = tmCntAgency(ilLoop).sKey    'Traffic!Contract.lbcAgency.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mMoveRecToCtrlErr
                gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
                On Error GoTo 0
                If slRecCode = slCode Then
                    blFound = True
                    Contract.lbcAgency.ListIndex = ilLoop + 2
                    smAgency = Contract.lbcAgency.List(ilLoop + 2)
                    Exit For
                End If
            Next ilLoop
            If Not blFound Then
                smCntAgencyTag = ""
                mAgencyPop
            Else
                Exit For
            End If
        Next ilPass
    End If
    '10/27/14: set 1 or 2 place rating
    smCntr1or2PlaceRating = gSet1or2PlaceRating(tgChfCntr.iAgfCode)
    
    If tmAgf.iCode > 0 Then
        mBuyerPop 1, tmAgf.iCode
    Else
        mBuyerPop 0, tmAdf.iCode
    End If
    'look up salesperson name from code number
    '5/18/12: Only include dormant salesperson if referenced in contract
    bmCheckSlfDormant = True
    Contract.mSPersonPop
    
    For ilOutLp = 0 To 9 Step 1
        imSlspIndex(ilOutLp) = -1
        smSlspName(ilOutLp) = ""
        smInitSlspName(ilOutLp) = ""
        'smSlspComm(ilOutLp) = ""
        'smInitSlspComm(ilOutLp) = ""
        smSlspField2(ilOutLp) = ""
        smInitSlspField2(ilOutLp) = ""
        smSlspField3(ilOutLp) = ""
        smInitSlspField3(ilOutLp) = ""
        smSlspField4(ilOutLp) = ""
        smInitSlspField4(ilOutLp) = ""
        slRecCode = Trim$(str$(tgChfCntr.iSlfCode(ilOutLp)))
        If tgChfCntr.iSlfCode(ilOutLp) <> 0 Then
        'If ilOutLp < 8 Then
            For ilLoop = 0 To UBound(tmCntSalesperson) - 1 Step 1  'Traffic!Contract.lbcSalesperson.ListCount - 1 Step 1
                slNameCode = tmCntSalesperson(ilLoop).sKey 'Traffic!Contract.lbcSalesperson.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mMoveRecToCtrlErr
                gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
                On Error GoTo 0
                If slRecCode = slCode Then
                    imSlspIndex(ilOutLp) = ilLoop + 2
                    smSlspName(ilOutLp) = Contract.lbcSPerson(0).List(ilLoop + 2)
                    smInitSlspName(ilOutLp) = smSlspName(ilOutLp)
                    'smSlspComm(ilOutLp) = gLongToStrDec(tgChfCntr.lComm(ilOutLp), 4)
                    'smInitSlspComm(ilOutLp) = smSlspComm(ilOutLp)
                    If tgSpf.sCommByCntr = "Y" Then
                        smSlspField2(ilOutLp) = gIntToStrDec(tgChfCntr.iSlspCommPct(ilOutLp), 2)
                        smInitSlspField2(ilOutLp) = smSlspField2(ilOutLp)
                        smSlspField3(ilOutLp) = gLongToStrDec(tgChfCntr.lComm(ilOutLp), 4)
                        smInitSlspField3(ilOutLp) = smSlspField3(ilOutLp)
                    ElseIf tgSpf.sSubCompany = "Y" Then
                        smSlspField3(ilOutLp) = gLongToStrDec(tgChfCntr.lComm(ilOutLp), 4)
                        smInitSlspField3(ilOutLp) = smSlspField3(ilOutLp)
                        smSlspField2(ilOutLp) = "[None]"
                        smInitSlspField2(ilOutLp) = smSlspField2(ilOutLp)
                        If tgChfCntr.iMnfSubCmpy(ilOutLp) > 0 Then
                            slRecCode = Trim$(str$(tgChfCntr.iMnfSubCmpy(ilOutLp)))
                            For ilVef = 0 To UBound(tmVehGp6Code) - 1 Step 1  'Traffic!Contract.lbcPotCode.ListCount - 1 Step 1
                                slNameCode = tmVehGp6Code(ilVef).sKey   'Traffic!Contract.lbcPotCode.List(ilLoop)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                On Error GoTo mMoveRecToCtrlErr
                                gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
                                On Error GoTo 0
                                If slRecCode = slCode Then
                                    ilRet = gParseItem(slNameCode, 1, "\", slName)
                                    On Error GoTo mMoveRecToCtrlErr
                                    gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 1)", Contract
                                    On Error GoTo 0
                                    smSlspField2(ilOutLp) = Trim$(slName)
                                    smInitSlspField2(ilOutLp) = smSlspField2(ilOutLp)
                                    Exit For
                                End If
                            Next ilVef
                        End If
                        smSlspField4(ilOutLp) = gIntToStrDec(tgChfCntr.iSlspCommPct(ilOutLp), 2)
                        smInitSlspField4(ilOutLp) = smSlspField4(ilOutLp)
                    Else
                        smSlspField2(ilOutLp) = gLongToStrDec(tgChfCntr.lComm(ilOutLp), 4)
                        smInitSlspField2(ilOutLp) = smSlspField2(ilOutLp)
                    End If
                    'Contract.lbcSPerson.ListIndex = ilLoop + 1
                    'smSPerson = Contract.lbcSPerson.List(ilLoop + 1)
                    Exit For
                End If
            Next ilLoop
        End If
    Next ilOutLp
    'look up competitive name from code number
    For ilOutLp = 0 To 1 Step 1
        Contract.lbcComp(ilOutLp).ListIndex = 1
        smComp(ilOutLp) = ""
        slRecCode = Trim$(str$(tgChfCntr.iMnfComp(ilOutLp)))
        For ilLoop = 0 To UBound(tgCompCode) - 1 Step 1 'Traffic!Contract.lbcCompCode.ListCount - 1 Step 1
            slNameCode = tgCompCode(ilLoop).sKey    'Traffic!Contract.lbcCompCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
            On Error GoTo 0
            If slRecCode = slCode Then
                Contract.lbcComp(ilOutLp).ListIndex = ilLoop + 2
                smComp(ilOutLp) = Contract.lbcComp(ilOutLp).List(ilLoop + 2)
                Exit For
            End If
        Next ilLoop
    Next ilOutLp
    smComp0 = smComp(0)
    smComp1 = smComp(1)
    For ilOutLp = 0 To 1 Step 1
        Contract.lbcExcl(ilOutLp).ListIndex = 1
        smExcl(ilOutLp) = ""
        slRecCode = Trim$(str$(tgChfCntr.iMnfExcl(ilOutLp)))
        For ilLoop = 0 To UBound(tgExclCode) - 1 Step 1 'Traffic!Contract.lbcExclCode.ListCount - 1 Step 1
            slNameCode = tgExclCode(ilLoop).sKey    'Traffic!Contract.lbcExclCode.List(ilLoop)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
            On Error GoTo 0
            If slRecCode = slCode Then
                Contract.lbcExcl(ilOutLp).ListIndex = ilLoop + 2
                smExcl(ilOutLp) = Contract.lbcExcl(ilOutLp).List(ilLoop + 2)
                Exit For
            End If
        Next ilLoop
    Next ilOutLp
    Contract.lbcBusCat.ListIndex = 1
    smBusCat = ""
    slRecCode = Trim$(str$(tgChfCntr.iMnfBus))
    For ilLoop = 0 To UBound(tgBusCatCode) - 1 Step 1 'Traffic!Contract.lbcBusCatCode.ListCount - 1 Step 1
        slNameCode = tgBusCatCode(ilLoop).sKey  'Traffic!Contract.lbcBusCatCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
        On Error GoTo 0
        If slRecCode = slCode Then
            Contract.lbcBusCat.ListIndex = ilLoop + 2
            smBusCat = Contract.lbcBusCat.List(ilLoop + 2)
            Exit For
        End If
    Next ilLoop
    smBuyer = Trim$(tgChfCntr.sBuyer)
    smOrigBuyer = smBuyer
    slRecCode = Trim$(str$(tgChfCntr.iPnfBuyer))
    'Look up buyers name
    For ilLoop = 0 To UBound(tmCntBuyerCode) - 1 Step 1 'Traffic!Contract.lbcBuyerCode.ListCount - 1 Step 1
        slNameCode = tmCntBuyerCode(ilLoop).sKey   'Traffic!Contract.lbcBuyerCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 1)", Contract
        On Error GoTo 0
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
        On Error GoTo 0
        'If Trim$(tgChfCntr.sBuyer) = Trim$(slName) Then
        If slRecCode = slCode Then
            Contract.lbcBuyer.ListIndex = ilLoop + 2
            smBuyer = Contract.lbcBuyer.List(ilLoop + 2)
            smOrigBuyer = smBuyer
            Exit For
        End If
    Next ilLoop

    gSetPhoneNo tgChfCntr.sPhone, Contract.mkcPhone
    'If (tgChfCntr.sPriceCntr = "N") And (tgChfCntr.sPriceInv = "N") Then
    '    Contract.lbcShowPrice.ListIndex = 0
    'ElseIf (tgChfCntr.sPriceCntr = "Y") And (tgChfCntr.sPriceInv = "N") Then
    '    Contract.lbcShowPrice.ListIndex = 1
    'ElseIf (tgChfCntr.sPriceCntr = "N") And (tgChfCntr.sPriceInv = "Y") Then
    '    Contract.lbcShowPrice.ListIndex = 2
    'ElseIf (tgChfCntr.sPriceCntr = "Y") And (tgChfCntr.sPriceInv = "Y") Then
    '    Contract.lbcShowPrice.ListIndex = 3
    'Else
    '    Contract.lbcShowPrice.ListIndex = -1
    'End If
    'gPDNToStr tgChfCntr.sPctBudget, 0, slStr
    'slStr = gIntToStrDec(tgChfCntr.iPctBudget, 0)
    'Contract.edcPctBud.Text = slStr
    'look up Share Budget Names from code number
    For ilOutLp = 0 To 6 Step 1
        imShareBudgetIndex(ilOutLp) = -1
        smShareBudget(ilOutLp) = ""
        smInitShareBudget(ilOutLp) = ""
        smShareBudgetPct(ilOutLp) = ""
        smInitShareBudgetPct(ilOutLp) = ""
        slRecCode = Trim$(str$(tgChfCntr.iMnfCmpy(ilOutLp)))
        If ilOutLp = 0 Then
            ilIndex = 0
            ilCount = UBound(tmShareBudgetCode0) - 1
        Else
            ilIndex = 1
            ilCount = -1
            On Error Resume Next
            ilCount = UBound(tmShareBudgetCode1) - 1
            On Error GoTo 0
        End If
        For ilLoop = 0 To ilCount Step 1
            If ilIndex = 0 Then
                slNameCode = tmShareBudgetCode0(ilLoop).sKey   'Traffic!Contract.lbcShareBudgetCode(ilIndex).List(ilLoop)
            Else
                slNameCode = tmShareBudgetCode1(ilLoop).sKey   'Traffic!Contract.lbcShareBudgetCode(ilIndex).List(ilLoop)
            End If
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
            On Error GoTo 0
            If slRecCode = slCode Then
                imShareBudgetIndex(ilOutLp) = ilLoop + 2
                smShareBudget(ilOutLp) = Contract.lbcShareBudget(ilIndex).List(ilLoop + 2)
                smInitShareBudget(ilOutLp) = smShareBudget(ilOutLp)
                smShareBudgetPct(ilOutLp) = gIntToStrDec(tgChfCntr.iCmpyPct(ilOutLp), 0)
                smInitShareBudgetPct(ilOutLp) = smShareBudgetPct(ilOutLp)
                Exit For
            End If
        Next ilLoop
    Next ilOutLp
    smComment(0) = smInitComment(0) 'Merchandising
    smComment(1) = smInitComment(1) 'Promotion
    'look up revenue sets from code number
    For ilRevSet = 0 To 4 Step 1
        Contract.lbcRevSet(ilRevSet).ListIndex = 1
        smRevSet(ilRevSet) = ""
        slRecCode = Trim$(str$(tgChfCntr.iMnfRevSet(ilRevSet)))
        Select Case ilRevSet
            Case 0
                ilCount = UBound(tmRevSetCode0) - 1
            Case 1
                ilCount = UBound(tmRevSetCode1) - 1
            Case 2
                ilCount = UBound(tmRevSetCode2) - 1
            Case 3
                ilCount = UBound(tmRevSetCode3) - 1
            Case 4
                ilCount = UBound(tmRevSetCode4) - 1
        End Select
        For ilLoop = 0 To ilCount Step 1
            'slNameCode = Traffic!Contract.lbcRevSetCode(ilRevSet).List(ilLoop)
            Select Case ilRevSet
                Case 0
                    slNameCode = tmRevSetCode0(ilLoop).sKey
                Case 1
                    slNameCode = tmRevSetCode1(ilLoop).sKey
                Case 2
                    slNameCode = tmRevSetCode2(ilLoop).sKey
                Case 3
                    slNameCode = tmRevSetCode3(ilLoop).sKey
                Case 4
                    slNameCode = tmRevSetCode4(ilLoop).sKey
            End Select
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            On Error GoTo mMoveRecToCtrlErr
            gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
            On Error GoTo 0
            If slRecCode = slCode Then
                Contract.lbcRevSet(ilRevSet).ListIndex = ilLoop + 2
                smRevSet(ilRevSet) = Contract.lbcRevSet(ilRevSet).List(ilLoop + 2)
                Exit For
            End If
        Next ilLoop
    Next ilRevSet
    'Segment
    If tgSpf.sCUseSegments = "Y" Then
        If Contract.lbcSeg.ListCount > 1 Then
            Contract.lbcSeg.ListIndex = 1
            smSeg = ""
            slRecCode = Trim$(str$(tgChfCntr.imnfSeg))
            For ilLoop = 0 To UBound(tgSegCode) - 1 Step 1  'Traffic!Contract.lbcPotCode.ListCount - 1 Step 1
                slNameCode = tgSegCode(ilLoop).sKey   'Traffic!Contract.lbcPotCode.List(ilLoop)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                On Error GoTo mMoveRecToCtrlErr
                gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
                On Error GoTo 0
                If slRecCode = slCode Then
                    Contract.lbcSeg.ListIndex = ilLoop + 1
                    smSeg = Contract.lbcSeg.List(ilLoop + 1)
                    Exit For
                End If
            Next ilLoop
        Else
            smSeg = ""
            Contract.lbcSeg.ListIndex = -1
        End If
    Else
        smSeg = ""
        Contract.lbcSeg.ListIndex = -1
    End If
    If (tgSpf.sGUsePropSys <> "Y") Then
        tgChfCntr.sCppCpm = "N"
        For ilLoop = 0 To 3 Step 1
            'tgChfCntr.sDemo(ilLoop) = ""
            tgChfCntr.iMnfDemo(ilLoop) = 0
            'slStr = ""
            'gStrToPDN slStr, 2, 4, tgChfCntr.sTarget(ilLoop)
            tgChfCntr.lTarget(ilLoop) = 0
        Next ilLoop
    Else
        Select Case tgChfCntr.sCppCpm
            Case "N"
                imCppCpm = 0 'cbcPriceType.ListIndex = 0
                For ilLoop = 0 To 3 Step 1
                    Contract.lbcDemo(ilLoop).ListIndex = -1
                    smSvDemo(ilLoop) = ""
                    smSvTarget(ilLoop) = ""
                    smTarget(ilLoop) = ""
                Next ilLoop
            Case "M"
                imCppCpm = 1 'cbcPriceType.ListIndex = 1
                For ilLoop = 0 To 3 Step 1
                    Contract.lbcDemo(ilLoop).ListIndex = -1
                    smSvDemo(ilLoop) = ""
                    smSvTarget(ilLoop) = ""
                    smTarget(ilLoop) = ""
                    slRecCode = Trim$(str$(tgChfCntr.iMnfDemo(ilLoop)))
                    For ilIndex = 0 To UBound(tgDemoCode) - 1 Step 1 'Traffic!Contract.lbcDemoCode.ListCount - 1 Step 1
                        slNameCode = tgDemoCode(ilIndex).sKey 'Traffic!Contract.lbcDemoCode.List(ilIndex)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        On Error GoTo mMoveRecToCtrlErr
                        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
                        On Error GoTo 0
                        If slRecCode = slCode Then
                            Contract.lbcDemo(ilLoop).ListIndex = ilIndex + 1
                            smSvDemo(ilLoop) = Contract.lbcDemo(ilLoop).List(ilIndex + 1)
                            slStr = gLongToStrDec(tgChfCntr.lTarget(ilLoop), 2)
                            smSvTarget(ilLoop) = slStr
                            smTarget(ilLoop) = slStr
                            Exit For
                        End If
                    Next ilIndex
                Next ilLoop
    
    
               '     gFindMatch tgChfCntr.sDemo(ilLoop), 1, Contract.lbcDemo(ilLoop)
               '     If gLastFound(Contract.lbcDemo(ilLoop)) >= 1 Then
               '         Contract.lbcDemo(ilLoop).ListIndex = gLastFound(Contract.lbcDemo(ilLoop))
               '         'gPDNToStr tgChfCntr.sTarget(ilLoop), 2, slStr
               '         slStr = gLongToStrDec(tgChfCntr.lTarget(ilLoop), 2)
               '         smSvTarget(ilLoop) = slStr
               '         smTarget(ilLoop) = slStr
               '     Else
               '         Contract.lbcDemo(ilLoop).ListIndex = -1
               '         smSvTarget(ilLoop) = ""
               '         smTarget(ilLoop) = ""
               '     End If
               ' Next ilLoop
            Case "P"
                imCppCpm = 2 'cbcPriceType.ListIndex = 2
                For ilLoop = 0 To 3 Step 1
                    Contract.lbcDemo(ilLoop).ListIndex = -1
                    smSvDemo(ilLoop) = ""
                    smSvTarget(ilLoop) = ""
                    smTarget(ilLoop) = ""
                    slRecCode = Trim$(str$(tgChfCntr.iMnfDemo(ilLoop)))
                    For ilIndex = 0 To UBound(tgDemoCode) - 1 Step 1 'Traffic!Contract.lbcDemoCode.ListCount - 1 Step 1
                        slNameCode = tgDemoCode(ilIndex).sKey  'Traffic!Contract.lbcDemoCode.List(ilIndex)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        On Error GoTo mMoveRecToCtrlErr
                        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
                        On Error GoTo 0
                        If slRecCode = slCode Then
                            Contract.lbcDemo(ilLoop).ListIndex = ilIndex + 1
                            smSvDemo(ilLoop) = Contract.lbcDemo(ilLoop).List(ilIndex + 1)
                            slStr = gLongToStrDec(tgChfCntr.lTarget(ilLoop), 2)
                            slStr = Left$(slStr, Len(slStr) - 3) 'Remove decimal point
                            smSvTarget(ilLoop) = slStr
                            smTarget(ilLoop) = slStr
                            Exit For
                        End If
                    Next ilIndex
                Next ilLoop
                'For ilLoop = 0 To 3 Step 1
                '    gFindMatch tgChfCntr.sDemo(ilLoop), 1, Contract.lbcDemo(ilLoop)
                '    If gLastFound(Contract.lbcDemo(ilLoop)) >= 1 Then
                '            Contract.lbcDemo(ilLoop).ListIndex = gLastFound(Contract.lbcDemo(ilLoop))
                '            'gPDNToStr tgChfCntr.sTarget(ilLoop), 2, slStr
                '            slStr = gLongToStrDec(tgChfCntr.lTarget(ilLoop), 2)
                '            smSvTarget(ilLoop) = slStr
                '            smTarget(ilLoop) = slStr
                '    Else
                '        Contract.lbcDemo(ilLoop).ListIndex = -1
                '        smSvTarget(ilLoop) = ""
                '        smTarget(ilLoop) = ""
                '    End If
                'Next ilLoop
            Case Else
                imCppCpm = -1 'cbcPriceType.ListIndex = 0
                For ilLoop = 0 To 3 Step 1
                    Contract.lbcDemo(ilLoop).ListIndex = -1
                    smSvDemo(ilLoop) = ""
                    smSvTarget(ilLoop) = ""
                    smTarget(ilLoop) = ""
                Next ilLoop
        End Select
    End If
    If tgSaf(0).sHideDemoOnBR = "Y" Then
        Select Case tgChfCntr.sHideDemo
            Case "N"
                imHideDemo = 0
            Case "Y"
                imHideDemo = 1
            Case Else
                imHideDemo = -1
        End Select
    Else
        imHideDemo = -1
    End If
    'gPDNToStr tgChfCntr.sPctTrade, 0, slStr
    slStr = gLongToStrDec(tgChfCntr.lGuar, 0)
    'Contract.edcGuar.Text = slStr
    smHdSave(2) = slStr
    slStr = gIntToStrDec(tgChfCntr.iPctTrade, 0)
    smPctTrade = slStr
    smOrigPctTrade = slStr
    imAllowCashTradeChgs = True
    smRvfTypeI_Adj = "0.00"
    smRvfTypeA_Adj = "0.00"
    If (tgChfCntr.lCntrNo > 0) And (imStep1Index > 1) Then
        imAllowCashTradeChgs = Not mRvfExistForCntr(hmRvf, hmPhf)
    End If
    mGetInstallAdjustments hmAgf, hmRvf, hmPhf
    If tgChfCntr.sAgyCTrade = "Y" Then    'Change name later
        imAgyComm = 0
    Else
        imAgyComm = 1   'No
    End If
    If tgSpf.sMktBase = "Y" Then
        '3/31/06- Handle case where no vehicles defined for contract
        'imMktFound = False
        If tgChfCntr.lVefCode <> 0 Then
            imMktFound = False
        Else
            imMktFound = True
        End If
    Else
        imMktFound = True
    End If
    'look up rate card from code number
    If imStep3Index < 0 Then
        Contract.lbcRateCard.ListIndex = 0
    Else
        Contract.lbcRateCard.ListIndex = -1
    End If
    smRateCard = ""
    tmRcf.iCode = 0
    slRecCode = Trim$(str$(tgChfCntr.iRcfCode))
    For ilLoop = 0 To UBound(tmCntRateCard) - 1 Step 1 'Traffic!Contract.lbcRateCardCode.ListCount - 1 Step 1
        slNameCode = tmCntRateCard(ilLoop).sKey   'Traffic!Contract.lbcRateCardCode.List(ilLoop)
        ilRet = gParseItem(slNameCode, 3, "\", slCode)
        On Error GoTo mMoveRecToCtrlErr
        gCPErrorMsg ilRet, "mMoveRecToCtrl (gParseItem field 2)", Contract
        On Error GoTo 0
        If slRecCode = slCode Then
            Contract.lbcRateCard.ListIndex = ilLoop '+ 1  'Traffic!Contract.lbcRateCardCode.ListCount - ilLoop - 1
            smRateCard = Contract.lbcRateCard.List(ilLoop) ' + 1)   'Traffic!Contract.lbcRateCardCode.ListCount - ilLoop - 1)
            smCurRC = smRateCard
            On Error GoTo mMoveRecToCtrlErr
            'gRCRead Contract, tgChfCntr.iRcfCode, tgRcf, tgRif(), tgRdf()
            'Build from tgMRcf; tgMRif; tgMRdf
            For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                If tgMRcf(ilRcf).iCode = tgChfCntr.iRcfCode Then
                    tmRcf = tgMRcf(ilRcf)
                    smBaseLen = Trim$(str$(tmRcf.iBaseLen)) 'Trim$(Str$(tgMRcf(ilRcf).iBaseLen))
                    Exit For
                End If
            Next ilRcf
            If imVefDormantAdd Then
                smVehicleCodeTag = ""
                imVefDormantAdd = False
            End If
            Contract.mVehPop True
            'Add any Missing dormant vehicles
            If tgChfCntr.lVefCode > 0 Then
                ilFound = False
                For ilIndex = 0 To UBound(tmVehicleCode) - 1 Step 1  'Contract.lbcVehicle.ListCount - 1 Step 1
                    slNameCode = tmVehicleCode(ilIndex).sKey    'Contract.lbcVehicle.List(ilTest)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If CInt(tgChfCntr.lVefCode) = Val(slCode) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilIndex
                'Check if NTR vehicle
                If (Not ilFound) And (tgSpf.sMktBase = "Y") Then
                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    ilVef = gBinarySearchVef(CInt(tgChfCntr.lVefCode))
                    If ilVef <> -1 Then
                        If tgMVef(ilVef).sType = "N" Then
                    '        If CInt(tgChfCntr.lVefCode) = tgMVef(ilVef).iCode Then
                                ilFound = True
                    '            Exit For
                    '        End If
                        End If
                    End If
                    'Next ilVef
                End If
                If (ilFound) And (tgSpf.sMktBase = "Y") Then
                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If CInt(tgChfCntr.lVefCode) = tgMVef(ilVef).iCode Then
                        ilVef = gBinarySearchVef(CInt(tgChfCntr.lVefCode))
                        If ilVef <> -1 Then
                            If (tgMVef(ilVef).sType <> "N") Or ((tgMVef(ilVef).sType = "N") And (tgMVef(ilVef).iMnfVehGp3Mkt > 0)) Then
                                For ilIndex = 0 To UBound(igCntrMktCode) - 1 Step 1
                                    If tgMVef(ilVef).iMnfVehGp3Mkt = igCntrMktCode(ilIndex) Then
                                        imMktFound = True
                                        Exit For
                                    End If
                                Next ilIndex
                            Else
                                imMktFound = True
                            End If
                    '        Exit For
                        End If
                    'Next ilVef
                End If
                If Not ilFound Then
                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If CInt(tgChfCntr.lVefCode) = tgMVef(ilVef).iCode Then
                        ilVef = gBinarySearchVef(CInt(tgChfCntr.lVefCode))
                        If ilVef <> -1 Then
                            If tgMVef(ilVef).sType <> "N" Then
                                slName = "999" & "|" & "999" & "|" & tgMVef(ilVef).sName & "\" & Trim$(str$(tgMVef(ilVef).iCode))
                                tmVehicleCode(UBound(tmVehicleCode)).sKey = slName
                                ReDim Preserve tmVehicleCode(0 To UBound(tmVehicleCode) + 1) As SORTCODE
                                Contract.lbcLnVehicle(igTabMapIndex).AddItem Trim$(tgMVef(ilVef).sName)
                                imVefDormantAdd = True
                            End If
                    '        Exit For
                        End If
                    'Next ilVef
                End If
            ElseIf tgChfCntr.lVefCode < 0 Then
                tmVsfSrchKey.lCode = -tgChfCntr.lVefCode
                ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    If tgSpf.sMktBase = "Y" Then
                        For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                            If (tmVsf.iFSCode(ilVsf) > 0) Then
                                ilFound = False
                                For ilIndex = 0 To UBound(tmVehicleCode) - 1 Step 1  'Contract.lbcVehicle.ListCount - 1 Step 1
                                    slNameCode = tmVehicleCode(ilIndex).sKey    'Contract.lbcVehicle.List(ilTest)
                                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                    If tmVsf.iFSCode(ilVsf) = Val(slCode) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilIndex
                                If (ilFound) Then
                                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                    '    If tmVsf.iFSCode(ilVsf) = tgMVef(ilVef).iCode Then
                                        ilVef = gBinarySearchVef(tmVsf.iFSCode(ilVsf))
                                        If ilVef <> -1 Then
                                            If (tgMVef(ilVef).sType <> "N") Or ((tgMVef(ilVef).sType = "N") And (tgMVef(ilVef).iMnfVehGp3Mkt > 0)) Then
                                                For ilIndex = 0 To UBound(igCntrMktCode) - 1 Step 1
                                                    If tgMVef(ilVef).iMnfVehGp3Mkt = igCntrMktCode(ilIndex) Then
                                                        imMktFound = True
                                                        Exit For
                                                    End If
                                                Next ilIndex
                                            Else
                                                imMktFound = True
                                            End If
                                    '        Exit For
                                        End If
                                    'Next ilVef
                                Else
                                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                    '    If tmVsf.iFSCode(ilVsf) = tgMVef(ilVef).iCode Then
                                        ilVef = gBinarySearchVef(tmVsf.iFSCode(ilVsf))
                                        If ilVef <> -1 Then
                                            If tgMVef(ilVef).sType = "N" Then
                                                If tgMVef(ilVef).iMnfVehGp3Mkt > 0 Then
                                                    For ilIndex = 0 To UBound(igCntrMktCode) - 1 Step 1
                                                        If tgMVef(ilVef).iMnfVehGp3Mkt = igCntrMktCode(ilIndex) Then
                                                            imMktFound = True
                                                            Exit For
                                                        End If
                                                    Next ilIndex
                                                Else
                                                    imMktFound = True
                                                End If
                                            End If
                                    '        Exit For
                                        End If
                                    'Next ilVef
                                End If
                                If imMktFound Then
                                    Exit For
                                End If
                            End If
                        Next ilVsf
                    End If
                    For ilVsf = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                        If (tmVsf.iFSCode(ilVsf) > 0) Then
                            ilFound = False
                            For ilIndex = 0 To UBound(tmVehicleCode) - 1 Step 1  'Contract.lbcVehicle.ListCount - 1 Step 1
                                slNameCode = tmVehicleCode(ilIndex).sKey    'Contract.lbcVehicle.List(ilTest)
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                If tmVsf.iFSCode(ilVsf) = Val(slCode) Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilIndex
                            If Not ilFound Then
                                'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                '    If tmVsf.iFSCode(ilVsf) = tgMVef(ilVef).iCode Then
                                    ilVef = gBinarySearchVef(tmVsf.iFSCode(ilVsf))
                                    If ilVef <> -1 Then
                                        If tgMVef(ilVef).sType <> "N" Then
                                            slName = "999" & "|" & "999" & "|" & tgMVef(ilVef).sName & "\" & Trim$(str$(tgMVef(ilVef).iCode))
                                            tmVehicleCode(UBound(tmVehicleCode)).sKey = slName
                                            ReDim Preserve tmVehicleCode(0 To UBound(tmVehicleCode) + 1) As SORTCODE
                                            Contract.lbcLnVehicle(igTabMapIndex).AddItem Trim$(tgMVef(ilVef).sName)
                                            imVefDormantAdd = True
                                        End If
                                '        Exit For
                                    End If
                                'Next ilVef
                            End If
                        End If
                    Next ilVsf
                End If
            End If
            Exit For
        End If
    Next ilLoop
    'gPDNToStr tgChfCntr.sInputGross, 2, slStr
    slStr = gLongToStrDec(tgChfCntr.lInputGross, 2)
    'Contract.edcGross.Text = slStr
    smHdSave(1) = slStr
    '8/16/12: Handle case where contract was set as Calendar but is not allowed at this time
    If (tgChfCntr.sBillCycle = "C") And ((Asc(tgSpf.sOptionFields) And OFCALENDARBILL) <> OFCALENDARBILL) Then
        Contract.mPopBillCycle True, True
    Else
        Contract.mPopBillCycle False, True
    End If
    Contract.lbcBillCycle.ListIndex = -1
    Select Case tgChfCntr.sBillCycle
        Case "S"  'Standard
            Contract.lbcBillCycle.ListIndex = 0
        Case "C"  'Calendar
            For ilLoop = 0 To Contract.lbcBillCycle.ListCount - 1 Step 1
                If Contract.lbcBillCycle.List(ilLoop) = "Calendar" Then
                    Contract.lbcBillCycle.ListIndex = ilLoop
                    Exit For
                End If
            Next ilLoop
        'Case "T"  'Fixed Price Standard
        '    Contract.lbcBillCycle.ListIndex = 2
        'Case "D"  'Fixed Price Calendar
        '    Contract.lbcBillCycle.ListIndex = 3
        Case "W"
            'If ((Asc(tgSpf.sOptionFields) And OFCALENDARBILL) = OFCALENDARBILL) Then
            '    Contract.lbcBillCycle.ListIndex = 2
            'Else
            '    Contract.lbcBillCycle.ListIndex = 1
            'End If
            For ilLoop = 0 To Contract.lbcBillCycle.ListCount - 1 Step 1
                If Contract.lbcBillCycle.List(ilLoop) = "Weekly" Then
                    Contract.lbcBillCycle.ListIndex = ilLoop
                    Exit For
                End If
            Next ilLoop
        Case Else
            Contract.lbcBillCycle.ListIndex = -1
    End Select
    Select Case tgChfCntr.sInvGp
        Case "A"  'Only one-all Spots
            Contract.lbcInvGroup.ListIndex = 0
        Case "P"  '1 per Product
            Contract.lbcInvGroup.ListIndex = 1
        Case "T"  '1 per Tag
            Contract.lbcInvGroup.ListIndex = 2
        Case Else
            Contract.lbcInvGroup.ListIndex = -1
    End Select
    'gPDNToStr tgChfCntr.sPctTag, 2, slStr
    '10043 iPctTag removed
'    slStr = gIntToStrDec(tgChfCntr.iPctTag, 2)
'    'Contract.edcProdTag.Text = slStr
'    smHdSave(3) = slStr
    smComment(2) = smInitComment(2) 'Internal
    smComment(3) = smInitComment(3) 'Cancel
    smComment(4) = smInitComment(4) 'Change Reason
    smComment(5) = smInitComment(5) 'Other
    '10910
    'For ilLoop = LBound(tmHdCtrls) To UBound(tmHdCtrls) Step 1
    For ilLoop = imLBCTCtrls To UBound(tmHdCtrls) Step 1
        tmHdCtrls(ilLoop).iChg = False
    Next ilLoop
    For ilLoop = imLBCTCtrls To UBound(tmCECtrls) Step 1
        tmCECtrls(ilLoop).iChg = False
    Next ilLoop
    For ilLoop = imLBCTCtrls To UBound(tmDmCtrls) Step 1
        tmDmCtrls(ilLoop).iChg = False
    Next ilLoop
    gUnpackDate tgChfCntr.iStartDate(0), tgChfCntr.iStartDate(1), smChfStartDate
    gUnpackDate tgChfCntr.iEndDate(0), tgChfCntr.iEndDate(1), smChfEndDate
    Contract.mPopRegion
    Exit Sub
mMoveRecToCtrlErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Public Sub mPBCDemoBarPaint(Index As Integer, imLnRowNo As Integer, imVefCode As Integer, imVpfIndex As Integer, imTerminate As Integer)
    Dim ilBox As Integer
    Dim llColor As Long
    Dim ilDate As Integer
    Dim llDate As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim llLnStartDate As Long
    Dim llLnEndDate As Long
    Dim ilWkIndex As Integer
    Dim ilIndex As Integer
    Dim llSpotCount As Long
    Dim slStr As String
    Dim slDate As String
    Dim ilAvail As Integer
    Dim llActPrice As Long
    Dim llRCPrice As Long
    Dim llGRP As Long
    Dim llGrImp As Long
    Dim ilLen As Integer
    Dim ilNo30 As Integer
    Dim slLineType As String
    Dim llSpotPrice As Long
    Dim ilFdIndex As Integer
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim ilClf As Integer
    Dim ilInclude As Integer
    Dim tmPopCtrl As FIELDAREA
    Dim blFontBold As Boolean
    Dim ilColTitle As Integer
    
    If Not Contract.pbcDemoBar(Index).Visible Then
        Exit Sub
    End If
    If (imLnRowNo < 1) Or (imLnRowNo > UBound(tgClfCntr) + 1) Then
        If imDBTotals < 5 Then
            Contract.pbcDemoBar(Index).Cls
            mPaintDemoBarGrid Index
        End If
        Exit Sub
    End If
    If tgClfCntr(imLnRowNo - 1).iCancel Then
        If imDBTotals < 5 Then
            Contract.pbcDemoBar(Index).Cls
            mPaintDemoBarGrid Index
            Exit Sub
        End If
    End If
    If (Trim$(smLnSave(1, imLnRowNo)) = "") Then
        If (imDBTotals < 5) Or (Index <= 1) Then
            Contract.pbcDemoBar(Index).Cls
            mPaintDemoBarGrid Index
            Exit Sub
        End If
    End If
    mPaintDemoBarGrid Index
    If Index = 0 Then
        'If (igLnSPChg) And (tgUrf(0).sUseComputeCMC = "I") Then
        '    Contract.pbcDemoBar(0).Cls
        'Else
            For ilBox = imLBCTCtrls To UBound(tmDBCtrls) - 1 Step 1
                Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(ilBox).fBoxX + fgBoxInsetX
                Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
                Contract.pbcDemoBar(0).Print tmDBCtrls(ilBox).sShow
            Next ilBox
        'End If
        'Paint Titles
        llColor = Contract.pbcDemoBar(0).ForeColor
        slFontName = Contract.pbcDemoBar(0).FontName
        flFontSize = Contract.pbcDemoBar(0).FontSize
        Contract.pbcDemoBar(0).ForeColor = BLUE
        Contract.pbcDemoBar(0).FontBold = False
        Contract.pbcDemoBar(0).FontSize = 7
        Contract.pbcDemoBar(0).FontName = "Arial"
        Contract.pbcDemoBar(0).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        Contract.pbcDemoBar(0).CurrentX = 3015   '3135   'tmDBCtrls(1).fBoxX + 30 '+ fgBoxInsetX
        Contract.pbcDemoBar(0).CurrentY = 95    '195    'tmDBCtrls(1).fBoxY - tmDBCtrls(1).fBoxY + 15'+ fgBoxInsetY
        slLineType = mGetLineType(imLnRowNo)
        Contract.pbcDemoBar(0).CurrentX = 30 '+ fgBoxInsetX
        Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(1).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
        Contract.pbcDemoBar(0).Print "Line"
        Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(4).fBoxX + tmDBCtrls(4).fBoxW + 30 '+ fgBoxInsetX
        Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(1).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
        If (slLineType = "H") Or (slLineType = "O") Or (slLineType = "A") Or (slLineType = "E") Then
            Contract.pbcDemoBar(0).Print "Pkg"
        Else
            Contract.pbcDemoBar(0).Print "Vehicle"
        End If
        Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(8).fBoxX + tmDBCtrls(8).fBoxW + 30 '+ fgBoxInsetX
        Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(1).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
        Contract.pbcDemoBar(0).Print "Total"
        Contract.pbcDemoBar(0).ForeColor = WHITE
        If imToggleCPPCPM = 0 Then
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(1).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(1).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "CPP"
            Contract.pbcDemoBar(0).ForeColor = BLUE
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(2).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(2).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "GRP"
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(3).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(3).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "%" 'GRP"
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(5).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(5).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "CPP"
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(6).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(6).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "GRP"
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(7).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(7).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "%" 'GRP"
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(9).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(9).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "CPP"
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(10).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(10).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "GRP"
        Else
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(1).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(1).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "CPM"
            Contract.pbcDemoBar(0).ForeColor = BLUE
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(2).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(2).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "GrImp"
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(3).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(3).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "%" 'GrImp"
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(5).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(5).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "CPM"
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(6).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(6).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "GrImp"
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(7).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(7).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "%" 'GrImp"
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(9).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(9).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "CPM"
            Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(10).fBoxX + 30 '+ fgBoxInsetX
            Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(10).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(0).Print "GrImp"
            
        End If
        Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(4).fBoxX + 30 '+ fgBoxInsetX
        Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(4).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
        Contract.pbcDemoBar(0).Print "Total Gross"
        Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(8).fBoxX + 30 '+ fgBoxInsetX
        Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(8).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
        Contract.pbcDemoBar(0).Print "Total Gross" 'GRP"
        Contract.pbcDemoBar(0).CurrentX = tmDBCtrls(11).fBoxX + 30 '+ fgBoxInsetX
        Contract.pbcDemoBar(0).CurrentY = tmDBCtrls(11).fBoxY - tmDBCtrls(1).fBoxY + 15 '+ fgBoxInsetY
        Contract.pbcDemoBar(0).Print "Total Gross"
        Contract.pbcDemoBar(0).FontSize = flFontSize
        Contract.pbcDemoBar(0).FontName = slFontName
        Contract.pbcDemoBar(0).FontSize = flFontSize
        Contract.pbcDemoBar(0).ForeColor = llColor
        Contract.pbcDemoBar(0).FontBold = True
    Else
        'If ((Index = 2) Or (Index = 3)) And igLnSPChg And (tgUrf(0).sUseComputeCMC = "I") Then
        '    Contract.pbcDemoBar(Index).Cls
        '    Exit Sub
        'End If
        'Build Header dates
        If (lmHSStartDate > 0) And (imHSCurWkNo > 0) Then
            llColor = Contract.pbcDemoBar(1).ForeColor
            slFontName = Contract.pbcDemoBar(1).FontName
            flFontSize = Contract.pbcDemoBar(1).FontSize
            Contract.pbcDemoBar(1).ForeColor = BLUE
            Contract.pbcDemoBar(1).FontBold = False
            Contract.pbcDemoBar(1).FontSize = 7
            Contract.pbcDemoBar(1).FontName = "Arial"
            Contract.pbcDemoBar(1).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
            'slPrevMonth = "x"
            For ilDate = 0 To 12 Step 1
                If imHSCurWkNo + ilDate > imHSEndWkNo Then
                    Exit For
                End If
                If ilDate = 0 Then
                    llDate = lmHSStartDate + 7 * (imHSCurWkNo - 1)
                End If
                If (ilDate = 0) Or (ilDate = 12) Then
                    slDate = Format$(llDate, "m/d")
                Else
                    slDate = Format$(llDate, "m/d/yy")
                End If
                'gObtainYearMonthDayStr slDate, False, slYear, slMonth, slDay
                'slMonth = gMonthName(slDate)
                'If slMonth <> slPrevMonth Then
                '    Contract.pbcDemoBar(Index).CurrentX = tmLnCtrls(WKSPOTSINDEX + ilDate).fBoxX + 15'fgBoxInsetX
                '    Contract.pbcDemoBar(Index).CurrentY = 15  '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
                '    Contract.pbcDemoBar(Index).Print Left$(slMonth, 3)
                '    slPrevMonth = slMonth
                'End If
                'If tgSpf.sCLnStdQt = "Y" Then
                    'Determine start of Quarter
                    gObtainMonthYear 0, Format$(llDate, "m/d/yy"), ilMonth, ilYear
                    If ilMonth <= 3 Then
                        Contract.pbcDemoBar(1).ForeColor = BLUE
                    ElseIf ilMonth <= 6 Then
                        Contract.pbcDemoBar(1).ForeColor = MAGENTA
                    ElseIf ilMonth <= 9 Then
                        Contract.pbcDemoBar(1).ForeColor = DARKGREEN
                    Else
                        Contract.pbcDemoBar(1).ForeColor = Red
                    End If
                'End If
                If ilDate = 0 Then
                    Contract.pbcDemoBar(1).CurrentX = tmDBSCCtrls(ilDate + 1).fBoxX + Contract.pbcLnWkArrow(2).Width + 30   '255 'fgBoxInsetX
                Else
                    Contract.pbcDemoBar(1).CurrentX = tmDBSCCtrls(ilDate + 1).fBoxX + 15 'fgBoxInsetX
                End If
                Contract.pbcDemoBar(1).CurrentY = 15
                Contract.pbcDemoBar(1).Print slDate
                llDate = llDate + 7
            Next ilDate
            Contract.pbcDemoBar(1).CurrentX = 45
            Contract.pbcDemoBar(1).CurrentY = tmDBRCtrls(1).fBoxY - 15 '+ fgBoxInsetY
            'If imDBRatio = 0 Then
            '    Contract.pbcDemoBar(Index).Print "Actual/Rate Card"
            'Else
            '    Contract.pbcDemoBar(Index).Print "Actual/Requeted"
            'End If
            Contract.pbcDemoBar(1).FontSize = flFontSize
            Contract.pbcDemoBar(1).FontName = slFontName
            Contract.pbcDemoBar(1).FontSize = flFontSize
            Contract.pbcDemoBar(1).ForeColor = llColor
            Contract.pbcDemoBar(1).FontBold = True
        End If
        llColor = Contract.pbcDemoBar(1).ForeColor
        slFontName = Contract.pbcDemoBar(1).FontName
        flFontSize = Contract.pbcDemoBar(1).FontSize
        Contract.pbcDemoBar(1).ForeColor = BLUE
        Contract.pbcDemoBar(1).FontBold = False
        Contract.pbcDemoBar(1).FontSize = 7
        Contract.pbcDemoBar(1).FontName = "Arial"
        Contract.pbcDemoBar(1).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
        If Index = 1 Then
            If imDBTotals < 5 Then
                Contract.pbcDemoBar(1).CurrentX = tmDBAvgCtrls(1).fBoxX + 30 '+ fgBoxInsetX
                Contract.pbcDemoBar(1).CurrentY = tmDBAPCtrls(1).fBoxY - 15 '+ fgBoxInsetY
                Contract.pbcDemoBar(1).Print "Actual Spot Price"
                Contract.pbcDemoBar(1).CurrentX = tmDBAvgCtrls(1).fBoxX + 30 '+ fgBoxInsetX
                Contract.pbcDemoBar(1).CurrentY = tmDBRCPCtrls(1).fBoxY - 15 '+ fgBoxInsetY
                Contract.pbcDemoBar(1).Print "Rate Card Price"
                Contract.pbcDemoBar(1).CurrentX = tmDBAvgCtrls(1).fBoxX + 30 '+ fgBoxInsetX
                Contract.pbcDemoBar(1).CurrentY = tmDBSCCtrls(1).fBoxY - 15 '+ fgBoxInsetY
                Contract.pbcDemoBar(1).Print "Spot Count"
                Contract.pbcDemoBar(1).CurrentX = tmDBCtrls(12).fBoxX + 30 '+ fgBoxInsetX
                Contract.pbcDemoBar(1).CurrentY = tmDBGRPCtrls(1).fBoxY - 15 '+ fgBoxInsetY
                Contract.pbcDemoBar(1).Print "GRP"
                Contract.pbcDemoBar(1).CurrentX = tmDBCtrls(12).fBoxX + 30 '+ fgBoxInsetX
                Contract.pbcDemoBar(1).CurrentY = tmDBGICtrls(1).fBoxY - 15 '+ fgBoxInsetY
                Contract.pbcDemoBar(1).Print "GrImp"
            End If
        Else
            If imDBTotals >= 5 Then
                gPaintArea Contract.pbcDemoBar(1), tmDBAvgCtrls(1).fBoxX, tmDBSCCtrls(1).fBoxY, tmDBAvgCtrls(1).fBoxW - 15, tmDBAvgCtrls(1).fBoxH - 15, WHITE
                Contract.pbcDemoBar(1).CurrentX = tmDBAvgCtrls(1).fBoxX + 30 '+ fgBoxInsetX
                Contract.pbcDemoBar(1).CurrentY = tmDBSCCtrls(1).fBoxY - 15 '+ fgBoxInsetY
                Contract.pbcDemoBar(1).Print "Actual Spot Price"
            End If
        End If
        If tgClfCntr(imLnRowNo - 1).iAvgRating > 0 Then
            '10/28/14: Handle 2 place Rating
            'slStr = "Rating " & gIntToStrDec(tgClfCntr(imLnRowNo - 1).iAvgRating, 1)
            If smCntr1or2PlaceRating <> "2" Then
                slStr = "Rating " & gIntToStrDec(tgClfCntr(imLnRowNo - 1).iAvgRating, 1)
            Else
                slStr = "Rating " & gIntToStrDec(tgClfCntr(imLnRowNo - 1).iAvgRating, 2)
            End If
        Else
            slStr = "Rating"
        End If
        gSetShow Contract.pbcDemoBar(1), slStr, tmDBAvgCtrls(1)
        Contract.pbcDemoBar(1).CurrentX = tmDBAvgCtrls(1).fBoxX + 30 '+ fgBoxInsetX
        Contract.pbcDemoBar(1).CurrentY = tmDBAvgCtrls(1).fBoxY - 15 '+ fgBoxInsetY
        Contract.pbcDemoBar(1).Print tmDBAvgCtrls(1).sShow
        If tgClfCntr(imLnRowNo - 1).lPop > 0 Then
'            If tgSpf.sSAudData <> "H" Then
'                slStr = "Pop " & gLongToStrDec(tgClfCntr(imLnRowNo - 1).lPop, 0)
'            Else
'                slStr = "Pop " & gLongToStrDec(tgClfCntr(imLnRowNo - 1).lPop, 1)
'            End If
            If tgSpf.sSAudData = "H" Then
                slStr = "Pop " & gLongToStrDec(tgClfCntr(imLnRowNo - 1).lPop, 1)
            ElseIf tgSpf.sSAudData = "N" Then
                slStr = "Pop " & gLongToStrDec(tgClfCntr(imLnRowNo - 1).lPop, 2)
            ElseIf tgSpf.sSAudData = "U" Then
                slStr = "Pop " & gLongToStrDec(tgClfCntr(imLnRowNo - 1).lPop, 3)
            Else
                slStr = "Pop " & gLongToStrDec(tgClfCntr(imLnRowNo - 1).lPop, 0)
            End If
            '4/26/10:  Added Mix to name if required
            For ilClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                '6/16/10: Bypass package line as the population might not be set.
                ilInclude = True
                slLineType = mGetLineType(ilClf + 1, True)
                If (slLineType = "H") And (tgSaf(0).sAudByPackage = "Y") Then
                    ilInclude = False
                End If
                If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") And (Not tgClfCntr(ilClf).iCancel) And ((tgClfCntr(ilClf).iStatus = 0) Or (tgClfCntr(ilClf).iStatus = 1)) And (ilInclude) Then
                    If (tgClfCntr(ilClf).lPop <> tgClfCntr(imLnRowNo - 1).lPop) And (Not tgClfCntr(imLnRowNo - 1).iCancel) And ((tgClfCntr(imLnRowNo - 1).iStatus = 0) Or (tgClfCntr(imLnRowNo - 1).iStatus = 1)) Then
                        slStr = slStr & " Mixed"
                        Exit For
                    End If
                End If
            Next ilClf
        Else
            slStr = "Pop"
        End If
        '4/26/10:  Added Mix to name if required
        'gSetShow Contract.pbcDemoBar(1), slStr, tmDBAvgCtrls(2)
        'Contract.pbcDemoBar(1).CurrentX = tmDBAvgCtrls(2).fBoxX + 30 '+ fgBoxInsetX
        'Contract.pbcDemoBar(1).CurrentY = 15 'tmDBAvgCtrls(2).fBoxY - 15 '+ fgBoxInsetY
        'Contract.pbcDemoBar(1).Print tmDBAvgCtrls(2).sShow
        tmPopCtrl = tmDBAvgCtrls(2)
        tmPopCtrl.fBoxW = tmDBAvgCtrls(2).fBoxW + tmDBAvgCtrls(2).fBoxW / 3
        gSetShow Contract.pbcDemoBar(1), slStr, tmPopCtrl
        Contract.pbcDemoBar(1).CurrentX = tmPopCtrl.fBoxX + 30 '+ fgBoxInsetX
        Contract.pbcDemoBar(1).CurrentY = 15 'tmDBAvgCtrls(2).fBoxY - 15 '+ fgBoxInsetY
        Contract.pbcDemoBar(1).Print tmPopCtrl.sShow
        If tgClfCntr(imLnRowNo - 1).lAvgAud > 0 Then
'            If tgSpf.sSAudData <> "H" Then
'                slStr = "Aud " & gLongToStrDec(tgClfCntr(imLnRowNo - 1).lAvgAud, 0)
'            Else
'                slStr = "Aud " & gLongToStrDec(tgClfCntr(imLnRowNo - 1).lAvgAud, 1)
'            End If
            If tgSpf.sSAudData = "H" Then
                slStr = "Aud " & gLongToStrDec(tgClfCntr(imLnRowNo - 1).lAvgAud, 1)
            ElseIf tgSpf.sSAudData = "N" Then
                slStr = "Aud " & gLongToStrDec(tgClfCntr(imLnRowNo - 1).lAvgAud, 2)
            ElseIf tgSpf.sSAudData = "U" Then
                slStr = "Aud " & gLongToStrDec(tgClfCntr(imLnRowNo - 1).lAvgAud, 3)
            Else
                slStr = "Aud " & gLongToStrDec(tgClfCntr(imLnRowNo - 1).lAvgAud, 0)
            End If
        Else
            slStr = "Aud"
        End If
        gSetShow Contract.pbcDemoBar(1), slStr, tmDBAvgCtrls(2)
        Contract.pbcDemoBar(1).CurrentX = tmDBAvgCtrls(2).fBoxX + 30 '+ fgBoxInsetX
        Contract.pbcDemoBar(1).CurrentY = tmDBAvgCtrls(2).fBoxY - 15 '+ fgBoxInsetY
        Contract.pbcDemoBar(1).Print tmDBAvgCtrls(2).sShow
        Contract.pbcDemoBar(1).FontSize = flFontSize
        Contract.pbcDemoBar(1).FontName = slFontName
        Contract.pbcDemoBar(1).FontSize = flFontSize
        Contract.pbcDemoBar(1).ForeColor = llColor
        Contract.pbcDemoBar(1).FontBold = True
        If imDBTotals = 6 Then
            llColor = Contract.pbcDemoBar(1).ForeColor
            slFontName = Contract.pbcDemoBar(1).FontName
            flFontSize = Contract.pbcDemoBar(1).FontSize
            Contract.pbcDemoBar(3).ForeColor = BLUE
            Contract.pbcDemoBar(3).FontBold = False
            Contract.pbcDemoBar(3).FontSize = 7
            Contract.pbcDemoBar(3).FontName = "Arial"
            Contract.pbcDemoBar(3).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
            Contract.pbcDemoBar(3).CurrentX = 30 'tmDBAvgCtrls(1).fBoxX + 30'+ fgBoxInsetX
            Contract.pbcDemoBar(3).CurrentY = 15    'tmDBSCCtrls(1).fBoxY - 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(3).Print "Weekly Dates"
            Contract.pbcDemoBar(3).FontSize = flFontSize
            Contract.pbcDemoBar(3).FontName = slFontName
            Contract.pbcDemoBar(3).FontSize = flFontSize
            Contract.pbcDemoBar(3).ForeColor = llColor
            Contract.pbcDemoBar(3).FontBold = True
        ElseIf imDBTotals = 7 Then
            llColor = Contract.pbcDemoBar(1).ForeColor
            slFontName = Contract.pbcDemoBar(1).FontName
            flFontSize = Contract.pbcDemoBar(1).FontSize
            Contract.pbcDemoBar(3).ForeColor = BLUE
            Contract.pbcDemoBar(3).FontBold = False
            Contract.pbcDemoBar(3).FontSize = 7
            Contract.pbcDemoBar(3).FontName = "Arial"
            Contract.pbcDemoBar(3).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
            Contract.pbcDemoBar(3).CurrentX = 30 'tmDBAvgCtrls(1).fBoxX + 30'+ fgBoxInsetX
            Contract.pbcDemoBar(3).CurrentY = 15    'tmDBSCCtrls(1).fBoxY - 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(3).Print "Quarterly Dates"
            Contract.pbcDemoBar(3).FontSize = flFontSize
            Contract.pbcDemoBar(3).FontName = slFontName
            Contract.pbcDemoBar(3).FontSize = flFontSize
            Contract.pbcDemoBar(3).ForeColor = llColor
            Contract.pbcDemoBar(3).FontBold = True
        ElseIf imDBTotals = 8 Then
            llColor = Contract.pbcDemoBar(1).ForeColor
            slFontName = Contract.pbcDemoBar(1).FontName
            flFontSize = Contract.pbcDemoBar(1).FontSize
            Contract.pbcDemoBar(3).ForeColor = BLUE
            Contract.pbcDemoBar(3).FontBold = False
            Contract.pbcDemoBar(3).FontSize = 7
            Contract.pbcDemoBar(3).FontName = "Arial"
            Contract.pbcDemoBar(3).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
            Contract.pbcDemoBar(3).CurrentX = 30 'tmDBAvgCtrls(1).fBoxX + 30'+ fgBoxInsetX
            Contract.pbcDemoBar(3).CurrentY = 15    'tmDBSCCtrls(1).fBoxY - 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(3).Print "Vehicles"
            Contract.pbcDemoBar(3).FontSize = flFontSize
            Contract.pbcDemoBar(3).FontName = slFontName
            Contract.pbcDemoBar(3).FontSize = flFontSize
            Contract.pbcDemoBar(3).ForeColor = llColor
            Contract.pbcDemoBar(3).FontBold = True
        ElseIf imDBTotals = 9 Then
            llColor = Contract.pbcDemoBar(1).ForeColor
            slFontName = Contract.pbcDemoBar(1).FontName
            flFontSize = Contract.pbcDemoBar(1).FontSize
            Contract.pbcDemoBar(3).ForeColor = BLUE
            Contract.pbcDemoBar(3).FontBold = False
            Contract.pbcDemoBar(3).FontSize = 7
            Contract.pbcDemoBar(3).FontName = "Arial"
            Contract.pbcDemoBar(3).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
            Contract.pbcDemoBar(3).CurrentX = 30 'tmDBAvgCtrls(1).fBoxX + 30'+ fgBoxInsetX
            Contract.pbcDemoBar(3).CurrentY = 15    'tmDBSCCtrls(1).fBoxY - 15 '+ fgBoxInsetY
            slStr = Contract.edcDBDisplay.Text
            If slStr <> "Daypart Group" Then
                Contract.pbcDemoBar(3).Print "Dayparts"
            Else
                Contract.pbcDemoBar(3).Print "Daypart Group"
            End If
            Contract.pbcDemoBar(3).FontSize = flFontSize
            Contract.pbcDemoBar(3).FontName = slFontName
            Contract.pbcDemoBar(3).FontSize = flFontSize
            Contract.pbcDemoBar(3).ForeColor = llColor
            Contract.pbcDemoBar(3).FontBold = True
        ElseIf imDBTotals = 10 Then
            llColor = Contract.pbcDemoBar(1).ForeColor
            slFontName = Contract.pbcDemoBar(1).FontName
            flFontSize = Contract.pbcDemoBar(1).FontSize
            Contract.pbcDemoBar(3).ForeColor = BLUE
            Contract.pbcDemoBar(3).FontBold = False
            Contract.pbcDemoBar(3).FontSize = 7
            Contract.pbcDemoBar(3).FontName = "Arial"
            Contract.pbcDemoBar(3).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
            Contract.pbcDemoBar(3).CurrentX = 30 'tmDBAvgCtrls(1).fBoxX + 30'+ fgBoxInsetX
            Contract.pbcDemoBar(3).CurrentY = 15    'tmDBSCCtrls(1).fBoxY - 15 '+ fgBoxInsetY
            Contract.pbcDemoBar(3).Print "Vehicles"
            Contract.pbcDemoBar(3).FontSize = flFontSize
            Contract.pbcDemoBar(3).FontName = slFontName
            Contract.pbcDemoBar(3).FontSize = flFontSize
            Contract.pbcDemoBar(3).ForeColor = llColor
            Contract.pbcDemoBar(3).FontBold = True
        End If
        If imDBTotals = 5 Then
            llColor = Contract.pbcDemoBar(1).ForeColor
            slFontName = Contract.pbcDemoBar(1).FontName
            flFontSize = Contract.pbcDemoBar(1).FontSize
            Contract.pbcDemoBar(2).ForeColor = BLUE
            Contract.pbcDemoBar(2).FontBold = False
            Contract.pbcDemoBar(2).FontSize = 7
            Contract.pbcDemoBar(2).FontName = "Arial"
            Contract.pbcDemoBar(2).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
            ilColTitle = 1
            For ilBox = imLBCTCtrls To UBound(tmSumCtrls) Step 1
                Contract.pbcDemoBar(2).CurrentX = tmSumCtrls(ilBox).fBoxX + fgBoxInsetX
                Contract.pbcDemoBar(2).CurrentY = 15
                If (ilColTitle = 1) Then Contract.pbcDemoBar(2).Print "#"
                If (ilColTitle = 2) Then Contract.pbcDemoBar(2).Print "Vehicle"
                If (ilColTitle = 3) Then Contract.pbcDemoBar(2).Print "Daypart"
                If (ilColTitle = 4) Then Contract.pbcDemoBar(2).Print "GRP"
                If (ilColTitle = 5) Then Contract.pbcDemoBar(2).Print "% GRP"
                If (ilColTitle = 6) Then Contract.pbcDemoBar(2).Print "GrImp"
                If (ilColTitle = 7) Then Contract.pbcDemoBar(2).Print "% GrImp"
                If (ilColTitle = 8) Then Contract.pbcDemoBar(2).Print "CPM"
                If (ilColTitle = 9) Then Contract.pbcDemoBar(2).Print "CPP"
                If (ilColTitle = 10) Then Contract.pbcDemoBar(2).Print "Avg Price"
                If (ilColTitle = 11) Then Contract.pbcDemoBar(2).Print "Total Gross"
                ilColTitle = ilColTitle + 1
            Next ilBox
            
            Contract.pbcDemoBar(2).FontSize = flFontSize
            Contract.pbcDemoBar(2).FontName = slFontName
            Contract.pbcDemoBar(2).FontSize = flFontSize
            Contract.pbcDemoBar(2).ForeColor = llColor
            Contract.pbcDemoBar(2).FontBold = True
        End If
        
        If imDBTotals >= 6 And imDBTotals <= 10 Then
            llColor = Contract.pbcDemoBar(1).ForeColor
            slFontName = Contract.pbcDemoBar(1).FontName
            flFontSize = Contract.pbcDemoBar(1).FontSize
            Contract.pbcDemoBar(3).ForeColor = BLUE
            Contract.pbcDemoBar(3).FontBold = False
            Contract.pbcDemoBar(3).FontSize = 7
            Contract.pbcDemoBar(3).FontName = "Arial"
            Contract.pbcDemoBar(3).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
            ilColTitle = 1
            For ilBox = imLBCTCtrls + 3 To UBound(tmSumCtrls) Step 1
                Contract.pbcDemoBar(3).CurrentX = tmSumCtrls(ilBox).fBoxX + fgBoxInsetX
                Contract.pbcDemoBar(3).CurrentY = 15
                If (ilColTitle = 1) Then Contract.pbcDemoBar(3).Print "GRP"
                If (ilColTitle = 2) Then Contract.pbcDemoBar(3).Print "% GRP"
                If (ilColTitle = 3) Then Contract.pbcDemoBar(3).Print "GrImp"
                If (ilColTitle = 4) Then Contract.pbcDemoBar(3).Print "% GrImp"
                If (ilColTitle = 5) Then Contract.pbcDemoBar(3).Print "CPM"
                If (ilColTitle = 6) Then Contract.pbcDemoBar(3).Print "CPP"
                If (ilColTitle = 7) Then Contract.pbcDemoBar(3).Print "Avg Price"
                If (ilColTitle = 8) Then Contract.pbcDemoBar(3).Print "Total Gross"
                ilColTitle = ilColTitle + 1
            Next ilBox
            
            Contract.pbcDemoBar(3).FontSize = flFontSize
            Contract.pbcDemoBar(3).FontName = slFontName
            Contract.pbcDemoBar(3).FontSize = flFontSize
            Contract.pbcDemoBar(3).ForeColor = llColor
            Contract.pbcDemoBar(3).FontBold = True
        End If
        If (imDBTotals < 5) Or (Trim$(smLnSave(1, imLnRowNo)) <> "") Then
        
            llColor = Contract.pbcDemoBar(1).ForeColor
            slFontName = Contract.pbcDemoBar(1).FontName
            flFontSize = Contract.pbcDemoBar(1).FontSize
            blFontBold = Contract.pbcDemoBar(1).FontBold
            Contract.pbcDemoBar(1).ForeColor = BLUE
            Contract.pbcDemoBar(1).FontBold = True 'False
            Contract.pbcDemoBar(1).FontSize = 7
            Contract.pbcDemoBar(1).FontName = "Arial"
            Contract.pbcDemoBar(1).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
            'Contract.pbcDemoBar(1).Line (30, 420)-Step(1395, 180), BLUE, BF
            Contract.pbcDemoBar(1).Line (tmAvToggleCtrls(1).fBoxX, tmAvToggleCtrls(1).fBoxY)-Step(tmAvToggleCtrls(1).fBoxW, tmAvToggleCtrls(1).fBoxH), BLUE, BF
            Contract.pbcDemoBar(1).ForeColor = WHITE
            Contract.pbcDemoBar(1).CurrentX = fgBoxInsetX + 15
            Contract.pbcDemoBar(1).CurrentY = tmDBAvCtrls(1).fBoxY
            If imToggleAvail = 0 Then
                Contract.pbcDemoBar(1).Print "Avails: # " & Trim$(smLnSave(16, imLnRowNo))
            ElseIf imToggleAvail = 1 Then
                Contract.pbcDemoBar(1).Print "Avails: 30/60"
            ElseIf imToggleAvail = 2 Then
                Contract.pbcDemoBar(1).Print "Avails: 10s break"
            Else
                Contract.pbcDemoBar(1).Print "Avails: None"
            End If
            Contract.pbcDemoBar(1).FontSize = flFontSize
            Contract.pbcDemoBar(1).FontName = slFontName
            Contract.pbcDemoBar(1).FontSize = flFontSize
            Contract.pbcDemoBar(1).ForeColor = llColor
            Contract.pbcDemoBar(1).FontBold = blFontBold
        
            slLineType = mGetLineType(imLnRowNo)
            If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                mCVehOptIndex imLnRowNo, imVefCode, imVpfIndex, imTerminate
                ilNo30 = 0
                ilLen = Val(smLnSave(16, imLnRowNo))
                If ilLen = 0 Then
                    If imLnRowNo > 1 Then
                        ilLen = Val(smLnSave(16, imLnRowNo - 1))
                    End If
                    If ilLen = 0 Then
                        ilLen = tgVpf(imVpfIndex).iSDLen
                    End If
                    If ilLen = 0 Then
                        ilNo30 = 1
                    End If
                End If
                If (tgVpf(imVpfIndex).sSSellOut = "B") Or (tgVpf(imVpfIndex).sSSellOut = "U") Then
                    If (ilLen Mod 30) = 0 Then
                        Do While ilLen >= 30
                            ilNo30 = ilNo30 + 1
                            ilLen = ilLen - 30
                        Loop
                    End If
                ElseIf tgVpf(imVpfIndex).sSSellOut = "M" Then
                    If (ilLen Mod 30) = 0 Then
                        Do While ilLen >= 30
                            ilNo30 = ilNo30 + 1
                            ilLen = ilLen - 30
                        Loop
                    End If
                ElseIf tgVpf(imVpfIndex).sSSellOut = "T" Then
                End If
            End If
            For ilBox = imLBCTCtrls To UBound(tmDBSCCtrls) Step 1
                llDate = lmHSStartDate + 7 * (imHSCurWkNo + ilBox - 2)
                llLnStartDate = gDateValue(gObtainPrevMonday(smLnSave(14, imLnRowNo)))
                llLnEndDate = gDateValue(gObtainPrevMonday(smLnSave(15, imLnRowNo)))
                ilWkIndex = (llDate - lmHSStartDate) \ 7 + 1
                slStr = ""
                llSpotCount = -1
                ilAvail = -29999
                llActPrice = -10
                llRCPrice = -1
                llGRP = -1
                llGrImp = -1
                ilFdIndex = False
                'If (llDate >= llLnStartDate) And (llDate <= llLnEndDate) Then
                    If ilWkIndex > 0 Then
                        ilIndex = imLnSave(10, imLnRowNo)
                        Do While ilIndex > 0
                            If ilWkIndex - 1 <= UBound(tmWkInfo(ilIndex).lSpotCount) Then
                                If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                                    ilFdIndex = True
                                    llSpotCount = tmWkInfo(ilIndex).lSpotCount(ilWkIndex - 1)
                                    ilAvail = tmWkInfo(ilIndex).iAvail(ilWkIndex - 1)
                                    llActPrice = tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1)
                                    llRCPrice = tmWkInfo(ilIndex).lRCPrice(ilWkIndex - 1)
                                    'If (tgClfCntr(imLnRowNo - 1).iSPChg = False) Or (tgUrf(0).sUseComputeCMC <> "I") Then
                                        llGRP = tmWkInfo(ilIndex).lGRP(ilWkIndex - 1)
                                        llGrImp = tmWkInfo(ilIndex).lGrImp(ilWkIndex - 1)
                                    'End If
                                Else
                                    llSpotCount = tmWkInfo(ilIndex).lSpotCount(ilWkIndex - 1)
                                    llActPrice = tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1)
                                    llGRP = tmWkInfo(ilIndex).lGRP(ilWkIndex - 1)
                                    llGrImp = tmWkInfo(ilIndex).lGrImp(ilWkIndex - 1)
                                End If
                                Exit Do
                            Else
                                ilWkIndex = ilWkIndex - UBound(tmWkInfo(ilIndex).lSpotCount) - 1
                                If ilWkIndex <= 0 Then
                                    Exit Do
                                End If
                            End If
                            ilIndex = tmWkInfo(ilIndex).iNextIndex
                        Loop
                    End If
                'End If
                llSpotCount = mGetFlightSpots(imLnRowNo, llDate, llSpotPrice)
                If llSpotPrice <> 0 Then
                    llActPrice = llSpotPrice
                    If ilFdIndex Then
                        tmWkInfo(ilIndex).lSpotCount(ilWkIndex - 1) = llSpotCount
                        tmWkInfo(ilIndex).lActPrice(ilWkIndex - 1) = llSpotPrice
                    End If
                End If
                If Index = 1 Then
                    If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                        Contract.pbcDemoBar(1).ForeColor = BLACK
                        If llSpotCount > 0 Then
                            slStr = Trim$(str$(llSpotCount))
                            If ilAvail <> -29999 Then
                                If ilNo30 > 0 Then
                                    If llSpotCount > ilAvail Then   '\ ilNo30 Then
                                        Contract.pbcDemoBar(1).ForeColor = Red
                                    End If
                                Else
                                    If llSpotCount > ilAvail Then
                                        Contract.pbcDemoBar(1).ForeColor = Red
                                    End If
                                End If
                            End If
                        Else
                            slStr = ""
                        End If
                    Else
                        Contract.pbcDemoBar(1).ForeColor = BLACK
                        If llSpotCount > 0 Then
                            slStr = Trim$(str$(llSpotCount))
                        Else
                            slStr = ""
                        End If
                    End If
                    gSetShow Contract.pbcDemoBar(1), slStr, tmDBSCCtrls(ilBox)
                    Contract.pbcDemoBar(1).CurrentX = tmDBSCCtrls(ilBox).fBoxX + fgBoxInsetX
                    Contract.pbcDemoBar(1).CurrentY = tmDBSCCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
                    Contract.pbcDemoBar(1).Print tmDBSCCtrls(ilBox).sShow
                    Contract.pbcDemoBar(1).ForeColor = BLACK
                    If (slLineType <> "O") And (slLineType <> "A") And (slLineType <> "E") Then
                        If ilAvail <> -29999 Then
                            If ilAvail < 0 Then
                                slStr = Trim$(str$(ilAvail))

                            Else
                                slStr = Trim$(str$(ilAvail))
                            End If
                            If tgUrf(0).iRemoteUserID > 0 Then
                                If ilAvail >= 255 Then
                                    slStr = "255+"
                                ElseIf ilAvail <= -50 Then
                                    slStr = "-50+"
                                End If
                            End If
                        Else
                            slStr = ""
                        End If
                    Else
                        slStr = ""
                    End If
                    gSetShow Contract.pbcDemoBar(1), slStr, tmDBAvCtrls(ilBox)
                    Contract.pbcDemoBar(1).CurrentX = tmDBAvCtrls(ilBox).fBoxX + fgBoxInsetX
                    Contract.pbcDemoBar(1).CurrentY = tmDBAvCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
                    Contract.pbcDemoBar(1).Print tmDBAvCtrls(ilBox).sShow
                    Contract.pbcDemoBar(1).ForeColor = BLACK
                End If
                If imDBTotals < 5 Then
                    gPaintArea Contract.pbcDemoBar(1), tmDBAPCtrls(ilBox).fBoxX, tmDBAPCtrls(ilBox).fBoxY, tmDBAPCtrls(ilBox).fBoxW - 15, tmDBAPCtrls(ilBox).fBoxH - 15, WHITE
                Else
                    gPaintArea Contract.pbcDemoBar(1), tmDBAPCtrls(ilBox).fBoxX, tmDBSCCtrls(ilBox).fBoxY, tmDBAPCtrls(ilBox).fBoxW - 15, tmDBAPCtrls(ilBox).fBoxH - 15, WHITE
                End If
                If llActPrice >= 0 Then
                    slStr = gLongToStrDec(llActPrice, 2)
                    'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 2, slStr
                    gFormatStr slStr, FMTLEAVEBLANK, 2, slStr
                Else
                    Select Case llActPrice
                        Case -1
                            slStr = Contract.lbcPriceType(1).List(1)
                        Case -2
                            slStr = Contract.lbcPriceType(1).List(2)
                        Case -3
                            slStr = Contract.lbcPriceType(1).List(3)
                        Case -4
                            slStr = Contract.lbcPriceType(1).List(4)
                        Case -5
                            slStr = Contract.lbcPriceType(1).List(5)
                        Case -6
                            slStr = Contract.lbcPriceType(1).List(6)
                        Case -7
                            slStr = Contract.lbcPriceType(1).List(7)
                        Case Else
                            slStr = ""
                    End Select
                End If
                gSetShow Contract.pbcDemoBar(1), slStr, tmDBAPCtrls(ilBox)
                Contract.pbcDemoBar(1).CurrentX = tmDBAPCtrls(ilBox).fBoxX + fgBoxInsetX
                If imDBTotals < 5 Then
                    Contract.pbcDemoBar(1).CurrentY = tmDBAPCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
                Else
                    Contract.pbcDemoBar(1).CurrentY = tmDBSCCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
                End If
                Contract.pbcDemoBar(1).Print tmDBAPCtrls(ilBox).sShow
                If imDBTotals < 5 Then
                    If (llRCPrice >= 0) And (slLineType <> "O") And (slLineType <> "A") Then
                        slStr = gLongToStrDec(llRCPrice, 0) '2)
                        'gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA, 0, slStr    '2, slStr
                        gFormatStr slStr, FMTLEAVEBLANK, 0, slStr    '2, slStr
                    Else
                        slStr = ""
                    End If
                    gSetShow Contract.pbcDemoBar(1), slStr, tmDBRCPCtrls(ilBox)
                    Contract.pbcDemoBar(1).CurrentX = tmDBRCPCtrls(ilBox).fBoxX + fgBoxInsetX
                    Contract.pbcDemoBar(1).CurrentY = tmDBRCPCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
                    Contract.pbcDemoBar(1).Print tmDBRCPCtrls(ilBox).sShow
                    gPaintArea Contract.pbcDemoBar(1), tmDBRCtrls(ilBox).fBoxX, tmDBRCtrls(ilBox).fBoxY, tmDBRCtrls(ilBox).fBoxW - 15, tmDBRCtrls(ilBox).fBoxH - 15, WHITE
                    slStr = ""
                    If imDBRatio = 0 Then
                        'Actual/Rate Card
                        If (llRCPrice > 0) And (llActPrice >= 0) Then
                            slStr = gLongToStrDec(llActPrice / llRCPrice, 2) '100 * llActPrice / llRCPrice, 2)
                        End If
                    Else
                        'Old Actual/Actual
                    End If
                    gSetShow Contract.pbcDemoBar(1), slStr, tmDBRCtrls(ilBox)
                    Contract.pbcDemoBar(1).CurrentX = tmDBRCtrls(ilBox).fBoxX + fgBoxInsetX
                    Contract.pbcDemoBar(1).CurrentY = tmDBRCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
                    Contract.pbcDemoBar(1).Print tmDBRCtrls(ilBox).sShow
                    slStr = ""
                    If llGRP > 0 Then
                        slStr = gLongToStrDec(llGRP, 1)
                    End If
                    gSetShow Contract.pbcDemoBar(1), slStr, tmDBGRPCtrls(ilBox)
                    Contract.pbcDemoBar(1).CurrentX = tmDBGRPCtrls(ilBox).fBoxX + fgBoxInsetX
                    Contract.pbcDemoBar(1).CurrentY = tmDBGRPCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
                    Contract.pbcDemoBar(1).Print tmDBGRPCtrls(ilBox).sShow
                    slStr = ""
                    If llGrImp > 0 Then
'                        If tgSpf.sSAudData <> "H" Then
'                            slStr = Trim$(Str$(llGrImp))
'                        Else
'                            slStr = gLongToStrDec(llGrImp, 1)
'                        End If
                        If tgSpf.sSAudData = "H" Then
                            slStr = gLongToStrDec(llGrImp, 1)
                        ElseIf tgSpf.sSAudData = "N" Then
                            slStr = gLongToStrDec(llGrImp, 2)
                        ElseIf tgSpf.sSAudData = "U" Then
                            slStr = gLongToStrDec(llGrImp, 3)
                        Else
                            slStr = Trim$(str$(llGrImp))
                        End If
                    End If
                    gSetShow Contract.pbcDemoBar(1), slStr, tmDBGICtrls(ilBox)
                    Contract.pbcDemoBar(1).CurrentX = tmDBGICtrls(ilBox).fBoxX + fgBoxInsetX
                    Contract.pbcDemoBar(1).CurrentY = tmDBGICtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
                    Contract.pbcDemoBar(1).Print tmDBGICtrls(ilBox).sShow
                End If
            Next ilBox
        End If
        If imDBTotals < 5 Then
            For ilBox = UBound(tmDBCtrls) To UBound(tmDBCtrls) Step 1
                Contract.pbcDemoBar(1).CurrentX = tmDBCtrls(ilBox).fBoxX + fgBoxInsetX
                Contract.pbcDemoBar(1).CurrentY = tmDBCtrls(ilBox).fBoxY - 30 '+ fgBoxInsetY
                Contract.pbcDemoBar(1).Print tmDBCtrls(ilBox).sShow
            Next ilBox
        ElseIf imDBTotals = 5 Then
            ilStartRow = Contract.vbcSummary.Value  'Top location
            ilEndRow = Contract.vbcSummary.Value + Contract.vbcSummary.LargeChange
            If ilEndRow > UBound(smLnSumShow, 2) - 1 Then
                ilEndRow = UBound(smLnSumShow, 2) - 1  'don't include blank row
            End If
            For ilRow = ilStartRow To ilEndRow Step 1
                For ilBox = imLBCTCtrls To UBound(tmSumCtrls) Step 1
                    If Trim$(smLnSumShow(12, ilRow)) = "T" Then
                        gPaintArea Contract.pbcDemoBar(2), tmSumCtrls(ilBox).fBoxX, tmSumCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmSumCtrls(ilBox).fBoxW - 15, tmSumCtrls(ilBox).fBoxH - 15, GRAY  'WHITE
                    End If
                    Contract.pbcDemoBar(2).CurrentX = tmSumCtrls(ilBox).fBoxX + fgBoxInsetX
                    Contract.pbcDemoBar(2).CurrentY = tmSumCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                    gSetShow Contract.pbcDemoBar(2), smLnSumShow(ilBox, ilRow), tmSumCtrls(ilBox)
                    Contract.pbcDemoBar(2).Print tmSumCtrls(ilBox).sShow
                Next ilBox
            Next ilRow
        ElseIf imDBTotals = 6 Then
            ilStartRow = Contract.vbcSummary.Value  'Top location
            ilEndRow = Contract.vbcSummary.Value + Contract.vbcSummary.LargeChange
            If ilEndRow > UBound(smWkSumShow, 2) - 1 Then
                ilEndRow = UBound(smWkSumShow, 2) - 1  'don't include blank row
            End If
            For ilRow = ilStartRow To ilEndRow Step 1
                Contract.pbcDemoBar(3).CurrentX = tmSumCtrls(1).fBoxX + fgBoxInsetX
                Contract.pbcDemoBar(3).CurrentY = tmSumCtrls(1).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                Contract.pbcDemoBar(3).Print smWkSumShow(1, ilRow)
                For ilBox = imLBCTCtrls + 3 To UBound(tmSumCtrls) Step 1
                    Contract.pbcDemoBar(3).CurrentX = tmSumCtrls(ilBox).fBoxX + fgBoxInsetX
                    Contract.pbcDemoBar(3).CurrentY = tmSumCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                    gSetShow Contract.pbcDemoBar(3), smWkSumShow(ilBox - 2, ilRow), tmSumCtrls(ilBox)
                    Contract.pbcDemoBar(3).Print tmSumCtrls(ilBox).sShow
                Next ilBox
            Next ilRow
        ElseIf imDBTotals = 7 Then
            ilStartRow = Contract.vbcSummary.Value  'Top location
            ilEndRow = Contract.vbcSummary.Value + Contract.vbcSummary.LargeChange
            If ilEndRow > UBound(smQSumShow, 2) - 1 Then
                ilEndRow = UBound(smQSumShow, 2) - 1  'don't include blank row
            End If
            For ilRow = ilStartRow To ilEndRow Step 1
                Contract.pbcDemoBar(3).CurrentX = tmSumCtrls(1).fBoxX + fgBoxInsetX
                Contract.pbcDemoBar(3).CurrentY = tmSumCtrls(1).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                Contract.pbcDemoBar(3).Print smQSumShow(1, ilRow)
                For ilBox = imLBCTCtrls + 3 To UBound(tmSumCtrls) Step 1
                    Contract.pbcDemoBar(3).CurrentX = tmSumCtrls(ilBox).fBoxX + fgBoxInsetX
                    Contract.pbcDemoBar(3).CurrentY = tmSumCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                    gSetShow Contract.pbcDemoBar(3), smQSumShow(ilBox - 2, ilRow), tmSumCtrls(ilBox)
                    Contract.pbcDemoBar(3).Print tmSumCtrls(ilBox).sShow
                Next ilBox
            Next ilRow
        ElseIf imDBTotals = 8 Then
            ilStartRow = Contract.vbcSummary.Value  'Top location
            ilEndRow = Contract.vbcSummary.Value + Contract.vbcSummary.LargeChange
            If ilEndRow > UBound(smVSumShow, 2) - 1 Then
                ilEndRow = UBound(smVSumShow, 2) - 1  'don't include blank row
            End If
            For ilRow = ilStartRow To ilEndRow Step 1
                If Trim$(smVSumShow(12, ilRow)) = "T" Then
                    gPaintArea Contract.pbcDemoBar(3), tmSumCtrls(1).fBoxX, tmSumCtrls(1).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmSumCtrls(1).fBoxW + tmSumCtrls(2).fBoxW + tmSumCtrls(3).fBoxW + 15, tmSumCtrls(1).fBoxH - 15, GRAY 'WHITE
                End If
                Contract.pbcDemoBar(3).CurrentX = tmSumCtrls(1).fBoxX + fgBoxInsetX
                Contract.pbcDemoBar(3).CurrentY = tmSumCtrls(1).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                Contract.pbcDemoBar(3).Print smVSumShow(1, ilRow)
                For ilBox = imLBCTCtrls + 3 To UBound(tmSumCtrls) Step 1
                    If Trim$(smVSumShow(12, ilRow)) = "T" Then
                        gPaintArea Contract.pbcDemoBar(3), tmSumCtrls(ilBox).fBoxX, tmSumCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmSumCtrls(ilBox).fBoxW - 15, tmSumCtrls(ilBox).fBoxH - 15, GRAY  'WHITE
                    End If
                    Contract.pbcDemoBar(3).CurrentX = tmSumCtrls(ilBox).fBoxX + fgBoxInsetX
                    Contract.pbcDemoBar(3).CurrentY = tmSumCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                    gSetShow Contract.pbcDemoBar(3), smVSumShow(ilBox - 2, ilRow), tmSumCtrls(ilBox)
                    Contract.pbcDemoBar(3).Print tmSumCtrls(ilBox).sShow
                Next ilBox
            Next ilRow
        ElseIf imDBTotals = 9 Then
            ilStartRow = Contract.vbcSummary.Value  'Top location
            ilEndRow = Contract.vbcSummary.Value + Contract.vbcSummary.LargeChange
            If ilEndRow > UBound(smDPSumShow, 2) - 1 Then
                ilEndRow = UBound(smDPSumShow, 2) - 1  'don't include blank row
            End If
            For ilRow = ilStartRow To ilEndRow Step 1
                If Trim$(smDPSumShow(12, ilRow)) = "T" Then
                    gPaintArea Contract.pbcDemoBar(3), tmSumCtrls(1).fBoxX, tmSumCtrls(1).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmSumCtrls(1).fBoxW + tmSumCtrls(2).fBoxW + tmSumCtrls(3).fBoxW + 15, tmSumCtrls(1).fBoxH - 15, GRAY 'WHITE
                End If
                Contract.pbcDemoBar(3).CurrentX = tmSumCtrls(1).fBoxX + fgBoxInsetX
                Contract.pbcDemoBar(3).CurrentY = tmSumCtrls(1).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                Contract.pbcDemoBar(3).Print smDPSumShow(1, ilRow)
                For ilBox = imLBCTCtrls + 3 To UBound(tmSumCtrls) Step 1
                    If Trim$(smDPSumShow(12, ilRow)) = "T" Then
                        gPaintArea Contract.pbcDemoBar(3), tmSumCtrls(ilBox).fBoxX, tmSumCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmSumCtrls(ilBox).fBoxW - 15, tmSumCtrls(ilBox).fBoxH - 15, GRAY  'WHITE
                    End If
                    Contract.pbcDemoBar(3).CurrentX = tmSumCtrls(ilBox).fBoxX + fgBoxInsetX
                    Contract.pbcDemoBar(3).CurrentY = tmSumCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                    gSetShow Contract.pbcDemoBar(3), smDPSumShow(ilBox - 2, ilRow), tmSumCtrls(ilBox)
                    Contract.pbcDemoBar(3).Print tmSumCtrls(ilBox).sShow
                Next ilBox
            Next ilRow
        ElseIf imDBTotals = 10 Then
            ilStartRow = Contract.vbcSummary.Value  'Top location
            ilEndRow = Contract.vbcSummary.Value + Contract.vbcSummary.LargeChange
            If ilEndRow > UBound(smVDSumShow, 2) - 1 Then
                ilEndRow = UBound(smVDSumShow, 2) - 1  'don't include blank row
            End If
            For ilRow = ilStartRow To ilEndRow Step 1
                Contract.pbcDemoBar(3).ForeColor = BLACK
                Contract.pbcDemoBar(3).CurrentX = tmSumCtrls(1).fBoxX + fgBoxInsetX
                Contract.pbcDemoBar(3).CurrentY = tmSumCtrls(1).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                Contract.pbcDemoBar(3).Print smVDSumShow(1, ilRow)
                For ilBox = imLBCTCtrls + 3 To UBound(tmSumCtrls) Step 1
                    Contract.pbcDemoBar(3).CurrentX = tmSumCtrls(ilBox).fBoxX + fgBoxInsetX
                    Contract.pbcDemoBar(3).CurrentY = tmSumCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15) - 30 '+ fgBoxInsetY
                    If InStr(smVDSumShow(ilBox - 2, ilRow), "-") > 0 Then
                        Contract.pbcDemoBar(3).ForeColor = Red
                    Else
                        Contract.pbcDemoBar(3).ForeColor = DARKGREEN
                    End If
                    gSetShow Contract.pbcDemoBar(3), smVDSumShow(ilBox - 2, ilRow), tmSumCtrls(ilBox)
                    Contract.pbcDemoBar(3).Print tmSumCtrls(ilBox).sShow
                Next ilBox
            Next ilRow
            Contract.pbcDemoBar(3).ForeColor = BLACK
        End If
    End If
    If (Index = 2) Or (Index = 3) Then
        For ilBox = imLBCTCtrls + 3 To UBound(tmSumCtrls) Step 1
            Contract.pbcDemoBar(Index).CurrentX = tmSumCtrls(ilBox).fBoxX + fgBoxInsetX
            Contract.pbcDemoBar(Index).CurrentY = tmSumCtrls(ilBox).fBoxY + (6) * (fgBoxGridH + 15) '+ fgBoxInsetY
            If imDBTotals = 10 Then
                If InStr(smTSumShow(ilBox - imLBCTCtrls - 2), "-") > 0 Then
                    Contract.pbcDemoBar(3).ForeColor = Red
                Else
                    Contract.pbcDemoBar(3).ForeColor = DARKGREEN
                End If
            End If
            Contract.pbcDemoBar(Index).Print smTSumShow(ilBox - imLBCTCtrls - 2)
        Next ilBox
        Contract.pbcDemoBar(Index).ForeColor = BLACK
    End If
End Sub


Public Sub mSetACTInfo(Ctrl As TextBox)
    Dim slStr As String
    slStr = "Lineup Code" & sgCR & sgLF
    slStr = slStr & "A list of stations making up a network in Act 1 is called a Lineup, and has its own unique Lineup Code.  The following 4 questions indicate lineup settings used by networks to calculate network audience.  Agencies need to use these same lineup settings in order to replicate the network-calculated audience." & sgCR & sgLF
    slStr = slStr & sgCR & sgLF
    slStr = slStr & "Stored Times" & sgCR & sgLF
    slStr = slStr & "Act 1 Lineups can be created with preassigned dayparts, which define the exact times that each station airs the network.  If left blank, it is assumed that the research for each station is based on a common daypart, which is inclusive of the time most stations air the network." & sgCR & sgLF
    slStr = slStr & sgCR & sgLF
    slStr = slStr & "Stored Spots" & sgCR & sgLF
    slStr = slStr & "Act 1 Lineups can be created where the number of spots airing on each station can vary.  If left blank (which is common for traditional networks) it is assumed that the same number of spots will air on each station, as indicated in the schedule/proposal." & sgCR & sgLF
    slStr = slStr & sgCR & sgLF
    slStr = slStr & "Stored Clear%" & sgCR & sgLF
    slStr = slStr & "Act 1 Lineups can be created with pre-assigned clearance percentages that vary by station, to account for the fact that some stations may not air their allocated number of spots.  If left blank it is assumed that all stations generally air all spots assigned to them." & sgCR & sgLF
    slStr = slStr & sgCR & sgLF
    slStr = slStr & "Daypart Filter" & sgCR & sgLF
    slStr = slStr & "If different stations in an Act 1 Lineup have different dayparts, this filter allows only research from stations that match the proposal’s daypart to be used, and ignore the rest.  If left blank it is assumed that all stations will be included no matter what dayparts are defined." & sgCR & sgLF
    Ctrl.Text = slStr
End Sub

Public Sub mResizeDemoBar()
    Dim llMax As Long
    Dim ilLoop As Integer
    Dim flAdjFactorW As Single
    
    flAdjFactorW = 3 / 2
'    If Screen.Width * 15 = 640 Then
'        flAdjFactorW = 1#
'        flAdjFactorW = 1#
'    Else
'        flAdjFactorW = (((lgPercentAdjW / 2) * ((Screen.Width) / (640 * 15 / Contract.pbcDemoBar(0).Width))) / 100) / Contract.pbcDemoBar(0).Width
'        If flAdjFactorW < 1# Then
'            flAdjFactorW = 1#
'        End If
'    End If
    llMax = 0
    For ilLoop = LBound(tmDBCtrls) + 1 To UBound(tmDBCtrls) Step 1
        If tmDBCtrls(ilLoop).fBoxX >= 0 Then
            tmDBCtrls(ilLoop).fBoxW = CLng(flAdjFactorW * tmDBCtrls(ilLoop).fBoxW)
            Do While (tmDBCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmDBCtrls(ilLoop).fBoxW = tmDBCtrls(ilLoop).fBoxW + 1
            Loop
            tmDBCtrls(ilLoop).fBoxX = CLng(flAdjFactorW * tmDBCtrls(ilLoop).fBoxX)
            Do While (tmDBCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmDBCtrls(ilLoop).fBoxX = tmDBCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmDBCtrls(ilLoop).fBoxX > 90 Then
                    If tmDBCtrls(ilLoop - 1).fBoxX + tmDBCtrls(ilLoop - 1).fBoxW + 15 < tmDBCtrls(ilLoop).fBoxX Then
                        tmDBCtrls(ilLoop - 1).fBoxW = tmDBCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmDBCtrls(ilLoop - 1).fBoxX + tmDBCtrls(ilLoop - 1).fBoxW + 15 > tmDBCtrls(ilLoop).fBoxX Then
                        tmDBCtrls(ilLoop - 1).fBoxW = tmDBCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmDBCtrls(ilLoop).fBoxX + tmDBCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmDBCtrls(ilLoop).fBoxX + tmDBCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    tmDBCtrls(11).fBoxW = tmDBCtrls(11).fBoxW + 15
    Contract.plcDemoBar.Width = llMax + 2 * fgBevelX + 15
    Contract.pbcDemoBar(0).Width = llMax
    
    llMax = 0
    For ilLoop = LBound(tmDBSCCtrls) + 1 To UBound(tmDBSCCtrls) Step 1
        If tmDBSCCtrls(ilLoop).fBoxX >= 0 Then
            tmDBSCCtrls(ilLoop).fBoxW = CLng(flAdjFactorW * tmDBSCCtrls(ilLoop).fBoxW)
            Do While (tmDBSCCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmDBSCCtrls(ilLoop).fBoxW = tmDBSCCtrls(ilLoop).fBoxW + 1
            Loop
            tmDBSCCtrls(ilLoop).fBoxX = CLng(flAdjFactorW * tmDBSCCtrls(ilLoop).fBoxX)
            Do While (tmDBSCCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmDBSCCtrls(ilLoop).fBoxX = tmDBSCCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmDBSCCtrls(ilLoop).fBoxX > 90 Then
                    If tmDBSCCtrls(ilLoop - 1).fBoxX + tmDBSCCtrls(ilLoop - 1).fBoxW + 15 < tmDBSCCtrls(ilLoop).fBoxX Then
                        tmDBSCCtrls(ilLoop - 1).fBoxW = tmDBSCCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmDBSCCtrls(ilLoop - 1).fBoxX + tmDBSCCtrls(ilLoop - 1).fBoxW + 15 > tmDBSCCtrls(ilLoop).fBoxX Then
                        tmDBSCCtrls(ilLoop - 1).fBoxW = tmDBSCCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmDBSCCtrls(ilLoop).fBoxX + tmDBSCCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmDBSCCtrls(ilLoop).fBoxX + tmDBSCCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    Contract.pbcDemoBar(1).Width = llMax
    Contract.pbcLnWkArrow(2).Move tmDBSCCtrls(1).fBoxX + 15, 30, Contract.pbcLnWkArrow(2).Width - 30, tmDBSCCtrls(1).fBoxH
    Contract.pbcLnWkArrow(3).Move tmDBSCCtrls(13).fBoxX + tmDBSCCtrls(13).fBoxW - Contract.pbcLnWkArrow(2).Width + 15, 30, Contract.pbcLnWkArrow(2).Width - 30, tmDBSCCtrls(1).fBoxH
    
    lmDBHotSpot(1, 1) = Contract.pbcLnWkArrow(2).Left    '3675  '3555  'Left
    lmDBHotSpot(1, 2) = 15    'Top
    lmDBHotSpot(1, 3) = Contract.pbcLnWkArrow(2).Left + 150  '3675 + 150 'Right
    lmDBHotSpot(1, 4) = 15 + 180  'Bottom
    lmDBHotSpot(2, 1) = Contract.pbcLnWkArrow(2).Left + 150  '3825  'Left
    lmDBHotSpot(2, 2) = 15    'Top
    lmDBHotSpot(2, 3) = Contract.pbcLnWkArrow(2).Left + 300  '3825 + 150 'Right
    lmDBHotSpot(2, 4) = 15 + 180  'Bottom
    lmDBHotSpot(3, 1) = Contract.pbcLnWkArrow(3).Left    '7470  'Left
    lmDBHotSpot(3, 2) = 15    'Top
    lmDBHotSpot(3, 3) = Contract.pbcLnWkArrow(3).Left + 150  '7470 + 150 'Right
    lmDBHotSpot(3, 4) = 15 + 180  'Bottom
    lmDBHotSpot(4, 1) = Contract.pbcLnWkArrow(3).Left + 150  '7620  'Left
    lmDBHotSpot(4, 2) = 15    'Top
    lmDBHotSpot(4, 3) = Contract.pbcLnWkArrow(3).Left + 300  '7620 + 150 'Right
    lmDBHotSpot(4, 4) = 15 + 180  'Bottom
    Contract.pbcLnWkArrow(2).Enabled = False
    Contract.pbcLnWkArrow(3).Enabled = False
    
    llMax = 0
    For ilLoop = LBound(tmDBAvgCtrls) + 1 To UBound(tmDBAvgCtrls) Step 1
        If tmDBAvgCtrls(ilLoop).fBoxX >= 0 Then
            tmDBAvgCtrls(ilLoop).fBoxW = CLng(flAdjFactorW * tmDBAvgCtrls(ilLoop).fBoxW)
            Do While (tmDBAvgCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmDBAvgCtrls(ilLoop).fBoxW = tmDBAvgCtrls(ilLoop).fBoxW + 1
            Loop
            tmDBAvgCtrls(ilLoop).fBoxX = CLng(flAdjFactorW * tmDBAvgCtrls(ilLoop).fBoxX)
            Do While (tmDBAvgCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmDBAvgCtrls(ilLoop).fBoxX = tmDBAvgCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmDBAvgCtrls(ilLoop).fBoxX > 90 Then
                    If tmDBAvgCtrls(ilLoop - 1).fBoxX + tmDBAvgCtrls(ilLoop - 1).fBoxW + 15 < tmDBAvgCtrls(ilLoop).fBoxX Then
                        tmDBAvgCtrls(ilLoop - 1).fBoxW = tmDBAvgCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmDBAvgCtrls(ilLoop - 1).fBoxX + tmDBAvgCtrls(ilLoop - 1).fBoxW + 15 > tmDBAvgCtrls(ilLoop).fBoxX Then
                        tmDBAvgCtrls(ilLoop - 1).fBoxW = tmDBAvgCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmDBAvgCtrls(ilLoop).fBoxX + tmDBAvgCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmDBAvgCtrls(ilLoop).fBoxX + tmDBAvgCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    
    llMax = 0
    For ilLoop = LBound(tmDBAPCtrls) + 1 To UBound(tmDBAPCtrls) Step 1
        If tmDBAPCtrls(ilLoop).fBoxX >= 0 Then
            tmDBAPCtrls(ilLoop).fBoxW = CLng(flAdjFactorW * tmDBAPCtrls(ilLoop).fBoxW)
            Do While (tmDBAPCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmDBAPCtrls(ilLoop).fBoxW = tmDBAPCtrls(ilLoop).fBoxW + 1
            Loop
            tmDBAPCtrls(ilLoop).fBoxX = CLng(flAdjFactorW * tmDBAPCtrls(ilLoop).fBoxX)
            Do While (tmDBAPCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmDBAPCtrls(ilLoop).fBoxX = tmDBAPCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmDBAPCtrls(ilLoop).fBoxX > 90 Then
                    If tmDBAPCtrls(ilLoop - 1).fBoxX + tmDBAPCtrls(ilLoop - 1).fBoxW + 15 < tmDBAPCtrls(ilLoop).fBoxX Then
                        tmDBAPCtrls(ilLoop - 1).fBoxW = tmDBAPCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmDBAPCtrls(ilLoop - 1).fBoxX + tmDBAPCtrls(ilLoop - 1).fBoxW + 15 > tmDBAPCtrls(ilLoop).fBoxX Then
                        tmDBAPCtrls(ilLoop - 1).fBoxW = tmDBAPCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmDBAPCtrls(ilLoop).fBoxX + tmDBAPCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmDBAPCtrls(ilLoop).fBoxX + tmDBAPCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    
    For ilLoop = LBound(tmDBRCPCtrls) + 1 To UBound(tmDBRCPCtrls) Step 1
        If tmDBRCPCtrls(ilLoop).fBoxX >= 0 Then
            tmDBRCPCtrls(ilLoop).fBoxW = CLng(flAdjFactorW * tmDBRCPCtrls(ilLoop).fBoxW)
            Do While (tmDBRCPCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmDBRCPCtrls(ilLoop).fBoxW = tmDBRCPCtrls(ilLoop).fBoxW + 1
            Loop
            tmDBRCPCtrls(ilLoop).fBoxX = CLng(flAdjFactorW * tmDBRCPCtrls(ilLoop).fBoxX)
            Do While (tmDBRCPCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmDBRCPCtrls(ilLoop).fBoxX = tmDBRCPCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmDBRCPCtrls(ilLoop).fBoxX > 90 Then
                    If tmDBRCPCtrls(ilLoop - 1).fBoxX + tmDBRCPCtrls(ilLoop - 1).fBoxW + 15 < tmDBRCPCtrls(ilLoop).fBoxX Then
                        tmDBRCPCtrls(ilLoop - 1).fBoxW = tmDBRCPCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmDBRCPCtrls(ilLoop - 1).fBoxX + tmDBRCPCtrls(ilLoop - 1).fBoxW + 15 > tmDBRCPCtrls(ilLoop).fBoxX Then
                        tmDBRCPCtrls(ilLoop - 1).fBoxW = tmDBRCPCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmDBRCPCtrls(ilLoop).fBoxX + tmDBRCPCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmDBRCPCtrls(ilLoop).fBoxX + tmDBRCPCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    
    For ilLoop = LBound(tmDBRCtrls) + 1 To UBound(tmDBRCtrls) Step 1
        If tmDBRCtrls(ilLoop).fBoxX >= 0 Then
            tmDBRCtrls(ilLoop).fBoxW = CLng(flAdjFactorW * tmDBRCtrls(ilLoop).fBoxW)
            Do While (tmDBRCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmDBRCtrls(ilLoop).fBoxW = tmDBRCtrls(ilLoop).fBoxW + 1
            Loop
            tmDBRCtrls(ilLoop).fBoxX = CLng(flAdjFactorW * tmDBRCtrls(ilLoop).fBoxX)
            Do While (tmDBRCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmDBRCtrls(ilLoop).fBoxX = tmDBRCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmDBRCtrls(ilLoop).fBoxX > 90 Then
                    If tmDBRCtrls(ilLoop - 1).fBoxX + tmDBRCtrls(ilLoop - 1).fBoxW + 15 < tmDBRCtrls(ilLoop).fBoxX Then
                        tmDBRCtrls(ilLoop - 1).fBoxW = tmDBRCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmDBRCtrls(ilLoop - 1).fBoxX + tmDBRCtrls(ilLoop - 1).fBoxW + 15 > tmDBRCtrls(ilLoop).fBoxX Then
                        tmDBRCtrls(ilLoop - 1).fBoxW = tmDBRCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmDBRCtrls(ilLoop).fBoxX + tmDBRCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmDBRCtrls(ilLoop).fBoxX + tmDBRCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    
    For ilLoop = LBound(tmDBGRPCtrls) + 1 To UBound(tmDBGRPCtrls) Step 1
        If tmDBGRPCtrls(ilLoop).fBoxX >= 0 Then
            tmDBGRPCtrls(ilLoop).fBoxW = CLng(flAdjFactorW * tmDBGRPCtrls(ilLoop).fBoxW)
            Do While (tmDBGRPCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmDBGRPCtrls(ilLoop).fBoxW = tmDBGRPCtrls(ilLoop).fBoxW + 1
            Loop
            tmDBGRPCtrls(ilLoop).fBoxX = CLng(flAdjFactorW * tmDBGRPCtrls(ilLoop).fBoxX)
            Do While (tmDBGRPCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmDBGRPCtrls(ilLoop).fBoxX = tmDBGRPCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmDBGRPCtrls(ilLoop).fBoxX > 90 Then
                    If tmDBGRPCtrls(ilLoop - 1).fBoxX + tmDBGRPCtrls(ilLoop - 1).fBoxW + 15 < tmDBGRPCtrls(ilLoop).fBoxX Then
                        tmDBGRPCtrls(ilLoop - 1).fBoxW = tmDBGRPCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmDBGRPCtrls(ilLoop - 1).fBoxX + tmDBGRPCtrls(ilLoop - 1).fBoxW + 15 > tmDBGRPCtrls(ilLoop).fBoxX Then
                        tmDBGRPCtrls(ilLoop - 1).fBoxW = tmDBGRPCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmDBGRPCtrls(ilLoop).fBoxX + tmDBGRPCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmDBGRPCtrls(ilLoop).fBoxX + tmDBGRPCtrls(ilLoop).fBoxW + 15
        End If
     Next ilLoop
    
     For ilLoop = LBound(tmDBGICtrls) + 1 To UBound(tmDBGICtrls) Step 1
        If tmDBGICtrls(ilLoop).fBoxX >= 0 Then
            tmDBGICtrls(ilLoop).fBoxW = CLng(flAdjFactorW * tmDBGICtrls(ilLoop).fBoxW)
            Do While (tmDBGICtrls(ilLoop).fBoxW Mod 15) <> 0
                tmDBGICtrls(ilLoop).fBoxW = tmDBGICtrls(ilLoop).fBoxW + 1
            Loop
            tmDBGICtrls(ilLoop).fBoxX = CLng(flAdjFactorW * tmDBGICtrls(ilLoop).fBoxX)
            Do While (tmDBGICtrls(ilLoop).fBoxX Mod 15) <> 0
                tmDBGICtrls(ilLoop).fBoxX = tmDBGICtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmDBGICtrls(ilLoop).fBoxX > 90 Then
                    If tmDBGICtrls(ilLoop - 1).fBoxX + tmDBGICtrls(ilLoop - 1).fBoxW + 15 < tmDBGICtrls(ilLoop).fBoxX Then
                        tmDBGICtrls(ilLoop - 1).fBoxW = tmDBGICtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmDBGICtrls(ilLoop - 1).fBoxX + tmDBGICtrls(ilLoop - 1).fBoxW + 15 > tmDBGICtrls(ilLoop).fBoxX Then
                        tmDBGICtrls(ilLoop - 1).fBoxW = tmDBGICtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmDBGICtrls(ilLoop).fBoxX + tmDBGICtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmDBGICtrls(ilLoop).fBoxX + tmDBGICtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
   
    llMax = 0
    For ilLoop = LBound(tmDBAvCtrls) + 1 To UBound(tmDBAvCtrls) Step 1
        If tmDBAvCtrls(ilLoop).fBoxX >= 0 Then
            tmDBAvCtrls(ilLoop).fBoxW = CLng(flAdjFactorW * tmDBAvCtrls(ilLoop).fBoxW)
            Do While (tmDBAvCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmDBAvCtrls(ilLoop).fBoxW = tmDBAvCtrls(ilLoop).fBoxW + 1
            Loop
            tmDBAvCtrls(ilLoop).fBoxX = CLng(flAdjFactorW * tmDBAvCtrls(ilLoop).fBoxX)
            Do While (tmDBAvCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmDBAvCtrls(ilLoop).fBoxX = tmDBAvCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmDBAvCtrls(ilLoop).fBoxX > 90 Then
                    If tmDBAvCtrls(ilLoop - 1).fBoxX + tmDBAvCtrls(ilLoop - 1).fBoxW + 15 < tmDBAvCtrls(ilLoop).fBoxX Then
                        tmDBAvCtrls(ilLoop - 1).fBoxW = tmDBAvCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmDBAvCtrls(ilLoop - 1).fBoxX + tmDBAvCtrls(ilLoop - 1).fBoxW + 15 > tmDBAvCtrls(ilLoop).fBoxX Then
                        tmDBAvCtrls(ilLoop - 1).fBoxW = tmDBAvCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmDBAvCtrls(ilLoop).fBoxX + tmDBAvCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmDBAvCtrls(ilLoop).fBoxX + tmDBAvCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
   
    llMax = 0
    For ilLoop = LBound(tmSumCtrls) + 1 To UBound(tmSumCtrls) Step 1
        If tmSumCtrls(ilLoop).fBoxX >= 0 Then
            tmSumCtrls(ilLoop).fBoxW = CLng(flAdjFactorW * tmSumCtrls(ilLoop).fBoxW)
            Do While (tmSumCtrls(ilLoop).fBoxW Mod 15) <> 0
                tmSumCtrls(ilLoop).fBoxW = tmSumCtrls(ilLoop).fBoxW + 1
            Loop
            tmSumCtrls(ilLoop).fBoxX = CLng(flAdjFactorW * tmSumCtrls(ilLoop).fBoxX)
            Do While (tmSumCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmSumCtrls(ilLoop).fBoxX = tmSumCtrls(ilLoop).fBoxX + 1
            Loop
            If ilLoop > 1 Then
                If tmSumCtrls(ilLoop).fBoxX > 90 Then
                    If tmSumCtrls(ilLoop - 1).fBoxX + tmSumCtrls(ilLoop - 1).fBoxW + 15 < tmSumCtrls(ilLoop).fBoxX Then
                        tmSumCtrls(ilLoop - 1).fBoxW = tmSumCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmSumCtrls(ilLoop - 1).fBoxX + tmSumCtrls(ilLoop - 1).fBoxW + 15 > tmSumCtrls(ilLoop).fBoxX Then
                        tmSumCtrls(ilLoop - 1).fBoxW = tmSumCtrls(ilLoop - 1).fBoxW - 15
                    End If
                End If
            End If
        End If
        If tmSumCtrls(ilLoop).fBoxX + tmSumCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmSumCtrls(ilLoop).fBoxX + tmSumCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
   
    Contract.pbcDemoBar(2).Width = llMax
    Contract.pbcDemoBar(3).Width = llMax
   
    Contract.pbcDemoBar(0).Picture = LoadPicture("")
    Contract.pbcDemoBar(1).Picture = LoadPicture("")
    Contract.pbcDemoBar(2).Picture = LoadPicture("")
    Contract.pbcDemoBar(3).Picture = LoadPicture("")
    
    ''Contract.vbcSummary.Move Contract.plcDemoBar.Width - Contract.vbcSummary.Width - fgBevelX, Contract.pbcDemoBar(2).Top - Contract.plcDemoBar.Top - 15, Contract.vbcSummary.Width, Contract.pbcDemoBar(2).Height
    'Contract.vbcSummary.Move Contract.pbcDemoBar(2).Left + Contract.pbcDemoBar(2).Width + 15, Contract.pbcDemoBar(2).Top - Contract.plcDemoBar.Top - 15, Contract.vbcSummary.Width, Contract.pbcDemoBar(2).Height
    Contract.vbcSummary.Move Contract.pbcDemoBar(2).Width + 30, Contract.pbcDemoBar(2).Top - Contract.plcDemoBar.Top - 15, Contract.vbcSummary.Width, Contract.pbcDemoBar(2).height
                
    'Book Name 30 Characters + Date
    Contract.lbcDBBook.Width = (Traffic.pbcArial.TextWidth("123456789012345678901234567890123456789012345"))
    
    Contract.edcDBDisplay.Width = Contract.lbcDBDisplay.Width - Contract.cmcDBDisplay.Width
    Contract.cmcDBDisplay.Left = Contract.edcDBDisplay.Left + Contract.edcDBDisplay.Width
    
    Contract.edcDBDemo.Left = Contract.cmcDBDisplay.Left + 2 * Contract.cmcDBDisplay.Width
    Contract.edcDBDemo.Width = Contract.lbcDBDemo.Width - Contract.cmcDBDemo.Width
    Contract.cmcDBDemo.Left = Contract.edcDBDemo.Left + Contract.edcDBDemo.Width
    
    Contract.edcDBBook.Left = Contract.cmcDBDemo.Left + 2 * Contract.cmcDBDemo.Width
    Contract.edcDBBook.Width = Contract.lbcDBBook.Width - Contract.cmcDBBook.Width
    Contract.cmcDBBook.Left = Contract.edcDBBook.Left + Contract.edcDBBook.Width
    
    Contract.edcDBSocEco.Left = Contract.cmcDBBook.Left + 2 * Contract.cmcDBBook.Width
    Contract.edcDBSocEco.Width = Contract.lbcDBSocEco.Width - Contract.cmcDBSocEco.Width
    Contract.cmcDBSocEco.Left = Contract.edcDBSocEco.Left + Contract.edcDBSocEco.Width
    
    'lbcDBSocEco.Height = gListBoxHeight(lbcDBSocEco.ListCount, 10)
    Contract.lbcDBDisplay.Move Contract.plcDemoBar.Left + Contract.edcDBDisplay.Left, Contract.plcDemoBar.Top + Contract.edcDBDisplay.Top + Contract.edcDBDisplay.height
    Contract.lacDBBookOvly.Move Contract.edcDBBook.Left, Contract.edcDBBook.Top, Contract.edcDBBook.Width, Contract.edcDBBook.height
    'lacDBBookOvly.ZOrder vbBringToFront   'Get it on top of edcDBBook as making edcDBBook invisble causes a problem in mGetResearchTotal and NO research values exist
    Contract.lbcDBBook.Move Contract.plcDemoBar.Left + Contract.edcDBBook.Left, Contract.plcDemoBar.Top + Contract.edcDBBook.Top + Contract.edcDBBook.height
    Contract.lbcDBDemo.Move Contract.plcDemoBar.Left + Contract.edcDBDemo.Left, Contract.plcDemoBar.Top + Contract.edcDBDemo.Top + Contract.edcDBDemo.height
    Contract.lbcDBSocEco.Move Contract.plcDemoBar.Left + Contract.edcDBSocEco.Left, Contract.plcDemoBar.Top + Contract.edcDBSocEco.Top + Contract.edcDBSocEco.height
        
        
End Sub

Public Sub mPaintDemoBarGrid(ilIndex As Integer)
    
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilHeightAdj As Integer
    Dim ilRow As Integer
    
    llColor = Contract.pbcDemoBar(ilIndex).ForeColor
    slFontName = Contract.pbcDemoBar(ilIndex).FontName
    flFontSize = Contract.pbcDemoBar(ilIndex).FontSize
    Contract.pbcDemoBar(ilIndex).ForeColor = BLUE
    Contract.pbcDemoBar(ilIndex).FontBold = False
    Contract.pbcDemoBar(ilIndex).FontSize = 7
    Contract.pbcDemoBar(ilIndex).FontName = "Arial"
    Contract.pbcDemoBar(ilIndex).FontSize = 7  'Font size done twice as indicated in FontSize property area in manual
    If ilIndex = 0 Then
        llTop = 15   'tmDBCtrls(1).fBoxY
        ilHeightAdj = 0
        ilRow = 1
        Do
            For ilLoop = LBound(tmDBCtrls) + 1 To UBound(tmDBCtrls) - 1 Step 1
                If ilRow = 1 And ilLoop = LBound(tmDBCtrls) + 1 Then
                    Contract.pbcDemoBar(ilIndex).Line (tmDBCtrls(ilLoop).fBoxX, llTop - 15)-Step(tmDBCtrls(ilLoop).fBoxW - 15, tmDBCtrls(ilLoop).fBoxH - 15), BLUE, BF
                Else
                    Contract.pbcDemoBar(ilIndex).Line (tmDBCtrls(ilLoop).fBoxX, llTop - 15)-Step(tmDBCtrls(ilLoop).fBoxW - 15, tmDBCtrls(ilLoop).fBoxH - 15), LIGHTERYELLOW, BF
                End If
                Contract.pbcDemoBar(ilIndex).Line (tmDBCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmDBCtrls(ilLoop).fBoxW + 15, tmDBCtrls(ilLoop).fBoxH + ilHeightAdj), BLUE, B
            Next ilLoop
            ilRow = ilRow + 1
            llTop = llTop + tmDBCtrls(1).fBoxH + ilHeightAdj
            ilHeightAdj = 15
        Loop While llTop + tmDBCtrls(1).fBoxH < Contract.pbcDemoBar(ilIndex).height
        Contract.pbcDemoBar(ilIndex).Line (15, 0)-Step(tmDBCtrls(1).fBoxX - 30, 2 * tmDBCtrls(1).fBoxH + 30), LIGHTERYELLOW, BF
        Contract.pbcDemoBar(ilIndex).Line (15, 0)-Step(tmDBCtrls(1).fBoxX - 30, 2 * tmDBCtrls(1).fBoxH + 30), BLUE, B
        Contract.pbcDemoBar(ilIndex).Line (tmDBCtrls(4).fBoxX + tmDBCtrls(4).fBoxW, 0)-(tmDBCtrls(5).fBoxX - 15, 2 * tmDBCtrls(1).fBoxH + 30), LIGHTERYELLOW, BF
        Contract.pbcDemoBar(ilIndex).Line (tmDBCtrls(4).fBoxX + tmDBCtrls(4).fBoxW, 0)-(tmDBCtrls(5).fBoxX - 15, 2 * tmDBCtrls(1).fBoxH + 30), BLUE, B
        Contract.pbcDemoBar(ilIndex).Line (tmDBCtrls(8).fBoxX + tmDBCtrls(8).fBoxW, 0)-(tmDBCtrls(9).fBoxX - 15, 2 * tmDBCtrls(1).fBoxH + 30), LIGHTERYELLOW, BF
        Contract.pbcDemoBar(ilIndex).Line (tmDBCtrls(8).fBoxX + tmDBCtrls(8).fBoxW, 0)-(tmDBCtrls(9).fBoxX - 15, 2 * tmDBCtrls(1).fBoxH + 30), BLUE, B
        
    End If
    If ilIndex = 1 Then
        llTop = 15  'tmDBSCCtrls(1).fBoxY
        ilHeightAdj = 15
        ilRow = 1
        Do
            For ilLoop = LBound(tmDBSCCtrls) + 1 To UBound(tmDBSCCtrls) Step 1
                If ilRow = 3 Or ilRow = 5 Or ilRow >= 7 Then
                    Contract.pbcDemoBar(ilIndex).Line (tmDBSCCtrls(ilLoop).fBoxX, llTop - 15)-Step(tmDBSCCtrls(ilLoop).fBoxW - 15, tmDBSCCtrls(ilLoop).fBoxH - 15), LIGHTERYELLOW, BF
                End If
                'Contract.pbcDemoBar(ilIndex).Line (tmDBSCCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmDBSCCtrls(ilLoop).fBoxW + 15, tmDBSCCtrls(ilLoop).fBoxH + ilHeightAdj), BLUE, B
                Contract.pbcDemoBar(ilIndex).Line (tmDBSCCtrls(ilLoop).fBoxX - 15, llTop)-Step(tmDBSCCtrls(ilLoop).fBoxW + 15, tmDBSCCtrls(ilLoop).fBoxH + ilHeightAdj), BLUE, B
            Next ilLoop
            If ilRow = 5 Then
                Contract.pbcDemoBar(ilIndex).Line (15, llTop)-Step(tmDBSCCtrls(1).fBoxX - 30, tmDBSCCtrls(1).fBoxH + ilHeightAdj), LIGHTERYELLOW, BF
            End If
            If ilRow <= 5 Then
                Contract.pbcDemoBar(ilIndex).Line (15, llTop)-Step(tmDBSCCtrls(1).fBoxX - 30, tmDBSCCtrls(1).fBoxH + ilHeightAdj), BLUE, B
            Else
                'Ratio (Index)
                Contract.pbcDemoBar(ilIndex).Line (15, llTop)-Step(tmDBCtrls(12).fBoxX - 30, tmDBCtrls(12).fBoxH + ilHeightAdj), BLUE, B
                Contract.pbcDemoBar(ilIndex).Line (15, llTop)-Step(tmDBCtrls(12).fBoxX + tmDBCtrls(12).fBoxW - 15, tmDBCtrls(12).fBoxH + ilHeightAdj), BLUE, B
            End If
            ilRow = ilRow + 1
            llTop = llTop + tmDBSCCtrls(1).fBoxH + ilHeightAdj
            ilHeightAdj = 15
        Loop While llTop + tmDBSCCtrls(1).fBoxH < Contract.pbcDemoBar(ilIndex).height
        Contract.pbcRatio.Width = tmDBCtrls(12).fBoxX - 45
        tmAvToggleCtrls(1).fBoxW = tmDBSCCtrls(1).fBoxX - 45
    End If
    If ilIndex = 2 Then
        llTop = 15  'tmDBSCCtrls(1).fBoxY
        ilHeightAdj = -15
        Do
            For ilLoop = LBound(tmSumCtrls) + 1 To UBound(tmSumCtrls) Step 1
                Contract.pbcDemoBar(ilIndex).Line (tmSumCtrls(ilLoop).fBoxX, llTop - 15)-Step(tmSumCtrls(ilLoop).fBoxW - 15, tmSumCtrls(ilLoop).fBoxH - 15), LIGHTERYELLOW, BF
                'Contract.pbcDemoBar(ilIndex).Line (tmDBSCCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmDBSCCtrls(ilLoop).fBoxW + 15, tmDBSCCtrls(ilLoop).fBoxH + ilHeightAdj), BLUE, B
                Contract.pbcDemoBar(ilIndex).Line (tmSumCtrls(ilLoop).fBoxX - 15, llTop)-Step(tmSumCtrls(ilLoop).fBoxW + 15, tmSumCtrls(ilLoop).fBoxH + ilHeightAdj), BLUE, B
            Next ilLoop
            llTop = llTop + tmSumCtrls(1).fBoxH + ilHeightAdj
            ilHeightAdj = 15
        Loop While llTop + tmSumCtrls(1).fBoxH < Contract.pbcDemoBar(ilIndex).height - tmSumCtrls(1).fBoxH
    End If
    
    If ilIndex = 3 Then
        llTop = 15  'tmDBSCCtrls(1).fBoxY
        ilHeightAdj = -15
        Do
            For ilLoop = LBound(tmSumCtrls) + 4 To UBound(tmSumCtrls) Step 1
                Contract.pbcDemoBar(ilIndex).Line (tmSumCtrls(ilLoop).fBoxX, llTop - 15)-Step(tmSumCtrls(ilLoop).fBoxW - 15, tmSumCtrls(ilLoop).fBoxH - 15), LIGHTERYELLOW, BF
                'Contract.pbcDemoBar(ilIndex).Line (tmDBSCCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmDBSCCtrls(ilLoop).fBoxW + 15, tmDBSCCtrls(ilLoop).fBoxH + ilHeightAdj), BLUE, B
                Contract.pbcDemoBar(ilIndex).Line (tmSumCtrls(ilLoop).fBoxX - 15, llTop)-Step(tmSumCtrls(ilLoop).fBoxW + 15, tmSumCtrls(ilLoop).fBoxH + ilHeightAdj), BLUE, B
            Next ilLoop
            Contract.pbcDemoBar(ilIndex).Line (0, llTop)-Step(tmSumCtrls(4).fBoxX - 15, tmSumCtrls(1).fBoxH - 45), LIGHTERYELLOW, BF
            Contract.pbcDemoBar(ilIndex).Line (15, llTop)-Step(tmSumCtrls(4).fBoxX - 30, tmSumCtrls(1).fBoxH + ilHeightAdj), BLUE, B
            llTop = llTop + tmSumCtrls(1).fBoxH + ilHeightAdj
            ilHeightAdj = 15
        Loop While llTop + tmSumCtrls(1).fBoxH < Contract.pbcDemoBar(ilIndex).height - tmSumCtrls(1).fBoxH
    End If
    If ilIndex >= 2 Then
        ilHeightAdj = 15
        llTop = 7 * tmSumCtrls(1).fBoxH + 6 * ilHeightAdj
        For ilLoop = LBound(tmSumCtrls) + 4 To UBound(tmSumCtrls) Step 1
            If ilIndex = 2 Then
                Contract.pbcDemoBar(ilIndex).Line (tmSumCtrls(ilLoop).fBoxX, llTop + 30)-Step(tmSumCtrls(ilLoop).fBoxW - 15, tmSumCtrls(ilLoop).fBoxH - 15), LIGHTERYELLOW, BF
                Contract.pbcDemoBar(ilIndex).Line (tmSumCtrls(ilLoop).fBoxX - 15, llTop + 30)-Step(tmSumCtrls(ilLoop).fBoxW + 15, tmSumCtrls(ilLoop).fBoxH + ilHeightAdj), BLUE, B
            Else
                Contract.pbcDemoBar(ilIndex).Line (tmSumCtrls(ilLoop).fBoxX, llTop + 30)-Step(tmSumCtrls(ilLoop).fBoxW - 15, tmSumCtrls(ilLoop).fBoxH - 30), LIGHTERYELLOW, BF
                Contract.pbcDemoBar(ilIndex).Line (tmSumCtrls(ilLoop).fBoxX - 15, llTop + 30)-Step(tmSumCtrls(ilLoop).fBoxW + 15, tmSumCtrls(ilLoop).fBoxH), BLUE, B
            End If
        Next ilLoop
        llTop = 15
        ilHeightAdj = 15
        Contract.pbcDemoBar(1).Line (15, llTop)-Step(tmDBSCCtrls(1).fBoxX - 30, tmDBSCCtrls(1).fBoxH + ilHeightAdj), BLUE, B
        llTop = llTop + tmDBSCCtrls(1).fBoxH + ilHeightAdj
        Contract.pbcDemoBar(1).Line (15, llTop)-Step(tmDBSCCtrls(1).fBoxX - 30, tmDBSCCtrls(1).fBoxH + ilHeightAdj), BLUE, B
    End If
    If ilIndex >= 2 Then
        llTop = 15  'tmDBSCCtrls(1).fBoxY
        ilHeightAdj = 15
        For ilRow = 1 To 2 Step 1
            For ilLoop = LBound(tmDBSCCtrls) + 1 To UBound(tmDBSCCtrls) Step 1
                'Contract.pbcDemoBar(ilIndex).Line (tmDBSCCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmDBSCCtrls(ilLoop).fBoxW + 15, tmDBSCCtrls(ilLoop).fBoxH + ilHeightAdj), BLUE, B
                Contract.pbcDemoBar(1).Line (tmDBSCCtrls(ilLoop).fBoxX - 15, llTop)-Step(tmDBSCCtrls(ilLoop).fBoxW + 15, tmDBSCCtrls(ilLoop).fBoxH + ilHeightAdj), BLUE, B
            Next ilLoop
            llTop = llTop + tmDBSCCtrls(1).fBoxH + ilHeightAdj
            ilHeightAdj = 15
        Next ilRow
    End If
    Contract.pbcDemoBar(ilIndex).FontSize = flFontSize
    Contract.pbcDemoBar(ilIndex).FontName = slFontName
    Contract.pbcDemoBar(ilIndex).FontSize = flFontSize
    Contract.pbcDemoBar(ilIndex).ForeColor = llColor
    Contract.pbcDemoBar(ilIndex).FontBold = True

    
End Sub


