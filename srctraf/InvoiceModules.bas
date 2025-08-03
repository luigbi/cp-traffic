Attribute VB_Name = "INVOICEMODULES"
' Copyright 1993 Counterpoint Software ® All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: InvoiceModules.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Initialize subs and functions
Option Explicit
Option Compare Text

Dim tmPcf As PCF
Dim tmIbf As IBF
Dim ibf_rst As ADODB.Recordset

'JW 01/27/2023 - Const's To help code readability in Invoices
Public Const INVTYPE_Commercial = 0     'Type/Commercial ckcType(0)
Public Const INVTYPE_PrintRep = 1       'Type/Print Rep Invoices ckcType(1)
Public Const INVTYPE_Installment = 2    'Type/Installment ckcType(2)
Public Const INVTYPE_NTR = 3            'Type/NTR ckcType(3)
Public Const INVTYPE_GenRepAR = 4       'Type/Gen Rep A/R ckcType(4)

Public Const INVGEN_Preliminary = 0     'Generate/Preliminary rbcType(0)
Public Const INVGEN_Final = 1           'Generate/Final rbcType(1)
Public Const INVGEN_Reprint = 2         'Generate/Reprint rbcType(2)
Public Const INVGEN_Aff = 3             'Generate/Aff rbcType(3)
Public Const INVGEN_Undo = 4            'Generate/Undo rbcType(4)
Public Const INVGEN_Archive = 5         'Generate/Archive rbcType(5)

Public Const INVCNTR_All = 0            'All Contracts rbcCntr(0)
Public Const INVCNTR_Selective = 1      'Selective Contracts rbcCntr(1)
Public Const INVCNTR_WOExpired = 2      'W/O Expired Contracts rbcCntr(2)
Public Const INVCNTR_Posted = 3         'Posted Contracts rbcCntr(3)

Public Const INVBILLCYCLE_STD = 0       'Bill Cycle/Std B'cast ckcBillCycle(0)
Public Const INVBILLCYCLE_Calendar = 1  'Bill Cycle/Calendar ckcBillCycle(1)
Public Const INVBILLCYCLE_Week = 2      'Bill Cycle/Week ckcBillCycle(2)

'tmIvr.iType: Note these values are also used in Invoice reports, so be careful changing these
Public Const IVRTYPE_NA = -1                'N/A - used during Reset IVR
Public Const IVRTYPE_Spot = 0               'Spot/Bonus Spot/Detail Airtime
Public Const IVRTYPE_Bonus = 1              '[Not used] seems 0 is used for Spots and Bonus Spots
Public Const IVRTYPE_TotalVefMkt = 2        'Vehicle or Market Subtotal
Public Const IVRTYPE_TotalAirtimeRep = 3    'AirTime or REP Total
Public Const IVRTYPE_AdServer = 4           'CPM Item
Public Const IVRTYPE_TotalAdServer = 5      'CPM Total
Public Const IVRTYPE_NTR = 6                'NTR Item
Public Const IVRTYPE_TotalNTR = 7           'NTR Total
Public Const IVRTYPE_TotalInstallment = 8   'Installment Total
Public Const IVRTYPE_TotalCntr = 9          'Contract Total

'*******************************************************
'*                                                     *
'*      Procedure Name:mCPM_ReadRec                    *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read CPM records               *
'*                      Pcf table                      *
'*                                                     *
'*******************************************************
Function mCPM_ReadRec() As Integer
'   iRet = mRepInv_ReadRec
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim slKey As String
    Dim slDate As String
    Dim slPcfStartDate As String
    Dim slPCFEndDate As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilAdd As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilFound As Integer
    Dim llEarliestDate As Long
    Dim ilIncludeCPM As Integer
    Dim ilAgePeriod As Integer
    Dim ilAgingYear As Integer
    Dim ilChfRecLen As Integer        'CHF record length
    Dim tlChf As CHF
    Dim tlChfSrchKey1 As CHFKEY1            'CHF record image

    ReDim tgCPMCntr(0 To 0) As CPMCNTR
    ReDim tmCPMCntStatus(0 To 0) As INVCNTRSTATUS
    
    imCPMStatusConflict = False
    If ((Asc(tgSaf(0).sFeatures8) And PODADSERVER) <> PODADSERVER) Then     'CPM Invoices
        mCPM_ReadRec = True
        Exit Function
    End If
    
    imPcfRecLen = Len(tmPcf)
    ilChfRecLen = Len(tlChf)
    imIbfRecLen = Len(tmIbf)
    If (Invoice.rbcType(INVGEN_Reprint).Value) Or (Invoice.rbcType(INVGEN_Archive).Value) Then        '11-15-16 reprint or archive
        'Need to have created an array of sbf records to reprint
        ilRet = mCPM_BuildReprint()
    ElseIf (Invoice.rbcType(INVGEN_Preliminary).Value) Or (Invoice.rbcType(INVGEN_Final).Value) Then
        'tmIbfSrchKey3.iBillYear = ilAgingYear
        'tmIbfSrchKey3.iBillMonth = ilAgePeriod
        'ilRet = btrGetGreaterOrEqual(hmIbf, tmIbf, imIbfRecLen, tmIbfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        gPackDate slStartDate, tmPcfSrchKey3.iEndDate(0), tmPcfSrchKey3.iEndDate(1)
        ilRet = btrGetGreaterOrEqual(hmPcf, tmPcf, imPcfRecLen, tmPcfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE)
            tmChfSrchKey.lCode = tmPcf.lChfCode
            ilRet = btrGetEqual(hmCHF, tlChf, ilChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)

            '-------------------------------------------
            'TTP 10697 - Ad server tab and working proposals status triggering billing alert
            'Check BillCycle
            If (Invoice.ckcBillCycle(INVBILLCYCLE_Calendar).Value = vbChecked And tlChf.sBillCycle = "C") Or (Invoice.ckcBillCycle(INVBILLCYCLE_STD).Value = vbChecked And tlChf.sBillCycle = "S") Then
                If Invoice.ckcBillCycle(INVBILLCYCLE_Calendar).Value = vbChecked And tlChf.sBillCycle = "C" Then    'Cal
                    slStartDate = smStartCal
                    slEndDate = smEndCal
                ElseIf Invoice.ckcBillCycle(INVBILLCYCLE_STD).Value = vbChecked And tlChf.sBillCycle = "S" Then
                    slStartDate = smStartStd
                    slEndDate = smEndStd
                End If
                '-------------------------------------------
                'Check for Rev Working - a contact that was scheduled then revised on the proposal screen (Final Only)
                'W=Working Proposal and P=Prevent scheduling
                'RE: v81 TTP 10697, per Jason: preliminary invoice process was supposed to check for unscheduled and rev working contracts.
                'If (ilRet = BTRV_ERR_NONE) And (tlChf.sDelete = "N") And (tlChf.sStatus = "W") And (tlChf.sSchStatus = "P") And (tlChf.iCntRevNo > 0) And Invoice.rbcType(INVGEN_Final).Value Then
                If (ilRet = BTRV_ERR_NONE) And (tlChf.sDelete = "N") And (tlChf.sStatus = "W") And (tlChf.sSchStatus = "P") And (tlChf.iCntRevNo > 0) Then
                    'Not Deleted and (Working Proposal and Prevent Schd and CntrRevNo > 0)
                    ilAdd = False
                    If (gDateValue(slPCFEndDate) >= gDateValue(slStartDate)) And (gDateValue(slPcfStartDate) <= gDateValue(slEndDate)) Then
                        ilAdd = True
                        'imCPMStatusConflict = True
                        For ilLoop = 1 To UBound(tmCPMCntStatus) - 1 Step 1
                            If tmCPMCntStatus(ilLoop).lCntrNo = tlChf.lCntrNo Then
                                ilAdd = False
                                '--------------------------------------------
                                If (tlChf.sSchStatus = "F") Or (tlChf.sSchStatus = "M") Then
                                    'Fully or Manually Scheduled
                                    tmCPMCntStatus(ilLoop).lChfCodeSchd = tlChf.lCode
                                    tmCPMCntStatus(ilLoop).lChfCodeAltered = 0
                                Else 'Not Fully or Manually Scheduled (Prevent Schd, Interupted Sched, New Cntr, Altered Cntr)
                                    tmCPMCntStatus(ilLoop).lChfCodeSchd = 0
                                    tmCPMCntStatus(ilLoop).lChfCodeAltered = tlChf.lCode
                                    tmCPMCntStatus(ilLoop).sSchStatus = tlChf.sSchStatus
                                End If
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If (ilAdd) Then
                        ilIndex = UBound(tmCPMCntStatus)
                        tmCPMCntStatus(ilIndex).lCntrNo = tlChf.lCntrNo
                        tmCPMCntStatus(ilIndex).iAdfCode = tlChf.iAdfCode
                        tmCPMCntStatus(ilIndex).sSchStatus = ""
                        tmCPMCntStatus(ilIndex).lEarliestDate = gDateValue(slStartDate)
                        tmCPMCntStatus(ilIndex).sBillCycle = tlChf.sBillCycle
                        '-------------------------------------------
                        If (tlChf.sSchStatus = "F") Or (tlChf.sSchStatus = "M") Then
                            'Fully or Manually Scheduled
                            tmCPMCntStatus(ilIndex).lChfCodeSchd = tlChf.lCode
                            tmCPMCntStatus(ilIndex).lChfCodeAltered = 0
                        Else
                            'Not Fully or Manually Scheduled (Prevent Schd, Interupted Sched, New Cntr, Altered Cntr)
                            tmCPMCntStatus(ilIndex).lChfCodeSchd = 0
                            tmCPMCntStatus(ilIndex).lChfCodeAltered = tlChf.lCode
                            tmCPMCntStatus(ilIndex).sSchStatus = tlChf.sSchStatus
                        End If
                        ReDim Preserve tmCPMCntStatus(0 To ilIndex + 1) As INVCNTRSTATUS
                    End If
                End If
                
                '-------------------------------------------
                'Check Unscheduled (Final & Prelim)
                'N=Unscheduled Order and A=Altered Contract
                If (ilRet = BTRV_ERR_NONE) And (tlChf.sDelete = "N") And (tlChf.sStatus = "N") And (tlChf.sSchStatus = "A") Then
                    'Not Deleted and (Unscheduled and Altered Cntr)
                    ilAdd = False
                    If (gDateValue(slPCFEndDate) >= gDateValue(slStartDate)) And (gDateValue(slPcfStartDate) <= gDateValue(slEndDate)) Then
                        ilAdd = True
                        'imCPMStatusConflict = True
                        For ilLoop = 1 To UBound(tmCPMCntStatus) - 1 Step 1
                            If tmCPMCntStatus(ilLoop).lCntrNo = tlChf.lCntrNo Then
                                ilAdd = False
                                '-------------------------------------------
                                If (tlChf.sSchStatus = "F") Or (tlChf.sSchStatus = "M") Then
                                    'Fully or Manually Scheduled
                                    tmCPMCntStatus(ilLoop).lChfCodeSchd = tlChf.lCode
                                    tmCPMCntStatus(ilLoop).lChfCodeAltered = 0
                                Else
                                    'Not Fully or Manually Scheduled (Prevent Schd, Interupted Sched, New Cntr, Altered Cntr)
                                    tmCPMCntStatus(ilLoop).lChfCodeSchd = 0
                                    tmCPMCntStatus(ilLoop).lChfCodeAltered = tlChf.lCode
                                    tmCPMCntStatus(ilLoop).sSchStatus = tlChf.sSchStatus
                                End If
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If (ilAdd) Then
                        ilIndex = UBound(tmCPMCntStatus)
                        tmCPMCntStatus(ilIndex).lCntrNo = tlChf.lCntrNo
                        tmCPMCntStatus(ilIndex).iAdfCode = tlChf.iAdfCode
                        tmCPMCntStatus(ilIndex).sSchStatus = ""
                        tmCPMCntStatus(ilIndex).lEarliestDate = gDateValue(slStartDate)
                        tmCPMCntStatus(ilIndex).sBillCycle = tlChf.sBillCycle
                        '-------------------------------------------
                        If (tlChf.sSchStatus = "F") Or (tlChf.sSchStatus = "M") Then
                            'Fully or Manually Scheduled
                            tmCPMCntStatus(ilIndex).lChfCodeSchd = tlChf.lCode
                            tmCPMCntStatus(ilIndex).lChfCodeAltered = 0
                        Else
                            'Not Fully or Manually Scheduled (Prevent Schd, Interupted Sched, New Cntr, Altered Cntr)
                            tmCPMCntStatus(ilIndex).lChfCodeSchd = 0
                            tmCPMCntStatus(ilIndex).lChfCodeAltered = tlChf.lCode
                            tmCPMCntStatus(ilIndex).sSchStatus = tlChf.sSchStatus
                        End If
                        ReDim Preserve tmCPMCntStatus(0 To ilIndex + 1) As INVCNTRSTATUS
                    End If
                End If
            End If
            
            '-------------------------------------------
            'Check for scheduled Orders bill-cycle
            ilFound = False
            If ilRet = BTRV_ERR_NONE And ((tlChf.sStatus = "O") Or (tlChf.sStatus = "H")) And ((tlChf.sSchStatus = "F") Or (tlChf.sSchStatus = "M")) And (tlChf.sDelete = "N") Then
                '(O=Order(Scheduled order) or H=Hold) and (F=Fully scheduled or M=Manually scheduled) and Not Deleted
                If Invoice.ckcBillCycle(INVBILLCYCLE_Calendar).Value = vbChecked And tlChf.sBillCycle = "C" Then    'Cal
                    ilAgePeriod = Month(gDateValue(smEndCal))
                    ilAgingYear = Year(gDateValue(smEndCal))
                    slStartDate = smStartCal
                    slEndDate = smEndCal
                    ilFound = True
                ElseIf Invoice.ckcBillCycle(INVBILLCYCLE_Week).Value = vbChecked And tlChf.sBillCycle = "W" Then    'Week
                ElseIf Invoice.ckcBillCycle(INVBILLCYCLE_STD).Value = vbChecked And tlChf.sBillCycle = "S" Then     'Std
                    ilAgePeriod = Month(gDateValue(smEndStd))
                    ilAgingYear = Year(gDateValue(smEndStd))
                    slStartDate = smStartStd
                    slEndDate = smEndStd
                    ilFound = True
                End If
            End If
            
            '-------------------------------------------
            If ilFound Then
                'Test pcf Date and Load IBF
                gUnpackDate tmPcf.iStartDate(0), tmPcf.iStartDate(1), slPcfStartDate
                gUnpackDate tmPcf.iEndDate(0), tmPcf.iEndDate(1), slPCFEndDate
                If (gDateValue(slPCFEndDate) >= gDateValue(slStartDate)) And (gDateValue(slPcfStartDate) <= gDateValue(slEndDate)) Then
                    For ilLoop = 1 To UBound(lmCPMBypassCntr) - 1 Step 1
                        If lmCPMBypassCntr(ilLoop) = tlChf.lCode Then
                            ilRet = Not BTRV_ERR_NONE
                            Exit For
                        End If
                    Next ilLoop
                    
                    If ilRet = BTRV_ERR_NONE Then
                        tmIbfSrchKey1.lCntrNo = tlChf.lCntrNo
                        tmIbfSrchKey1.iPodCPMID = tmPcf.iPodCPMID
                        ilRet = btrGetEqual(hmIbf, tmIbf, imIbfRecLen, tmIbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmIbf.lCntrNo = tlChf.lCntrNo) And (tmIbf.iPodCPMID = tmPcf.iPodCPMID)
                            If (tmIbf.iBillMonth = ilAgePeriod) And (tmIbf.iBillYear = ilAgingYear) Then
                                Exit Do
                            End If
                            ilRet = btrGetNext(hmIbf, tmIbf, imIbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If (ilRet <> BTRV_ERR_NONE) Or (tmIbf.lCntrNo <> tlChf.lCntrNo) Or (tmIbf.iPodCPMID <> tmPcf.iPodCPMID) Or (tmIbf.iBillMonth <> ilAgePeriod) Or (tmIbf.iBillYear <> ilAgingYear) Then
                            If tmPcf.sPriceType = "F" Then  'Create missing record as Flat Rate does Not require any posting
                                mAddIbf tlChf
                                ilRet = BTRV_ERR_NONE
                            Else
                                'Missing billing info
                                tmIbf.lCode = 0
                                tmIbf.sBilled = "N"
                                tmIbf.sBillCycle = tlChf.sBillCycle
                                ilRet = BTRV_ERR_NONE
                            End If
                        End If
                        
                        'If (tmIbf.sBilled <> "Y") And (((Invoice.ckcBillCycle(INVBILLCYCLE_Calendar).Value = vbChecked) And (tmIbf.sBillCycle = "C")) Or ((Invoice.ckcBillCycle(INVBILLCYCLE_STD).Value = vbChecked) And (tmIbf.sBillCycle = "S"))) Then
                        If (tmIbf.sBilled <> "Y") Then
                            ilAdd = True
                            If Invoice.rbcCntr(INVCNTR_Selective).Value Or Invoice.rbcCntr(INVCNTR_WOExpired).Value Or Invoice.rbcCntr(INVCNTR_Posted).Value Then
                                ilAdd = False
                                For ilLoop = 0 To UBound(lmSelCntrCode) - 1 Step 1
                                    If tmPcf.lChfCode = lmSelCntrCode(ilLoop) Then
                                        ilAdd = True
                                        Exit For
                                    End If
                                Next ilLoop
                            End If
                            If ilAdd Then
                                If (tgSpf.sInvVehSel = "Y") And (Invoice.ckcAll.Value = vbUnchecked) Then
                                    If Not Invoice.mAllVehSel("F", tmPcf.lChfCode) Then
                                        ilAdd = False
                                    End If
                                End If
                            End If
                            If ilAdd Then
                                If (Invoice.ckcType(INVTYPE_Commercial).Value = vbUnchecked) And (Invoice.ckcType(INVTYPE_NTR).Value = vbChecked) Then
                                    ''Only allow contract if no air time
                                    'ilRet = gObtainChfClf(hmCHF, hmClf, tmPcf.lChfCode, False, tlChf, tgClfInv())
                                    'If tlChf.sInstallDefined = "Y" Then
                                    If tlChf.sInstallDefined = "Y" Then
                                        ilAdd = False
                                    End If
                                End If
                            End If
                            'If ilAdd Then
                            '    If mSetNTRBillFlagIfMissing() Then
                            '        ilAdd = False
                            '    End If
                            'End If
                            If ilAdd Then
                                slKey = Trim$(str$(tmPcf.lChfCode))
                                Do While Len(slKey) < 8
                                    slKey = "0" & slKey
                                Loop
                                'tmChfSrchKey.lCode = tmPcf.lChfCode
                                'ilRet = btrGetEqual(hmCHF, tlChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                '---------------------------------------------
                                '(W=Working Proposal or C=Completed Proposal)
                                If (ilRet = BTRV_ERR_NONE) And (tlChf.sDelete = "N") And (tlChf.iCntRevNo = 0) And ((tlChf.sStatus = "W") Or (tlChf.sStatus = "C")) Then
                                    tlChf.sDelete = "Y"
                                End If

                                '---------------------------------------------
                                'Not D=Rejected
                                If (ilRet = BTRV_ERR_NONE) And (tlChf.sDelete = "N") And (tlChf.sStatus <> "D") And (tmPcf.sType <> "P") Then
                                    'Not Deleted and (Cntr Not Rejected and Line type <> Package)
                                    ilFound = False
                                    For ilIndex = 0 To UBound(tmCPMCntStatus) - 1 Step 1
                                        If tmCPMCntStatus(ilIndex).lCntrNo = tlChf.lCntrNo Then
                                            ilFound = True
                                            'gUnpackDateLongError tmPcf.iStartDate(0), tmPcf.iStartDate(1), llEarliestDate, "2:mNTR_ReadRec " & tmPcf.lCode
                                            'If llEarliestDate < tmCPMCntStatus(ilIndex).lEarliestDate Then
                                            '    tmCPMCntStatus(ilIndex).lEarliestDate = llEarliestDate
                                            'End If
                                            '-------------------------------------------
                                            'If F=Fully scheduled or M=Manually scheduled
                                            If (tlChf.sSchStatus = "F") Or (tlChf.sSchStatus = "M") Then 'Fully or Manually Scheduled
                                                tmCPMCntStatus(ilIndex).lChfCodeSchd = tlChf.lCode
                                                'If tmCPMCntStatus(ilIndex).lChfCodeAltered <> 0 Then
                                                '    imCPMStatusConflict = True
                                                'End If
                                            Else 'Not Fully or Manually Scheduled (Prevent Schd, Interupted Sched, New Cntr, Altered Cntr)
                                                tmCPMCntStatus(ilIndex).lChfCodeAltered = tlChf.lCode
                                                tmCPMCntStatus(ilIndex).sSchStatus = tlChf.sSchStatus
                                                'If tmCPMCntStatus(ilIndex).lChfCodeSchd <> 0 Then
                                                '    imCPMStatusConflict = True
                                                'End If
                                            End If
                                        End If
                                    Next ilIndex
                                    If (Not ilFound) Then
                                        ilIndex = UBound(tmCPMCntStatus)
                                        tmCPMCntStatus(ilIndex).lCntrNo = tlChf.lCntrNo
                                        tmCPMCntStatus(ilIndex).iAdfCode = tlChf.iAdfCode
                                        tmCPMCntStatus(ilIndex).sSchStatus = ""
                                        'gUnpackDateLongError tmPcf.iStartDate(0), tmPcf.iStartDate(1), tmCPMCntStatus(ilIndex).lEarliestDate, "3:mNTR_ReadRec " & tmPcf.lCode
                                        tmCPMCntStatus(ilIndex).lEarliestDate = gDateValue(slStartDate)
                                        tmCPMCntStatus(ilIndex).sBillCycle = tlChf.sBillCycle
                                        '-------------------------------------------
                                        'F=Fully scheduled or M=Manually scheduled
                                        If (tlChf.sSchStatus = "F") Or (tlChf.sSchStatus = "M") Then 'Fully or Manually Scheduled
                                            tmCPMCntStatus(ilIndex).lChfCodeSchd = tlChf.lCode
                                            tmCPMCntStatus(ilIndex).lChfCodeAltered = 0
                                        Else 'Not Fully or Manually Scheduled (Prevent Schd, Interupted Sched, New Cntr, Altered Cntr)
                                            tmCPMCntStatus(ilIndex).lChfCodeSchd = 0
                                            tmCPMCntStatus(ilIndex).lChfCodeAltered = tlChf.lCode
                                            tmCPMCntStatus(ilIndex).sSchStatus = tlChf.sSchStatus
                                        End If
                                        ReDim Preserve tmCPMCntStatus(0 To ilIndex + 1) As INVCNTRSTATUS
                                    End If
                                End If
                                
                                '-------------------------------------------
                                'Not D=Rejected And (F=Fully scheduled or M=Manually scheduled)
                                If (ilRet = BTRV_ERR_NONE) And (tlChf.sDelete = "N") And (tlChf.sStatus <> "D") And ((tlChf.sSchStatus = "F") Or (tlChf.sSchStatus = "M")) Then
                                    tlChf.lVefCode = tmPcf.iVefCode
                                    slKey = mCreateField1Key(tlChf) & slKey
                                    'gUnpackDate tmPcf.iStartDate(0), tmPcf.iStartDate(1), slDate
                                    slDate = slStartDate
                                    slDate = Trim$(str$(gDateValue(slDate)))
                                    Do While Len(slDate) < 5
                                        slDate = "0" & slDate
                                    Loop
                                    'For GenIvr to work, package must sort after hidden lines
                                    If tmPcf.sType = "P" Then
                                        tgCPMCntr(UBound(tgCPMCntr)).sKey = slKey & "00000000" + slDate & "B"
                                    Else
                                        tgCPMCntr(UBound(tgCPMCntr)).sKey = slKey & "00000000" + slDate & "A"
                                    End If
                                    tgCPMCntr(UBound(tgCPMCntr)).lSDate = Val(slDate)
                                    tgCPMCntr(UBound(tgCPMCntr)).lEDate = gDateValue(slEndDate)
                                    tgCPMCntr(UBound(tgCPMCntr)).tPcf = tmPcf
                                    tgCPMCntr(UBound(tgCPMCntr)).tIbf = tmIbf
                                    ReDim Preserve tgCPMCntr(0 To UBound(tgCPMCntr) + 1) As CPMCNTR
                                    'TTP 10515 - NTR Invoices - "NTR INVOICE and AFFIDAVIT" not displaying correct on NTR Invoices it shows "INVOICES and AFFIDAVIT" started with V81
                                    Invoice.mSetAirNTRStatus tlChf.lCntrNo, "AIR", IIF(tlChf.iAgfCode = 0, tlChf.iAdfCode, tlChf.iAgfCode)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            ilRet = btrGetNext(hmPcf, tmPcf, imPcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    
    'Sort by Contract and date- process one contract/date at a time
    If UBound(tgCPMCntr) > 0 Then
        ArraySortTyp fnAV(tgCPMCntr(), 0), UBound(tgCPMCntr), 0, LenB(tgCPMCntr(0)), 0, LenB(tgCPMCntr(0).sKey), 0
    End If
    
    imCPMStatusConflict = False
    For ilIndex = 0 To UBound(tmCPMCntStatus) - 1 Step 1
        If mCompareChfInPast("C", tmCPMCntStatus(ilIndex).lChfCodeSchd, tmCPMCntStatus(ilIndex).lChfCodeAltered, tmCPMCntStatus(ilIndex).lEarliestDate, tmCPMCntStatus(ilIndex).sBillCycle) Then
            tmCPMCntStatus(ilIndex).lChfCodeSchd = 0
            tmCPMCntStatus(ilIndex).lChfCodeAltered = 0
        Else
            imCPMStatusConflict = True
        End If
    Next ilIndex
    mCPM_ReadRec = True
    Exit Function
End Function

Public Function mCompareChfInPast(slType As String, llSchChfCode As Long, llAlteredChfCode As Long, llEarliestDate As Long, slBillCycle As String) As Integer
'    slType(I):  A=Air Time; N=NTR; I =Install; C=CPM
    Dim ilRet As Integer
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llSvStartDate As Long
    Dim llSvEndDate As Long
    Dim ilAlter As Integer
    Dim ilSch As Integer
    Dim llAlterStartDate As Long
    Dim llAlterEndDate As Long
    Dim llSchStartDate As Long
    Dim llSchEndDate As Long
    Dim ilAlterSpots As Integer
    Dim llAlterPrice As Long
    Dim ilSchSpots As Integer
    Dim llSchPrice As Long
    Dim ilDay As Integer
    Dim llCffStartDate As Long
    Dim llCffEndDate As Long
    Dim ilCff As Integer
    Dim ilLnFd As Integer

    mCompareChfInPast = True
    If (llSchChfCode = 0) And (llAlteredChfCode = 0) Then
        Exit Function
    End If
    If (llSchChfCode <> 0) And (llAlteredChfCode = 0) Then
        Exit Function
    End If
    If (llAlteredChfCode <> 0) And (llSchChfCode = 0) Then
        If slType = "A" Then 'A=Air Time
            ilRet = gObtainCntrPlusGame(hmCHF, Invoice.hmClf, Invoice.hmCff, Invoice.hmCgf, llAlteredChfCode, False, tmAlterChf, tmAlterClf(), tmAlterCff(), tmAlterCgf())
            If Not ilRet Then
                mCompareChfInPast = False
                Exit Function
            End If
            If tmAlterChf.sBillCycle = "C" Then
                llStartDate = lmSvStartCal
                llEndDate = lmSvEndCal
            ElseIf tmAlterChf.sBillCycle = "W" Then
                llStartDate = lmSvStartWk
                llEndDate = lmSvEndWk
            Else
                llStartDate = lmSvStartStd
                llEndDate = lmSvEndStd
            End If
            For ilAlter = LBound(tmAlterClf) To UBound(tmAlterClf) - 1 Step 1
                'gUnpackDateLong tmAlterClf(ilAlter).ClfRec.iStartDate(0), tmAlterClf(ilAlter).ClfRec.iStartDate(1), llAlterStartDate
                gUnpackDateLongError tmAlterClf(ilAlter).ClfRec.iStartDate(0), tmAlterClf(ilAlter).ClfRec.iStartDate(1), llAlterStartDate, "55:mCompareChfInPast " & tmAlterClf(ilAlter).ClfRec.lCode
                'gUnpackDateLong tmAlterClf(ilAlter).ClfRec.iEndDate(0), tmAlterClf(ilAlter).ClfRec.iEndDate(1), llAlterEndDate
                gUnpackDateLongError tmAlterClf(ilAlter).ClfRec.iEndDate(0), tmAlterClf(ilAlter).ClfRec.iEndDate(1), llAlterEndDate, "56:mCompareChfInPast " & tmAlterClf(ilAlter).ClfRec.lCode
                If (llAlterStartDate <= llEndDate) And (llAlterEndDate >= llStartDate) Then
                    mCompareChfInPast = False
                    Exit Function
                End If
            Next ilAlter
            Exit Function
        End If
        If slType = "N" Then 'N=NTR
            If llEarliestDate <= lmNTRDate Then
                mCompareChfInPast = False
            End If
        End If
        If slType = "I" Then 'I=Install
            If slBillCycle = "C" Then
                If (llEarliestDate <= lmEndCal) And (llEarliestDate >= lmStartCal) Then
                    mCompareChfInPast = False
                End If
            ElseIf slBillCycle = "W" Then
                If (llEarliestDate <= lmEndWk) And (llEarliestDate >= lmStartWk) Then
                    mCompareChfInPast = False
                End If
            Else
                If (llEarliestDate <= lmEndStd) And (llEarliestDate >= lmStartStd) Then
                    mCompareChfInPast = False
                End If
            End If
        End If
        If slType = "C" Then 'C=CPM
            If slBillCycle = "C" Then
                If (llEarliestDate <= lmEndCal) And (llEarliestDate >= lmStartCal) Then
                    mCompareChfInPast = False
                End If
            ElseIf slBillCycle = "W" Then
                If (llEarliestDate <= lmEndWk) And (llEarliestDate >= lmStartWk) Then
                    mCompareChfInPast = False
                End If
            Else
                If (llEarliestDate <= lmEndStd) And (llEarliestDate >= lmStartStd) Then
                    mCompareChfInPast = False
                End If
            End If
        End If
        Exit Function
    End If
    If slType = "A" Then 'A=Air Time
        ilRet = gObtainCntrPlusGame(hmCHF, Invoice.hmClf, Invoice.hmCff, Invoice.hmCgf, llSchChfCode, False, tmSchChf, tmSchClf(), tmSchCff(), tmSchCgf())
        If Not ilRet Then
            mCompareChfInPast = False
            Exit Function
        End If
        ilRet = gObtainCntrPlusGame(hmCHF, Invoice.hmClf, Invoice.hmCff, Invoice.hmCgf, llAlteredChfCode, False, tmAlterChf, tmAlterClf(), tmAlterCff(), tmAlterCgf())
        If Not ilRet Then
            mCompareChfInPast = False
            Exit Function
        End If
        If tmSchChf.sBillCycle = "C" Then
            llStartDate = lmSvStartCal
            llEndDate = lmSvEndCal
        ElseIf tmSchChf.sBillCycle = "W" Then
            llStartDate = lmSvStartWk
            llEndDate = lmSvEndWk
        Else
            llStartDate = lmSvStartStd
            llEndDate = lmSvEndStd
        End If
        llSvStartDate = llStartDate
        llSvEndDate = llEndDate
        'Compare only air values prior or equal to end date
        'Check for Weeks added, removed or changed
        For ilAlter = LBound(tmAlterClf) To UBound(tmAlterClf) - 1 Step 1
            ilLnFd = False
            llStartDate = llSvStartDate
            llEndDate = llSvEndDate
            For ilSch = LBound(tmSchClf) To UBound(tmSchClf) - 1 Step 1
                If tmAlterClf(ilAlter).ClfRec.iLine = tmSchClf(ilSch).ClfRec.iLine Then
                    ilLnFd = True
                    'gUnpackDateLong tmAlterClf(ilAlter).ClfRec.iStartDate(0), tmAlterClf(ilAlter).ClfRec.iStartDate(1), llAlterStartDate
                    gUnpackDateLongError tmAlterClf(ilAlter).ClfRec.iStartDate(0), tmAlterClf(ilAlter).ClfRec.iStartDate(1), llAlterStartDate, "57:mCompareChfInPast " & tmAlterClf(ilAlter).ClfRec.lCode
                    'gUnpackDateLong tmAlterClf(ilAlter).ClfRec.iEndDate(0), tmAlterClf(ilAlter).ClfRec.iEndDate(1), llAlterEndDate
                    gUnpackDateLongError tmAlterClf(ilAlter).ClfRec.iEndDate(0), tmAlterClf(ilAlter).ClfRec.iEndDate(1), llAlterEndDate, "58:mCompareChfInPast " & tmAlterClf(ilAlter).ClfRec.lCode
                    'gUnpackDateLong tmSchClf(ilSch).ClfRec.iStartDate(0), tmSchClf(ilSch).ClfRec.iStartDate(1), llSchStartDate
                    gUnpackDateLongError tmSchClf(ilSch).ClfRec.iStartDate(0), tmSchClf(ilSch).ClfRec.iStartDate(1), llSchStartDate, "59:mCompareChfInPast " & tmSchClf(ilSch).ClfRec.lCode
                    'gUnpackDateLong tmSchClf(ilSch).ClfRec.iEndDate(0), tmSchClf(ilSch).ClfRec.iEndDate(1), llSchEndDate
                    gUnpackDateLongError tmSchClf(ilSch).ClfRec.iEndDate(0), tmSchClf(ilSch).ClfRec.iEndDate(1), llSchEndDate, "60:mCompareChfInPast " & tmSchClf(ilSch).ClfRec.lCode
                    'If (llAlterStartDate <= llEndDate) Or (llSchStartDate <= llEndDate) Then
                    If (llAlterStartDate <= llEndDate) Or ((llSchStartDate <= llEndDate) And (llSchStartDate <= llSchEndDate)) Then
                        If llAlterStartDate <> llSchStartDate Then
                            mCompareChfInPast = False
                            Exit Function
                        End If
                        
                        'L.Bianchi 06/23/2021 TTP 10204
                        If llAlterEndDate <> llSchEndDate Then
                            mCompareChfInPast = False
                            Exit Function
                        End If
                        
                        If llSchStartDate > llStartDate Then
                            llStartDate = llSchStartDate
                        End If
                        For llDate = llStartDate To llEndDate Step 7
                            ilAlterSpots = 0
                            llAlterPrice = -1
                            ilCff = tmAlterClf(ilAlter).iFirstCff
                            Do While ilCff <> -1
                                llCffStartDate = tmAlterCff(ilCff).lStartDate
                                llCffEndDate = tmAlterCff(ilCff).lEndDate
                                If (llDate >= llCffStartDate) And (llDate <= llCffEndDate) Then
                                    If tmAlterCff(ilCff).CffRec.sDyWk = "D" Then
                                        For ilDay = 0 To 6 Step 1
                                            ilAlterSpots = ilAlterSpots + tmAlterCff(ilCff).CffRec.iDay(ilDay)
                                        Next ilDay
                                    Else
                                        ilAlterSpots = tmAlterCff(ilCff).CffRec.iSpotsWk + tmAlterCff(ilCff).CffRec.iXSpotsWk
                                    End If
                                    llAlterPrice = tmAlterCff(ilCff).CffRec.lActPrice
                                    Exit Do
                                End If
                                ilCff = tmAlterCff(ilCff).iNextCff
                            Loop
                            ilSchSpots = 0
                            llSchPrice = -1
                            ilCff = tmSchClf(ilSch).iFirstCff
                            Do While ilCff <> -1
                                llCffStartDate = tmSchCff(ilCff).lStartDate
                                llCffEndDate = tmSchCff(ilCff).lEndDate
                                If (llDate >= llCffStartDate) And (llDate <= llCffEndDate) Then
                                    If tmSchCff(ilCff).CffRec.sDyWk = "D" Then
                                        For ilDay = 0 To 6 Step 1
                                            ilSchSpots = ilSchSpots + tmSchCff(ilCff).CffRec.iDay(ilDay)
                                        Next ilDay
                                    Else
                                        ilSchSpots = tmSchCff(ilCff).CffRec.iSpotsWk + tmSchCff(ilCff).CffRec.iXSpotsWk
                                    End If
                                    llSchPrice = tmSchCff(ilCff).CffRec.lActPrice
                                    Exit Do
                                End If
                                ilCff = tmSchCff(ilCff).iNextCff
                            Loop
                            If (llAlterPrice <> llSchPrice) Or (ilAlterSpots <> ilSchSpots) Then
                                mCompareChfInPast = False
                                Exit Function
                            End If
                        Next llDate
                    End If
                End If
            Next ilSch
            If Not ilLnFd Then
                'gUnpackDateLong tmAlterClf(ilAlter).ClfRec.iStartDate(0), tmAlterClf(ilAlter).ClfRec.iStartDate(1), llAlterStartDate
                gUnpackDateLongError tmAlterClf(ilAlter).ClfRec.iStartDate(0), tmAlterClf(ilAlter).ClfRec.iStartDate(1), llAlterStartDate, "61:mCompareChfInPast " & tmAlterClf(ilAlter).ClfRec.lCode
                'gUnpackDateLong tmAlterClf(ilAlter).ClfRec.iEndDate(0), tmAlterClf(ilAlter).ClfRec.iEndDate(1), llAlterEndDate
                gUnpackDateLongError tmAlterClf(ilAlter).ClfRec.iEndDate(0), tmAlterClf(ilAlter).ClfRec.iEndDate(1), llAlterEndDate, "62:mCompareChfInPast " & tmAlterClf(ilAlter).ClfRec.lCode
                If llAlterEndDate >= llAlterStartDate Then
                    If (llAlterEndDate >= llStartDate) And (llAlterStartDate <= llEndDate) Then
                        mCompareChfInPast = False
                        Exit Function
                    End If
                End If
            End If
        Next ilAlter
        For ilSch = LBound(tmSchClf) To UBound(tmSchClf) - 1 Step 1
            ilLnFd = False
            For ilAlter = LBound(tmAlterClf) To UBound(tmAlterClf) - 1 Step 1
                If tmAlterClf(ilAlter).ClfRec.iLine = tmSchClf(ilSch).ClfRec.iLine Then
                    ilLnFd = True
                End If
            Next ilAlter
            If Not ilLnFd Then
                'gUnpackDateLong tmSchClf(ilSch).ClfRec.iStartDate(0), tmSchClf(ilSch).ClfRec.iStartDate(1), llSchStartDate
                gUnpackDateLongError tmSchClf(ilSch).ClfRec.iStartDate(0), tmSchClf(ilSch).ClfRec.iStartDate(1), llSchStartDate, "63:mCompareChfInPast " & tmSchClf(ilSch).ClfRec.lCode
                'gUnpackDateLong tmSchClf(ilSch).ClfRec.iEndDate(0), tmSchClf(ilSch).ClfRec.iEndDate(1), llSchEndDate
                gUnpackDateLongError tmSchClf(ilSch).ClfRec.iEndDate(0), tmSchClf(ilSch).ClfRec.iEndDate(1), llSchEndDate, "64:mCompareChfInPast " & tmSchClf(ilSch).ClfRec.lCode
                If llSchEndDate >= llSchStartDate Then
                    If (llSchEndDate >= llStartDate) And (llSchStartDate <= llEndDate) Then
                        mCompareChfInPast = False
                        Exit Function
                    End If
                End If
            End If
        Next ilSch

    ElseIf slType = "N" Then 'N=NTR
        '1/15/09:  Force as if not matching.  This is required as if a contract needs to be scheduled, then two images of sbf will exist and can't
        '          invoice the original.  This would cause double billing
'        tmChfSrchKey.lCode = llSchChfCode
'        ilRet = btrGetEqual(hmChf, tmSchChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet <> BTRV_ERR_NONE Then
'            mCompareChfInPast = False
'            Exit Function
'        End If
'        If tmSchChf.sBillCycle = "C" Then
'            llStartDate = lmSvStartCal
'            llEndDate = lmSvEndCal
'        Else
'            llStartDate = lmSvStartStd
'            llEndDate = lmSvEndStd
'        End If
'        ReDim tmAlterIBSbf(0 To 0) As SBFLIST
'        ilUpper = 0                             '11-29-07
'        tmSbfSrchKey0.lChfCode = llAlteredChfCode
'        tmSbfSrchKey0.iDate(0) = 0
'        tmSbfSrchKey0.iDate(1) = 0
'        tmSbfSrchKey0.sTranType = " "
'        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = llAlteredChfCode)
'            If tmSbf.sTranType = "I" Then   'Items Billing (sBill is ignored)
'                gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
'                If (llDate >= llStartDate) And (llDate <= llEndDate) Then
'                    tmAlterIBSbf(ilUpper).SbfRec = tmSbf
'                    tmAlterIBSbf(ilUpper).lRecPos = 0
'                    tmAlterIBSbf(ilUpper).iStatus = 1
'                    ilUpper = ilUpper + 1
'                    ReDim Preserve tmAlterIBSbf(0 To ilUpper) As SBFLIST
'                End If
'            End If
'            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'        ReDim tmSchIBSbf(0 To 0) As SBFLIST
'        ilUpper = 0
'        tmSbfSrchKey0.lChfCode = llSchChfCode
'        tmSbfSrchKey0.iDate(0) = 0
'        tmSbfSrchKey0.iDate(1) = 0
'        tmSbfSrchKey0.sTranType = " "
'        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = llSchChfCode)
'            If tmSbf.sTranType = "I" Then   'Items Billing (sBill is ignored)
'                gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
'                If (llDate >= llStartDate) And (llDate <= llEndDate) Then
'                    tmSchIBSbf(ilUpper).SbfRec = tmSbf
'                    tmSchIBSbf(ilUpper).lRecPos = 0
'                    tmSchIBSbf(ilUpper).iStatus = 1
'                    ilUpper = ilUpper + 1
'                    ReDim Preserve tmSchIBSbf(0 To ilUpper) As SBFLIST
'                End If
'            End If
'            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'        For ilAlter = 0 To UBound(tmAlterIBSbf) - 1 Step 1
'            ilMatch = False
'            For ilSch = 0 To UBound(tmSchIBSbf) - 1 Step 1
'                If (tmAlterIBSbf(ilAlter).SbfRec.iDate(0) = tmSchIBSbf(ilSch).SbfRec.iDate(0)) And (tmAlterIBSbf(ilAlter).SbfRec.iDate(1) = tmSchIBSbf(ilSch).SbfRec.iDate(1)) Then
'                    If (tmAlterIBSbf(ilAlter).SbfRec.iBillVefCode = tmSchIBSbf(ilSch).SbfRec.iBillVefCode) And (tmAlterIBSbf(ilAlter).SbfRec.iAirVefCode = tmSchIBSbf(ilSch).SbfRec.iAirVefCode) Then
'                        If (tmAlterIBSbf(ilAlter).SbfRec.iMnfItem = tmSchIBSbf(ilSch).SbfRec.iMnfItem) Then
'                            If (tmAlterIBSbf(ilAlter).SbfRec.iNoItems = tmSchIBSbf(ilSch).SbfRec.iNoItems) And (tmAlterIBSbf(ilAlter).SbfRec.lGross = tmSchIBSbf(ilSch).SbfRec.lGross) Then
'                                If (tmAlterIBSbf(ilAlter).SbfRec.sAgyComm = tmSchIBSbf(ilSch).SbfRec.sAgyComm) And (tmAlterIBSbf(ilAlter).SbfRec.iTrfCode = tmSchIBSbf(ilSch).SbfRec.iTrfCode) And (tmAlterIBSbf(ilAlter).SbfRec.lAcquisitionCost = tmSchIBSbf(ilSch).SbfRec.lAcquisitionCost) Then
'                                    ilMatch = True
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
'            Next ilSch
'            If Not ilMatch Then
'                mCompareChfInPast = False
'                Exit Function
'            End If
'        Next ilAlter
'        For ilSch = 0 To UBound(tmSchIBSbf) - 1 Step 1
'            ilMatch = False
'            For ilAlter = 0 To UBound(tmAlterIBSbf) - 1 Step 1
'                If (tmAlterIBSbf(ilAlter).SbfRec.iDate(0) = tmSchIBSbf(ilSch).SbfRec.iDate(0)) And (tmAlterIBSbf(ilAlter).SbfRec.iDate(1) = tmSchIBSbf(ilSch).SbfRec.iDate(1)) Then
'                    If (tmAlterIBSbf(ilAlter).SbfRec.iBillVefCode = tmSchIBSbf(ilSch).SbfRec.iBillVefCode) And (tmAlterIBSbf(ilAlter).SbfRec.iAirVefCode = tmSchIBSbf(ilSch).SbfRec.iAirVefCode) Then
'                        If (tmAlterIBSbf(ilAlter).SbfRec.iMnfItem = tmSchIBSbf(ilSch).SbfRec.iMnfItem) Then
'                            If (tmAlterIBSbf(ilAlter).SbfRec.iNoItems = tmSchIBSbf(ilSch).SbfRec.iNoItems) And (tmAlterIBSbf(ilAlter).SbfRec.lGross = tmSchIBSbf(ilSch).SbfRec.lGross) Then
'                                If (tmAlterIBSbf(ilAlter).SbfRec.sAgyComm = tmSchIBSbf(ilSch).SbfRec.sAgyComm) And (tmAlterIBSbf(ilAlter).SbfRec.iTrfCode = tmSchIBSbf(ilSch).SbfRec.iTrfCode) And (tmAlterIBSbf(ilAlter).SbfRec.lAcquisitionCost = tmSchIBSbf(ilSch).SbfRec.lAcquisitionCost) Then
'                                    ilMatch = True
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
'            Next ilAlter
'            If Not ilMatch Then
'                mCompareChfInPast = False
'                Exit Function
'            End If
'        Next ilSch
        mCompareChfInPast = False
    ElseIf slType = "I" Then 'I=Install
        '1/15/09:  Force as if not matching.  This is required as if a contract needs to be scheduled, then two images of sbf will exist and can't
        '          invoice the original.  This would cause double billing
'        tmChfSrchKey.lCode = llSchChfCode
'        ilRet = btrGetEqual(hmChf, tmSchChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'        If ilRet <> BTRV_ERR_NONE Then
'            mCompareChfInPast = False
'            Exit Function
'        End If
'        If tmSchChf.sBillCycle = "C" Then
'            llStartDate = lmSvStartCal
'            llEndDate = lmSvEndCal
'        Else
'            llStartDate = lmSvStartStd
'            llEndDate = lmSvEndStd
'        End If
'        ReDim tmAlterFBSbf(0 To 0) As SBFLIST
'        ilUpper = 0
'        tmSbfSrchKey0.lChfCode = llAlteredChfCode
'        tmSbfSrchKey0.iDate(0) = 0
'        tmSbfSrchKey0.iDate(1) = 0
'        tmSbfSrchKey0.sTranType = " "
'        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = llAlteredChfCode)
'            If tmSbf.sTranType = "F" Then   'Items Billing
'                gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
'                If (llDate >= llStartDate) And (llDate <= llEndDate) Then
'                    tmAlterFBSbf(ilUpper).SbfRec = tmSbf
'                    tmAlterFBSbf(ilUpper).lRecPos = 0
'                    tmAlterFBSbf(ilUpper).iStatus = 1
'                    ilUpper = ilUpper + 1
'                    ReDim Preserve tmAlterFBSbf(0 To ilUpper) As SBFLIST
'                End If
'            End If
'            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'        ReDim tmSchFBSbf(0 To 0) As SBFLIST
'        ilUpper = 0
'        tmSbfSrchKey0.lChfCode = llSchChfCode
'        tmSbfSrchKey0.iDate(0) = 0
'        tmSbfSrchKey0.iDate(1) = 0
'        tmSbfSrchKey0.sTranType = " "
'        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = llSchChfCode)
'            If tmSbf.sTranType = "F" Then   'Fix Billing
'                gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
'                If (llDate >= llStartDate) And (llDate <= llEndDate) Then
'                    tmSchFBSbf(ilUpper).SbfRec = tmSbf
'                    tmSchFBSbf(ilUpper).lRecPos = 0
'                    tmSchFBSbf(ilUpper).iStatus = 1
'                    ilUpper = ilUpper + 1
'                    ReDim Preserve tmSchFBSbf(0 To ilUpper) As SBFLIST
'                End If
'            End If
'            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'        For ilAlter = 0 To UBound(tmAlterFBSbf) - 1 Step 1
'            ilMatch = False
'            For ilSch = 0 To UBound(tmSchFBSbf) - 1 Step 1
'                If (tmAlterFBSbf(ilAlter).SbfRec.iDate(0) = tmSchFBSbf(ilSch).SbfRec.iDate(0)) And (tmAlterFBSbf(ilAlter).SbfRec.iDate(1) = tmSchFBSbf(ilSch).SbfRec.iDate(1)) Then
'                    If (tmAlterFBSbf(ilAlter).SbfRec.iBillVefCode = tmSchFBSbf(ilSch).SbfRec.iBillVefCode) And (tmAlterFBSbf(ilAlter).SbfRec.iAirVefCode = tmSchFBSbf(ilSch).SbfRec.iAirVefCode) Then
'                        If (tmAlterFBSbf(ilAlter).SbfRec.lGross = tmSchFBSbf(ilSch).SbfRec.lGross) Then
'                            ilMatch = True
'                        End If
'                    End If
'                End If
'            Next ilSch
'            If Not ilMatch Then
'                mCompareChfInPast = False
'                Exit Function
'            End If
'        Next ilAlter
'        For ilSch = 0 To UBound(tmSchFBSbf) - 1 Step 1
'            ilMatch = False
'            For ilAlter = 0 To UBound(tmAlterFBSbf) - 1 Step 1
'                If (tmAlterFBSbf(ilAlter).SbfRec.iDate(0) = tmSchFBSbf(ilSch).SbfRec.iDate(0)) And (tmAlterFBSbf(ilAlter).SbfRec.iDate(1) = tmSchFBSbf(ilSch).SbfRec.iDate(1)) Then
'                    If (tmAlterFBSbf(ilAlter).SbfRec.iBillVefCode = tmSchFBSbf(ilSch).SbfRec.iBillVefCode) And (tmAlterFBSbf(ilAlter).SbfRec.iAirVefCode = tmSchFBSbf(ilSch).SbfRec.iAirVefCode) Then
'                        If (tmAlterFBSbf(ilAlter).SbfRec.iMnfItem = tmSchFBSbf(ilSch).SbfRec.iMnfItem) Then
'                            If (tmAlterFBSbf(ilAlter).SbfRec.iNoItems = tmSchFBSbf(ilSch).SbfRec.iNoItems) And (tmAlterFBSbf(ilAlter).SbfRec.lGross = tmSchFBSbf(ilSch).SbfRec.lGross) Then
'                                If (tmAlterFBSbf(ilAlter).SbfRec.sAgyComm = tmSchFBSbf(ilSch).SbfRec.sAgyComm) And (tmAlterFBSbf(ilAlter).SbfRec.iTrfCode = tmSchFBSbf(ilSch).SbfRec.iTrfCode) And (tmAlterFBSbf(ilAlter).SbfRec.lAcquisitionCost = tmSchFBSbf(ilSch).SbfRec.lAcquisitionCost) Then
'                                    ilMatch = True
'                                End If
'                            End If
'                        End If
'                    End If
'                End If
'            Next ilAlter
'            If Not ilMatch Then
'                mCompareChfInPast = False
'                Exit Function
'            End If
'        Next ilSch
        mCompareChfInPast = False
    ElseIf slType = "C" Then 'C=CPM
        mCompareChfInPast = False
    End If
    Exit Function
End Function

Public Function mCreateField1Key(tlChf As CHF) As String
    Dim ilVef As Integer
    Dim slGpSort As String
    Dim slVehSort As String
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slAgyAdvtName As String
    Dim slVefSort As String
    Dim ilRet As Integer
    Dim slSort As String
    Dim slKey As String
    Dim ilVefCode As Integer

    slVefSort = ""
    If (Asc(tgSpf.sUsingFeatures4) And INVSORTBYVEHICLE) = INVSORTBYVEHICLE Then
        ilVefCode = 0
        If tlChf.lVefCode > 0 Then
            ilVefCode = tlChf.lVefCode
        Else
            tmVsfSrchKey.lCode = -tlChf.lVefCode
            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            Do While ilRet = BTRV_ERR_NONE
                For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                    If tmVsf.iFSCode(ilLoop) > 0 Then
                        ilVefCode = tmVsf.iFSCode(ilLoop)
                        Exit Do
                    End If
                Next ilLoop
                If ilVefCode > 0 Then
                    Exit Do
                End If
                tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
                ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
        For ilVef = 0 To Invoice.lbcSortVehicle.ListCount - 1 Step 1
            If ilVefCode = Invoice.lbcSortVehicle.ItemData(ilVef) Then
                slStr = Trim$(str$(ilVef))
                Do While Len(slStr) < 5
                    slStr = "0" & slStr
                Loop
                ilRet = gBinarySearchVef(ilVefCode)
                If ilRet <> -1 Then
                    slVehSort = Trim$(str$(tgMVef(ilRet).iSort))
                    tmMnfSrchKey.iCode = tgMVef(ilRet).iOwnerMnfCode
                    If tmMnf.iCode <> tmMnfSrchKey.iCode Then
                        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            tmMnf.iGroupNo = 999
                        End If
                    End If
                    slGpSort = "000"    'Trim$(Str$(tmMnf.iGroupNo))
                    Do While Len(slGpSort) < 3
                        slGpSort = "0" & slGpSort
                    Loop
                    Do While Len(slVehSort) < 3
                        slVehSort = "0" & slVehSort
                    Loop
                Else
                    slGpSort = "999"
                    slVehSort = "999"
                End If
                slVefSort = slGpSort & ":" & slVehSort & ":" & slStr
                Exit For
            End If
        Next ilVef
    End If
    If tlChf.iAgfCode = 0 Then  'Obtain advertiser- direct bill
        ilLoop = gBinarySearchAdf(tlChf.iAdfCode)
        If ilLoop <> -1 Then
            If (tgCommAdf(ilLoop).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") Then
                slAgyAdvtName = Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID)
            Else
                slAgyAdvtName = Trim$(tgCommAdf(ilLoop).sName) & "/Direct"
            End If
            tmMnfSrchKey.iCode = tgCommAdf(ilLoop).iMnfSort
        End If
    Else    'Obtain agency
        ilLoop = gBinarySearchAgf(tlChf.iAgfCode)
        If ilLoop <> -1 Then
            slAgyAdvtName = Trim$(tgCommAgf(ilLoop).sName) & "/" & Trim$(tgCommAgf(ilLoop).sCityID)
            tmMnfSrchKey.iCode = tgCommAgf(ilLoop).iMnfSort
        End If
        If tmMnfSrchKey.iCode <= 0 Then
            ilLoop = gBinarySearchAdf(tlChf.iAdfCode)
            If ilLoop <> -1 Then
                tmMnfSrchKey.iCode = tgCommAdf(ilLoop).iMnfSort
            End If
        End If
    End If
    If (Asc(tgSpf.sUsingFeatures4) And INVSORTBYVEHICLE) = INVSORTBYVEHICLE Then
        For ilLoop = 0 To Invoice.lbcSortAgyAdvtDir.ListCount - 1 Step 1
            If StrComp(Trim$(Invoice.lbcSortAgyAdvtDir.List(ilLoop)), Trim$(slAgyAdvtName), vbTextCompare) = 0 Then
                slAgyAdvtName = Trim$(str$(ilLoop))
                Exit For
            End If
        Next ilLoop
    End If
    If tmMnfSrchKey.iCode > 0 Then
        If tmMnf.iCode <> tmMnfSrchKey.iCode Then
            ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Else
            ilRet = BTRV_ERR_NONE
        End If
        If ilRet <> BTRV_ERR_NONE Then
            tmMnf.iCode = 0
            tmMnf.iGroupNo = 0
            'Sort PSA and Promos to end of list
            If (tlChf.sType = "S") Or (tlChf.sType = "M") Then
                slSort = "99999"
            Else
                slSort = "99998"
            End If
        Else
            If tmMnf.iGroupNo > 0 Then
                slSort = Trim$(str$(tmMnf.iGroupNo))
                Do While Len(slSort) < 5
                    slSort = "0" & slSort
                Loop
            Else
                'Sort PSA and Promos to end of list
                If (tlChf.sType = "S") Or (tlChf.sType = "M") Then
                    slSort = "99999"
                Else
                    slSort = "99998"
                End If
            End If
        End If
    Else
        'Sort PSA and Promos to end of list
        If (tlChf.sType = "S") Or (tlChf.sType = "M") Then
            slSort = "99999"
        Else
            slSort = "99998"
        End If
    End If
    If slVefSort <> "" Then
        slKey = slVefSort & ":" & slSort & ":" & slAgyAdvtName
    Else
        slKey = slSort & slAgyAdvtName
    End If
    Do While Len(slKey) < 51
        slKey = slKey & " "
    Loop
    If Len(slKey) > 51 Then
        slKey = Left$(slKey, 51)
    End If
    mCreateField1Key = slKey
End Function

'********************************************************
'*                                                      *
'*      Procedure Name:mCPM_GenIvr                      *
'*                                                      *
'*             Created:6/13/93       By:D. LeVine       *
'*            Modified:              By:                *
'*                                                      *
'*            Comments:Generate IVR records for AdServer*
'*                                                      *
'********************************************************
Sub mCPM_GenIvr(llInvoiceNo As Long)
    Dim ilCpm As Integer
    Dim ilImpactPrt As Integer
    Dim ilPass As Integer
    Dim ilFound As Integer
    Dim slPctTrade As String
    Dim ilVeh As Integer
    Dim slForm2Key As String
    Dim ilONoSpots As Integer
    Dim ilANoSpots As Integer
    Dim llOGross As Long
    Dim llAGross As Long
    Dim llANet As Long
    Dim slTGross As String
    Dim slTNet As String
    Dim llTImpressions As Long
    Dim slProduct As String
    Dim ilRvf As Integer
    Dim ilSRvf As Integer
    Dim slAgyRate As String
    Dim slGross As String
    Dim slNet As String
    Dim slAmount As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llChfCode As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim llInvNo As Long
    Dim llCPMDate As Long
    Dim llTax1 As Long
    Dim llTax2 As Long
    Dim slSDate As String
    Dim slEDate As String
    Dim ilIncInvNo As Integer
    Dim llReprintInvNo As Long
    Dim ilNowTime(0 To 1) As Integer        '10-31-01
    Dim ilTax1 As Integer                   '1-24-05
    Dim ilTax2 As Integer                   '1-24-05
    Dim llTax1Rate As Long
    Dim llTax2Rate As Long
    Dim slGrossNet As String
    Dim llAirCPMInvNo As Long
    Dim llIvrCode As Long
    Dim ilTest As Integer
    Dim ilEDIFlag As Integer
    Dim ilAirNTRCombineIndex As Integer
    Dim tlIvr As IVR
    Dim ilRdf As Integer
    Dim llUB As Long
    Dim ilPkg As Integer
    Dim ilHidden As Integer
    Dim slPcfStartDate As String 'TTP 10827
    Dim slPCFEndDate As String
    Dim llRet As Long
    Dim llPkgImpressions As Long
    Dim llPkgCost As Long
    Dim slRvfGross As String
    Dim slSQLQuery As String
    Dim mnf_rst As ADODB.Recordset
    Dim raf_rst As ADODB.Recordset
    Dim rvf_rst As ADODB.Recordset
    Dim phf_rst As ADODB.Recordset

    imIbfRecLen = Len(tmIbf)
    imPcfRecLen = Len(tmPcf)
    
    Invoice.smRemoteGenDate = Invoice.smGenDate
    Invoice.lmRemoteGenDate = gDateValue(Invoice.smRemoteGenDate)
    Invoice.smRemoteGenTime = Invoice.smGenTime
    Invoice.lmRemoteGenTime = gTimeToLong(Invoice.smRemoteGenTime, False)
    gPackDate Invoice.smGenDate, tmIvr.iGenDate(0), tmIvr.iGenDate(1)
    gPackTime Invoice.smGenTime, ilNowTime(0), ilNowTime(1)
    gUnpackTimeLong ilNowTime(0), ilNowTime(1), False, tmIvr.lGenTime

    mCheckForOverDelivery
    
    ilIncInvNo = False
    ilEDIFlag = False
    imArfPDFEMailCode = -1

    tmIvr.lDisclaimer = 0
    ilImpactPrt = False
    tmIvr.lSpotKeyNo = 0
    
    ilCpm = LBound(tgCPMCntr)
    'Compute package impressions
    For ilPkg = 0 To UBound(tgCPMCntr) - 1 Step 1
        If tgCPMCntr(ilPkg).tPcf.sType = "P" Then
            llPkgImpressions = 0
            llPkgCost = 0
            For ilHidden = 0 To UBound(tgCPMCntr) - 1 Step 1
                If tgCPMCntr(ilHidden).tPcf.sType = "H" Then
                    If (tgCPMCntr(ilPkg).tPcf.iPodCPMID = tgCPMCntr(ilHidden).tPcf.iPkCPMID) And (tgCPMCntr(ilPkg).tPcf.lChfCode = tgCPMCntr(ilHidden).tPcf.lChfCode) Then
                        llPkgImpressions = llPkgImpressions + tgCPMCntr(ilHidden).tIbf.lBilledImpression
                        llPkgCost = llPkgCost + tgCPMCntr(ilHidden).tIbf.lBilledImpression * tgCPMCntr(ilHidden).tPcf.lPodCPM / 1000
                    End If
                End If
            Next ilHidden
            tgCPMCntr(ilPkg).tIbf.lBilledImpression = llPkgImpressions
            If tgCPMCntr(ilPkg).tPcf.sPriceType <> "F" Then
                tgCPMCntr(ilPkg).tPcf.lTotalCost = llPkgCost
            Else
                tmChfSrchKey.lCode = llChfCode
                ilRet = btrGetEqual(hmCHF, tgChfInv, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    If (Not Invoice.rbcType(INVGEN_Reprint).Value) And (Not Invoice.rbcType(INVGEN_Archive).Value) Then        '11-15-16 reprint or archive
                        slSQLQuery = "Select sum(rvfGross) as TotalGross from RVF_Receivables where rvfCntrNo = " & tgChfInv.lCntrNo & " And rvfPkLineNo = " & tgCPMCntr(ilPkg).tPcf.iPodCPMID & " And (rvfTranType = 'IN' or rvfTranType = 'HI')"
                        Set rvf_rst = gSQLSelectCall(slSQLQuery)
                        If Not rvf_rst.EOF And Not IsNull(rvf_rst!TotalGross) Then
                            slRvfGross = rvf_rst!TotalGross
                        Else
                            slRvfGross = "0.00"
                        End If
                        slSQLQuery = "Select sum(phfGross) as TotalGross from phf_Payment_History where phfCntrNo = " & tgChfInv.lCntrNo & " And phfPkLineNo = " & tgCPMCntr(ilPkg).tPcf.iPodCPMID & " And (phfTranType = 'IN' or phfTranType = 'HI')"
                        Set phf_rst = gSQLSelectCall(slSQLQuery)
                        If Not phf_rst.EOF And Not IsNull(phf_rst!TotalGross) Then
                            slRvfGross = gAddStr(phf_rst!TotalGross, slRvfGross)
                        End If
                        slRvfGross = gSubStr(gLongToStrDec(tgCPMCntr(ilPkg).tPcf.lTotalCost, 2), slRvfGross)
                        gUnpackDate tgCPMCntr(ilPkg).tPcf.iEndDate(0), tgCPMCntr(ilPkg).tPcf.iEndDate(1), slPCFEndDate
                        tgCPMCntr(ilPkg).tPcf.lTotalCost = gStrDecToLong(slRvfGross, 2) / mDetermineRemainingMonths(slPCFEndDate, tgChfInv.sBillCycle)
                    Else
                        slSQLQuery = "Select sum(rvfGross) as TotalGross from RVF_Receivables where rvfCntrNo = " & tgChfInv.lCntrNo & " And rvfPkLineNo = " & tgCPMCntr(ilPkg).tPcf.iPodCPMID & " And rvfPcfCode <> " & 0 & " And rvfInvNo = " & tgCPMCntr(ilPkg).lInvNo & " And (rvfTranType = 'IN' or rvfTranType = 'HI')"
                        Set rvf_rst = gSQLSelectCall(slSQLQuery)
                        If Not rvf_rst.EOF And Not IsNull(rvf_rst!TotalGross) Then
                            slRvfGross = rvf_rst!TotalGross
                        Else
                            slRvfGross = "0.00"
                        End If
                        slSQLQuery = "Select sum(phfGross) as TotalGross from phf_Payment_History where phfCntrNo = " & tgChfInv.lCntrNo & " And phfPkLineNo = " & tgCPMCntr(ilPkg).tPcf.iPodCPMID & " And phfPcfCode <> " & 0 & " And phfInvNo = " & tgCPMCntr(ilPkg).lInvNo & " And (phfTranType = 'IN' or phfTranType = 'HI')"
                        Set phf_rst = gSQLSelectCall(slSQLQuery)
                        If Not phf_rst.EOF And Not IsNull(phf_rst!TotalGross) Then
                            slRvfGross = gAddStr(phf_rst!TotalGross, slRvfGross)
                        End If
                        If InStr(1, slRvfGross, ".") = 0 Then
                            slGross = slRvfGross & ".00"
                        Else
                            slGross = slRvfGross
                        End If
                        tgCPMCntr(ilPkg).tPcf.lTotalCost = gStrDecToLong(slGross, 2)
                    End If
                End If
            End If
            tgCPMCntr(ilPkg).tIbf.iVefCode = tgCPMCntr(ilPkg).tPcf.iVefCode
        End If
    Next ilPkg
    
    Do While ilCpm <= UBound(tgCPMCntr) - 1
        tmPcf = tgCPMCntr(ilCpm).tPcf
        tmIbf = tgCPMCntr(ilCpm).tIbf
        If tmIbf.lCode > 0 Then
            llChfCode = tmPcf.lChfCode
            llReprintInvNo = tgCPMCntr(ilCpm).lInvNo
            llCPMDate = tgCPMCntr(ilCpm).lEDate
            tmChfSrchKey.lCode = llChfCode
            ilRet = btrGetEqual(hmCHF, tgChfInv, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                'See if the dates can be avoided- if not then set to earliest and latest date of sbf
                If tgChfInv.sBillCycle = "C" Then
                    gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slSDate
                    slSDate = gIncOneDay(slSDate)
                    slEDate = gObtainEndStd(slSDate)
                    llSDate = gDateValue(slSDate)
                    llEDate = gDateValue(slEDate)
                    If llCPMDate < llSDate Then
                        llCPMDate = llSDate
                        If Invoice.rbcType(INVGEN_Final).Value Then    'Final Invoice
                            Do
                                tmIbfSrchKey0.lCode = tmIbf.lCode
                                ilRet = btrGetEqual(hmIbf, tmIbf, imIbfRecLen, tmIbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                If ilRet = BTRV_ERR_NONE Then
                                    tmIbf.iBillMonth = tmIbf.iBillMonth + 1
                                    If tmIbf.iBillMonth > 12 Then
                                        tmIbf.iBillMonth = 1
                                        tmIbf.iBillYear = tmIbf.iBillYear + 1
                                    End If
                                    ilRet = btrUpdate(hmIbf, tmIbf, imIbfRecLen)
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            If ilRet <> BTRV_ERR_NONE Then
                                If ilRet >= 30000 Then
                                    ilRet = csiHandleValue(0, 7)
                                End If
                                Print #hmMsg, "mCPM_GenIvr: btrUpdate-Point 1, Sbf Error # " & Trim$(str$(ilRet))
                            End If
                        End If
                    End If
                    smEndCal = Format$(llCPMDate, "m/d/yy")
                    lmEndCal = llCPMDate
                    If Invoice.rbcType(INVGEN_Reprint).Value Or Invoice.rbcType(INVGEN_Aff).Value Or Invoice.rbcType(INVGEN_Undo).Value Or Invoice.rbcType(INVGEN_Archive).Value Then     '11-15-16 reprint, aff, undo or archive Determine invoice number
                        lmEndCal = tgCPMCntr(ilCpm).lInvDate
                        smEndCal = Format$(lmEndCal, "m/d/yy")
                        llCPMDate = lmEndCal
                    End If
                ElseIf tgChfInv.sBillCycle = "W" Then
                    gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slSDate
                    slSDate = gIncOneDay(slSDate)
                    slEDate = gObtainNextSunday(slSDate)
                    llSDate = gDateValue(slSDate)
                    llEDate = gDateValue(slEDate)
                    smEndWk = Format$(llCPMDate, "m/d/yy")
                    lmEndWk = llCPMDate
                    If Invoice.rbcType(INVGEN_Reprint).Value Or Invoice.rbcType(INVGEN_Aff).Value Or Invoice.rbcType(INVGEN_Undo).Value Or Invoice.rbcType(INVGEN_Archive).Value Then     '11-15-16 reprint, aff, undo or archive Determine invoice number
                        lmEndWk = tgCPMCntr(ilCpm).lInvDate
                        smEndWk = Format$(lmEndWk, "m/d/yy")
                        llCPMDate = lmEndWk
                    End If
                Else
                    gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slSDate
                    slSDate = gIncOneDay(slSDate)
                    slEDate = gObtainEndStd(slSDate)
                    llSDate = gDateValue(slSDate)
                    llEDate = gDateValue(slEDate)
                    If llCPMDate < llSDate Then
                        llCPMDate = llSDate
                        If Invoice.rbcType(INVGEN_Final).Value Then    'Final Invoice
                            Do
                                tmIbfSrchKey0.lCode = tmIbf.lCode
                                ilRet = btrGetEqual(hmIbf, tmIbf, imIbfRecLen, tmIbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                                If ilRet = BTRV_ERR_NONE Then
                                    tmIbf.iBillMonth = tmIbf.iBillMonth + 1
                                    If tmIbf.iBillMonth > 12 Then
                                        tmIbf.iBillMonth = 1
                                        tmIbf.iBillYear = tmIbf.iBillYear + 1
                                    End If
                                    ilRet = btrUpdate(hmIbf, tmIbf, imIbfRecLen)
                                End If
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            If ilRet <> BTRV_ERR_NONE Then
                                If ilRet >= 30000 Then
                                    ilRet = csiHandleValue(0, 7)
                                End If
                                Print #hmMsg, "mCPM_GenIvr: btrUpdate-Point 1, Sbf Error # " & Trim$(str$(ilRet))
                            End If
                        End If
                    End If
                    
                    smEndStd = Format$(llCPMDate, "m/d/yy")
                    lmEndStd = llCPMDate
                    If Invoice.rbcType(INVGEN_Reprint).Value Or Invoice.rbcType(INVGEN_Aff).Value Or Invoice.rbcType(INVGEN_Undo).Value Or Invoice.rbcType(INVGEN_Archive).Value Then     '11-15-16 reprint,aff, undo, archuve Determine invoice number
                        lmEndStd = tgCPMCntr(ilCpm).lInvDate
                        smEndStd = Format$(lmEndStd, "m/d/yy")
                        llCPMDate = lmEndStd
                    End If
                End If
                tgSortKey(UBound(tgSortKey)).lCntrNo = tgChfInv.lCntrNo
                tgSortKey(UBound(tgSortKey)).lChfCode = tgChfInv.lCode
                'Use todays date so that all sbf records show on same invoice
                tgSortKey(UBound(tgSortKey)).lCPMDate = llCPMDate
                tgSortKey(UBound(tgSortKey)).iType = 11
                tgSortKey(UBound(tgSortKey)).iBilled = False
                If tgCPMCntr(ilCpm).tPcf.sType <> "P" Then
                    ReDim Preserve tgSortKey(0 To UBound(tgSortKey) + 1) As TYPESORTKEY
                End If
                slPctTrade = gIntToStrDec(tgChfInv.iPctTrade, 0)
                slProduct = Trim$(tgChfInv.sProduct)
                For ilPass = 1 To 2 Step 1
                    ilFound = False
                    If (ilPass = 1) And (tgChfInv.iPctTrade <> 100) Then
                        ilFound = True
                    End If
                    If (ilPass = 2) And (tgChfInv.iPctTrade <> 0) Then
                        ilFound = True
                    End If
                    If ilFound Then
                        'Determine Aired Gross, Ordered number of spots, Aired number of Spots
                        llOGross = 0
                        llAGross = 0
                        llANet = 0
                        ilONoSpots = 0
                        ilANoSpots = 0
                        llChfCode = tgChfInv.lCode
                        ilFound = True
                        ilIncInvNo = False
                        ilEDIFlag = False
                        imArfPDFEMailCode = -1
                        ilAirNTRCombineIndex = -1
                        If Invoice.rbcType(INVGEN_Reprint).Value Or Invoice.rbcType(INVGEN_Aff).Value Or Invoice.rbcType(INVGEN_Undo).Value Or Invoice.rbcType(INVGEN_Archive).Value Then      '11-15-16 repirnt, aff, undo, archive Determine invoice number
                            ilFound = True
                            llInvoiceNo = llReprintInvNo
                            llAirCPMInvNo = llReprintInvNo
                            llIvrCode = 0
                            'If imCombineAirAndNTR Then
                                For ilTest = LBound(tmAirNTRCombine) To UBound(tmAirNTRCombine) - 1 Step 1
                                    If tmAirNTRCombine(ilTest).lChfCode = tmPcf.lChfCode Then
                                        If ((ilPass = 1) And (tmAirNTRCombine(ilTest).sCashTrade = "C")) Or ((ilPass = 2) And (tmAirNTRCombine(ilTest).sCashTrade = "T")) Then
                                            llIvrCode = tmAirNTRCombine(ilTest).lIvrCode
                                            ilEDIFlag = tmAirNTRCombine(ilTest).iEDIFlag
                                            imArfPDFEMailCode = tmAirNTRCombine(ilTest).iArfPDFEMailCode
                                            ilAirNTRCombineIndex = ilTest
                                            Exit For
                                        End If
                                    End If
                                Next ilTest
                            'End If
                        Else
                            'Determine invoice number
                            llAirCPMInvNo = -1
                            llIvrCode = 0
                            For ilTest = LBound(tmAirNTRCombine) To UBound(tmAirNTRCombine) - 1 Step 1
                                If tmAirNTRCombine(ilTest).lChfCode = tmPcf.lChfCode Then
                                    If ((ilPass = 1) And (tmAirNTRCombine(ilTest).sCashTrade = "C")) Or ((ilPass = 2) And (tmAirNTRCombine(ilTest).sCashTrade = "T")) Then
                                        llAirCPMInvNo = tmAirNTRCombine(ilTest).lInvNo
                                        llIvrCode = tmAirNTRCombine(ilTest).lIvrCode
                                        ilEDIFlag = tmAirNTRCombine(ilTest).iEDIFlag
                                        imArfPDFEMailCode = tmAirNTRCombine(ilTest).iArfPDFEMailCode
                                        ilAirNTRCombineIndex = ilTest
                                        Exit For
                                    End If
                                End If
                            Next ilTest
                            
                            If ilPass = 1 Then
                                ilFound = Not Invoice.mRvfExist(llInvNo, 1, tmPcf.lCode, "C")
                            Else
                                ilFound = Not Invoice.mRvfExist(llInvNo, 1, tmPcf.lCode, "T")
                            End If
                            If ilFound Then
                                If llAirCPMInvNo = -1 Then
                                    ilIncInvNo = True
                                    llAirCPMInvNo = llInvoiceNo
                                End If
                            End If
                        End If
    
                        If ilFound Then
                            'TTP 10706 - Digital only reprint invoice: contract header invoice comment not shown if comment added after invoicing, then invoice is reprinted
                            'tmChf = tgChfInv
                            tmChfSrchKey1.lCntrNo = tgChfInv.lCntrNo
                            tmChfSrchKey1.iCntRevNo = 32000
                            tmChfSrchKey1.iPropVer = 32000
                            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            '-------------------------------------------
                            'Not F=Fully scheduled
                            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tgChfInv.lCntrNo) And (tmChf.sSchStatus <> "F")
                                ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            Loop
                            Invoice.mGetComment
                            tmChf = tgChfInv
                            'End of 10706
                            Invoice.mGetAdfAgfSlf
                            llUB = UBound(tmAirNTRCombine)
                            'TTP 10515 - NTR Invoices - "NTR INVOICE and AFFIDAVIT" not displaying correct on NTR Invoices it shows "INVOICES and AFFIDAVIT" started with V81; Add Contract # so that the contract can be checked for AirTime/CPM & NTR to determine if Header should show NTR
                            Invoice.mGenHeader ilImpactPrt, llAirCPMInvNo, ilPass, slPctTrade, 1, True, ilEDIFlag, tmChf.lCntrNo
                            If llUB <> UBound(tmAirNTRCombine) Then
                                For ilTest = LBound(tmAirNTRCombine) To UBound(tmAirNTRCombine) - 1 Step 1
                                    If tmAirNTRCombine(ilTest).lChfCode = tmPcf.lChfCode Then
                                        If ((ilPass = 1) And (tmAirNTRCombine(ilTest).sCashTrade = "C")) Or ((ilPass = 2) And (tmAirNTRCombine(ilTest).sCashTrade = "T")) Then
                                            llIvrCode = tmAirNTRCombine(ilTest).lIvrCode
                                            ilEDIFlag = tmAirNTRCombine(ilTest).iEDIFlag
                                            imArfPDFEMailCode = tmAirNTRCombine(ilTest).iArfPDFEMailCode
                                            ilAirNTRCombineIndex = ilTest
                                            Exit For
                                        End If
                                    End If
                                Next ilTest
                            End If
                            slForm2Key = Invoice.mRepInv_BuildKey()
                            slTNet = ".00"
                            slTGross = ".00"
                            llTImpressions = 0
                            llTax1 = 0
                            llTax2 = 0
                            ilTax1 = False      '1-24-05 assume no taxes apply
                            ilTax2 = False
                            For ilLoop = ilCpm To UBound(tgCPMCntr) - 1 Step 1
                                tmPcf = tgCPMCntr(ilLoop).tPcf
                                tmIbf = tgCPMCntr(ilLoop).tIbf
                                If (tmIbf.lCode > 0) Or (tmPcf.sType = "P") Then
                                    If Invoice.rbcType(INVGEN_Reprint).Value Or Invoice.rbcType(INVGEN_Aff).Value Or Invoice.rbcType(INVGEN_Undo).Value Or Invoice.rbcType(INVGEN_Archive).Value Then      '11-15-16 reprint, aff, undo, archive Determine invoice number
                                        If (llChfCode <> tgCPMCntr(ilLoop).tPcf.lChfCode) Or (llReprintInvNo <> tgCPMCntr(ilLoop).lInvNo) Then
                                            Exit For
                                        End If
                                    Else
                                        If (llChfCode <> tgCPMCntr(ilLoop).tPcf.lChfCode) Then
                                            Exit For
                                        End If
                                    End If
                                    If (tmChf.iAgfCode > 0) And (ilPass = 1) Then   '(Val(slPctTrade) <> 100) Then
                                        slAgyRate = gIntToStrDec(tmAgf.iComm, 2)
                                        tmIvr.iPctComm = tmAgf.iComm
                                    ElseIf (tmChf.iAgfCode > 0) And (ilPass = 2) Then     '(Val(slPctTrade) <> 100) Then
                                        slAgyRate = gIntToStrDec(tmAgf.iComm, 2)
                                        tmIvr.iPctComm = tmAgf.iComm
                                    Else
                                        slAgyRate = ""
                                        tmIvr.iPctComm = 0
                                    End If
                                    tmIvr.iType = IVRTYPE_AdServer 'CPM Detail
                                    tmIvr.lChfCode = tmPcf.lChfCode
                                    tmIvr.iLineNo = tmPcf.iPodCPMID
                                    tmIvr.sADayDate = ""
                                    tmIvr.iPrgEnfCode = 0
                                    tmIvr.sOVehName = "Vehicle Name Missing"
                                    tmIvr.sAVehName = "Vehicle Name Missing"
                                    ilVeh = gBinarySearchVef(tmPcf.iVefCode)
                                    If ilVeh <> -1 Then
                                        tmIvr.sOVehName = tgMVef(ilVeh).sName
                                    End If
                                    ilVeh = gBinarySearchVef(tmIbf.iVefCode)
                                    If ilVeh <> -1 Then
                                        tmIvr.sAVehName = tgMVef(ilVeh).sName
                                    End If
                                    tmIvr.iRdfCode = tmPcf.iRdfCode
                                    tmIvr.sODPName = "Daypart Name Missing"
                                    ilRdf = gBinarySearchRdf(tmPcf.iRdfCode)
                                    If ilRdf <> -1 Then
                                        tmIvr.sODPName = tgMRdf(ilRdf).sName
                                    End If
                                    'Description will be obtained directly from the sbf record
                                    tmIvr.lSpotKeyNo = tmPcf.iPodCPMID
                                    If tmPcf.sPriceType <> "F" Then
                                        'Not Flat Rate Digital
                                        If tmPcf.sType = "P" Then
                                            tmIvr.sARate = gLongToStrDec((tmPcf.lTotalCost * 1000) / tmIbf.lBilledImpression, 2)
                                            'Amount/item
                                            slGross = gLongToStrDec(tmPcf.lTotalCost, 2)
                                        Else
                                            tmIvr.sARate = gLongToStrDec(tmPcf.lPodCPM, 2)
                                            'Amount/item
                                            'TTP 10514 - Invoice: overflow error when impressions too high
                                            'slGross = gLongToStrDec(tmIbf.lBilledImpression * tmPcf.lPodCPM / 1000, 2)
                                            'TTP 10644 - Invoices: digital line - a CPM of $0 results in incorrect invoice amount: Reset slGross
                                            slGross = gLongToStrDec(tmIbf.lBilledImpression * (tmPcf.lPodCPM / 1000), 2)
                                        End If
                                    Else
                                        'Flat Rate Digital
                                        tmIvr.sARate = ""
                                        If tmPcf.sType = "P" Then
                                            slGross = gLongToStrDec(tmPcf.lTotalCost, 2)
                                        Else
                                            'Not a package
                                            slRvfGross = "0.00"
                                            If (Not Invoice.rbcType(INVGEN_Reprint).Value) And (Not Invoice.rbcType(INVGEN_Archive).Value) Then        '11-15-16 reprint or archive
                                                'slSQLQuery = "Select Sum(rvfGross) as TotalGross from RVF_Receivables Left Outer Join pcf_Pod_CPM_Cntr On rvfPcfCode = pcfCode where rvfCntrNo = " & tgChfInv.lCntrNo & " And rvfPcfCode <> " & 0 & " And pcfPodCPMID = " & tmPcf.iPodCPMID & " And (rvfTranType = 'IN' or rvfTranType = 'AN'  or rvfTranType = 'HI')"
                                                slSQLQuery = "Select Sum(rvfGross) as TotalGross from RVF_Receivables Left Outer Join pcf_Pod_CPM_Cntr On rvfPcfCode = pcfCode where rvfCntrNo = " & tgChfInv.lCntrNo & " And rvfPcfCode <> " & 0 & " And pcfPodCPMID = " & tmPcf.iPodCPMID & " And (rvfTranType = 'IN' or rvfTranType = 'HI')"
                                                Set rvf_rst = gSQLSelectCall(slSQLQuery)
                                                If Not rvf_rst.EOF And Not IsNull(rvf_rst!TotalGross) Then
                                                    slRvfGross = rvf_rst!TotalGross
                                                Else
                                                    slRvfGross = "0.00"
                                                End If
                                                slSQLQuery = "Select Sum(phfGross) as TotalGross from phf_Payment_History Left Outer Join pcf_Pod_CPM_Cntr On phfPcfCode = pcfCode where phfCntrNo = " & tgChfInv.lCntrNo & " And phfPcfCode <> " & 0 & " And pcfPodCPMID = " & tmPcf.iPodCPMID & " And (phfTranType = 'IN' or phfTranType = 'HI')"
                                                Set phf_rst = gSQLSelectCall(slSQLQuery)
                                                If Not phf_rst.EOF And Not IsNull(phf_rst!TotalGross) Then
                                                    slRvfGross = gAddStr(phf_rst!TotalGross, slRvfGross)
                                                End If
                                                slRvfGross = gSubStr(gLongToStrDec(tmPcf.lTotalCost, 2), slRvfGross)  'Amount remaining
                                                
                                                gUnpackDate tmPcf.iStartDate(0), tmPcf.iStartDate(1), slPcfStartDate 'TTP 10827
                                                gUnpackDate tmPcf.iEndDate(0), tmPcf.iEndDate(1), slPCFEndDate
                                                
                                                'TTP 10827 - Boostr Phase 2: change flat rate invoice method
                                                If tgSpfx.iLineCostType = 1 And tmPcf.sPriceType = "F" Then
                                                    slGross = gLongToStrDec(mDeterminePeriodAmountByDaily(slPcfStartDate, slPCFEndDate, tmChf.sBillCycle, slRvfGross) * 100, 2)
                                                Else
                                                    slGross = gDivStr(slRvfGross, str(mDetermineRemainingMonths(slPCFEndDate, tmChf.sBillCycle)))
                                                End If
                                            Else
                                                'Reprint or Archive
                                                slSQLQuery = "Select Sum(rvfGross) as TotalGross from RVF_Receivables Left Outer Join pcf_Pod_CPM_Cntr On rvfPcfCode = pcfCode where rvfCntrNo = " & tgChfInv.lCntrNo & " And rvfPcfCode <> " & 0 & " And pcfPodCPMID = " & tmPcf.iPodCPMID & " And rvfInvNo = " & tgCPMCntr(ilCpm).lInvNo & " And  (rvfTranType = 'IN' or rvfTranType = 'HI')"
                                                Set rvf_rst = gSQLSelectCall(slSQLQuery)
                                                If Not rvf_rst.EOF And Not IsNull(rvf_rst!TotalGross) Then
                                                    slRvfGross = rvf_rst!TotalGross
                                                Else
                                                    slRvfGross = "0.00"
                                                End If
                                                slSQLQuery = "Select Sum(phfGross) as TotalGross from phf_Payment_History Left Outer Join pcf_Pod_CPM_Cntr On phfPcfCode = pcfCode where phfCntrNo = " & tgChfInv.lCntrNo & " And phfPcfCode <> " & 0 & " And pcfPodCPMID = " & tmPcf.iPodCPMID & " And phfInvNo = " & tgCPMCntr(ilCpm).lInvNo & "  And (phfTranType = 'IN' or phfTranType = 'HI')"
                                                Set phf_rst = gSQLSelectCall(slSQLQuery)
                                                If Not phf_rst.EOF And Not IsNull(phf_rst!TotalGross) Then
                                                    slRvfGross = gAddStr(phf_rst!TotalGross, slRvfGross)
                                                End If
                                                If InStr(1, slRvfGross, ".") = 0 Then
                                                    slGross = slRvfGross & ".00"
                                                Else
                                                    slGross = slRvfGross
                                                End If
                                            End If
                                        End If
                                    End If
                                    If ilPass = 1 Then
                                        If Val(slPctTrade) <> 100 Then
                                            slGross = gDivStr(gMulStr(slGross, gSubStr("100", slPctTrade)), "100")
                                        End If
                                        tmIvr.lATotalGross = gStrDecToLong(slGross, 2)
                                        If slAgyRate = "" Then
                                            slStr = slGross
                                        Else
                                            slStr = gDivStr(gMulStr(slGross, gSubStr("100.00", slAgyRate)), "100.00")
                                        End If
                                    ElseIf ilPass = 2 Then
                                        If Val(slPctTrade) <> 100 Then
                                            slGross = gSubStr(slGross, gDivStr(gMulStr(slGross, gSubStr("100", slPctTrade)), "100"))
                                        End If
                                        tmIvr.lATotalGross = gStrDecToLong(slGross, 2)
                                        If slAgyRate = "" Then
                                            slStr = slGross
                                        Else
                                            slStr = gDivStr(gMulStr(slGross, gSubStr("100.00", slAgyRate)), "100.00")
                                        End If
                                    End If
                                    slNet = slStr
                                    If tmPcf.sType <> "P" Then
                                        slTNet = gAddStr(slTNet, slStr)
                                    End If
                                    tmIvr.lATotalNet = gStrDecToLong(slStr, 2)
                                    tgCPMCntr(ilLoop).lCPMNet = gStrDecToLong(slStr, 2)
                                    If tmPcf.sType <> "P" Then
                                        slTGross = gAddStr(slTGross, slGross)
                                    End If
                                    tgCPMCntr(ilLoop).lCPMGross = gStrDecToLong(slGross, 2)
                                    
                                    'Price type
                                    '09/27/2022 - JW - Invoice: show Baked-in Flat Rate lines as "Baked-in" Price Type - pcf_pod_cpm_cntr/pcfDeliveryType
                                    If tmPcf.iDeliveryType = 1 Then
                                        tmIvr.sACopy(0) = "Baked-in"
                                    Else
                                        If tmPcf.sPriceType = "F" Then
                                            tmIvr.sACopy(0) = "Flat Rate"
                                        Else
                                            tmIvr.sACopy(0) = "CPM"
                                        End If
                                    End If
                                    'Units/Item
                                    tmIvr.sRRemark = ""
                                    If tmPcf.sType = "P" Then
                                        tmIvr.lOTotalSpots = tmIbf.lBilledImpression
                                    Else
                                        tmIvr.lOTotalSpots = tmPcf.lImpressionGoal
                                    End If
                                    tmIvr.lOTotalGross = tmIbf.lCode
                                    '09/27/2022 - JW - Invoice:  In addition, on the invoice, the Impression value will be suppressed for Baked-in lines.
                                    If tmPcf.iDeliveryType = 1 Then
                                        tmIvr.lATotalSpots = 0
                                    Else
                                        tmIvr.lATotalSpots = tmIbf.lBilledImpression
                                        If tmPcf.sType <> "P" Then
                                            llTImpressions = llTImpressions + tmIbf.lBilledImpression
                                        End If
                                    End If
                                    tmIvr.lRTotalGross = tmPcf.lRafCode
                                    tmIvr.sACopy(1) = ""
                                    'TTP 10920 - Invoice: "ad server targeting" field is limited to 45 characters but the targeting field can be much longer
                                    tmIvr.sACopy(2) = ""
                                    If tmPcf.lRafCode > 0 Then
                                        slSQLQuery = "Select rafName From raf_Region_Area Where rafCode = " & tmPcf.lRafCode
                                        Set raf_rst = gSQLSelectCall(slSQLQuery)
                                        If Not raf_rst.EOF Then
                                            tmIvr.sACopy(1) = Mid(Trim$(raf_rst!rafName), 1, 45)
                                            'TTP 10920 - Invoice: "ad server targeting" field is limited to 45 characters but the targeting field can be much longer
                                            tmIvr.sACopy(2) = Mid(Trim$(raf_rst!rafName), 46)
                                        End If
                                    End If
                                    tmIvr.sKey = slForm2Key & "4" & tmIvr.sOVehName
                                        
                                    If tgChfInv.sBillCycle = "C" Then
                                        gPackDateLong lmSvStartCal, tmIvr.iInvStartDate(0), tmIvr.iInvStartDate(1)
                                        gPackDateLong lmSvEndCal, tmIvr.iInvDate(0), tmIvr.iInvDate(1)
                                    ElseIf tgChfInv.sBillCycle = "W" Then
                                        gPackDateLong lmSvStartWk, tmIvr.iInvStartDate(0), tmIvr.iInvStartDate(1)
                                        gPackDateLong lmSvEndWk, tmIvr.iInvDate(0), tmIvr.iInvDate(1)
                                    Else
                                        gPackDateLong lmSvStartStd, tmIvr.iInvStartDate(0), tmIvr.iInvStartDate(1)
                                        gPackDateLong lmSvEndStd, tmIvr.iInvDate(0), tmIvr.iInvDate(1)
                                    End If
                                    
                                    tmIvr.iFormType = 1  'Force to Air
                                    tmIvr.lCode = 0
                                    tmIvr.lClfCxfCode = tmPcf.lCxfCode
                                    'Boostr Phase 2 issues - for Joel - Issue 15, use the latest PCF Code. It should show the ad length from the line. This way the updated ad lengths for the lines that didn't originally have ad lengths will show ad lengths.
                                    'tmIvr.lClfCode = tmPcf.lCode
                                    tmIvr.lClfCode = mGetLatestPCFCode(tgChfInv.lCntrNo, tmPcf.iPodCPMID)
                                    If tmIvr.lClfCode = 0 Then tmIvr.lClfCode = tmPcf.lCode
                                    mGetInstallmentCounts tmChf ', tmIvr.iInstallInvoiceNo, tmIvr.iTotalInstallInvs
                                    tmIvr.iInstallInvoiceNo = imInstallInvoiceNo
                                    tmIvr.iTotalInstallInvs = imTotalInstallInvs
                                    tmIvr.sInstallCntr = tmChf.sInstallDefined
                                    tmIvr.sEDIComment = ""
                                    tmIvr.lDisclaimer = 0
                                    If (tmPcf.sType <> "P") And (tmIbf.lCode > 0) Then
                                        slSQLQuery = "Update ibf_Impression_Bill Set ibfBilledImpression = " & tmIbf.lBilledImpression & " Where ibfCode = " & tmIbf.lCode
                                        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                                    End If
                                    If tmPcf.sType <> "H" Then
                                        ilRet = mDetermineIfEDIRequired()
                                        ilRet = mInsertIVRandClearDollars(ilEDIFlag, tmIvr)
                                        'TTP 10517 - Invoices: if "ad server" option is not checked on, and "commercial and NTR" invoices are set to be separate, the air time portion of the invoice does not print
                                        If (tmIvr.iType = IVRTYPE_TotalAirtimeRep) Or (tmIvr.iType = IVRTYPE_TotalAdServer) Or (tmIvr.iType = IVRTYPE_TotalNTR) Or (tmIvr.iType = IVRTYPE_TotalInstallment) Or (tmIvr.iType = IVRTYPE_TotalCntr) Then
                                            LSet tmImr = tmIvr
                                            tmImr.lIvrCode = tmIvr.lCode
                                            tmImr.sUnused = ""
                                            'ilRet = btrInsert(hmImr, tmImr, imImrRecLen, INDEXKEY0)
                                            ilRet = mWriteImr(ilEDIFlag, tmImr)
                                            If ilRet <> BTRV_ERR_NONE Then
                                                If ilRet >= 30000 Then
                                                    ilRet = csiHandleValue(0, 7)
                                                End If
                                                Print #hmMsg, "mNTR_GenIvr: btrInsert-Point 1, Imr Error # " & Trim$(str$(ilRet))
                                            End If
                                        End If
                                    End If
                                End If
                            Next ilLoop
                            'Now form type 3 record
Debug.Print "mInvoiceRpt: " & tmPcf.lCode & " (TotalAdServer) " & gStrDecToLong(slTGross, 2)
                            tmIvr.iType = IVRTYPE_TotalAdServer 'CPM Total
                            tmIvr.sOVehName = ""
                            tmIvr.sORate = ""
                            tmIvr.sRRemark = ""
                            tmIvr.sRAmount = ""
                            tmIvr.lOTotalSpots = 0  'ilONoSpots
                            tmIvr.lOTotalGross = 0  'llOGross
                            tmIvr.lATotalSpots = llTImpressions '0  'ilANoSpots
                            tmIvr.lATotalGross = gStrDecToLong(slTGross, 2)
                            tmIvr.lRTotalGross = 0
                            tmIvr.lATotalNet = gStrDecToLong(slTNet, 2)
                            tmIvr.lTax1 = llTax1
                            tmIvr.lTax2 = llTax2
                            tmIvr.lSpotKeyNo = 0
                            tmIvr.sKey = slForm2Key & "5"
                            If tgChfInv.sBillCycle = "C" Then
                                gPackDateLong lmSvStartCal, tmIvr.iInvStartDate(0), tmIvr.iInvStartDate(1)
                                gPackDateLong lmSvEndCal, tmIvr.iInvDate(0), tmIvr.iInvDate(1)
                            ElseIf tgChfInv.sBillCycle = "W" Then
                                gPackDateLong lmSvStartWk, tmIvr.iInvStartDate(0), tmIvr.iInvStartDate(1)
                                gPackDateLong lmSvEndWk, tmIvr.iInvDate(0), tmIvr.iInvDate(1)
                            Else
                                gPackDateLong lmSvStartStd, tmIvr.iInvStartDate(0), tmIvr.iInvStartDate(1)
                                gPackDateLong lmSvEndStd, tmIvr.iInvDate(0), tmIvr.iInvDate(1)
                            End If
                            tmIvr.lClfCxfCode = 0
                            tmIvr.lClfCode = 0
                            tmIvr.iFormType = 1  'Force to Air
                            tmIvr.lCode = 0
                            tmIvr.sARate = ""
                            tmIvr.sODPName = ""
                            tmIvr.iRdfCode = 0
                            tmIvr.iLineNo = 0
                            mGetInstallmentCounts tmChf ', tmIvr.iInstallInvoiceNo, tmIvr.iTotalInstallInvs
                            tmIvr.iInstallInvoiceNo = imInstallInvoiceNo
                            tmIvr.iTotalInstallInvs = imTotalInstallInvs
                            tmIvr.sInstallCntr = tmChf.sInstallDefined
                            tmIvr.sEDIComment = ""
                            'TTP 10716 - Invoice disclaimer: not shown on CPM/Digital-only invoice
                            'tmIvr.lDisclaimer = 0
                            tmIvr.lDisclaimer = tgSpf.lBCxfDisclaimer
                            ilRet = mInsertIVRandClearDollars(ilEDIFlag, tmIvr)
                            'TTP 10517 - Invoices: if "ad server" option is not checked on, and "commercial and NTR" invoices are set to be separate, the air time portion of the invoice does not print
                            If (tmIvr.iType = IVRTYPE_TotalAirtimeRep) Or (tmIvr.iType = IVRTYPE_TotalAdServer) Or (tmIvr.iType = IVRTYPE_TotalNTR) Or (tmIvr.iType = IVRTYPE_TotalInstallment) Or (tmIvr.iType = IVRTYPE_TotalCntr) Then
                                LSet tmImr = tmIvr
                                tmImr.lIvrCode = tmIvr.lCode
                                tmImr.sUnused = ""
                                ilRet = mWriteImr(ilEDIFlag, tmImr)
                                If ilRet <> BTRV_ERR_NONE Then
                                    If ilRet >= 30000 Then
                                        ilRet = csiHandleValue(0, 7)
                                    End If
                                    Print #hmMsg, "mNTR_GenIvr: btrInsert-Point 2, Imr Error # " & Trim$(str$(ilRet))
                                End If
                            End If
Debug.Print "mCPM_GenIvr: " & tmChf.lCntrNo & " (TotalCntr) " & tmIvr.lATotalGross + tlIvr.lATotalGross
                            tmIvr.iType = IVRTYPE_TotalCntr 'Contract Total
                            If llIvrCode > 0 Then
                                tmIvrSrchKey1.lCode = llIvrCode
                                ilRet = btrGetEqual(hmIvr, tlIvr, imIvrRecLen, tmIvrSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    tmIvr.lATotalGross = tmIvr.lATotalGross + tlIvr.lATotalGross
                                    tmIvr.lATotalNet = tmIvr.lATotalNet + tlIvr.lATotalNet
                                    tmIvr.lTax1 = tmIvr.lTax1 + tlIvr.lTax1
                                    tmIvr.lTax2 = tmIvr.lTax2 + tlIvr.lTax2
                                    tmIvr.lDisclaimer = tlIvr.lDisclaimer
                                    tmIvr.lComment(0) = tlIvr.lComment(0)
                                    tmIvr.lComment(1) = tlIvr.lComment(1)
                                    tmIvr.lComment(2) = tlIvr.lComment(2)
                                    tmIvr.lComment(3) = tlIvr.lComment(3)
                                    tmIvr.sEDIComment = tlIvr.sEDIComment
                                    tmIvr.lClfCxfCode = tlIvr.lClfCxfCode
                                    tmIvr.lClfCode = tlIvr.lClfCode
                                End If
                            End If
                            tmIvr.sARate = ""
                            tmIvr.sODPName = ""
                            tmIvr.iRdfCode = 0
                            tmIvr.iLineNo = 0
                            tmIvr.lCode = 0
                            mGetInstallmentCounts tmChf ', tmIvr.iInstallInvoiceNo, tmIvr.iTotalInstallInvs
                            tmIvr.iInstallInvoiceNo = imInstallInvoiceNo
                            tmIvr.iTotalInstallInvs = imTotalInstallInvs
                            tmIvr.sInstallCntr = tmChf.sInstallDefined
                            'Replace totals with Installment amounts
                            mResetIVRForInstallment tmIvr, ilPass
                            If (ilAirNTRCombineIndex = -1) Or (ilAirNTRCombineIndex >= 0 And tlIvr.iType = IVRTYPE_TotalAirtimeRep And llIvrCode > 0) Or (llIvrCode <= 0) Then
                                ilRet = mWriteIvr(ilEDIFlag, tmIvr, tmChf)
                                If ((ilAirNTRCombineIndex >= 0 And tlIvr.iType = IVRTYPE_TotalAirtimeRep) And llIvrCode > 0) Or ((ilAirNTRCombineIndex >= 0) And (llIvrCode <= 0)) Then
                                    tmAirNTRCombine(ilAirNTRCombineIndex).lIvrCode = tmIvr.lCode
                                End If
                            Else
                                tmIvr.lCode = tlIvr.lCode
                                ilRet = btrUpdate(hmIvr, tmIvr, imIvrRecLen)
                            End If
                            If ilRet <> BTRV_ERR_NONE Then
                                If ilRet >= 30000 Then
                                    ilRet = csiHandleValue(0, 7)
                                End If
                                Print #hmMsg, "mCPM_GenIvr: btrInsert-Point 3, Ivr Error # " & Trim$(str$(ilRet))
                            End If
                            'TTP 10517 - Invoices: if "ad server" option is not checked on, and "commercial and NTR" invoices are set to be separate, the air time portion of the invoice does not print
                            If (tmIvr.iType = IVRTYPE_TotalAirtimeRep) Or (tmIvr.iType = IVRTYPE_TotalAdServer) Or (tmIvr.iType = IVRTYPE_TotalNTR) Or (tmIvr.iType = IVRTYPE_TotalInstallment) Or (tmIvr.iType = IVRTYPE_TotalCntr) Then
                                LSet tmImr = tmIvr
                                tmImr.lIvrCode = tmIvr.lCode
                                tmImr.sUnused = ""
                                ilRet = mWriteImr(ilEDIFlag, tmImr)
                                If ilRet <> BTRV_ERR_NONE Then
                                    If ilRet >= 30000 Then
                                        ilRet = csiHandleValue(0, 7)
                                    End If
                                    Print #hmMsg, "mNTR_GenIvr: btrInsert-Point 4, Imr Error # " & Trim$(str$(ilRet))
                                End If
                            End If
                            
                            'Create RVF record
                            If Invoice.rbcType(INVGEN_Final).Value Then
                                tgSortKey(UBound(tgSortKey) - 1).iBilled = True
                                ilRvf = UBound(tgRvf)
                                ilSRvf = ilRvf
                                tgRvf(ilRvf).lCode = 0
                                'Product- update prf now instead of after all invoices generated
                                tgRvf(ilRvf).lPrfCode = 0
                                If (Trim$(slProduct) <> "") Then
                                    ilFound = False
                                    tmPrfSrchKey.iAdfCode = tmChf.iAdfCode
                                    ilRet = btrGetGreaterOrEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmPrf.iAdfCode = tmChf.iAdfCode) 'tmRvf.iAdfCode)
                                        If StrComp(Trim$(Trim$(slProduct)), Trim$(tmPrf.sName), 1) = 0 Then
                                            ilFound = True
                                            tgRvf(ilRvf).lPrfCode = tmPrf.lCode
                                            Exit Do
                                        End If
                                        ilRet = btrGetNext(hmPrf, tmPrf, imPrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                    If (Not ilFound) And (Invoice.rbcType(INVGEN_Final).Value) Then
                                        Do  'Loop until record updated or added
                                            tmPrf.lCode = 0
                                            tmPrf.iAdfCode = tmChf.iAdfCode
                                            tmPrf.sName = Trim$(slProduct)
                                            tmPrf.iMnfComp(0) = 0
                                            tmPrf.iMnfComp(1) = 0
                                            tmPrf.iMnfExcl(0) = 0
                                            tmPrf.iMnfExcl(1) = 0
                                            tmPrf.iUrfCode = tgUrf(0).iCode
                                            tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                                            tmPrf.lAutoCode = tmPrf.lCode
                                            ilRet = btrInsert(hmPrf, tmPrf, imPrfRecLen, INDEXKEY0)
                                        Loop While ilRet = BTRV_ERR_CONFLICT
                                        If ilRet <> BTRV_ERR_NONE Then
                                            If ilRet >= 30000 Then
                                                ilRet = csiHandleValue(0, 7)
                                            End If
                                            Print #hmMsg, "mNTR_GenIvr: btrInsert-Point 5, Prf Error # " & Trim$(str$(ilRet))
                                        End If
                                        If ilRet = BTRV_ERR_NONE Then
                                            tgRvf(ilRvf).lPrfCode = tmPrf.lCode
                                            Do
                                                tmPrf.iRemoteID = tgUrf(0).iRemoteUserID
                                                tmPrf.lAutoCode = tmPrf.lCode
                                                tmPrf.iSourceID = tgUrf(0).iRemoteUserID
                                                gPackDate Format(Now, "ddddd"), tmPrf.iSyncDate(0), tmPrf.iSyncDate(1)
                                                gPackTime Format(Now, "ttttt"), tmPrf.iSyncTime(0), tmPrf.iSyncTime(1)
                                                ilRet = btrUpdate(hmPrf, tmPrf, imPrfRecLen)
                                            Loop While ilRet = BTRV_ERR_CONFLICT
                                            If ilRet <> BTRV_ERR_NONE Then
                                                If ilRet >= 30000 Then
                                                    ilRet = csiHandleValue(0, 7)
                                                End If
                                                Print #hmMsg, "mNTR_GenIvr: btrUpdate-Point 3, Prf Error # " & Trim$(str$(ilRet))
                                            End If
                                        Else
                                            tgRvf(ilRvf).lPrfCode = 0
                                        End If
                                    End If
                                End If
                                tgRvf(ilRvf).iAgfCode = tmChf.iAgfCode
                                tgRvf(ilRvf).iAdfCode = tmChf.iAdfCode
                                tgRvf(ilRvf).iSlfCode = tmChf.iSlfCode(0)
                                tgRvf(ilRvf).lCntrNo = tmChf.lCntrNo
                                tgRvf(ilRvf).lInvNo = llAirCPMInvNo
                                tgRvf(ilRvf).lRefInvNo = 0
                                ''6/7/15: Check number changed to string
                                ''tgRvf(ilRvf).lCheckNo = llCPMDate   'should be 0 but temporary store llCPMDate so that RVF can be tied to tgSortKey
                                'tgRvf(ilRvf).sCheckNo = Trim$(str$(llCPMDate))   'should be 0 but temporary store llCPMDate so that RVF can be tied to tgSortKey
                                tgRvf(ilRvf).sCheckNo = ""
                                If tmChf.sBillCycle = "C" Then
                                    gPackDate smEndCal, tgRvf(ilRvf).iTranDate(0), tgRvf(ilRvf).iTranDate(1)
                                ElseIf tmChf.sBillCycle = "W" Then
                                    gPackDate smEndWk, tgRvf(ilRvf).iTranDate(0), tgRvf(ilRvf).iTranDate(1)
                                Else
                                    gPackDate smEndStd, tgRvf(ilRvf).iTranDate(0), tgRvf(ilRvf).iTranDate(1)
                                End If
                                tgRvf(ilRvf).sTranType = "IN"
                                tgRvf(ilRvf).sAction = " "
                                If tmChf.sBillCycle = "C" Then
                                    tgRvf(ilRvf).iAgePeriod = Month(gDateValue(smEndCal))
                                    tgRvf(ilRvf).iAgingYear = Year(gDateValue(smEndCal))
                                ElseIf tmChf.sBillCycle = "W" Then
                                    tgRvf(ilRvf).iAgePeriod = Month(gDateValue(smEndWk))
                                    tgRvf(ilRvf).iAgingYear = Year(gDateValue(smEndWk))
                                Else
                                    slStr = gObtainEndStd(smEndStd)
                                    tgRvf(ilRvf).iAgePeriod = Month(gDateValue(slStr))
                                    tgRvf(ilRvf).iAgingYear = Year(gDateValue(slStr))
                                End If
                                gPackDate "", tgRvf(ilRvf).iPurgeDate(0), tgRvf(ilRvf).iPurgeDate(1)
                                If ilPass = 2 Then
                                    tgRvf(ilRvf).sCashTrade = "T"
                                Else
                                    tgRvf(ilRvf).sCashTrade = "C"
                                End If
                                tgRvf(ilRvf).iPkLineNo = tmPcf.iPkCPMID
                                tgRvf(ilRvf).iInvDate(0) = tgRvf(ilRvf).iTranDate(0)
                                tgRvf(ilRvf).iInvDate(1) = tgRvf(ilRvf).iTranDate(1)
                                gPackDate Format$(gNow(), "m/d/yy"), tgRvf(ilRvf).iDateEntrd(0), tgRvf(ilRvf).iDateEntrd(1)
                                tgRvf(ilRvf).iUrfCode = tgUrf(0).iCode
                                tgRvf(ilRvf).iRemoteID = 0
                                tgRvf(ilRvf).iMnfGroup = 0  'Participant MnfCode, set in mUpdateRvf if Sales Source is Ask
                                tgRvf(ilRvf).lCefCode = 0
                                tgRvf(ilRvf).sInCollect = "N"
                                tgRvf(ilRvf).iRemoteID = 0
                                tgRvf(ilRvf).iBacklogTrfCode = 0
                                '1/17/09: Added buyer
                                tgRvf(ilRvf).iPnfBuyer = tmChf.iPnfBuyer
                                For ilLoop = ilCpm To UBound(tgCPMCntr) - 1 Step 1
                                    tmPcf = tgCPMCntr(ilLoop).tPcf
                                    tmIbf = tgCPMCntr(ilLoop).tIbf
                                    If (tmIbf.lCode > 0) And (tmPcf.sType <> "P") Then
                                        ilRvf = UBound(tgRvf)
                                        tgRvf(ilRvf).iBillVefCode = tmPcf.iVefCode
                                        If tmPcf.sType = "H" Then
                                            For ilPkg = 0 To UBound(tgCPMCntr) - 1 Step 1
                                                If tgCPMCntr(ilPkg).tPcf.sType = "P" Then
                                                    If (tgCPMCntr(ilPkg).tPcf.iPodCPMID = tmPcf.iPkCPMID) And (tgCPMCntr(ilPkg).tPcf.lChfCode = tmPcf.lChfCode) Then
                                                        tgRvf(ilRvf).iBillVefCode = tgCPMCntr(ilPkg).tPcf.iVefCode
                                                        Exit For
                                                    End If
                                                End If
                                            Next ilPkg
                                        End If
                                        tgRvf(ilRvf).iAirVefCode = tmIbf.iVefCode
                                        tgRvf(ilRvf).lRefInvNo = 0  'tgCPMCntr(ilLoop).tSbf.lRefInvNo
                                        tgRvf(ilRvf).iMnfItem = 0   'tgCPMCntr(ilLoop).tPcf.iMnfItem
                                        tgRvf(ilRvf).lSbfCode = 0
                                        tgRvf(ilRvf).lPcfCode = tmPcf.lCode
                                        If Invoice.rbcType(INVGEN_Reprint).Value Or Invoice.rbcType(INVGEN_Aff).Value Or Invoice.rbcType(INVGEN_Undo).Value Or Invoice.rbcType(INVGEN_Archive).Value Then      '11-15-16 reprint, aff, undo , archive Determine invoice number
                                            If (llChfCode <> tgCPMCntr(ilLoop).tPcf.lChfCode) Or (llReprintInvNo <> tgCPMCntr(ilLoop).lInvNo) Then
                                                Exit For
                                            End If
                                        Else
                                            If (llChfCode <> tgCPMCntr(ilLoop).tPcf.lChfCode) Then
                                                Exit For
                                            End If
                                        End If
                                        slGross = gLongToStrDec(tgCPMCntr(ilLoop).lCPMGross, 2)
                                        slNet = gLongToStrDec(tgCPMCntr(ilLoop).lCPMNet, 2)
                                        gStrToPDN slGross, 2, 6, tgRvf(ilRvf).sGross
                                        gStrToPDN slNet, 2, 6, tgRvf(ilRvf).sNet
                                        tgRvf(ilRvf).lTax1 = 0
                                        tgRvf(ilRvf).lTax2 = 0
                                        tgCPMCntr(ilLoop).lTax1 = 0
                                        tgCPMCntr(ilLoop).lTax2 = 0
                                        tgRvf(ilRvf).lAcquisitionCost = 0
                                        ReDim Preserve tgRvf(0 To UBound(tgRvf) + 1) As RVF
                                        tgRvf(ilRvf + 1) = tgRvf(ilRvf)
                                    End If
                                Next ilLoop
                            End If
                            If ilIncInvNo Then
                                llInvoiceNo = llInvoiceNo + 1
                                mCheckNextInvNo llInvoiceNo
                            End If
                        End If
                    End If
                Next ilPass
                'Advance ilCpm to next contract
                Do
                    If Invoice.rbcType(INVGEN_Reprint).Value Or Invoice.rbcType(INVGEN_Aff).Value Or Invoice.rbcType(INVGEN_Undo).Value Or Invoice.rbcType(INVGEN_Archive).Value Then     '11-15-16 reprint, aff, undo, archive, Determine invoice number
                        If (llChfCode <> tgCPMCntr(ilCpm).tPcf.lChfCode) Or (llReprintInvNo <> tgCPMCntr(ilCpm).lInvNo) Then
                            Exit Do
                        End If
                    Else
                        If (llChfCode <> tgCPMCntr(ilCpm).tPcf.lChfCode) Then
                            Exit Do
                        End If
                    End If
                    ilCpm = ilCpm + 1
                Loop While ilCpm <= UBound(tgCPMCntr) - 1
            End If
        Else
            ilCpm = ilCpm + 1
        End If
    Loop
    On Error Resume Next
    mnf_rst.Close
    raf_rst.Close
    rvf_rst.Close
    phf_rst.Close
    ibf_rst.Close
End Sub

Public Sub mGetInstallmentCounts(tlChf As CHF)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slDate                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    Dim llDate As Long
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilAdd As Integer
    Dim llStartCalDate As Long
    Dim llStartWkDate As Long
    Dim llStartStdDate As Long
    Dim llEndCalDate As Long
    Dim llEndWkDate As Long
    Dim llEndStdDate As Long
    Dim llCountInBillPeriod As Long
    
    If tlChf.lCntrNo = lmInstallCntrNo Then
        Exit Sub
    End If
    lmInstallCntrNo = tlChf.lCntrNo
    llEndCalDate = 0
    llEndWkDate = 0
    llEndStdDate = 0
    If lmEndCal > 0 Then
        llEndCalDate = gDateValue(gObtainEndCal(Format(lmEndCal, "ddddd")))
        llStartCalDate = gDateValue(gObtainStartCal(Format(lmEndCal, "ddddd")))
    End If
    If lmEndWk > 0 Then
        llEndWkDate = gDateValue(gObtainNextSunday(Format(lmEndWk, "ddddd")))
        llStartWkDate = gDateValue(gObtainPrevMonday(Format(lmEndWk, "ddddd")))
    End If
    If lmEndStd > 0 Then
        llEndStdDate = gDateValue(gObtainEndStd(Format(lmEndStd, "ddddd")))
        llStartStdDate = gDateValue(gObtainStartStd(Format(lmEndStd, "ddddd")))
    End If
    
    ReDim llInstallInvoiceNoDate(0 To 0) As Long
    ReDim ilTotalInstallInvsDate(0 To 0) As Long
    imInstallInvoiceNo = 0
    imTotalInstallInvs = 0
    llCountInBillPeriod = 0
    tmSbfSrchKey0.lChfCode = tlChf.lCode
    tmSbfSrchKey0.iDate(0) = 0
    tmSbfSrchKey0.iDate(1) = 0
    tmSbfSrchKey0.sTranType = "F"
    ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = tlChf.lCode)
        If tmSbf.sTranType = "F" Then
            ilAdd = False
            gUnpackDateLongError tmSbf.iDate(0), tmSbf.iDate(1), llDate, "69:mGetInstallmentCounts " & tmSbf.lCode
            If tlChf.sBillCycle = "C" Then
                If (llDate <= llEndCalDate) Then
                    ilAdd = True
                End If
                If Invoice.rbcType(INVGEN_Reprint).Value Then
                    If (tmSbf.sBilled = "Y") And (llDate >= llStartCalDate) And (llDate <= llEndCalDate) Then
                        llCountInBillPeriod = llCountInBillPeriod + 1
                    End If
                Else
                    If (tmSbf.sBilled <> "Y") And (llDate >= llStartCalDate) And (llDate <= llEndCalDate) Then
                        llCountInBillPeriod = llCountInBillPeriod + 1
                    End If
                End If
            ElseIf tlChf.sBillCycle = "W" Then
                If (llDate <= llEndWkDate) Then
                    ilAdd = True
                End If
                If Invoice.rbcType(INVGEN_Reprint).Value Then
                    If tmSbf.sBilled = "Y" And (llDate >= llStartWkDate) And (llDate <= llEndWkDate) Then
                        llCountInBillPeriod = llCountInBillPeriod + 1
                    End If
                Else
                    If tmSbf.sBilled <> "Y" And (llDate >= llStartWkDate) And (llDate <= llEndWkDate) Then
                        llCountInBillPeriod = llCountInBillPeriod + 1
                    End If
                End If
            Else
                If (llDate <= llEndStdDate) Then
                    ilAdd = True
                End If
                If Invoice.rbcType(INVGEN_Reprint).Value Then
                    If tmSbf.sBilled = "Y" And (llDate >= llStartStdDate) And (llDate <= llEndStdDate) Then
                        llCountInBillPeriod = llCountInBillPeriod + 1
                    End If
                Else
                    If tmSbf.sBilled <> "Y" And (llDate >= llStartStdDate) And (llDate <= llEndStdDate) Then
                        llCountInBillPeriod = llCountInBillPeriod + 1
                    End If
                End If
            End If
            If ilAdd Then
                ilFound = False
                For ilLoop = 0 To UBound(llInstallInvoiceNoDate) - 1 Step 1
                    If llInstallInvoiceNoDate(ilLoop) = llDate Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    llInstallInvoiceNoDate(UBound(llInstallInvoiceNoDate)) = llDate
                    ReDim Preserve llInstallInvoiceNoDate(0 To UBound(llInstallInvoiceNoDate) + 1) As Long
                End If
            End If
            ilFound = False
            For ilLoop = 0 To UBound(ilTotalInstallInvsDate) - 1 Step 1
                If ilTotalInstallInvsDate(ilLoop) = llDate Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                ilTotalInstallInvsDate(UBound(ilTotalInstallInvsDate)) = llDate
                ReDim Preserve ilTotalInstallInvsDate(0 To UBound(ilTotalInstallInvsDate) + 1) As Long
            End If
        End If
        ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (llCountInBillPeriod = 0) And (UBound(ilTotalInstallInvsDate) > 0) Then
        imInstallInvoiceNo = -1
        imTotalInstallInvs = UBound(ilTotalInstallInvsDate)
    Else
        imInstallInvoiceNo = UBound(llInstallInvoiceNoDate)
        imTotalInstallInvs = UBound(ilTotalInstallInvsDate)
    End If
End Sub

Public Function mDetermineIfEDIRequired() As Integer
    Dim ilRet As Integer
    Dim ilEDIFlag As Integer
    Dim ilArfInvCode As Integer
    Dim blFound As Boolean
    Dim ilEDI As Integer
    Dim ilSvEDIFlag As Integer

    '4/3/12:
    imEDIIndex = -1
    If (tgSpf.sAEDII = "Y") Then    'Test if agency using EDI
        If (tmChf.sType = "S") Or (tmChf.sType = "M") Then
            ilEDIFlag = False
            imArfPDFEMailCode = -1
        Else
            If tmChf.iAgfCode > 0 Then
                '12/17/08:  Added to check if EDI required within mCreateSortRec
                If tmChf.iAgfCode <> tmAgf.iCode Then
                    ilRet = gBinarySearchAgf(tmChf.iAgfCode)
                    If ilRet <> -1 Then
                        ilArfInvCode = tgCommAgf(ilRet).iArfInvCode
                    Else
                        tmAgfSrchKey.iCode = tmChf.iAgfCode
                        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            ilArfInvCode = tmAgf.iArfInvCode
                        End If
                    End If
                Else
                    ilArfInvCode = tmAgf.iArfInvCode
                End If
                If ilArfInvCode > 0 Then
                    If Not mFindEDIService(ilArfInvCode) Then
                        'if PDF Email, that also returns false
                        tmEDIArf.iCode = -1
                        tmEDIArf.sNmAd(0) = "Missing"
                        tmEDIArf.sNmAd(1) = " "
                        tmEDIArf.sNmAd(2) = " "
                        tmEDIArf.sNmAd(3) = " "
                        tmEDIArf.sEDIMediaType = ""
                        If imArfPDFEMailCode > 0 Then           '12-15-16 its a pdf email entry
                            ilEDIFlag = False
                        End If
                    Else
                        ilEDIFlag = True
                    End If
                Else
                    ilEDIFlag = False
                    imArfPDFEMailCode = -1
                End If
            Else
                '12/17/08:  Added to check if EDI required within mCreateSortRec
                If tmChf.iAdfCode <> tmAdf.iCode Then
                    ilRet = gBinarySearchAdf(tmChf.iAdfCode)
                    If ilRet <> -1 Then
                        ilArfInvCode = tgCommAdf(ilRet).iArfInvCode
                    Else
                        tmAdfSrchKey.iCode = tmChf.iAdfCode
                        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            ilArfInvCode = tmAdf.iArfInvCode
                        End If
                    End If
                Else
                    ilArfInvCode = tmAdf.iArfInvCode
                End If
                If ilArfInvCode > 0 Then
                    If Not mFindEDIService(ilArfInvCode) Then
                        'if PDF Email, that also returns false
                        tmEDIArf.iCode = -1
                        tmEDIArf.sNmAd(0) = "Missing"
                        tmEDIArf.sNmAd(1) = " "
                        tmEDIArf.sNmAd(2) = " "
                        tmEDIArf.sNmAd(3) = " "
                        tmEDIArf.sEDIMediaType = ""
                        If imArfPDFEMailCode > 0 Then           '12-15-16 its a pdf email entry
                            ilEDIFlag = False
                        End If
                    Else
                        ilEDIFlag = True
                    End If
                Else
                    ilEDIFlag = False
                    imArfPDFEMailCode = -1
                End If
            End If
        End If
    Else
        ilEDIFlag = False
        imArfPDFEMailCode = -1
    End If
    '1/9/18: If Installment contract ignore EDI
    ilSvEDIFlag = ilEDIFlag
    If ilEDIFlag And (tmChf.sInstallDefined = "Y") Then
        ilEDIFlag = False
        imArfPDFEMailCode = -1
        bmEDIInstallBypassed = True
        blFound = False
        For ilEDI = 0 To UBound(lmEDIInstallBypass) - 1 Step 1
            If tmChf.lCntrNo = lmEDIInstallBypass(ilEDI) Then
                blFound = True
                Exit For
            End If
        Next ilEDI
        If Not blFound Then
            lmEDIInstallBypass(UBound(lmEDIInstallBypass)) = tmChf.lCntrNo
            ReDim Preserve lmEDIInstallBypass(0 To UBound(lmEDIInstallBypass) + 1) As Long
            Print #hmMsg, "Installment Contract " & tmChf.lCntrNo & ":  Printed Invoice generated instead of EDI"
        End If
    End If
    If ilSvEDIFlag And (tmChf.sNTRDefined = "Y") And (imCombineAirAndNTR) Then
        ilEDIFlag = False
        imArfPDFEMailCode = -1
        bmEDINTRBypassed = True
        blFound = False
        For ilEDI = 0 To UBound(lmEDINTRBypass) - 1 Step 1
            If tmChf.lCntrNo = lmEDINTRBypass(ilEDI) Then
                blFound = True
                Exit For
            End If
        Next ilEDI
        If Not blFound Then
            lmEDINTRBypass(UBound(lmEDINTRBypass)) = tmChf.lCntrNo
            ReDim Preserve lmEDINTRBypass(0 To UBound(lmEDINTRBypass) + 1) As Long
            Print #hmMsg, "NTR Contract " & tmChf.lCntrNo & ":  Printed Invoice generated instead of EDI"
        End If
    End If
    If ilSvEDIFlag And (tmChf.sAdServerDefined = "Y") Then
        ilEDIFlag = False
        imArfPDFEMailCode = -1
        bmEDICPMBypassed = True
        blFound = False
        For ilEDI = 0 To UBound(lmEDICPMBypass) - 1 Step 1
            If tmChf.lCntrNo = lmEDICPMBypass(ilEDI) Then
                blFound = True
                Exit For
            End If
        Next ilEDI
        If Not blFound Then
            lmEDICPMBypass(UBound(lmEDICPMBypass)) = tmChf.lCntrNo
            ReDim Preserve lmEDICPMBypass(0 To UBound(lmEDICPMBypass) + 1) As Long
            Print #hmMsg, "Digital Contract " & tmChf.lCntrNo & ":  Printed Invoice generated instead of EDI"
        End If
    End If
    mDetermineIfEDIRequired = ilEDIFlag
End Function

Public Function mInsertIVRandClearDollars(ilEDIFlag As Integer, tlIvr As IVR) As Integer
    Dim ilRet As Integer
    Dim llCode As Long
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim ilTimeAdj As Integer
    Dim ilZone As Integer
    Dim slDate As String
    Dim llTime As Long
    Dim slTime As String
    Dim slVehName As String

    tmSvIvr = tlIvr
    If (tlIvr.sInstallCntr = "Y") And (tlIvr.iType <> 6) Then
        tlIvr.lATotalGross = 0
        tlIvr.lATotalNet = 0
        tlIvr.lOTotalGross = 0
        tlIvr.lRTotalGross = 0
        tlIvr.lTax1 = 0
        tlIvr.lTax2 = 0
        tlIvr.sARate = ""
        tlIvr.sORate = ""
        tlIvr.sRAmount = ""
    End If
    ilTimeAdj = 0
    If (tlIvr.iType = IVRTYPE_Spot) Or (tlIvr.iType = IVRTYPE_Bonus) Then
        'Change Air date and time is required
        If (smInvSpotTimeZone = "E") Or (smInvSpotTimeZone = "C") Or (smInvSpotTimeZone = "M") Or (smInvSpotTimeZone = "P") Then
            slVehName = Trim$(tlIvr.sAVehName)
            If slVehName = "" Then
                slVehName = Trim$(tlIvr.sOVehName)
            End If
            For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If StrComp(Trim$(tgMVef(ilVef).sName), slVehName, vbTextCompare) = 0 Then
                    ilVpf = gBinarySearchVpf(tgMVef(ilVef).iCode)
                    If ilVpf <> -1 Then
                        'For ilZone = 1 To 5 Step 1
                        For ilZone = 0 To 4 Step 1
                            If Left$(tgVpf(ilVpf).sGZone(ilZone), 1) = smInvSpotTimeZone Then
                                ilTimeAdj = tgVpf(ilVpf).iGLocalAdj(ilZone)
                                Exit For
                            End If
                        Next ilZone
                    End If
                    Exit For
                End If
            Next ilVef
            If ilTimeAdj <> 0 Then
                slDate = Trim$(Mid$(tlIvr.sADayDate, 4))
                slTime = tlIvr.sATime
                llTime = gTimeToLong(slTime, False) + (CLng(ilTimeAdj) * 3600)
                If llTime < 0 Then
                    '9442
'                    llTime = 86400 - llTime
                    llTime = 86400 + llTime
                    slDate = gDecOneDay(slDate)
                ElseIf llTime > 86399 Then
                    llTime = llTime - 86400
                    slDate = gIncOneDay(slDate)
                End If
                tlIvr.sADayDate = Trim$(Left$(Format$(slDate, "ddd"), 2) & ", " & slDate)
                tlIvr.sATime = gFormatTimeLong(llTime, "A", "1")
            End If
        End If
    End If
    ilRet = mWriteIvr(ilEDIFlag, tlIvr, tmChf)
    If ilRet <> BTRV_ERR_NONE Then
        If ilRet >= 30000 Then
            ilRet = csiHandleValue(0, 7)
        End If
        Print #hmMsg, "mInsertIVRandClearDollars: btrInsert-Point 1, Ivr Error # " & Trim$(str$(ilRet))
    End If
    llCode = tlIvr.lCode
    tlIvr = tmSvIvr
    tlIvr.lCode = llCode
    mInsertIVRandClearDollars = ilRet
End Function

Public Function mWriteImr(ilEDIFlag As Integer, tlImr As IMR) As Integer
    '11-14-16 if not archive operation (rbctype), and not archive checked on for final, continue to test the edi flag so its not created
    'otherwise, always create the record for printing
    If (Not (Invoice.rbcType(INVGEN_Archive).Value)) And (Invoice.ckcArchive.Value = vbUnchecked) Then
        If (ilEDIFlag) And (Invoice.ckcIncludeEDI.Value = vbUnchecked) Then
            mWriteImr = BTRV_ERR_NONE
            Exit Function
        End If
    End If
    mWriteImr = btrInsert(hmImr, tlImr, imImrRecLen, INDEXKEY0)
    Exit Function
End Function

Public Sub mResetIVRForInstallment(tlIvr As IVR, ilPass As Integer)
    Dim ilLoop As Integer
    Dim slTGross As String
    Dim slTNet As String
    Dim llTax1 As Long
    Dim llTax2 As Long
    Dim slNet As String
    Dim slGross As String
    Dim slStr As String
    Dim slAgyRate As String
    Dim slPctTrade As String
    Dim llTax1Rate As Long
    Dim llTax2Rate As Long
    Dim slGrossNet As String

    If tlIvr.sInstallCntr <> "Y" Then
        Exit Sub
    End If
    tlIvr.lATotalGross = 0
    tlIvr.lATotalNet = 0
    tlIvr.lTax1 = 0
    tlIvr.lTax2 = 0
    If Invoice.ckcType(INVTYPE_Commercial).Value <> vbChecked Then
        Exit Sub
    End If
    slTNet = ".00"
    slTGross = ".00"
    llTax1 = 0
    llTax2 = 0
    If tmIvr.iPctComm = 0 Then
        slAgyRate = ""
    Else
        slAgyRate = gIntToStrDec(tlIvr.iPctComm, 2)
    End If
    slPctTrade = gIntToStrDec(tmChf.iPctTrade, 0)
    For ilLoop = LBound(tgSbfInstall) To UBound(tgSbfInstall) - 1 Step 1
        If tlIvr.lChfCode = tgSbfInstall(ilLoop).tSbf.lChfCode Then
            slGross = gLongToStrDec(tgSbfInstall(ilLoop).tSbf.lGross, 2)
            If ilPass = 1 Then
                If Val(slPctTrade) <> 100 Then
                    slGross = gDivStr(gMulStr(slGross, gSubStr("100", slPctTrade)), "100")
                End If
                If slAgyRate = "" Then
                    slStr = slGross
                Else
                    slStr = gDivStr(gMulStr(slGross, gSubStr("100.00", slAgyRate)), "100.00")
                End If
            ElseIf ilPass = 2 Then
                If Val(slPctTrade) <> 100 Then
                    slGross = gSubStr(slGross, gDivStr(gMulStr(slGross, gSubStr("100", slPctTrade)), "100"))
                End If
                If slAgyRate = "" Then
                    slStr = slGross
                Else
                    slStr = gDivStr(gMulStr(slGross, gSubStr("100.00", slAgyRate)), "100.00")
                End If
            End If
            slNet = slStr
            slTNet = gAddStr(slTNet, slStr)
            If Invoice.rbcType(INVGEN_Reprint).Value Or Invoice.rbcType(INVGEN_Aff).Value Or Invoice.rbcType(INVGEN_Archive).Value Then       '11-15-16 reprint, aff or archive Determine invoice number
                llTax1 = llTax1 + tgSbfInstall(ilLoop).lTax1
                llTax2 = llTax2 + tgSbfInstall(ilLoop).lTax2
            Else
                tgSbfInstall(ilLoop).lTax1 = 0
                tgSbfInstall(ilLoop).lTax2 = 0
                If ilPass = 1 Then
                    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
                        If ((ilPass = 1) And ((Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA)) Or ((Asc(tgSpf.sUsingFeatures4) And TAXBYCANADA) = TAXBYCANADA) Then
                            gGetNTRTaxRates tgSbfInstall(ilLoop).tSbf.iTrfCode, llTax1Rate, llTax2Rate, slGrossNet
                            If llTax1Rate > 0 Then
                                '12/17/06-Change to tax by agency or vehicle
                                'tgSbfInstall(ilLoop).tSbf.lTax1 = gStrDecToLong(gRoundStr(gMulStr(slStr, gIntToStrDec(tgSpf.iBTax(0), 4)), ".01", 2), 2)
                                'llTax1 = llTax1 + tgSbfInstall(ilLoop).tSbf.lTax1
                                ''Don't set tax amount into sbf because only one record and cash/trade might not be 50% so tax amount would be different
                                ''tgSbfInstall(ilLoop).tSbf.lTax1 = 0
                                'If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
                                If slGrossNet = "G" Then
                                    'Compute from Gross amount
                                    slStr = slGross
                                'ElseIf (Asc(tgSpf.sUsingFeatures4) And TAXBYCANADA) = TAXBYCANADA Then
                                ElseIf slGrossNet = "N" Then
                                    'Compute from Net amount
                                    slStr = slNet
                                End If
                                tgSbfInstall(ilLoop).lTax1 = gStrDecToLong(gRoundStr(gDivStr(gMulStr(slStr, gLongToStrDec(llTax1Rate, 4)), "100."), ".01", 2), 2)
                                llTax1 = llTax1 + tgSbfInstall(ilLoop).lTax1
                            Else
                                '12/17/06-Change to tax by agency or vehicle
                                'tgSbfInstall(ilLoop).tSbf.lTax1 = 0
                            End If
                            '12/17/06-Change to tax by agency or vehicle
                            'If (tgSpf.iBTax(1) <> 0) And (tgSbfInstall(ilLoop).tSbf.sSlsTax = "Y") And (ilTax2 = True) Then
                            If llTax2Rate > 0 Then
                                '12/17/06-Change to tax by agency or vehicle
                                'tgSbfInstall(ilLoop).tSbf.lTax2 = gStrDecToLong(gRoundStr(gMulStr(slStr, gIntToStrDec(tgSpf.iBTax(1), 4)), ".01", 2), 2)
                                'llTax2 = llTax2 + tgSbfInstall(ilLoop).tSbf.lTax2
                                ''tgSbfInstall(ilLoop).tSbf.lTax2 = 0
                                'If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
                                If slGrossNet = "G" Then
                                    'Compute from Gross amount
                                    slStr = slGross
                                'ElseIf (Asc(tgSpf.sUsingFeatures4) And TAXBYCANADA) = TAXBYCANADA Then
                                ElseIf slGrossNet = "N" Then
                                    'Compute from Net amount
                                    slStr = slNet
                                End If
                                tgSbfInstall(ilLoop).lTax2 = gStrDecToLong(gRoundStr(gDivStr(gMulStr(slStr, gLongToStrDec(llTax2Rate, 4)), "100."), ".01", 2), 2)
                                llTax2 = llTax2 + tgSbfInstall(ilLoop).lTax2
                            Else
                                '12/17/06-Change to tax by agency or vehicle
                                'tgSbfInstall(ilLoop).tSbf.lTax2 = 0
                            End If
                        End If
                    End If
                End If
            End If
            slTGross = gAddStr(slTGross, slGross)
        End If
    Next ilLoop
    tmIvr.lATotalGross = gStrDecToLong(slTGross, 2)
    tmIvr.lATotalNet = gStrDecToLong(slTNet, 2)
    tmIvr.lTax1 = llTax1
    tmIvr.lTax2 = llTax2
End Sub

Public Function mWriteIvr(ilEDIFlag As Integer, tlIvr As IVR, tlChf As CHF, Optional blItsREP As Boolean = False) As Integer
    Dim ilRet As Integer
    
    gInvExport_GatherDetail imCombineAirAndNTR, hmInvExportSpots, hmInvExportNTR, hmSbf, blItsREP, tlIvr, tmSmf, tmSbf     'always generate the invoice export if feature is set in site
    '11-14-16 if not archive operation (rbctype), and not archive checked on for final, continue to test the edi flag so its not created
    'otherwise, always create the record for printing
    If (Not (Invoice.rbcType(INVGEN_Archive).Value)) And (Invoice.ckcArchive.Value = vbUnchecked) Then
        If (ilEDIFlag) And (Invoice.ckcIncludeEDI.Value = vbUnchecked) Then
            mWriteIvr = BTRV_ERR_NONE
            Exit Function
        End If
    End If
    
    tlIvr.sPDFType = "0"
    tlIvr.iAdfOrAgfCode = 0
    If imArfPDFEMailCode > 0 Then
        If tlChf.iAgfCode > 0 Then
            If tmPDFAgf.iCode <> tlChf.iAgfCode Then
                If tmAgf.iCode <> tlChf.iAgfCode Then
                    tmAgfSrchKey.iCode = tlChf.iAgfCode
                    ilRet = btrGetEqual(hmAgf, tmPDFAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet <> BTRV_ERR_NONE Then
                        tmPDFAgf.iArfInvCode = -1
                    End If
                Else
                    tmPDFAgf = tmAgf
                End If
            End If
            If tmPDFAgf.iArfInvCode = imArfPDFEMailCode Then
                tlIvr.sPDFType = "2"
                tlIvr.iAdfOrAgfCode = tmPDFAgf.iCode
            End If
        Else
            If tmPDFAdf.iCode <> tlChf.iAdfCode Then
                If tmAdf.iCode <> tlChf.iAdfCode Then
                    tmAdfSrchKey.iCode = tlChf.iAdfCode
                    ilRet = btrGetEqual(hmAdf, tmPDFAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet <> BTRV_ERR_NONE Then
                        tmPDFAdf.iArfInvCode = -1
                    End If
                Else
                    tmPDFAdf = tmAdf
                End If
            End If
            If tmPDFAdf.iArfInvCode = imArfPDFEMailCode Then
                tlIvr.sPDFType = "1"
                tlIvr.iAdfOrAgfCode = tmPDFAdf.iCode
            End If
        End If
    End If
    
    mWriteIvr = btrInsert(hmIvr, tlIvr, imIvrRecLen, INDEXKEY1)
    Exit Function
End Function

Public Sub mCheckNextInvNo(llInvoiceNo As Long)
    Dim tlRvf As RVF
    Dim ilRet As Integer

    If Invoice.rbcType(INVGEN_Reprint).Value Or Invoice.rbcType(INVGEN_Aff).Value Or Invoice.rbcType(INVGEN_Undo).Value Or Invoice.rbcType(INVGEN_Archive).Value Then        '11-15-16 reprint, aff, undo or archive
        Exit Sub
    End If
    Do
        tmRvfSrchKey5.lInvNo = llInvoiceNo
        ilRet = btrGetEqual(hmRvf, tlRvf, imRvfRecLen, tmRvfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
            tmRvfSrchKey5.lInvNo = llInvoiceNo
            ilRet = btrGetEqual(hmPhf, tlRvf, imRvfRecLen, tmRvfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_KEY_NOT_FOUND Then
                Exit Sub
            End If
        End If
        llInvoiceNo = llInvoiceNo + 1
        If llInvoiceNo > tgSpf.lBHighestNo Then
            llInvoiceNo = tgSpf.lBLowestNo
        End If
    Loop
End Sub

Public Function mFindEDIService(ilArfInvCode As Integer) As Integer
    Dim ilEDI As Integer
    imEDIIndex = -1
    imArfPDFEMailCode = -1
    For ilEDI = 0 To UBound(tgEDIServiceInfo) - 1 Step 1
        If ilArfInvCode = tgEDIServiceInfo(ilEDI).tArf.iCode Then
            If Trim$(tgEDIServiceInfo(ilEDI).tArf.sID) <> "PDF EMail" Then
                tmEDIArf = tgEDIServiceInfo(ilEDI).tArf
                hmEDI = tgEDIServiceInfo(ilEDI).hEDI
                tgEDIServiceInfo(ilEDI).sUsed = "Y"
                '4/3/12
                imEDIIndex = ilEDI
                mFindEDIService = True
                Exit Function
            Else        '12-15'16
                imArfPDFEMailCode = ilArfInvCode
                Exit Function
            End If
        End If
    Next ilEDI
    mFindEDIService = False
End Function

Public Sub mCPM_BuildBypassCntr(ilFullCal As Integer, llCalEndDate As Long, ilFullWk As Integer, llWkEndDate As Long, ilFullStd As Integer, llStdEndDate As Long)
    Dim ilRet As Integer    'Return status
    Dim ilAgePeriod As Integer
    Dim ilAgingYear As Integer
    Dim slPcfStartDate As String
    Dim slPCFEndDate As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilFound As Integer
    Dim ilBillCycle As Integer
    Dim ilLoop As Integer
    Dim ilChfRecLen As Integer        'CHF record length
    Dim tlChf As CHF
    Dim llImpressionsPosted As Long 'TTP 10720 - Invoices: red unbilled screen when digital contract was fully billed prior to line end date in prior month
    Dim ilInclude As Integer 'TTP 10960 - Invoice: digital lines will get partially invoiced if final invoices are run before the lines are expired/month is over
    Dim slDate As String
    'Dim imFullCal As Integer
    'Dim llCalEndDate As Long
    'Dim imFullWk As Integer
    'Dim llWkEndDate As Long
    'Dim imFullStd As Integer
    
    
    If ((Asc(tgSaf(0).sFeatures8) And PODADSERVER) <> PODADSERVER) Then     'CPM Invoices
        ReDim lmCPMBypassCntr(0 To 1) As Long
        lmCPMBypassCntr(0) = -1
        Exit Sub
    End If
    If UBound(lmCPMBypassCntr) > LBound(lmCPMBypassCntr) Then
        Exit Sub
    End If
    ReDim lmCPMBypassCntr(0 To 1) As Long
    lmCPMBypassCntr(0) = -1
    imPcfRecLen = Len(tmPcf)
    ilChfRecLen = Len(tlChf)
    imIbfRecLen = Len(tmIbf)
    If (Invoice.rbcType(INVGEN_Reprint).Value) Or (Invoice.rbcType(INVGEN_Archive).Value) Then        '11-15-16 reprint or archive
        'Need to have created an array of sbf records to reprint
        'ilRet = mCPM_BuildReprint()
    ElseIf (Invoice.rbcType(INVGEN_Preliminary).Value) Or (Invoice.rbcType(INVGEN_Final).Value) Then 'Preliminary or final
        For ilBillCycle = 0 To 2 Step 1
            ilFound = False
            If Invoice.ckcBillCycle(INVBILLCYCLE_Calendar).Value = vbChecked Then   'Cal
                ilAgePeriod = Month(gDateValue(smEndCal))
                ilAgingYear = Year(gDateValue(smEndCal))
                slStartDate = smStartCal
                slEndDate = smEndCal
                ilFound = True
            ElseIf Invoice.ckcBillCycle(INVBILLCYCLE_Week).Value = vbChecked Then    'Week
                ilFound = False
            ElseIf Invoice.ckcBillCycle(INVBILLCYCLE_STD).Value = vbChecked Then
                ilAgePeriod = Month(gDateValue(smEndStd))
                ilAgingYear = Year(gDateValue(smEndStd))
                slStartDate = smStartStd
                slEndDate = smEndStd
                ilFound = True
            End If
            If ilFound Then
                gPackDate slStartDate, tmPcfSrchKey3.iEndDate(0), tmPcfSrchKey3.iEndDate(1)
                ilRet = btrGetGreaterOrEqual(hmPcf, tmPcf, imPcfRecLen, tmPcfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE)
                    'Test start date
                    gUnpackDate tmPcf.iStartDate(0), tmPcf.iStartDate(1), slPcfStartDate
                    gUnpackDate tmPcf.iEndDate(0), tmPcf.iEndDate(1), slPCFEndDate
                    If (gDateValue(slPCFEndDate) >= gDateValue(slStartDate)) And (gDateValue(slPcfStartDate) <= gDateValue(slEndDate)) Then
                        'Test if Not a package
                        If tmPcf.sType <> "P" Then
                            'Test contract
                            ilFound = False
                            tmChfSrchKey.lCode = tmPcf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tlChf, ilChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If (ilRet = BTRV_ERR_NONE) And (tlChf.sDelete <> "Y") Then
                                If ilBillCycle = 1 And tlChf.sBillCycle = "C" Then
                                    ilFound = True
                                End If
                                If ilBillCycle = 0 And tlChf.sBillCycle = "S" Then
                                    ilFound = True
                                End If
                                'TTP 10697 - Ad server tab and working proposals status triggering billing alert
                                '1. Working, completed, unapproved, and rejected proposals should not be checked at the end of the invoice process
                                '   if they have unbilled digital lines because those proposal types have never been scheduled.
                                '   Air time proposals are not checked either.
                                '-------------------------------------------
                                'W=Working Proposal or C=Completed Proposal or I=Unapproved or D=Rejected
                                If tlChf.sStatus = "W" Or tlChf.sStatus = "C" Or tlChf.sStatus = "I" Or tlChf.sStatus = "D" Then
                                    ilFound = False
                                End If
                            End If
                            
                            'TTP 10960 - Invoice: digital lines will get partially invoiced if final invoices are run before the lines are expired/month is over
                            If ilFound Then
                                ilInclude = True
                                If tlChf.sBillCycle = "C" Then      'Calendar month
                                    If Not ilFullCal Then
                                        gUnpackDate tlChf.iEndDate(0), tlChf.iEndDate(1), slDate
                                        If gDateValue(slDate) > llCalEndDate Then
                                            ilInclude = False
                                        End If
                                    End If
                                ElseIf tmChf.sBillCycle = "W" Then  'Weekly
                                    If Not ilFullWk Then
                                        gUnpackDate tlChf.iEndDate(0), tlChf.iEndDate(1), slDate
                                        If gDateValue(slDate) > llWkEndDate Then
                                            ilInclude = False
                                        End If
                                    End If
                                Else                    'Std
                                    If Not ilFullStd Then
                                        gUnpackDate tlChf.iEndDate(0), tlChf.iEndDate(1), slDate
                                        If gDateValue(slDate) > llStdEndDate Then
                                            ilInclude = False
                                        End If
                                    End If
                                End If
                                'if this contract should not be included because we're invoicing a partial month and this digital line is beyond the end date, add it to the lmCPMBypassCntr array
                                If ilInclude = False Then
                                    ilFound = False
                                    For ilLoop = 1 To UBound(lmCPMBypassCntr) - 1 Step 1
                                        If lmCPMBypassCntr(ilLoop) = tlChf.lCode Then
                                            ilInclude = True
                                            Exit For
                                        End If
                                    Next ilLoop
                                    If Not ilInclude Then
                                        lmCPMBypassCntr(UBound(lmCPMBypassCntr)) = tlChf.lCode
                                        ReDim Preserve lmCPMBypassCntr(0 To UBound(lmCPMBypassCntr) + 1) As Long
                                    End If
                                End If
                            End If

                            If ilFound Then
                                If tmPcf.sPriceType <> "F" Then 'Flat rate contracts don't require posting
                                    ''TTP 10480 - V81 ONLY: Preliminary/Final selective invoices: contracts with any ad server lines with no posted impressions don't appear in the list of contracts to run preliminary invoices for
                                    'revert TTP 10480 - replaced with TTP 10597 - Ad server/digital show Unposted Digital Contracts in Selective List in RED color
                                    ilFound = False
                                    llImpressionsPosted = 0
                                    tmIbfSrchKey1.lCntrNo = tlChf.lCntrNo
                                    tmIbfSrchKey1.iPodCPMID = tmPcf.iPodCPMID
                                    ilRet = btrGetEqual(hmIbf, tmIbf, imIbfRecLen, tmIbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmIbf.lCntrNo = tlChf.lCntrNo) And (tmIbf.iPodCPMID = tmPcf.iPodCPMID)
                                        If (tmIbf.iBillMonth = ilAgePeriod) And (tmIbf.iBillYear = ilAgingYear) Then
                                            ilFound = True
                                            Exit Do
                                        End If
                                        'TTP 10720 - Invoices: red unbilled screen when digital contract was fully billed prior to line end date in prior month
                                        'detect that impressions have been fully posted for this line and skip this line when determining whether the show the red screen or not.
                                        llImpressionsPosted = llImpressionsPosted + tmIbf.lImpressions
                                        If llImpressionsPosted >= tmPcf.lImpressionGoal Then
                                            ilFound = True
                                            Exit Do
                                        End If
                                        ilRet = btrGetNext(hmIbf, tmIbf, imIbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                    If Not ilFound Then
                                        If (ilRet <> BTRV_ERR_NONE) Or (tmIbf.lCntrNo <> tlChf.lCntrNo) Or (tmIbf.iPodCPMID <> tmPcf.iPodCPMID) Or (tmIbf.iBillMonth <> ilAgePeriod) Or (tmIbf.iBillYear <> ilAgingYear) Then
                                            ilFound = False
                                            For ilLoop = 1 To UBound(lmCPMBypassCntr) - 1 Step 1
                                                If lmCPMBypassCntr(ilLoop) = tlChf.lCode Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilLoop
                                            If Not ilFound Then
                                                lmCPMBypassCntr(UBound(lmCPMBypassCntr)) = tlChf.lCode
                                                ReDim Preserve lmCPMBypassCntr(0 To UBound(lmCPMBypassCntr) + 1) As Long
                                            End If
                                        End If
                                    End If
                                End If 'End of PriceType <> "F"
                            End If
                        End If
                    End If
                    ilRet = btrGetNext(hmPcf, tmPcf, imPcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
        Next ilBillCycle
    End If
End Sub

Public Function mSetCPMBillFlag(llSortKeyIndex As Long) As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim llChfCode As Long
    Dim llEndDate As Long
    
    If Not Invoice.rbcType(INVGEN_Final).Value Then
        mSetCPMBillFlag = BTRV_ERR_NONE
        Exit Function
    End If
    llChfCode = tgSortKey(llSortKeyIndex).lChfCode
    llEndDate = tgSortKey(llSortKeyIndex).lCPMDate
    For ilLoop = LBound(tgCPMCntr) To UBound(tgCPMCntr) - 1 Step 1
        If (tgCPMCntr(ilLoop).tPcf.lChfCode = llChfCode) And (tgCPMCntr(ilLoop).lEDate = llEndDate) Then
            If tgCPMCntr(ilLoop).tIbf.lCode > 0 Then
                Do
                    tmIbfSrchKey0.lCode = tgCPMCntr(ilLoop).tIbf.lCode
                    ilRet = btrGetEqual(hmIbf, tmIbf, imIbfRecLen, tmIbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                    If ilRet <> BTRV_ERR_NONE Then
                        If ilRet >= 30000 Then
                            ilRet = csiHandleValue(0, 7)
                        End If
                        mSetCPMBillFlag = ilRet
                        Exit Function
                    End If
                    tmIbf.sBilled = "Y"
                    ilRet = btrUpdate(hmIbf, tmIbf, imIbfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    Print #hmMsg, "mUpdateInv: btrUpdate-Point 5, Ibf Error # " & Trim$(str$(ilRet))
                End If
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet >= 30000 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    mSetCPMBillFlag = ilRet
                    Exit Function
                End If
            End If
        End If
    Next ilLoop
    mSetCPMBillFlag = BTRV_ERR_NONE
End Function

Private Function mCPM_BuildReprint() As Integer
    Dim ilPass As Integer
    Dim hlFile As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim ilMnfSort As Integer
    Dim llPcfCode As Long
    Dim ilOffSet As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim slKey As String
    Dim llRecPos As Long
    Dim slDate As String
    Dim slInvNo As String
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim llMergeInvNo As Long
    Dim ilMerge As Integer
    Dim ilIncludeCPM As Integer
    Dim ilAgePeriod As Integer
    Dim ilAgingYear As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilPkg As Integer
    Dim tlPcf As PCF
    Dim tlIbf As IBF
    Dim ilBillCycle As Integer

    imIbfRecLen = Len(tmIbf)
    imPcfRecLen = Len(tmPcf)
    For ilBillCycle = 0 To 2 Step 1
        ilFound = False
        If Invoice.ckcBillCycle(INVBILLCYCLE_Calendar).Value = vbChecked And ilBillCycle = 1 Then   'Cal
            ilAgePeriod = Month(gDateValue(smEndCal))
            ilAgingYear = Year(gDateValue(smEndCal))
            slStartDate = smStartCal
            slEndDate = smEndCal
            ilFound = True
        ElseIf Invoice.ckcBillCycle(INVBILLCYCLE_Week).Value = vbChecked And ilBillCycle = 2 Then   'Week
        ElseIf Invoice.ckcBillCycle(INVBILLCYCLE_STD).Value = vbChecked And ilBillCycle = 0 Then
            ilAgePeriod = Month(gDateValue(smEndStd))
            ilAgingYear = Year(gDateValue(smEndStd))
            slStartDate = smStartStd
            slEndDate = smEndStd
            ilFound = True
        End If
        If ilFound Then
            For ilPass = 0 To 2 Step 1  'leave the 2 because Two passes in PHF (IN and HI)
                If ilPass = 0 Then
                    hlFile = hmRvf
                Else
                    hlFile = hmPhf
                End If
                ilExtLen = Len(tmRvf)  'Extract operation record size
                llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlRvf) 'Obtain number of records
                btrExtClear hlFile   'Clear any previous extend operation
                ilRet = btrGetFirst(hlFile, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_END_OF_FILE Then
                    If ilPass = 0 Then
                        Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "RVF", "") 'Set extract limits (all records)
                    Else
                        Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "PHF", "") 'Set extract limits (all records)
                    End If
                    If ilPass <> 2 Then
                        tlCharTypeBuff.sType = "I"
                    Else
                        tlCharTypeBuff.sType = "H"
                    End If
                    ilOffSet = gFieldOffset("Rvf", "RvfTranType")
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                    If ilRet <> BTRV_ERR_NONE Then
                        mCPM_BuildReprint = ilRet
                        Exit Function
                    End If
                    If ilPass <> 2 Then
                        tlCharTypeBuff.sType = "N"
                    Else
                        tlCharTypeBuff.sType = "I"
                    End If
                    ilOffSet = gFieldOffset("Rvf", "RvfTranType") + 1
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                    If ilRet <> BTRV_ERR_NONE Then
                        mCPM_BuildReprint = ilRet
                        Exit Function
                    End If
                    ilMnfSort = 0
                    ilOffSet = gFieldOffset("Rvf", "RvfMnfItem")
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, ilMnfSort, 2)
                    If ilRet <> BTRV_ERR_NONE Then
                        mCPM_BuildReprint = ilRet
                        Exit Function
                    End If
                    llPcfCode = 0
                    ilOffSet = gFieldOffset("Rvf", "RvfPcfCode")
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GT, BTRV_EXT_LAST_TERM, llPcfCode, 4)
                    If ilRet <> BTRV_ERR_NONE Then
                        mCPM_BuildReprint = ilRet
                        Exit Function
                    End If
                    ilOffSet = 0
                    ilRet = btrExtAddField(hlFile, ilOffSet, ilExtLen)  'Extract First Name field
                    If ilRet <> BTRV_ERR_NONE Then
                        Invoice.imBuildingAdvt = False
                        Exit Function
                    End If
                    ilRet = btrExtGetNext(hlFile, tmRvf, ilExtLen, llRecPos)
                    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                            Invoice.imBuildingAdvt = False
                            Exit Function
                        End If
                        ilExtLen = Len(tmRvf)  'Extract operation record size
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hlFile, tmRvf, ilExtLen, llRecPos)
                        Loop
                        Do While ilRet = BTRV_ERR_NONE
                            ilFound = False
                            llMergeInvNo = -1
                            If ilPass > 1 Then
                                For ilMerge = 0 To UBound(tgAdvanceBillMergeInfo) - 1 Step 1
                                    If tmRvf.lInvNo = tgAdvanceBillMergeInfo(ilMerge).lPhfInvNo Then
                                        llMergeInvNo = tgAdvanceBillMergeInfo(ilMerge).lRvfInvNo
                                        Exit For
                                    End If
                                Next ilMerge
                            End If
                            For ilLoop = 0 To UBound(tmRPInfo) - 1 Step 1
                                If (tmRPInfo(ilLoop).iMnfItem = 0) Or (tmRPInfo(ilLoop).iMnfItem = -1) Then
                                    If (tmRvf.lCntrNo = tmRPInfo(ilLoop).lCntrNo) And ((tmRvf.lInvNo = tmRPInfo(ilLoop).lInvNo) Or (llMergeInvNo = tmRPInfo(ilLoop).lInvNo)) Then
                                        If tmRvf.lPcfCode > 0 Then
                                            llPcfCode = tmRvf.lPcfCode
                                            ilFound = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next ilLoop
                            If ilFound Then
                                ilFound = False
                                For ilLoop = LBound(tgCPMCntr) To UBound(tgCPMCntr) - 1 Step 1
                                    If tgCPMCntr(ilLoop).tPcf.lCode = tmRvf.lPcfCode Then
                                        ilFound = True
                                        tgCPMCntr(ilLoop).lTax1 = tgCPMCntr(ilLoop).lTax1 + tmRvf.lTax1
                                        tgCPMCntr(ilLoop).lTax2 = tgCPMCntr(ilLoop).lTax2 + tmRvf.lTax2
                                        Exit For
                                    End If
                                Next ilLoop
                                If Not ilFound Then
                                    tmPcfSrchKey0.lCode = tmRvf.lPcfCode
                                    ilRet = btrGetEqual(hmPcf, tmPcf, imPcfRecLen, tmPcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If (ilRet = BTRV_ERR_NONE) Then
                                        tmChfSrchKey.lCode = tmPcf.lChfCode
                                        ilRet = btrGetEqual(hmCHF, tgChfInv, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If (ilRet = BTRV_ERR_NONE) Then
                                            ilRet = Not BTRV_ERR_NONE
                                            If ilBillCycle = 1 And tgChfInv.sBillCycle = "C" Then    'Cal
                                                ilRet = BTRV_ERR_NONE
                                            End If
                                            If ilBillCycle = 0 And tgChfInv.sBillCycle = "S" Then    'Cal
                                                ilRet = BTRV_ERR_NONE
                                            End If
                                        End If
                                        If (ilRet = BTRV_ERR_NONE) Then
                                        
                                            'Get IBF
                                            tmIbfSrchKey1.lCntrNo = tgChfInv.lCntrNo
                                            tmIbfSrchKey1.iPodCPMID = tmPcf.iPodCPMID
                                            ilRet = btrGetEqual(hmIbf, tmIbf, imIbfRecLen, tmIbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            Do While (ilRet = BTRV_ERR_NONE) And (tmIbf.lCntrNo = tgChfInv.lCntrNo) And (tmIbf.iPodCPMID = tmPcf.iPodCPMID)
                                                If (tmIbf.iBillMonth = ilAgePeriod) And (tmIbf.iBillYear = ilAgingYear) Then
                                                    Exit Do
                                                End If
                                                ilRet = btrGetNext(hmIbf, tmIbf, imIbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                            Loop
                                            If (ilRet <> BTRV_ERR_NONE) Or (tmIbf.lCntrNo <> tgChfInv.lCntrNo) Or (tmIbf.iPodCPMID <> tmPcf.iPodCPMID) Or (tmIbf.iBillMonth <> ilAgePeriod) Or (tmIbf.iBillYear <> ilAgingYear) Then
                                                'Missing billing info
                                                tmIbf.lCode = 0
                                                tmIbf.sBilled = "Y"
                                                tmIbf.sBillCycle = tgChfInv.sBillCycle
                                                ilRet = BTRV_ERR_NONE
                                            End If
                                            If (tmIbf.sBilled = "Y") Then
                                                slKey = Trim$(str$(tmPcf.lChfCode))
                                                Do While Len(slKey) < 8
                                                    slKey = "0" & slKey
                                                Loop
                                                slInvNo = Trim(str$(tmRvf.lInvNo))
                                                Do While Len(slInvNo) < 8
                                                    slInvNo = "0" & slInvNo
                                                Loop
                                                slDate = Trim$(str$(gDateValue(slStartDate)))
                                                Do While Len(slDate) < 5
                                                    slDate = "0" & slDate
                                                Loop
                                                tgCPMCntr(UBound(tgCPMCntr)).sKey = slKey & slInvNo & slDate & "A"
                                                tgCPMCntr(UBound(tgCPMCntr)).lSDate = Val(slDate)
                                                tgCPMCntr(UBound(tgCPMCntr)).lEDate = gDateValue(slEndDate)
                                                tgCPMCntr(UBound(tgCPMCntr)).tPcf = tmPcf
                                                tgCPMCntr(UBound(tgCPMCntr)).tIbf = tmIbf
                                                If llMergeInvNo = -1 Then
                                                    tgCPMCntr(UBound(tgCPMCntr)).lInvNo = tmRvf.lInvNo
                                                Else
                                                    tgCPMCntr(UBound(tgCPMCntr)).lInvNo = llMergeInvNo
                                                End If
                                                tgCPMCntr(UBound(tgCPMCntr)).lTax1 = tmRvf.lTax1
                                                tgCPMCntr(UBound(tgCPMCntr)).lTax2 = tmRvf.lTax2
                                                gUnpackDateLongError tmRvf.iTranDate(0), tmRvf.iTranDate(1), tgCPMCntr(UBound(tgCPMCntr)).lInvDate, "48:mCPM_BuildReprint " & tmRvf.lCode
                                                ReDim Preserve tgCPMCntr(0 To UBound(tgCPMCntr) + 1) As CPMCNTR
                                                If tmPcf.sType = "H" Then
                                                    ilFound = False
                                                    For ilPkg = 0 To UBound(tgCPMCntr) - 1 Step 1
                                                        If (tmPcf.iPkCPMID = tgCPMCntr(ilPkg).tPcf.iPodCPMID) And (tmPcf.lChfCode = tgCPMCntr(ilPkg).tPcf.lChfCode) Then
                                                            ilFound = True
                                                            Exit For
                                                        End If
                                                    Next ilPkg
                                                    If Not ilFound Then
                                                        tmPcfSrchKey1.lChfCode = tmPcf.lChfCode
                                                        tmPcfSrchKey1.iPodCPMID = tmPcf.iPkCPMID
                                                        tmPcfSrchKey1.iCntRevNo = 0
                                                        tmPcfSrchKey1.iPropVer = 0
                                                        ilRet = btrGetGreaterOrEqual(hmPcf, tlPcf, imPcfRecLen, tmPcfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                                        Do While (ilRet = BTRV_ERR_NONE)
                                                            If (tlPcf.lChfCode = tmPcf.lChfCode) And (tlPcf.iPodCPMID = tmPcf.iPkCPMID) Then
                                                                LSet tgCPMCntr(UBound(tgCPMCntr)) = tgCPMCntr(UBound(tgCPMCntr) - 1)
                                                                tgCPMCntr(UBound(tgCPMCntr)).sKey = slKey & slInvNo & slDate & "B"
                                                                tgCPMCntr(UBound(tgCPMCntr)).tPcf = tlPcf
                                                                tlIbf.lCode = 0
                                                                tlIbf.sBilled = "Y"
                                                                tlIbf.sBillCycle = tgChfInv.sBillCycle
                                                                tgCPMCntr(UBound(tgCPMCntr)).tIbf = tlIbf
                                                                ReDim Preserve tgCPMCntr(0 To UBound(tgCPMCntr) + 1) As CPMCNTR
                                                                Exit Do
                                                            End If
                                                            ilRet = btrGetNext(hmPcf, tmPcf, imPcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        Loop
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            ilRet = btrExtGetNext(hlFile, tmRvf, ilExtLen, llRecPos)
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hlFile, tmRvf, ilExtLen, llRecPos)
                            Loop
                        Loop
                    End If
                End If
            Next ilPass
        End If
    Next ilBillCycle
    mCPM_BuildReprint = BTRV_ERR_NONE
End Function

Public Function mDetermineRemainingMonths(slEndDate As String, slBillCycle As String) As Integer
    Dim ilMonth As Integer
    Dim ilYear As Integer
    Dim ilAgePeriod As Integer
    Dim ilAgingYear As Integer
    
    mDetermineRemainingMonths = 0
    If Invoice.ckcBillCycle(INVBILLCYCLE_Calendar).Value = vbChecked And slBillCycle = "C" Then   'Cal
        ilAgePeriod = Month(gDateValue(smEndCal))
        ilAgingYear = Year(gDateValue(smEndCal))
        gObtainMonthYear 1, slEndDate, ilMonth, ilYear
    ElseIf Invoice.ckcBillCycle(INVBILLCYCLE_Week).Value = vbChecked And slBillCycle = "W" Then   'Week
        Exit Function
    Else 'Std
        ilAgePeriod = Month(gDateValue(smEndStd))
        ilAgingYear = Year(gDateValue(smEndStd))
        gObtainMonthYear 0, slEndDate, ilMonth, ilYear
    End If
    If ilAgingYear = ilYear Then
        mDetermineRemainingMonths = ilMonth - ilAgePeriod + 1
    Else
        mDetermineRemainingMonths = (13 - ilAgePeriod) + ilMonth + 12 * (ilYear - ilAgingYear - 1)
    End If
End Function

'TTP 10827 - Boostr Phase 2: change flat rate invoice method, add support for digital line ad length
Public Function mDeterminePeriodAmountByDaily(slLineStartDate As String, slLineEndDate As String, slBillCycle As String, slRemainingAmount As String) As Double
    Dim dlDailyAmount As Double 'The daily $ Amount
    Dim ilNumberOfDaysRemaining As Integer 'How many days from Invoice Start Date to Line EndDate
    Dim ilNumberOfDaysBeingInvoiced As Integer 'How many days are being invoiced
    Dim dStartDate As Date 'Temp Start Date
    Dim dEndDate As Date 'Temp End
    Dim bLastPeriod As Boolean 'If this is the last period to be invoiced
    
    If Invoice.ckcBillCycle(INVBILLCYCLE_Calendar).Value = vbChecked And slBillCycle = "C" Then   'Cal
        'Determine how many days remain of this line (beyond what's been billed)
        dStartDate = IIF(DateValue(slLineStartDate) > gDateValue(smStartCal), gDateValue(slLineStartDate), gDateValue(smStartCal))
        dEndDate = IIF(DateValue(slLineEndDate) > gDateValue(smEndCal), gDateValue(slLineEndDate), gDateValue(smEndCal))
        ilNumberOfDaysRemaining = DateDiff("d", dStartDate, dEndDate) + 1
        
        'Determine how many days of this Line are being invoiced
        dStartDate = IIF(DateValue(slLineStartDate) > gDateValue(smStartCal), gDateValue(slLineStartDate), gDateValue(smStartCal))
        dEndDate = IIF(DateValue(slLineEndDate) < gDateValue(smEndCal), gDateValue(slLineEndDate), gDateValue(smEndCal))
        ilNumberOfDaysBeingInvoiced = DateDiff("d", dStartDate, dEndDate) + 1
        
        'Determine if this is the last invoice
        bLastPeriod = IIF(DateValue(slLineEndDate) <= gDateValue(smEndCal), True, False)
        
    ElseIf Invoice.ckcBillCycle(INVBILLCYCLE_Week).Value = vbChecked And slBillCycle = "W" Then   'Week
        Exit Function
        
    Else 'Std
        'Determine how many days remain of this line (beyond what's been billed)
        dStartDate = IIF(DateValue(slLineStartDate) > gDateValue(smStartStd), gDateValue(slLineStartDate), gDateValue(smStartStd))
        dEndDate = IIF(DateValue(slLineEndDate) > gDateValue(smEndStd), gDateValue(slLineEndDate), gDateValue(smEndStd))
        ilNumberOfDaysRemaining = DateDiff("d", dStartDate, dEndDate) + 1
        
        'Determine how many days of this Line are being invoiced
        dStartDate = IIF(DateValue(slLineStartDate) > gDateValue(smStartStd), gDateValue(slLineStartDate), gDateValue(smStartStd))
        dEndDate = IIF(DateValue(slLineEndDate) < gDateValue(smEndStd), gDateValue(slLineEndDate), gDateValue(smEndStd))
        ilNumberOfDaysBeingInvoiced = DateDiff("d", dStartDate, dEndDate) + 1
        
        'Determine if this is the last invoice
        bLastPeriod = IIF(DateValue(slLineEndDate) <= gDateValue(smEndStd), True, False)
    End If
    
    If bLastPeriod Then
        ilNumberOfDaysRemaining = ilNumberOfDaysBeingInvoiced
    End If
    dlDailyAmount = Val(slRemainingAmount) / ilNumberOfDaysRemaining
    mDeterminePeriodAmountByDaily = dlDailyAmount * ilNumberOfDaysBeingInvoiced
    
Debug.Print "mDeterminePeriodAmountByDaily: "
If bLastPeriod Then Debug.Print " -> Final Invoice"
Debug.Print " -> slRemainingAmount: " & slRemainingAmount
Debug.Print " -> ilNumberOfDaysRemaining: " & ilNumberOfDaysRemaining
Debug.Print " -> ilNumberOfDaysBeingInvoiced: " & ilNumberOfDaysBeingInvoiced
Debug.Print " -> dlDailyAmount: " & dlDailyAmount
Debug.Print " -> Month Amount: " & mDeterminePeriodAmountByDaily
End Function

Private Sub mAddIbf(tlChf As CHF)
    Dim ilAgePeriod As Integer
    Dim ilAgingYear As Integer
    Dim slSQLQuery As String
    Dim llRet As Long
    
    tmIbf.lCode = 0
    tmIbf.lCntrNo = tlChf.lCntrNo
    tmIbf.iPodCPMID = tmPcf.iPodCPMID
    tmIbf.iVefCode = tmPcf.iVefCode
    tmIbf.sBillCycle = tlChf.sBillCycle
    If Invoice.ckcBillCycle(INVBILLCYCLE_Calendar).Value = vbChecked And tlChf.sBillCycle = "C" Then   'Cal
        ilAgePeriod = Month(gDateValue(smEndCal))
        ilAgingYear = Year(gDateValue(smEndCal))
    ElseIf Invoice.ckcBillCycle(INVBILLCYCLE_Week).Value = vbChecked And tlChf.sBillCycle = "W" Then   'Week
        tmIbf.lCode = 0
        Exit Sub
    ElseIf Invoice.ckcBillCycle(INVBILLCYCLE_STD).Value = vbChecked And tlChf.sBillCycle = "S" Then
        ilAgePeriod = Month(gDateValue(smEndStd))
        ilAgingYear = Year(gDateValue(smEndStd))
    End If
    tmIbf.iBillYear = ilAgingYear
    tmIbf.iBillMonth = ilAgePeriod
    tmIbf.lImpressions = 0
    tmIbf.sBilled = "N"
    tmIbf.iUrfCode = tgUrf(0).iCode
    tmIbf.sSource = "I"
    tmIbf.sUnused = ""

    slSQLQuery = "Insert Into ibf_Impression_Bill ( "
    slSQLQuery = slSQLQuery & "ibfCode, "
    slSQLQuery = slSQLQuery & "ibfCntrNo, "
    slSQLQuery = slSQLQuery & "ibfPodCPMID, "
    slSQLQuery = slSQLQuery & "ibfVefCode, "
    slSQLQuery = slSQLQuery & "ibfBillCycle, "
    slSQLQuery = slSQLQuery & "ibfBillYear, "
    slSQLQuery = slSQLQuery & "ibfBillMonth, "
    slSQLQuery = slSQLQuery & "ibfImpressions, "
    slSQLQuery = slSQLQuery & "ibfBilled, "
    slSQLQuery = slSQLQuery & "ibfUrfCode, "
    slSQLQuery = slSQLQuery & "ibfSource, "
    slSQLQuery = slSQLQuery & "ibfUnused "
    slSQLQuery = slSQLQuery & ") "
    slSQLQuery = slSQLQuery & "Values ( "
    slSQLQuery = slSQLQuery & "Replace" & ", "
    slSQLQuery = slSQLQuery & tmIbf.lCntrNo & ", "
    slSQLQuery = slSQLQuery & tmIbf.iPodCPMID & ", "
    slSQLQuery = slSQLQuery & tmIbf.iVefCode & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tmIbf.sBillCycle) & "', "
    slSQLQuery = slSQLQuery & tmIbf.iBillYear & ", "
    slSQLQuery = slSQLQuery & tmIbf.iBillMonth & ", "
    slSQLQuery = slSQLQuery & tmIbf.lImpressions & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tmIbf.sBilled) & "', "
    slSQLQuery = slSQLQuery & tmIbf.iUrfCode & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tmIbf.sSource) & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(tmIbf.sUnused) & "' "
    slSQLQuery = slSQLQuery & ") "
    llRet = gInsertAndReturnCode(slSQLQuery, "ibf_Impression_Bill", "ibfCode", "Replace")
    If llRet <= 0 Then
        tmIbf.lCode = 0
        Exit Sub
    Else
        tmIbf.lCode = llRet
    End If
End Sub

Private Sub mCheckForOverDelivery()
    Dim ilHidden As Integer
    Dim llImpressions As Long
    Dim llTotalBilledImpressions As Long
    Dim llTotalImpressions As Long
    Dim slSQLQuery As String
    Dim hlOver As Integer
    
    gLogMsgWODT "O", hlOver, sgDBPath & "Messages\" & "ContractOverDelivered_" & Format(Now, "mmddyyyy") & ".txt"
    gLogMsgWODT "W", hlOver, "Contracts Over-Delivered Digitial Impression as of " & Format(Now, "m/d/yy") & " " & Format(Now, "h:mm:ssAM/PM")
    For ilHidden = 0 To UBound(tgCPMCntr) - 1 Step 1
        tmPcf = tgCPMCntr(ilHidden).tPcf
        tmIbf = tgCPMCntr(ilHidden).tIbf
        If tmPcf.sType <> "P" Then
            If (Asc(tgSaf(0).sFeatures7) And PODBILLOVERDELIVERED) <> PODBILLOVERDELIVERED Then 'PODBILLOVERDELIVERED Disabled
                If tmPcf.sPriceType <> "F" Then 'PriceType NOT Flat Rate (CPM)
                    If (Not Invoice.rbcType(INVGEN_Reprint).Value) And (Not Invoice.rbcType(INVGEN_Archive).Value) Then  'NOT Reprint and NOT Archive
                        slSQLQuery = "Select Sum(ibfBilledImpression) as TotalBilledImpressions from ibf_Impression_Bill Where ibfBilled = " & "'Y'" & " And ibfCntrNo = " & tmIbf.lCntrNo & " And ibfPodCPMID = " & tmPcf.iPodCPMID
                        Set ibf_rst = gSQLSelectCall(slSQLQuery)
                        If Not ibf_rst.EOF And Not IsNull(ibf_rst!TotalBilledImpressions) Then
                            llTotalBilledImpressions = ibf_rst!TotalBilledImpressions
                        Else
                            llTotalBilledImpressions = 0
                        End If
                        llTotalImpressions = tmIbf.lImpressions + llTotalBilledImpressions
                        If llTotalImpressions > tmPcf.lImpressionGoal Then
                            llImpressions = tmPcf.lImpressionGoal - llTotalBilledImpressions
                            If llImpressions < 0 Then
                                llImpressions = 0
                            End If
                            gLogMsgWODT "W", hlOver, "Contracts " & tgCPMCntr(ilHidden).tIbf.lCntrNo & " ID " & tgCPMCntr(ilHidden).tIbf.iPodCPMID & " Total Impressions " & llTotalImpressions & " Impression Goal " & tgCPMCntr(ilHidden).tPcf.lImpressionGoal
                        Else
                            llImpressions = tmIbf.lImpressions
                        End If
                        tmIbf.lBilledImpression = llImpressions
                    Else    'Reprint or archive
                        If tmIbf.lBilledImpression <= 0 Then
                            tmIbf.lBilledImpression = tmIbf.lImpressions
                        End If
                    End If
                Else    'PriceType = Flat Rate
                    'TTP 10479 - V81 ONLY: Flat rate ad server lines: after running prelim invoices, if the posted impressions for the month are changed in Post Log, then final invoices are run, it uses the old posted impression count, not the new one
                    'If tmIbf.lBilledImpression <= 0 Then
                    tmIbf.lBilledImpression = tmIbf.lImpressions
                    'End If
                End If
            Else    'PODBILLOVERDELIVERED Enabled
                'TTP 10479 - V81 ONLY:  - Flat rate ad server lines: after running prelim invoices, if the posted impressions for the month are changed in Post Log, then final invoices are run, it uses the old posted impression count, not the new one
                'If tmIbf.lBilledImpression <= 0 Then
                tmIbf.lBilledImpression = tmIbf.lImpressions
                'End If
            End If
            LSet tgCPMCntr(ilHidden).tIbf = tmIbf
        End If
    Next ilHidden
    gLogMsgWODT "C", hlOver, ""
End Sub

'Boostr Phase 2 issues - for Joel - Issue 15, use the latest PCF Code. It should show the ad length from the line. This way the updated ad lengths for the lines that didn't originally have ad lengths will show ad lengths.
Function mGetLatestPCFCode(lCntrNo As Long, iPodCPMID As Integer) As Long
    Dim slSQLQuery As String
    Dim rst As ADODB.Recordset
    mGetLatestPCFCode = 0
    
    slSQLQuery = "SELECT TOP 1 pcfCode "
    slSQLQuery = slSQLQuery & "FROM pcf_Pod_CPM_Cntr "
    'TTP 10996 - Invoice reprint: SQL error when attempting to reprint digital invoice that requires scheduling
    'slSQLQuery = slSQLQuery & "WHERE pcfChfCode = "
    slSQLQuery = slSQLQuery & "WHERE pcfChfCode in "
    slSQLQuery = slSQLQuery & "("
    slSQLQuery = slSQLQuery & " SELECT chfCode "
    slSQLQuery = slSQLQuery & " FROM CHF_Contract_Header "
    slSQLQuery = slSQLQuery & " WHERE chfCntrNo = " & lCntrNo
    slSQLQuery = slSQLQuery & " AND chfDelete <> 'Y' "
    slSQLQuery = slSQLQuery & " AND chfStatus in ('H','O','G','N') "
    slSQLQuery = slSQLQuery & ") "
    slSQLQuery = slSQLQuery & "AND pcfPodCPMID = " & iPodCPMID
    
    Set rst = gSQLSelectCall(slSQLQuery)
    If Not rst.EOF And Not IsNull(rst!pcfCode) Then
        mGetLatestPCFCode = rst!pcfCode
    End If
End Function

