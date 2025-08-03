Attribute VB_Name = "RPTSLCOM"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptslcom.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Type SLFINDEX
    iSlfCode As Integer        'slsp #
    'iStartIndex As Integer     'starting index of slsp comm entries by vehicle
    'iEndIndex As Integer       'ending index of slp comm entries by vehicle
    '8-8-12 SCF table exceeds 32000
    lStartIndex As Long     'starting index of slsp comm entries by vehicle
    lEndIndex As Long       'ending index of slp comm entries by vehicle
End Type
Function gObtainScf(RptForm As Form, hlScf As Integer, tlScf() As SCF, llStartDate As Long, llEndDate As Long, tlSlfIndex() As SLFINDEX) As Integer
'*******************************************************
'*      <input>  hlscf - Salesperson handle (file must
'*                       open
'*               llStartDate - earliest date to include commission
'                       percentages
'                llEnddate -latest date to include comm pcts
'*      <I/O>    tlscf() - array of matching scf recds
'*
'*             Created:3/14/00       By:D. Hosaka
'*
'*            Comments: Read all of salesperson comm.
'*                      records.  Build array of records
'*                      whose start and end dates fall
'                       within the requested period
'*
'*******************************************************
'
'    gObtainscf (hlscf,  tlscf())
'
    Dim ilRet As Integer    'Return status
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilscfUpper As Integer
    Dim llTempStartDate As Long
    Dim llTempEndDate As Long
    Dim tmScf As SCF
    Dim ilscfRecLen As Integer
    Dim ilSlsp As Integer
    Dim ilOneFound As Integer
    Dim ilLowLimitScf As Integer
    Dim ilLowLimitSlf As Integer

    On Error GoTo gObtainScfErr
    ilRet = 0
    ilLowLimitScf = LBound(tlScf)
    If ilRet <> 0 Then
        ilLowLimitScf = 1
    End If
    ilRet = 0
    ilLowLimitSlf = LBound(tlSlfIndex)
    If ilRet <> 0 Then
        ilLowLimitSlf = 1
    End If
    On Error GoTo 0

    'llMinDate = gDateValue(slStartDate)     'convert the string dates to long forcomparisons
    'llMaxDate = gDateValue(slEndDate)       'convert the string latest date to long for comparison
    ReDim tlSlfIndex(ilLowLimitSlf To ilLowLimitSlf) As SLFINDEX     'table of slsp commission by vehicle
    ilSlsp = -1

    ReDim tlScf(ilLowLimitScf To ilLowLimitScf) As SCF
    btrExtClear hlScf   'Clear any previous extend operation
    ilExtLen = Len(tlScf(ilLowLimitScf))  'Extract operation record size
    ilscfRecLen = Len(tmScf)
    ilscfUpper = UBound(tlScf)
    'ilRet = btrGetFirst(hlscf, tlscf(0), ilscfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ilRet = btrGetFirst(hlScf, tmScf, ilscfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hlScf, llNoRec, -1, "UC", "SCF", "") '"EG") 'Set extract limits (all records)
        ilRet = btrExtAddField(hlScf, 0, ilExtLen)  'Extract the whole record
        On Error GoTo mScfErr
        gBtrvErrorMsg ilRet, "gObtainscf (btrExtAddField):" & "scf.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlScf, tmScf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mScfErr
            gBtrvErrorMsg ilRet, "gObtainscf (btrExtGetNextExt):" & "scf.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmScf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlScf, tmScf, ilExtLen, llRecPos)
            Loop
            ilOneFound = False
            Do While ilRet = BTRV_ERR_NONE
                gUnpackDateLong tmScf.iStartDate(0), tmScf.iStartDate(1), llTempStartDate
                gUnpackDateLong tmScf.iEndDate(0), tmScf.iEndDate(1), llTempEndDate
                If llTempEndDate = 0 Then                   'if tfn, set max value for comparison
                    llTempEndDate = gDateValue("12/31/2069")
                End If
                'filter out the commission records that are completely outside the requested range
                If llEndDate > llTempStartDate And llStartDate < llTempEndDate Then
                    tlScf(UBound(tlScf)) = tmScf           'save entire record
                    'build array of the unique slsp entries and their upper & lower bounds to search (speed up search when processing)
                    If ilSlsp <> tmScf.iSlfCode Then
                        ilOneFound = True           'at least one slsp defined with sub-company
                        tlSlfIndex(UBound(tlSlfIndex)).iSlfCode = tmScf.iSlfCode
                        'tlSlfIndex(UBound(tlSlfIndex)).iStartIndex = UBound(tlScf)
                        tlSlfIndex(UBound(tlSlfIndex)).lStartIndex = UBound(tlScf)          '8-8-12 chg to long
                        ReDim Preserve tlSlfIndex(ilLowLimitSlf To UBound(tlSlfIndex) + 1) As SLFINDEX
                        ilSlsp = tmScf.iSlfCode
                    Else
                       ' tlSlfIndex(UBound(tlSlfIndex) - 1).iEndIndex = UBound(tlScf)
                        tlSlfIndex(UBound(tlSlfIndex) - 1).lEndIndex = UBound(tlScf)            '8-8-12 chg to long
                    End If
                    ReDim Preserve tlScf(ilLowLimitScf To UBound(tlScf) + 1) As SCF
                End If
                ilRet = btrExtGetNext(hlScf, tmScf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlScf, tmScf, ilExtLen, llRecPos)
                Loop
            Loop
            If ilOneFound Then          'set the hi record for the last slsp processed
                'tlSlfIndex(UBound(tlSlfIndex) - 1).iEndIndex = UBound(tlScf) - 1
                tlSlfIndex(UBound(tlSlfIndex) - 1).lEndIndex = UBound(tlScf) - 1            '8-8-12 chg to long
            End If
        End If
    End If
    gObtainScf = True
    Exit Function
gObtainScfErr:
    ilRet = 0
    Resume Next
mScfErr:
    On Error GoTo 0
    gObtainScf = False
    Exit Function

End Function
Sub gGetCommByDates(ilSlfCode() As Integer, ilslfcomm() As Integer, ilslfremnant() As Integer, tlSlfIndex() As SLFINDEX, tlScf() As SCF, tlSlf() As SLF, ilVefCode As Integer, llDate As Long, tlChf As CHF)
Dim ilLoop As Integer
'Dim ilLo As Integer
'Dim ilHi As Integer
Dim llLo As Long                '8-8-12 scf table exceeds 32000
Dim llHi As Long
'Dim ilCount As Integer
Dim llCount As Long             '8-8-12 scf table exceeds 32000
'Dim ilWhichSlsp As Integer
Dim llWhichSlsp As Long         '8-8-12 scf table exceeds 32000
'Dim ilVehPct As Integer
Dim llVehPct As Long            '8-8-12 scf table exceeds 32000
Dim llScfStartDate As Long
Dim llScfEndDate As Long

    'for each slsp, look up the default commission percent to use from his slsp recd
    For ilLoop = 0 To 9     'loop on the slsp to process for the sub-co
        If ilSlfCode(ilLoop) > 0 Then
            '8-8-12 all occurences of ilcount changed to long
            'For ilCount = LBound(tlSlf) To UBound(tlSlf)    'find the salespersons record for this default commission
            For llCount = LBound(tlSlf) To UBound(tlSlf)    'find the salespersons record for this default commission
                If ilSlfCode(ilLoop) = tlSlf(llCount).iCode Then
                    ilslfcomm(ilLoop) = tlSlf(llCount).iUnderComm
                    ilslfremnant(ilLoop) = tlSlf(llCount).iRemUnderComm

                    If tgSpf.sCommByCntr = "Y" Then            '7-14-00
                        If tlChf.iSlspCommPct(ilLoop) <> 0 Then
                            ilslfcomm(ilLoop) = CLng(tlChf.iSlspCommPct(ilLoop))
                        End If
                    End If
                        'now see if there is a comm percentage defined for this vehicle
                        'For ilWhichSlsp = 1 To UBound(tlSlfIndex)      '8-8-12 all occurences of ilwhichslsp chged to long
                        For llWhichSlsp = LBound(tlSlfIndex) To UBound(tlSlfIndex)
                            If ilSlfCode(ilLoop) = tlSlfIndex(llWhichSlsp).iSlfCode Then
                                'ilLo = tlSlfIndex(llWhichSlsp).iStartIndex
                                'ilHi = tlSlfIndex(llWhichSlsp).iEndIndex
                                llLo = tlSlfIndex(llWhichSlsp).lStartIndex
                                llHi = tlSlfIndex(llWhichSlsp).lEndIndex
                                'For ilVehPct = ilLo To ilHi
                                For llVehPct = llLo To llHi
                                    If tlScf(llVehPct).iVefCode = ilVefCode Then        'test DATES HERE
                                        gUnpackDateLong tlScf(llVehPct).iStartDate(0), tlScf(llVehPct).iStartDate(1), llScfStartDate
                                        gUnpackDateLong tlScf(llVehPct).iEndDate(0), tlScf(llVehPct).iEndDate(1), llScfEndDate
                                        If llScfEndDate = 0 Then
                                            llScfEndDate = 62093            'force to highest date allowed (12/31/2069)
                                        End If
                                        If (llDate >= llScfStartDate And llDate <= llScfEndDate) Then
                                            ilslfcomm(ilLoop) = tlScf(llVehPct).iUnderComm
                                            ilslfremnant(ilLoop) = tlScf(llVehPct).iRemUnderComm
                                            Exit For
                                        End If
                                    End If
                                Next llVehPct

                                Exit For
                            End If
                        Next llWhichSlsp    'for llWhichSlsp = 1 to UBound(tlScfIndex)
                    Exit For
                End If
            Next llCount                'for llCount = lBound(tlSlf) to uBound(tlSlf)
        Else
            Exit For
        End If
    Next ilLoop                         'for ilLoop = 0 to 9

End Sub
'
'           gGetSubCmpy -
'
'       2-2-04 For clients using commission sub-companies, they made have slsp comm over 100%, and no
'               revenue share % defined in the header.  If that case, on certain reports the slsp comm
'               has to be computed but the gross cant be included in totals.  Need to maintain the
'               slsp Rev splits as well as the comm splits to compare
'               These percentages are used to calculate what % of the $ should be given to a slsp to calc the slsp % from.
'               This is basically the slsp gross share
'
'               contract header using sub-companies:
'                                         Slsp       SubCo        Revshare%         commshare%    (use chfslscommPct)
'                                         these columns are shown in contract header screen
'                                         Joe A      AAA          100%         100%
'                                         Joe A      BBB           75%          75%
'                                         Joe B      BBB           25%          25%
'                                         Joe C      BBB            0           100%
'               if not using subcompany, but A/E commissions (slsp)  (use chfcomm)
'                                       these columns shown in contract header screen
'                                        Slsp       Rev/Comm Share%
'                                        Joe A      70%
'                                        Joe B      30%
'               is using comm by contract,
'                                       Slsp       Comm        Rev/Comm Share    (use chfcomm to get slsp gross share, then use chfslscommPct to get slsp Comm % for comm earned)
'                                                  chfslscomm  chfcomm
'                                       Joe A       70%        75%
'                                       Joe B       30%        25%
Function gGetSubCmpy(tlChf As CHF, ilSlfCode() As Integer, llSlfSplit() As Long, ilVefCode As Integer, ilUseSlsComm As Integer, llSlfSplitRev() As Long) As Integer
Dim ilLoop As Integer
Dim ilMnfSubCo As Integer
Dim ilCount As Integer
Dim ilSubCmpyDefined As Integer
    'find the sub-company if used, based on the vehicle to process
    ilMnfSubCo = 0
    If tgSpf.sSubCompany = "Y" Then     '4-20-00
        'For ilLoop = LBound(tgMVef) To UBound(tgMVef)
        '    If tgMVef(ilLoop).iCode = ilVefCode Then
            ilLoop = gBinarySearchVef(ilVefCode)
            If ilLoop <> -1 Then
                ilMnfSubCo = tgMVef(ilLoop).iMnfVehGp6Sub
        '        Exit For
            End If
        'Next ilLoop
    End If

    If ilUseSlsComm Then            'one of the commission reports
        'if nonzero, found a subcompany to use.  Find only those slsp that should use
        'the matching sub-company
        If ilMnfSubCo > 0 Then
            ilCount = 0
            ilSubCmpyDefined = 0
            For ilLoop = 0 To 9
                If tlChf.iMnfSubCmpy(ilLoop) <> 0 Then  '4-13-00
                    ilSubCmpyDefined = 1
                End If
                If ilMnfSubCo = tlChf.iMnfSubCmpy(ilLoop) Then
                    ilSlfCode(ilCount) = tlChf.iSlfCode(ilLoop)
                    llSlfSplit(ilCount) = CLng(tlChf.iSlspCommPct(ilLoop)) * 100
                    llSlfSplitRev(ilCount) = tlChf.lComm(ilLoop)            '2-2-04
                    ilCount = ilCount + 1
                End If
            Next ilLoop
            If ilCount = 0 Then
                For ilLoop = 0 To 9
                    llSlfSplit(ilLoop) = CLng(tlChf.iSlspCommPct(ilLoop)) * 100
                    'llSlfSplitRev(ilCount) = tlChf.lComm(ilLoop)            '2-2-04
                    llSlfSplitRev(ilLoop) = tlChf.lComm(ilLoop)         '1-17-13 wrong index used
                    ilSlfCode(ilLoop) = tlChf.iSlfCode(ilLoop)
                Next ilLoop
            End If
        Else            'no subcompany, use the percentage stored in header
            For ilLoop = 0 To 9
                llSlfSplit(ilLoop) = CLng(tlChf.iSlspCommPct(ilLoop)) * 100
                llSlfSplitRev(ilLoop) = tlChf.lComm(ilLoop)            '2-2-04
                ilSlfCode(ilLoop) = tlChf.iSlfCode(ilLoop)
            Next ilLoop
        End If
    Else                                'other than commission, use the normal revenue splits
        'if nonzero, found a subcompany to use.  Find only those slsp that should use
        'the matching sub-company
        If ilMnfSubCo > 0 Then
            ilCount = 0
            ilSubCmpyDefined = 0

            For ilLoop = 0 To 9
                If tlChf.iMnfSubCmpy(ilLoop) <> 0 Then  '4-13-00
                    ilSubCmpyDefined = 1
                End If

                If ilMnfSubCo = tlChf.iMnfSubCmpy(ilLoop) Then
                    ilSlfCode(ilCount) = tlChf.iSlfCode(ilLoop)
                    llSlfSplit(ilCount) = tlChf.lComm(ilLoop)
                    llSlfSplitRev(ilCount) = tlChf.lComm(ilLoop)            '2-2-04
                    ilCount = ilCount + 1
                End If
            Next ilLoop
            If ilCount = 0 Then
                For ilLoop = 0 To 9
                    llSlfSplit(ilLoop) = tlChf.lComm(ilLoop)
                    llSlfSplitRev(ilLoop) = tlChf.lComm(ilLoop)            '2-2-04
                    ilSlfCode(ilLoop) = tlChf.iSlfCode(ilLoop)
                Next ilLoop
            Else
                If tlChf.lComm(0) = 1000000 And tlChf.iSlfCode(0) > 0 Then                     'no splits
                    llSlfSplit(0) = tlChf.lComm(0)
                    llSlfSplitRev(0) = tlChf.lComm(0)
                    ilSlfCode(0) = tlChf.iSlfCode(0)
                End If
            End If
        Else            'no subcompany, use the percentage stored in header
            ilCount = 0
            For ilLoop = 0 To 9
                llSlfSplit(ilLoop) = tlChf.lComm(ilLoop)
                llSlfSplitRev(ilLoop) = tlChf.lComm(ilLoop)            '2-2-04
                ilSlfCode(ilLoop) = tlChf.iSlfCode(ilLoop)
                If tlChf.iSlfCode(ilLoop) > 0 Then
                    ilCount = ilCount + 1
                End If
            Next ilLoop
            If ilCount = 1 And tlChf.lComm(0) = 0 Then          'got to have at least 1 slsp comm pct
                llSlfSplit(0) = 1000000
                llSlfSplitRev(0) = 1000000
            End If
        End If
    End If
    If tgSpf.sSubCompany = "Y" Then     'if using subcompanies, see how they are defined for this contract
        If ilMnfSubCo > 0 Then
            If ilCount = 0 And ilSubCmpyDefined > 0 Then    'subcompany defined for vehicle, but it didnot find any that matched in header;
                                                            'but there are other subcompanies defined with the slsp
                ilMnfSubCo = -1                             'flag error for report
            End If
        Else
            'vehicle didnt have a subcompany defined, if none of the slsp in header have subcompany defined, its OK.
            'if at least one has it defined, flag as error on report
            For ilLoop = 0 To 9
                If tlChf.iMnfSubCmpy(ilLoop) <> 0 Then      '4-13-00
                    ilMnfSubCo = -1
                End If
            Next ilLoop
        End If
    End If
    gGetSubCmpy = ilMnfSubCo
End Function
