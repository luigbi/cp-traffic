Attribute VB_Name = "POPSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of POPSUBS.BAS on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  imAdvtCode                                                                            *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  gObtainLatestRcf              gPopAASBoxViaCntr             gObtainPIFDates_ForVef    *
'*                                                                                        *
'******************************************************************************************

' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PopSubs.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Population subs and functions
Option Explicit
Option Compare Text


Dim tmVsf As VSF

Dim tmLRecBuff As LPOPREC
Dim tmIRecBuff As IPOPREC

'Comment record-Header/Line
Dim tmCxf As CXF            'CXF record image
Dim tmCxfSrchKey As LONGKEY0  'CXF key record image
Dim hmCxf As Integer        'CXF Handle
Dim imCxfRecLen As Integer      'CXF record length
'Advertiser
Dim tmAdf As ADF            'ADF record image
Dim tmAdfSrchKey As INTKEY0  'ADF key record image
Dim hmAdf As Integer        'ADF Handle
Dim imAdfRecLen As Integer      'ADF record length
'Agency
Dim tmAgf As AGF            'AGF record image
Dim tmAgfSrchKey As INTKEY0  'AGF key record image
Dim hmAgf As Integer        'AGF Handle
Dim imAgfRecLen As Integer      'AGF record length
'Vehicle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0  'VEF key record image
Dim hmVef As Integer        'VEF Handle
Dim imVefRecLen As Integer      'VEF record length
'Vehicle Preferences
Dim tmVpf As VPF            'VPF record image
Dim tmVpfSrchKey As INTKEY0  'VPF key record image
Dim hmVpf As Integer        'VPF Handle
Dim imVpfRecLen As Integer      'VPF record length

Dim smSdfExtStamp As String   'gObtainCntrSpot Sort selection

Dim lmWCINGCntrNo() As Long

Dim imCntrSpotsVefCode() As Integer
Dim lmInvStartDate As Long 'gPopCntrForAASBox when ilCurrent = 4; Start date stored into tag field.
Dim lmInvEndDate As Long
'11/26/17: Add parameters used to determine if populate is required
Public lgShttLastDateChgd As Long
Public lgShttLastTimeChgd As Long
Type FCTCHGDINFO
    lLastDateChgd As Long
    lLastTimeChgd As Long
End Type
Public tgFctChgdInfo(0 To 4) As FCTCHGDINFO
Const SHTTINDEX = 0
Const VEFINDEX = 1
Const VPFINDEX = 2
Const CPTTINDEX = 3
Const ADFINDEX = 4
Public fct_rst As ADODB.Recordset
'8886
Public Const FILEEXISTS = 0
Public Const FILEEXISTSNOT = 1
Type VEFNAME
    sName As String * 40
    iCode As Integer
End Type
Public tgVefName() As VEFNAME
Public Type BOOKNAMEINFO
    iDnfCode As Integer
    sName As String * 30
    sBookDate As String * 12
End Type
Public tgBookNameInfo() As BOOKNAMEINFO

Type PODITEMCATEGORY
    sName As String * 20
    iCode As Integer
End Type
Public tgPodCategoryItems() As PODITEMCATEGORY

Type PODITEMS
    ItemName As String * 20
    iCode As Integer
    iMnfCode As Integer
End Type

Public tgPodItems() As PODITEMS

'************************************************************************
'*      Procedure Name:gPopCategoryItems
'*      Created: 12/21/2020     By: L. Bianchi
'*      Modified:              By:
'*      Comments: populate dropdown with Pod-Category Items
'************************************************************************
Public Function gPopCategoryItems() As Integer
    Dim llUpper As Long
    Dim rst As ADODB.Recordset
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim slStamp As String
    Dim llMax As Long

    slStamp = gFileDateTime(sgDBPath & "mnf.Mkd")
    If Not gFileChgd("mnf.mkd") Then
        gPopCategoryItems = True
        Exit Function
    End If
    
    SQLQuery = "SELECT mnfCode, mnfType,mnfName FROM MNF_Multi_Names WHERE mnfType = '5'"
    Set rst = gSQLSelectCall(SQLQuery)
    llUpper = 0
    llMax = 100
    ReDim tgPodCategoryItems(0 To llMax) As PODITEMCATEGORY
    
    While Not rst.EOF
        tgPodCategoryItems(llUpper).iCode = rst!mnfCode
        tgPodCategoryItems(llUpper).sName = Trim$(rst!mnfName)
        llUpper = llUpper + 1
        If llUpper >= llMax Then
            llMax = llMax + 100
            ReDim Preserve tgPodCategoryItems(0 To llMax) As PODITEMCATEGORY
        End If
        rst.MoveNext
    Wend
    ReDim Preserve tgPodCategoryItems(0 To llUpper) As PODITEMCATEGORY
    gFileChgdUpdate "mnf.mkd", False
    
    gPopCategoryItems = True
    rst.Close
End Function

'************************************************************************
'*      Procedure Name:gPodItems
'*      Created: 12/21/2020     By:L. Bianchi
'*      Modified:              By:
'*      Comments: Load list of Pod-Item for a given category
'************************************************************************
Public Function gPodItems(ItemCategory As Integer) As Integer
    Dim llUpper As Long
    Dim rst As ADODB.Recordset
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim slStamp As String
    Dim llMax As Long

    slStamp = gFileDateTime(sgDBPath & "tif.Mkd")
    If Not gFileChgd("tif.mkd") Then
        gPodItems = True
        Exit Function
    End If
    
    SQLQuery = "SELECT thfCode, thfName from thf_Target_Header WHERE thfCategoryMnfCode = " & ItemCategory & " order by thfName"
    Set rst = gSQLSelectCall(SQLQuery)
    llUpper = 0
    llMax = 100
    ReDim tgPodItems(0 To llMax) As PODITEMS
    
    While Not rst.EOF
        tgPodItems(llUpper).iCode = rst!thfCode
        tgPodItems(llUpper).ItemName = Trim$(rst!thfName)
        llUpper = llUpper + 1
        If llUpper >= llMax Then
            llMax = llMax + 100
            ReDim Preserve tgPodItems(0 To llMax) As PODITEMS
        End If
        rst.MoveNext
    Wend
    ReDim Preserve tgPodItems(0 To llUpper) As PODITEMS
    gFileChgdUpdate "tif.mkd", False
    
    gPodItems = True
    rst.Close
End Function

'****************************************************************************************
'*
'
'           gPopITFBox - populate list box with multimedia type defintions
'           <input> FRM - form calling the populuate
'                   lbcLocal - list box to contain the media definitions
'                   tlSortCode() - array of media defintion keys containing the desc & internal code
Public Sub gPopITFBox(frm As Form, lbcLocal As Control, tlITF() As ITF, tlSortCode() As SORTCODE)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************
    Dim ilSortCode As Integer
    Dim illoop As Integer
    Dim slName As String

    lbcLocal.Clear
    ilSortCode = 0
    ReDim tlSortCode(0 To UBound(tlITF) - LBound(tlITF)) As SORTCODE
    For illoop = LBound(tlITF) To UBound(tlITF) - 1 Step 1
        slName = tlITF(illoop).sName & "\" & Trim$(str$(tlITF(illoop).iCode))
        tlSortCode(ilSortCode).sKey = slName
        lbcLocal.AddItem tlITF(illoop).sName  'add name to list box

        ilSortCode = ilSortCode + 1
    Next illoop
    If UBound(tlSortCode) - 1 > 0 Then
        ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
    End If
    Exit Sub
End Sub

'           gPopMCFBox - populate list box with media code defintions
'           <input> FRM - form calling the populuate
'                   lbcLocal - list box to contain the media definitions
'                   tlSortCode() - array of media defintion keys containing the desc & internal code
'                   slSortCodeTag - string to determine if file should be repopulated
'           return - true = populate OK
Public Function gPopMCFBox(frm As Form, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
    Dim ilRet As Integer
    Dim ilSortCode As Integer
    Dim illoop As Integer
    Dim slName As String

    gPopMCFBox = True
    ilRet = gObtainMCF()
    If ilRet = False Then
        gPopMCFBox = False
        Exit Function
    End If
    ilSortCode = 0
    ReDim tlSortCode(0 To UBound(tgMCF) - LBound(tgMCF)) As SORTCODE
    For illoop = LBound(tgMCF) To UBound(tgMCF) - 1 Step 1
        slName = tgMCF(illoop).sName & "\" & Trim$(str$(tgMCF(illoop).iCode))
        tlSortCode(ilSortCode).sKey = slName
        lbcLocal.AddItem tgMCF(illoop).sName  'add name to list box

        ilSortCode = ilSortCode + 1
    Next illoop
    If UBound(tlSortCode) - 1 > 0 Then
        ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
    End If
    Exit Function
End Function

'           Read the media Definitions table (MCF) and return in
'           Global array TGMCF()
'           Return - true if read OK
Public Function gObtainMCF() As Integer
'
'   ilRet = gObtainMCF ()
'   Where:
'       tgMCF() MCF record structure
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlMcf As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlMcf As MCF
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer

    slStamp = gFileDateTime(sgDBPath & "Mcf.Btr")

    On Error GoTo gObtainMCFErr2
    ilRet = 0
    If ilRet <> 0 Then
        sgMcfTagRpt = ""
    End If
    On Error GoTo 0

    If sgMcfTagRpt <> "" Then
        If StrComp(slStamp, sgMcfTagRpt, 1) = 0 Then
            'If UBound(tgMCF) > 1 Then
                gObtainMCF = True
                Exit Function
            'End If
        End If
    End If
    'ReDim tgMCF(1 To 1) As MCF
    ReDim tgMCF(0 To 0) As MCF
    hlMcf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainMCF = False
        ilRet = btrClose(hlMcf)
        btrDestroy hlMcf
        Exit Function
    End If
    ilRecLen = Len(tlMcf) 'btrRecordLength(hlMcf)  'Get and save record length
    sgMcfTagRpt = slStamp
    ilUpperBound = UBound(tgMCF)
    ilExtLen = Len(tgMCF(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlMcf) 'Obtain number of records
    btrExtClear hlMcf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlMcf, tlMcf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlMcf)
        btrDestroy hlMcf
        gObtainMCF = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainMCF = False
            ilRet = btrClose(hlMcf)
            btrDestroy hlMcf
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlMcf, llNoRec, -1, "UC", "MCF", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlMcf, ilOffSet, ilRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainMCF = False
        ilRet = btrClose(hlMcf)
        btrDestroy hlMcf
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlMcf)    'Extract record
    ilUpperBound = UBound(tgMCF)
    ilRet = btrExtGetNext(hlMcf, tgMCF(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainMCF = False
            ilRet = btrClose(hlMcf)
            btrDestroy hlMcf
            Exit Function
        End If
        ilUpperBound = UBound(tgMCF)
        ilExtLen = Len(tgMCF(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlMcf, tgMCF(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlMcf, tgMCF(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilUpperBound = ilUpperBound + 1
            'ReDim Preserve tgMCF(1 To ilUpperBound) As MCF
            ReDim Preserve tgMCF(0 To ilUpperBound) As MCF
            ilRet = btrExtGetNext(hlMcf, tgMCF(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlMcf, tgMCF(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlMcf)
    btrDestroy hlMcf
    gObtainMCF = True
    Exit Function
gObtainMCFErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainPopDnfInfo               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     Book Name and Date              *
'*                                                     *
'*******************************************************
Function gObtainPopDnfInfo() As Integer
'
'   ilRet = gObtainPopDnfInfo ()
'   Where:
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'Dnf date/time stamp
    Dim hlDnf As Integer        'Dnf handle
    Dim ilDnfRecLen As Integer     'Record length
    Dim tlDnf As DNF
    Dim llNoRec As Long         'Number of records in Sof
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim ilOffSet As Integer
    Dim ilTest As Integer
    Dim ilFound As Integer
    Dim ilLowLimit As Integer

    slStamp = gFileDateTime(sgDBPath & "Dnf.Btr")

    'On Error GoTo gObtainPopDnfInfoErr2
    'ilRet = 0
    'ilLowLimit = LBound(tgPopDnfInfo)
    'If ilRet <> 0 Then
    '    sgPopDnfInfoTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tgPopDnfInfo).Ptr <> 0 Then
        ilLowLimit = LBound(tgPopDnfInfo)
    Else
        sgPopDnfInfoTag = ""
        ilLowLimit = 0
    End If

    If sgPopDnfInfoTag <> "" Then
        If StrComp(slStamp, sgPopDnfInfoTag, 1) = 0 Then
            gObtainPopDnfInfo = CP_MSG_NOPOPREQ
            Exit Function
        End If
    End If
    gObtainPopDnfInfo = CP_MSG_POPREQ
    sgPopDnfInfoTag = slStamp
    'ReDim tgPopDnfInfo(1 To 1) As POPDNFINFO
    ReDim tgPopDnfInfo(ilLowLimit To ilLowLimit) As POPDNFINFO
    hlDnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Function
    End If
    ilDnfRecLen = Len(tlDnf) 'btrRecordLength(hlDnf)  'Get and save record length
    ilExtLen = Len(tlDnf)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlDnf) 'Obtain number of records
    btrExtClear hlDnf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlDnf, tlDnf, ilDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlDnf)
        btrDestroy hlDnf
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hlDnf)
            btrDestroy hlDnf
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlDnf, llNoRec, -1, "UC", "DNF", "") 'Set extract limits (all records including first)
    ilOffSet = 0
    ilRet = btrExtAddField(hlDnf, ilOffSet, ilDnfRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlDnf)
        btrDestroy hlDnf
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlDnf)    'Extract record
    ilRet = btrExtGetNext(hlDnf, tlDnf, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            ilRet = btrClose(hlDnf)
            btrDestroy hlDnf
            Exit Function
        End If
        ilExtLen = Len(tlDnf)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlDnf, tlDnfExt, ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlDnf, tlDnf, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            tgPopDnfInfo(UBound(tgPopDnfInfo)).iDnfCode = tlDnf.iCode
            tgPopDnfInfo(UBound(tgPopDnfInfo)).iPopDnfCode = tlDnf.iPopDnfCode
            tgPopDnfInfo(UBound(tgPopDnfInfo)).iQualPopDnfCode = tlDnf.iQualPopDnfCode
            'ReDim Preserve tgPopDnfInfo(1 To UBound(tgPopDnfInfo) + 1) As POPDNFINFO
            ReDim Preserve tgPopDnfInfo(ilLowLimit To UBound(tgPopDnfInfo) + 1) As POPDNFINFO
            ilRet = btrExtGetNext(hlDnf, tlDnf, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlDnf, tlDnf, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    'Check that reference book exist
    'For ilLoop = 1 To UBound(tgPopDnfInfo) - 1 Step 1
    For illoop = LBound(tgPopDnfInfo) To UBound(tgPopDnfInfo) - 1 Step 1
        ilFound = False
        If tgPopDnfInfo(illoop).iPopDnfCode > 0 Then
            'For ilTest = 1 To UBound(tgPopDnfInfo) - 1 Step 1
            For ilTest = LBound(tgPopDnfInfo) To UBound(tgPopDnfInfo) - 1 Step 1
                If tgPopDnfInfo(illoop).iPopDnfCode = tgPopDnfInfo(ilTest).iDnfCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilTest
        End If
        If Not ilFound Then
            tgPopDnfInfo(illoop).iPopDnfCode = 0
        End If
    Next illoop
    'For ilLoop = 1 To UBound(tgPopDnfInfo) - 1 Step 1
    For illoop = LBound(tgPopDnfInfo) To UBound(tgPopDnfInfo) - 1 Step 1
        ilFound = False
        If tgPopDnfInfo(illoop).iQualPopDnfCode > 0 Then
            'For ilTest = 1 To UBound(tgPopDnfInfo) - 1 Step 1
            For ilTest = LBound(tgPopDnfInfo) To UBound(tgPopDnfInfo) - 1 Step 1
                If tgPopDnfInfo(illoop).iQualPopDnfCode = tgPopDnfInfo(ilTest).iDnfCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilTest
        End If
        If Not ilFound Then
            tgPopDnfInfo(illoop).iQualPopDnfCode = 0
        End If
    Next illoop
    'Remove records without a Pop reference book
    illoop = UBound(tgPopDnfInfo) - 1
    'Do While ilLoop >= 1
    Do While illoop >= LBound(tgPopDnfInfo)
        If (tgPopDnfInfo(illoop).iPopDnfCode <= 0) And (tgPopDnfInfo(illoop).iQualPopDnfCode <= 0) Then
            For ilTest = illoop To UBound(tgPopDnfInfo) - 1 Step 1
                tgPopDnfInfo(illoop) = tgPopDnfInfo(ilTest)
            Next ilTest
            'ReDim Preserve tgPopDnfInfo(1 To UBound(tgPopDnfInfo) - 1) As POPDNFINFO
            ReDim Preserve tgPopDnfInfo(ilLowLimit To UBound(tgPopDnfInfo) - 1) As POPDNFINFO
        End If
        illoop = illoop - 1
    Loop
    'Sort then output new headers and lines
    'If UBound(tgPopDnfInfo) - 1 > 1 Then
    If UBound(tgPopDnfInfo) - 1 > ilLowLimit Then
        'ArraySortTyp fnAV(tgPopDnfInfo(), 1), UBound(tgPopDnfInfo) - 1, 0, LenB(tgPopDnfInfo(0)), 0, -1, 0
        ArraySortTyp fnAV(tgPopDnfInfo(), ilLowLimit), UBound(tgPopDnfInfo) - ilLowLimit, 0, LenB(tgPopDnfInfo(ilLowLimit)), 0, -1, 0
    End If
    On Error GoTo 0
    ilRet = btrClose(hlDnf)
    btrDestroy hlDnf
    Exit Function

    ilRet = btrClose(hlDnf)
    btrDestroy hlDnf
    gDbg_HandleError "PopSubs: gObtainPopDnfInfoBox"
gObtainPopDnfInfoErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gAddItemToSortCode              *
'*                                                     *
'*             Created:6/03/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add item to Sort Code List      *
'*                                                     *
'*******************************************************
Sub gAddItemToSortCode(slName As String, tlSortCode() As SORTCODE, ilResort As Integer)
    tlSortCode(UBound(tlSortCode)).sKey = slName
    ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 1) As SORTCODE
    If ilResort Then
        If UBound(tlSortCode) - 1 > 0 Then
            ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
        End If
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gCntrForActiveOHD              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain an array of contracts    *
'*                     salesperson code                *
'*     8-21-05 add pct trade to array                  *
'*******************************************************
Function gCntrForActiveOHD(frm As Form, slAStartDate As String, slAEndDate As String, slOHDStartDate As String, slOHDEndDate As String, slStatus As String, slCntrType As String, ilHOType As Integer, tlChfAdvtExt() As CHFADVTEXT, Optional ilAdfCode As Integer = -1) As Integer
'
'   ilRet = gCntrForActiveOHD (MainForm, slAStartDate, slAEndDate, slOHDStartDate, slOHDEndDate, slStatus, slCntrType, ilHOType, tlChfAdvtExt() )
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       slAStartDate(I)- Active Start Date
'       slAEndDate(I)- Active End Date
'       slOHDStartDate(I)- OHD Start Date
'       slOHDEndDate(I)- OHD End Date
'       slStatus (I)- chfStatus value or blank
'                         W=Working; D=Rejected; C=Completed; I=Unapproved; H=Hold; O=Order
'                         Multiple status can be specified (WDI)
'       slCntrType (I)- chfType value or blank
'                       C=Standard; V=Reservation; T=Remnant; R=DR; Q=PI; S=PSA; M=Promo
'       ilHOType (I)-  1=H or O only; 2=H or O or G or N (if G or N exists show it over H or O);
'                      3=H or O or G or N or W or C or I (if G or N or W or C or I exists show it over H or O)
'                        Note: G or N can't exist at the same time as W or C or I for an order
'                              G or N or W or C or I CntrRev > 0
'                      4=Return all matching contracts not deleted
'                      5=Return Original contract (smallest extrevno, ignoring delete flag)
'       tlChfAdvtExt(O)- Array of contracts of the structure CHFADVTEXT which match selection
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'CHF date/time stamp
    Dim hlChf As Integer        'CHF handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlChf As CHF
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilFound As Integer
    'Dim ilLoop As Integer
    Dim llLoop As Long 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
    Dim llTodayDate As Long
    Dim slDate As String
    Dim hlVef As Integer        'Vef handle
    Dim tlVef As VEF
    Dim ilVefRecLen As Integer     'Record length
    Dim hlVsf As Integer        'Vsf handle
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilOffSet As Integer
    Dim ilOper As Integer
    Dim slStr As String
    'Dim ilCntUpper As Integer
    Dim llCntUpper As Long 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
    Dim ilSlfCode As Integer
    Dim ilCurrent As Integer
    Dim ilTestCntrNo As Integer
    Dim slCntrStatus As String
    Dim slHOStatus As String
    Dim ilLowLimit As Integer

    If slStatus = "" Then
        slCntrStatus = "WCIDHO"
    Else
        slCntrStatus = slStatus
    End If
    slHOStatus = ""
    If ilHOType = 1 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "H"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "O"
        End If
    ElseIf ilHOType = 2 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
    ElseIf ilHOType = 3 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
        If (InStr(1, slCntrStatus, "H", 1) <> 0) Or (InStr(1, slCntrStatus, "O", 1) <> 0) Then
            slHOStatus = slHOStatus & "WCI"
        End If
    ElseIf ilHOType = 4 Then
        slHOStatus = slCntrStatus
    ElseIf ilHOType = 5 Then
        slHOStatus = slCntrStatus
    End If

    'On Error GoTo gPopCntrForActiveOHDErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlChfAdvtExt)
    'If ilRet <> 0 Then
    '    sgCntrForDateStamp = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlChfAdvtExt).Ptr <> 0 Then
        ilLowLimit = LBound(tlChfAdvtExt)
    Else
        sgCntrForDateStamp = ""
        ilLowLimit = 0
    End If

    slStamp = gFileDateTime(sgDBPath & "Chf.Btr") & slAStartDate & slAEndDate & slOHDStartDate & slOHDEndDate & Trim$(slCntrStatus) & Trim$(slCntrType) & Trim$(str$(ilCurrent)) & Trim$(str$(ilHOType))
    If sgCntrForDateStamp <> "" Then
        If StrComp(slStamp, sgCntrForDateStamp, 1) = 0 Then
            gCntrForActiveOHD = CP_MSG_NOPOPREQ
            Exit Function
        End If
    End If
    gCntrForActiveOHD = CP_MSG_POPREQ
    ilCurrent = 1   'All
    llTodayDate = gDateValue(gNow())
    'gObtainVehComboList
    hlChf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrOpen):" & "Chf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlChf) 'btrRecordLength(hlChf)  'Get and save record length
    hlVsf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrOpen):" & "Vsf.Btr", frm
    On Error GoTo 0
    ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length
    hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrOpen):" & "Vef.Btr", frm
    On Error GoTo 0
    ilVefRecLen = Len(tlVef) 'btrRecordLength(hlSlf)  'Get and save record length
    sgCntrForDateStamp = slStamp
    'ReDim tlChfAdvtExt(1 To 1) As CHFADVTEXT
    ReDim tlChfAdvtExt(ilLowLimit To ilLowLimit) As CHFADVTEXT
    'ilCntUpper = 1
    'ilCntUpper = ilLowLimit
    llCntUpper = ilLowLimit 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
    'ilExtLen = Len(tlChfAdvtExt(ilCntUpper))  'Extract operation record size
    ilExtLen = Len(tlChfAdvtExt(llCntUpper))  'Extract operation record size'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlChf) 'Obtain number of records
    btrExtClear hlChf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlChf, tlChf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
        ilRet = btrClose(hlChf)
        On Error GoTo gCntrForActiveOHDErr
        gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrReset):" & "Chf.Btr", frm
        On Error GoTo 0
        btrDestroy hlChf
        ilRet = btrClose(hlVsf)
        On Error GoTo gCntrForActiveOHDErr
        gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrReset):" & "Vsf.Btr", frm
        On Error GoTo 0
        btrDestroy hlVsf
        Exit Function
    Else
        On Error GoTo gCntrForActiveOHDErr
        gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrGetFirst):" & "Chf.Btr", frm
        On Error GoTo 0
    End If
    Call btrExtSetBounds(hlChf, llNoRec, -1, "UC", "CHFADVTEXTPK", CHFADVTEXTPK) 'Set extract limits (all records)
    ilSlfCode = tgUrf(0).iSlfCode
    If (tgUrf(0).iGroupNo > 0) Then     'And (tgUrf(0).iSlfCode <= 0) Then
        ilRet = gObtainUrf()
        ilRet = gObtainSalesperson()
    End If
    ' chfEndDate >= InputStartDate And chfStartDate <= InputEndDate
    If slAStartDate = "" Then
        slDate = "1/1/1970"
    Else
        slDate = slAStartDate
    End If
    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Chf", "ChfEndDate")
    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    If slAEndDate = "" Then
        slDate = "12/31/2069"
    Else
        slDate = slAEndDate
    End If
    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Chf", "ChfStartDate")
    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    If slOHDStartDate = "" Then
        slDate = "1/1/1970"
    Else
        slDate = slOHDStartDate
    End If
    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Chf", "ChfOHDDate")
    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    If slOHDEndDate = "" Then
        slDate = "12/31/2069"
    Else
        slDate = slOHDEndDate
    End If
    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Chf", "ChfOHDDate")
    If (ilHOType = 5) And (slCntrType = "") Then
        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
    Else
        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        tlCharTypeBuff.sType = "Y"
        ilOffSet = gFieldOffset("Chf", "ChfDelete")
    End If
    'If (slCntrStatus = "") And (slCntrType = "") Then
    If (slCntrType = "") Then
        If (ilHOType <> 5) Then
            ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        End If
    Else
        If (ilHOType <> 5) Then
            ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
        End If
        'If slCntrStatus <> "" Then
        '    ilOper = BTRV_EXT_OR
        '    slStr = slCntrStatus
        '    Do While slStr <> ""
        '        If Len(slStr) = 1 Then
        '            If slCntrType <> "" Then
        '                ilOper = BTRV_EXT_AND
        '            Else
        '                ilOper = BTRV_EXT_LAST_TERM
        '            End If
        '        End If
        '        tlCharTypeBuff.sType = Left$(slStr, 1)
        '        ilOffset = gFieldOffset("Chf", "ChfStatus")
        '        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
        '        slStr = Mid$(slStr, 2)
        '    Loop
        'End If
        If slCntrType <> "" Then
            ilOper = BTRV_EXT_OR
            slStr = slCntrType
            Do While slStr <> ""
                If Len(slStr) = 1 Then
                    ilOper = BTRV_EXT_LAST_TERM
                End If
                tlCharTypeBuff.sType = Left$(slStr, 1)
                ilOffSet = gFieldOffset("Chf", "ChfType")
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
                slStr = Mid$(slStr, 2)
            Loop
        End If
    End If
    ilOffSet = gFieldOffset("Chf", "ChfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract iCode field
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfCntrNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract Contract number
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfExtRevNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfCntRevNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfType")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfProduct")
    ilRet = btrExtAddField(hlChf, ilOffSet, 35) 'Extract Product
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfSlfCode1")
    ilRet = btrExtAddField(hlChf, ilOffSet, 20) 'Extract salesperson code
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfMnfDemo1")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract salesperson code
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfCxfInt")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfPropVer")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract end date
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfStatus")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfMnfPotnType")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract SellNet
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfStartDate")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfEndDate")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract end date
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfVefCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Chf", "ChfSifCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    '8-21-05 add pct of trade to array
    ilOffSet = gFieldOffset("Chf", "ChfPctTrade")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'pct trade
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    '7/12/10
    ilOffSet = gFieldOffset("Chf", "ChfCBSOrder")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract CBSOrder
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    '2/24/12
    ilOffSet = gFieldOffset("Chf", "ChfBillCycle")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract BillCycle
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    '2/17/18
    ilOffSet = gFieldOffset("Chf", "ChfSource")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Source
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    '2/9/21
    ilOffSet = gFieldOffset("Chf", "chfAdServerDefined")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract AdServerDefined
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    
    'ilRet = btrExtGetNextExt(hlChf)    'Extract record
    'ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos)
    ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(llCntUpper), ilExtLen, llRecPos) 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo gCntrForActiveOHDErr
        gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtGetNextExt):" & "Chf.Btr", frm
        On Error GoTo 0
        'ilExtLen = Len(tlChfAdvtExt(ilCntUpper))  'Extract operation record size
        ilExtLen = Len(tlChfAdvtExt(llCntUpper))  'Extract operation record size'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
        'ilRet = btrExtGetFirst(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            'ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos)
            ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(llCntUpper), ilExtLen, llRecPos) 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
        Loop
        Do While ilRet = BTRV_ERR_NONE
            'If (ilAdfCode <> -1) And (ilAdfCode <> tlChfAdvtExt(ilCntUpper).iAdfCode) Then
            If (ilAdfCode <> -1) And (ilAdfCode <> tlChfAdvtExt(llCntUpper).iAdfCode) Then 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
            Else
                ilFound = True
                ilTestCntrNo = False
                'For Proposal CntRevNo = 0; For Orders CntRevNo >= 0 (for W, C, I CntRevNo > 0)
                'If (tlChfAdvtExt(ilCntUpper).iCntRevNo = 0) And ((tlChfAdvtExt(ilCntUpper).sStatus <> "H") And (tlChfAdvtExt(ilCntUpper).sStatus <> "O") And (tlChfAdvtExt(ilCntUpper).sStatus <> "G") And (tlChfAdvtExt(ilCntUpper).sStatus <> "N")) Then  'Proposal
                If (tlChfAdvtExt(llCntUpper).iCntRevNo = 0) And ((tlChfAdvtExt(llCntUpper).sStatus <> "H") And (tlChfAdvtExt(llCntUpper).sStatus <> "O") And (tlChfAdvtExt(llCntUpper).sStatus <> "G") And (tlChfAdvtExt(llCntUpper).sStatus <> "N")) Then  'Proposal'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                    'If (InStr(1, slCntrStatus, tlChfAdvtExt(ilCntUpper).sStatus) = 0) Then
                    If (InStr(1, slCntrStatus, tlChfAdvtExt(llCntUpper).sStatus) = 0) Then 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                        ilFound = False
                    End If
                Else    'Order
                    'If (InStr(1, slHOStatus, tlChfAdvtExt(ilCntUpper).sStatus) <> 0) Then
                    If (InStr(1, slHOStatus, tlChfAdvtExt(llCntUpper).sStatus) <> 0) Then 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                        If (ilHOType = 2) Or (ilHOType = 3) Or (ilHOType = 5) Then
                            ilTestCntrNo = True
                        End If
                    Else
                        ilFound = False
                    End If
                End If
                If ilFound Then
                    '2/17/18: Remove Programmatic filter
                    'tlChfAdvtExt(ilCntUpper).sSource = ""
                    tlChfAdvtExt(llCntUpper).sSource = "" 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                    'ilFound = gTestChfAdvtExt(frm, ilSlfCode, tlChfAdvtExt(ilCntUpper), hlVsf, ilCurrent)
                    ilFound = gTestChfAdvtExt(frm, ilSlfCode, tlChfAdvtExt(llCntUpper), hlVsf, ilCurrent) 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                End If
                If ilFound Then
                    ilFound = False
                    If ilTestCntrNo Then
                        'For ilLoop = 1 To ilCntUpper - 1 Step 1
                        'For ilLoop = ilLowLimit To ilCntUpper - 1 Step 1
                        For llLoop = ilLowLimit To llCntUpper - 1 Step 1 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                            If (ilHOType <> 5) Then
                                'If tlChfAdvtExt(ilLoop).lCntrNo = tlChfAdvtExt(ilCntUpper).lCntrNo Then
                                If tlChfAdvtExt(llLoop).lCntrNo = tlChfAdvtExt(llCntUpper).lCntrNo Then 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                                    'If tlChfAdvtExt(ilLoop).iCntRevNo < tlChfAdvtExt(ilCntUpper).iCntRevNo Then
                                    If tlChfAdvtExt(llLoop).iCntRevNo < tlChfAdvtExt(llCntUpper).iCntRevNo Then 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                                        'tlChfAdvtExt(ilLoop) = tlChfAdvtExt(ilCntUpper)
                                        tlChfAdvtExt(llLoop) = tlChfAdvtExt(llCntUpper) 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                                    End If
                                    ilFound = True
                                    Exit For
                                End If
                            Else
                                'If tlChfAdvtExt(ilLoop).lCntrNo = tlChfAdvtExt(ilCntUpper).lCntrNo Then
                                If tlChfAdvtExt(llLoop).lCntrNo = tlChfAdvtExt(llCntUpper).lCntrNo Then 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                                    'If tlChfAdvtExt(ilLoop).iExtRevNo > tlChfAdvtExt(ilCntUpper).iExtRevNo Then
                                    If tlChfAdvtExt(llLoop).iExtRevNo > tlChfAdvtExt(llCntUpper).iExtRevNo Then 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                                        'tlChfAdvtExt(ilLoop) = tlChfAdvtExt(ilCntUpper)
                                        tlChfAdvtExt(llLoop) = tlChfAdvtExt(llCntUpper) 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                                    End If
                                    ilFound = True
                                    Exit For
                                End If
                            End If
                        'Next ilLoop
                        Next llLoop 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                    End If
                    If Not ilFound Then
                        'ilCntUpper = ilCntUpper + 1
                        llCntUpper = llCntUpper + 1 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                        'ReDim Preserve tlChfAdvtExt(1 To ilCntUpper) As CHFADVTEXT
                        'ReDim Preserve tlChfAdvtExt(ilLowLimit To ilCntUpper) As CHFADVTEXT
                        ReDim Preserve tlChfAdvtExt(ilLowLimit To llCntUpper) As CHFADVTEXT 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
                    End If
                End If
            End If
            'ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos)
            ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(llCntUpper), ilExtLen, llRecPos) 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                'ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos)
                ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(llCntUpper), ilExtLen, llRecPos) 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
            Loop
        Loop
    End If
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlChf)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrReset):" & "Chf.Btr", frm
    On Error GoTo 0
    btrDestroy hlChf
    Exit Function
gCntrForActiveOHDErr:
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlChf)
    btrDestroy hlChf
    gDbg_HandleError "PopSubs: gCntrForActiveOHD"
'    gCntrForActiveOHD = CP_MSG_NOSHOW
'    Exit Function
gPopCntrForActiveOHDErr2:
    ilRet = 1
    Resume Next
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:gCntrDetailForDateRange         *
'*                                                     *
'*             Created:6/13/93       By:J. White       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain an array of contracts    *
'*                     For TTP 10937 - Dashboard       *
'*******************************************************
Function gCntrDetailForDateRange(frm As Form, slAStartDate As String, slAEndDate As String, tlChfDetailList() As CHFDETAILLIST) As Integer
'
'   ilRet = gCntrDetailForDateRange (MainForm, slAStartDate, slAEndDate,  slCntrType,  tlChf() )
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       slAStartDate(I)- Active Start Date
'       slAEndDate(I)- Active End Date
'       tlChfDetailList(O)- Array of contracts of the structure CHFDETAILLIST() which match selection
'       ilRet (O)- Error code (0 if no error)

    Dim slStamp As String    'CHF date/time stamp
    Dim hlChf As Integer        'CHF handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlChf As CHF
    
    Dim hlClf As Integer        'CLF handle
    Dim ilClfRecLen As Integer
    Dim tlClf() As CLFLIST
    Dim tlClfSrchKey As CLFKEY0            'CLF record image
    
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim llLoop As Long 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
    Dim llTodayDate As Long
    Dim slDate As String
    Dim hlVsf As Integer        'Vsf handle
    Dim ilVsfReclen As Integer     'Record length
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilOffSet As Integer
    Dim ilOper As Integer
    Dim slStr As String
    Dim llCntUpper As Long 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
    Dim ilSlfCode As Integer
    Dim ilCurrent As Integer
    Dim ilTestCntrNo As Integer
    Dim slCntrStatus As String
    Dim slHOStatus As String
    Dim ilLowLimit As Integer
    Dim tlChfAdvTmp As CHFADVTEXT
    
    slCntrStatus = "WCIDHO"
    If PeekArray(tlChfDetailList).Ptr <> 0 Then
        ilLowLimit = LBound(tlChfDetailList)
    Else
        sgCntrForDateStamp = ""
        ilLowLimit = 0
    End If

    slStamp = gFileDateTime(sgDBPath & "Chf.Btr") & slAStartDate & slAEndDate & Trim$(slCntrStatus) & Trim$(str$(ilCurrent))
    If sgCntrForDateStamp <> "" Then
        If StrComp(slStamp, sgCntrForDateStamp, 1) = 0 Then
            gCntrDetailForDateRange = CP_MSG_NOPOPREQ
            Exit Function
        End If
    End If
        
    gCntrDetailForDateRange = CP_MSG_POPREQ
    ilCurrent = 1   'All
    llTodayDate = gDateValue(gNow())
    hlChf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrOpen):" & "Chf.Btr", frm
    On Error GoTo 0
    
    hlClf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrOpen):" & "Clf.Btr", frm
    On Error GoTo 0
    
    ilRecLen = Len(tlChf) 'btrRecordLength(hlChf)  'Get and save record length
    hlVsf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrOpen):" & "Vsf.Btr", frm
    On Error GoTo 0
    ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrOpen):" & "Vef.Btr", frm
    On Error GoTo 0
    sgCntrForDateStamp = slStamp
    ReDim tlChfDetailList(ilLowLimit To ilLowLimit) As CHFDETAILLIST
    llCntUpper = ilLowLimit 'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
    ilExtLen = Len(tlChfDetailList(llCntUpper))  'Extract operation record size'TTP 10602 - Invoice overflow error in gCntrForActiveOHD, gObtainCntrForDate when count of contracts exceeds 32767
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlChf) 'Obtain number of records
    btrExtClear hlChf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlChf, tlChf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
        ilRet = btrClose(hlChf)
        On Error GoTo gCntrForActiveOHDErr
        gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrReset):" & "Chf.Btr", frm
        On Error GoTo 0
        btrDestroy hlChf
        ilRet = btrClose(hlVsf)
        On Error GoTo gCntrForActiveOHDErr
        gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrReset):" & "Vsf.Btr", frm
        On Error GoTo 0
        btrDestroy hlVsf
        Exit Function
    Else
        On Error GoTo gCntrForActiveOHDErr
        gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrGetFirst):" & "Chf.Btr", frm
        On Error GoTo 0
    End If
    Call btrExtSetBounds(hlChf, llNoRec, -1, "UC", "CHFDETAILLISTPK", CHFDETAILLISTPK) 'Set extract limits (all records)
    ilSlfCode = tgUrf(0).iSlfCode
    If (tgUrf(0).iGroupNo > 0) Then
        ilRet = gObtainUrf()
        ilRet = gObtainSalesperson()
    End If
    'Start Date
    If slAStartDate = "" Then
        slDate = "1/1/1970"
    Else
        slDate = slAStartDate
    End If
    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Chf", "ChfEndDate")
    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    'End Date
    If slAEndDate = "" Then
        slDate = "12/31/2069"
    Else
        slDate = slAEndDate
    End If
    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
    ilOffSet = gFieldOffset("Chf", "ChfStartDate")
    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    
    'Type
    ilOffSet = gFieldOffset("Chf", "ChfType")
    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
    
    'ChfCode
    ilOffSet = gFieldOffset("Chf", "ChfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfCntrNo
    ilOffSet = gFieldOffset("Chf", "ChfCntrNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfExtRevNo
    ilOffSet = gFieldOffset("Chf", "ChfExtRevNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfCntRevNo
    ilOffSet = gFieldOffset("Chf", "ChfCntRevNo")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfType
    ilOffSet = gFieldOffset("Chf", "ChfType")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfAdfCode
    ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfProduct
    ilOffSet = gFieldOffset("Chf", "ChfProduct")
    ilRet = btrExtAddField(hlChf, ilOffSet, 35)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfAgfCode
    ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfSlfCode1
    ilOffSet = gFieldOffset("Chf", "ChfSlfCode1")
    ilRet = btrExtAddField(hlChf, ilOffSet, 20)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfMnfDemo1
    ilOffSet = gFieldOffset("Chf", "ChfMnfDemo1")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfCxfInt
    ilOffSet = gFieldOffset("Chf", "ChfCxfInt")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfPropVer
    ilOffSet = gFieldOffset("Chf", "ChfPropVer")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfStatus
    ilOffSet = gFieldOffset("Chf", "ChfStatus")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfMnfPotnType
    ilOffSet = gFieldOffset("Chf", "ChfMnfPotnType")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfStartDate
    ilOffSet = gFieldOffset("Chf", "ChfStartDate")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfEndDate
    ilOffSet = gFieldOffset("Chf", "ChfEndDate")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfVefCode
    ilOffSet = gFieldOffset("Chf", "ChfVefCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfSifCode
    ilOffSet = gFieldOffset("Chf", "ChfSifCode")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfPctTrade
    ilOffSet = gFieldOffset("Chf", "ChfPctTrade")
    ilRet = btrExtAddField(hlChf, ilOffSet, 2)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfCBSOrder
    ilOffSet = gFieldOffset("Chf", "ChfCBSOrder")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfBillCycle
    ilOffSet = gFieldOffset("Chf", "ChfBillCycle")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'ChfSource
    ilOffSet = gFieldOffset("Chf", "ChfSource")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'chfAdServerDefined
    ilOffSet = gFieldOffset("Chf", "chfAdServerDefined")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'HasAir: PodSpotDefined
    ilOffSet = gFieldOffset("Chf", "ChfPodSpotDefined")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'HasNTR: NTRDefined
    ilOffSet = gFieldOffset("Chf", "ChfNTRDefined")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'SCHEDULE Status: schStatus
    ilOffSet = gFieldOffset("Chf", "ChfSchStatus")
    ilRet = btrExtAddField(hlChf, ilOffSet, 1)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'LastChanged: OHDDate
    ilOffSet = gFieldOffset("Chf", "ChfOHDDate")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'Gross: inputGross
    ilOffSet = gFieldOffset("Chf", "ChfInputGross")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    'AdservDlvyStatus
    ilOffSet = gFieldOffset("Chf", "ChfAdservDlvyStatus")
    ilRet = btrExtAddField(hlChf, ilOffSet, 4)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrDetailForDateRange (btrExtAddField):" & "Chf.Btr", frm
    On Error GoTo 0
    
    ilRet = btrExtGetNext(hlChf, tlChfDetailList(llCntUpper), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo gCntrForActiveOHDErr
        gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrExtGetNextExt):" & "Chf.Btr", frm
        On Error GoTo 0
        ilExtLen = Len(tlChfDetailList(llCntUpper))  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlChf, tlChfDetailList(llCntUpper), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilFound = True
            ReDim tlClf(0) As CLFLIST
            '----------------------------------
            'If this is a contract, set TestCntrNo =true so that we get the latest revision only
            ilTestCntrNo = False
            If tlChfDetailList(llCntUpper).sStatus = "H" Or tlChfDetailList(llCntUpper).sStatus = "O" Or tlChfDetailList(llCntUpper).sStatus = "G" Or tlChfDetailList(llCntUpper).sStatus = "N" Then
                ilTestCntrNo = True
            End If
            'v81 TTP 10937 - testing 2/6/24 1:31 PM - Issue 3
            'Rev Complete
            If tlChfDetailList(llCntUpper).iCntRevNo > 0 And tlChfDetailList(llCntUpper).sStatus = "C" And tlChfDetailList(llCntUpper).sSchStatus = "P" Then
                ilTestCntrNo = True
            End If
            'Rev Working
            If tlChfDetailList(llCntUpper).iCntRevNo > 0 And tlChfDetailList(llCntUpper).sStatus = "W" And tlChfDetailList(llCntUpper).sSchStatus = "P" Then
                ilTestCntrNo = True
            End If
            'Rev Unapproved
            If tlChfDetailList(llCntUpper).iCntRevNo > 0 And tlChfDetailList(llCntUpper).sStatus = "I" And tlChfDetailList(llCntUpper).sSchStatus = "P" Then
                ilTestCntrNo = True
            End If
            'Approved
            If tlChfDetailList(llCntUpper).iCntRevNo > 0 And tlChfDetailList(llCntUpper).sStatus = "N" And tlChfDetailList(llCntUpper).sSchStatus = "A" Then
                ilTestCntrNo = True
            End If
            
            '----------------------------------
            'Check if Contract okay for the user
            tlChfDetailList(llCntUpper).sSource = ""
            tlChfAdvTmp.lVefCode = tlChfDetailList(llCntUpper).lVefCode
            tlChfAdvTmp.iSlfCode(0) = tlChfDetailList(llCntUpper).iSlfCode(0)
            tlChfAdvTmp.iSlfCode(1) = tlChfDetailList(llCntUpper).iSlfCode(1)
            tlChfAdvTmp.iSlfCode(2) = tlChfDetailList(llCntUpper).iSlfCode(2)
            tlChfAdvTmp.iSlfCode(3) = tlChfDetailList(llCntUpper).iSlfCode(3)
            tlChfAdvTmp.iSlfCode(4) = tlChfDetailList(llCntUpper).iSlfCode(4)
            tlChfAdvTmp.iSlfCode(5) = tlChfDetailList(llCntUpper).iSlfCode(5)
            tlChfAdvTmp.iSlfCode(6) = tlChfDetailList(llCntUpper).iSlfCode(6)
            tlChfAdvTmp.iSlfCode(7) = tlChfDetailList(llCntUpper).iSlfCode(7)
            tlChfAdvTmp.iSlfCode(8) = tlChfDetailList(llCntUpper).iSlfCode(8)
            tlChfAdvTmp.iSlfCode(9) = tlChfDetailList(llCntUpper).iSlfCode(9)
            tlChfAdvTmp.sSource = tlChfDetailList(llCntUpper).sSource
            tlChfAdvTmp.iEndDate(0) = tlChfDetailList(llCntUpper).iEndDate(0)
            tlChfAdvTmp.iEndDate(1) = tlChfDetailList(llCntUpper).iEndDate(1)
            tlChfAdvTmp.iStartDate(0) = tlChfDetailList(llCntUpper).iStartDate(0)
            tlChfAdvTmp.iStartDate(1) = tlChfDetailList(llCntUpper).iStartDate(1)
            ilFound = gTestChfAdvtExt(frm, ilSlfCode, tlChfAdvTmp, hlVsf, ilCurrent)
        
            'Check for Airtime Line
            tlChfDetailList(llCntUpper).sAirtimeDefined = "N"
            ilClfRecLen = Len(tlClf(0).ClfRec)
            btrExtClear hlClf   'Clear any previous extend operation
            tlClfSrchKey.lChfCode = tlChfDetailList(llCntUpper).lCode
            tlClfSrchKey.iLine = 0
            tlClfSrchKey.iCntRevNo = tlChfDetailList(llCntUpper).iCntRevNo ' 0 show latest version
            tlClfSrchKey.iPropVer = tlChfDetailList(llCntUpper).iPropVer ' 0 show latest version
            ilRet = btrGetGreaterOrEqual(hlClf, tlClf(0).ClfRec, ilClfRecLen, tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (tlClf(0).ClfRec.lChfCode = tlChfDetailList(llCntUpper).lCode) Then
                tlChfDetailList(llCntUpper).sAirtimeDefined = "Y"
            End If
            If tlChfDetailList(llCntUpper).sAirtimeDefined <> "Y" Then tlChfDetailList(llCntUpper).sAirtimeDefined = "N"
            If tlChfDetailList(llCntUpper).sNTRDefined <> "Y" Then tlChfDetailList(llCntUpper).sNTRDefined = "N"
            
            '----------------------------------
            'Add Contract to array?
            If ilFound Then
                '----------------------------------
                'Check for latest contract revision?
                ilFound = False
                If ilTestCntrNo Then
                    For llLoop = ilLowLimit To llCntUpper - 1 Step 1
                        'If (ilHOType <> 5) Then
                            If tlChfDetailList(llLoop).lCntrNo = tlChfDetailList(llCntUpper).lCntrNo Then
                                If tlChfDetailList(llLoop).iCntRevNo < tlChfDetailList(llCntUpper).iCntRevNo Then
                                    tlChfDetailList(llLoop) = tlChfDetailList(llCntUpper)
                                End If
                                ilFound = True
                                Exit For
                            End If
                        'Else
                        '    If tlChfDetailList(llLoop).lCntrNo = tlChfDetailList(llCntUpper).lCntrNo Then
                        '        If tlChfDetailList(llLoop).iExtRevNo > tlChfDetailList(llCntUpper).iExtRevNo Then
                        '            tlChfDetailList(llLoop) = tlChfDetailList(llCntUpper)
                        '
                        '        End If
                        '        ilFound = True
                        '        Exit For
                        '    End If
                        'End If
                    Next llLoop
                End If
                If Not ilFound Then
                    llCntUpper = llCntUpper + 1
                    ReDim Preserve tlChfDetailList(ilLowLimit To llCntUpper) As CHFDETAILLIST
                End If
            End If
            
            '----------------------------------
            'Get next record
            ilRet = btrExtGetNext(hlChf, tlChfDetailList(llCntUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlChf, tlChfDetailList(llCntUpper), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    'ilRet = btrClose(hlVef)
    'btrDestroy hlVef
    ilRet = btrClose(hlChf)
    ilRet = btrClose(hlClf)
    On Error GoTo gCntrForActiveOHDErr
    gBtrvErrorMsg ilRet, "gCntrForActiveOHD (btrReset):" & "Chf.Btr", frm
    On Error GoTo 0
    btrDestroy hlChf
    Exit Function
gCntrForActiveOHDErr:
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    'ilRet = btrClose(hlVef)
    'btrDestroy hlVef
    ilRet = btrClose(hlChf)
    ilRet = btrClose(hlClf)
    btrDestroy hlChf
    gDbg_HandleError "PopSubs: gCntrForActiveOHD"
gPopCntrForActiveOHDErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gDatePop                        *
'*                                                     *
'*             Created:8/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate list box with dates   *
'*                                                     *
'*******************************************************
Sub gDatePop(slInputDate As String, ilNoWeeks As Integer, lbc As Control)
'
'   gDatePop slDate, ilNoWeeks, lbc
'   Where:
'       slDate (I)- Start date
'       ilNoWeeks (I)- Number of weeks
'       lbc (I)- list box to be populated
'

    Dim llDate As Long
    Dim illoop As Integer
    Dim slDate As String
    Dim llLen As Long

    lbc.Clear
    llDate = gDateValue(slInputDate)
    llLen = 0
    For illoop = 1 To ilNoWeeks Step 1
        'slDate = Format$(llDate, "m/d/yy")
        slDate = Format$(llDate, "m/d/yy")
        If Not gOkAddStrToListBox(slDate, llLen, True) Then
            Exit For
        End If
        lbc.AddItem slDate
        llDate = llDate + 7
    Next illoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gDayPop                         *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list box with the Days *
'*                   of the week (starting with Monday)*
'*                                                     *
'*******************************************************
Sub gDayPop(Ctrl As Control)
'
'   gDayPop lbc
'   Where:
'       lbc (I)- list box to be populated
'
    Ctrl.Clear
    Ctrl.AddItem "Mo"
    Ctrl.AddItem "Tu"
    Ctrl.AddItem "We"
    Ctrl.AddItem "Th"
    Ctrl.AddItem "Fr"
    Ctrl.AddItem "Sa"
    Ctrl.AddItem "Su"
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gDemoPop                        *
'*                                                     *
'*             Created:8/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list box with standard *
'*                     demographics                    *
'*                                                     *
'*                     Click on Custom Demos on List   *
'*                     screen to get these demos loaded*
'*                     into MNF                        *
'*                                                     *
'*******************************************************
Sub gDemoPop(lbc As Control)
'
'   gDemoPop lbc
'   Where:
'       lbc (I)- list box to be populated
'
    lbc.Clear
    lbc.AddItem "N/A"
    lbc.AddItem "M12-17"
    lbc.AddItem "M12-20"
    lbc.AddItem "M12-24"
    lbc.AddItem "M12-34"
    lbc.AddItem "M12-44"
    lbc.AddItem "M12-49"
    lbc.AddItem "M12-54"
    lbc.AddItem "M12-64"
    lbc.AddItem "M12+"
    lbc.AddItem "M18-24"
    lbc.AddItem "M18-34"
    lbc.AddItem "M18-44"
    lbc.AddItem "M18-49"
    lbc.AddItem "M18-54"
    lbc.AddItem "M18-64"
    lbc.AddItem "M18+"
    lbc.AddItem "M21-34"
    lbc.AddItem "M21-44"
    lbc.AddItem "M21-49"
    lbc.AddItem "M21-54"
    lbc.AddItem "M21-64"
    lbc.AddItem "M21+"
    lbc.AddItem "M25-34"
    lbc.AddItem "M25-44"
    lbc.AddItem "M25-49"
    lbc.AddItem "M25-54"
    lbc.AddItem "M25-64"
    lbc.AddItem "M25+"
    lbc.AddItem "M35-44"
    lbc.AddItem "M35-49"
    lbc.AddItem "M35-54"
    lbc.AddItem "M35-64"
    lbc.AddItem "M35+"
    lbc.AddItem "M45-49"
    lbc.AddItem "M45-54"
    lbc.AddItem "M45-64"
    lbc.AddItem "M45+"
    lbc.AddItem "M50-54"
    lbc.AddItem "M50-64"
    lbc.AddItem "M50+"
    lbc.AddItem "M55-64"
    lbc.AddItem "M55+"
    lbc.AddItem "M65+"
    lbc.AddItem "P12-17"
    lbc.AddItem "P12-20"
    lbc.AddItem "P12-24"
    lbc.AddItem "P12-34"
    lbc.AddItem "P12-44"
    lbc.AddItem "P12-49"
    lbc.AddItem "P12-54"
    lbc.AddItem "P12-64"
    lbc.AddItem "P12+"
    lbc.AddItem "A18-24"
    lbc.AddItem "A18-34"
    lbc.AddItem "A18-44"
    lbc.AddItem "A18-49"
    lbc.AddItem "A18-54"
    lbc.AddItem "A18-64"
    lbc.AddItem "A18+"
    lbc.AddItem "A21-34"
    lbc.AddItem "A21-44"
    lbc.AddItem "A21-49"
    lbc.AddItem "A21-54"
    lbc.AddItem "A21-64"
    lbc.AddItem "A21+"
    lbc.AddItem "A25-34"
    lbc.AddItem "A25-44"
    lbc.AddItem "A25-49"
    lbc.AddItem "A25-54"
    lbc.AddItem "A25-64"
    lbc.AddItem "A25+"
    lbc.AddItem "A35-44"
    lbc.AddItem "A35-49"
    lbc.AddItem "A35-54"
    lbc.AddItem "A35-64"
    lbc.AddItem "A35+"
    lbc.AddItem "A45-49"
    lbc.AddItem "A45-54"
    lbc.AddItem "A45-64"
    lbc.AddItem "A45+"
    lbc.AddItem "A50-54"
    lbc.AddItem "A50-64"
    lbc.AddItem "A50+"
    lbc.AddItem "A55-64"
    lbc.AddItem "A55+"
    lbc.AddItem "A65+"
    lbc.AddItem "W12-17"
    lbc.AddItem "W12-20"
    lbc.AddItem "W12-24"
    lbc.AddItem "W12-34"
    lbc.AddItem "W12-44"
    lbc.AddItem "W12-49"
    lbc.AddItem "W12-54"
    lbc.AddItem "W12-64"
    lbc.AddItem "W12+"
    lbc.AddItem "W18-24"
    lbc.AddItem "W18-34"
    lbc.AddItem "W18-44"
    lbc.AddItem "W18-49"
    lbc.AddItem "W18-54"
    lbc.AddItem "W18-64"
    lbc.AddItem "W18+"
    lbc.AddItem "W21-34"
    lbc.AddItem "W21-44"
    lbc.AddItem "W21-49"
    lbc.AddItem "W21-54"
    lbc.AddItem "W21-64"
    lbc.AddItem "W21+"
    lbc.AddItem "W25-34"
    lbc.AddItem "W25-44"
    lbc.AddItem "W25-49"
    lbc.AddItem "W25-54"
    lbc.AddItem "W25-64"
    lbc.AddItem "W25+"
    lbc.AddItem "W35-44"
    lbc.AddItem "W35-49"
    lbc.AddItem "W35-54"
    lbc.AddItem "W35-64"
    lbc.AddItem "W35+"
    lbc.AddItem "W45-49"
    lbc.AddItem "W45-54"
    lbc.AddItem "W45-64"
    lbc.AddItem "W45+"
    lbc.AddItem "W50-54"
    lbc.AddItem "W50-64"
    lbc.AddItem "W50+"
    lbc.AddItem "W55-64"
    lbc.AddItem "W55+"
    lbc.AddItem "W65+"
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gEndCalDatePop                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate list box with calendar*
'*                      month end dates                *
'*                                                     *
'*******************************************************
Sub gEndCalDatePop(slInputDate As String, ilNoPeriods As Integer, lbc As Control)
'
'   gEndCalDatePop slDate, ilNoPeriods, lbc
'   Where:
'       slDate (I)- First calendar month end date to be stored in the list box
'       ilNoPeriods (I)- Number of calendar months to be stored into the list box
'       lbc (I)- list box to be populated
'

    Dim llDate As Long
    Dim illoop As Integer
    Dim slDate As String
    Dim llLen As Long

    llLen = 0
    lbc.Clear
    llDate = gDateValue(slInputDate)
    For illoop = 1 To ilNoPeriods Step 1
        'slDate = Format$(llDate, "m/d/yy")
        slDate = Format$(llDate, "m/d/yy")
        If Not gOkAddStrToListBox(slDate, llLen, True) Then
            Exit For
        End If
        lbc.AddItem slDate
        llDate = llDate + 27
        Do While Month(llDate) = Month(llDate + 1)
            llDate = llDate + 1
        Loop
    Next illoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gEndStdDatePop                  *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate list box with standard*
'*                      month end dates                *
'*                                                     *
'*******************************************************
Sub gEndStdDatePop(slInputDate As String, ilNoPeriods As Integer, lbc As Control)
'
'   gEndStdDatePop slDate, ilNoPeriods, lbc
'   Where:
'       slDate (I)- First standard month end date to be stored in the list box
'       ilNoPeriods (I)- Number of standard months to be stored into the list box
'       lbc (I)- list box to be populated
'

    Dim llDate As Long
    Dim illoop As Integer
    Dim slDate As String
    Dim llLen As Long

    llLen = 0
    lbc.Clear
    llDate = gDateValue(slInputDate)
    For illoop = 1 To ilNoPeriods Step 1
        'slDate = Format$(llDate, "m/d/yy")
        slDate = Format$(llDate, "m/d/yy")
        lbc.AddItem slDate
        If Not gOkAddStrToListBox(slDate, llLen, True) Then
            Exit For
        End If
        llDate = llDate + 21
        Do While Month(llDate) = Month(llDate + 7)
            llDate = llDate + 7
        Loop
    Next illoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gEndWkDatePop                   *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate list box with standard*
'*                      week end dates                *
'*                                                     *
'*******************************************************
Sub gEndWkDatePop(slInputDate As String, ilNoPeriods As Integer, lbc As Control)
'
'   gEndStdDatePop slDate, ilNoPeriods, lbc
'   Where:
'       slDate (I)- First standard month end date to be stored in the list box
'       ilNoPeriods (I)- Number of standard months to be stored into the list box
'       lbc (I)- list box to be populated
'

    Dim llDate As Long
    Dim illoop As Integer
    Dim slDate As String
    Dim llLen As Long

    llLen = 0
    lbc.Clear
    llDate = gDateValue(slInputDate)
    For illoop = 1 To ilNoPeriods Step 1
        'slDate = Format$(llDate, "m/d/yy")
        slDate = Format$(llDate, "m/d/yy")
        lbc.AddItem slDate
        If Not gOkAddStrToListBox(slDate, llLen, True) Then
            Exit For
        End If
        llDate = llDate + 6
        Do While Weekday(llDate) <> vbSunday
            llDate = llDate + 1
        Loop
    Next illoop
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gFileDateTime                   *
'*                                                     *
'*             Created:10/22/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain file time stamp          *
'*                                                     *
'*******************************************************
Function gFileDateTime(slPathFile As String) As String
    Dim ilRet As Integer

    ilRet = 0
    'On Error GoTo gFileDateTimeErr
    If igUsingTimeStamp Then
        ilRet = gFileExist(slPathFile)
        If ilRet = 0 Then
            gFileDateTime = FileDateTime(slPathFile)
        Else
            gFileDateTime = Format$(Now, "m/d/yy") & " " & Format$(Now, "h:mm:ssAM/PM")
        End If
    Else
        gFileDateTime = "NOTUSED"
    End If
    On Error GoTo 0
    Exit Function
gFileDateTimeErr:
    ilRet = err.Number
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gFileExist                      *
'*                                                     *
'*             Created:1/4/16        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determina if file exist         *
'*                                                     *
'*******************************************************
Function gFileExist(slPathFile As String) As Integer
    Dim fs As New FileSystemObject

    If fs.FILEEXISTS(slPathFile) Then
        gFileExist = 0
    Else
        gFileExist = 1
    End If
End Function

Function gFolderExist(slPathFolder As String) As Boolean
    '8886
    Dim fs As New FileSystemObject

    If fs.FolderExists(slPathFolder) Then
        gFolderExist = True
    Else
        gFolderExist = False
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gFileOpen                       *
'*                                                     *
'*             Created:1/4/16      By:D. LeVine        *
'*            Modified:            By:D. Smith         *
'*                                                     *
'*            Comments:Open a File                     *
'*                                                     *
'*******************************************************
Function gFileOpen(slPathFile As String, slForClause As String, hlHandle As Integer) As Integer
    Dim slCase As String

    gFileOpen = 0
    slForClause = UCase$(slForClause)
    On Error GoTo gFileOpenErr
    hlHandle = FreeFile
    slCase = ""
    
    If InStr(1, slForClause, "APPEND", vbTextCompare) > 0 Then
        slCase = "A"
    ElseIf InStr(1, slForClause, "BINARY", vbTextCompare) > 0 Then
        slCase = "B"
    ElseIf InStr(1, slForClause, "INPUT", vbTextCompare) > 0 Then
        slCase = "I"
    ElseIf InStr(1, slForClause, "OUTPUT", vbTextCompare) > 0 Then
        slCase = "O"
    ElseIf InStr(1, slForClause, "RANDOM", vbTextCompare) > 0 Then
        slCase = "R"
    Else
        slCase = "R"
    End If
    
    If InStr(1, slForClause, "ACCESS", vbTextCompare) > 0 Then
        If InStr(1, slForClause, "READ WRITE", vbTextCompare) > 0 Then
            slCase = slCase & "-B"
        ElseIf InStr(1, slForClause, "READ", vbTextCompare) > 0 Then
            slCase = slCase & "-R"
        ElseIf InStr(1, slForClause, "WRITE", vbTextCompare) > 0 Then
            slCase = slCase & "-W"
        End If
    End If
    
    If InStr(1, slForClause, "SHARED", vbTextCompare) > 0 Then
        slCase = slCase & ":S"
    ElseIf InStr(1, slForClause, "LOCK READ WRITE", vbTextCompare) > 0 Then
        slCase = slCase & ":B"
    ElseIf InStr(1, slForClause, "LOCK READ", vbTextCompare) > 0 Then
        slCase = slCase & ":R"
    ElseIf InStr(1, slForClause, "LOCK WRITE", vbTextCompare) > 0 Then
        slCase = slCase & ":W"
    End If
    
    Select Case slCase
        Case "A"
            Open slPathFile For Append As hlHandle
        Case "B"
            Open slPathFile For Binary As hlHandle
        Case "I"
            Open slPathFile For Input As hlHandle
        Case "O"
            Open slPathFile For Output As hlHandle
        Case "R"
            Open slPathFile For Random As hlHandle
           
        Case "A-B"
            Open slPathFile For Append Access Read Write As hlHandle
        Case "A-W"
            Open slPathFile For Append Access Write As hlHandle
            
        Case "B-B"
            Open slPathFile For Binary Access Read Write As hlHandle
        Case "B-R"
            Open slPathFile For Binary Access Read As hlHandle
        Case "B-W"
            Open slPathFile For Binary Access Write As hlHandle
            
        Case "I-R"
            Open slPathFile For Input Access Read As hlHandle
            
        Case "O-W"
            Open slPathFile For Output Access Write As hlHandle
            
        Case "R-B"
            Open slPathFile For Random Access Read Write As hlHandle
        Case "R-R"
            Open slPathFile For Random Access Read As hlHandle
        Case "R-W"
            Open slPathFile For Random Access Write As hlHandle


        Case "A:S"
            Open slPathFile For Append Shared As hlHandle
        Case "A:B"
            Open slPathFile For Append Lock Read Write As hlHandle
        Case "A:R"
            Open slPathFile For Append Lock Read As hlHandle
        Case "A:W"
            Open slPathFile For Append Lock Write As hlHandle

        Case "B:S"
            Open slPathFile For Binary Shared As hlHandle
        Case "B:B"
            Open slPathFile For Binary Lock Read Write As hlHandle
        Case "B:R"
            Open slPathFile For Binary Lock Read As hlHandle
        Case "B:W"
            Open slPathFile For Binary Lock Write As hlHandle

        Case "I:S"
            Open slPathFile For Input Shared As hlHandle
        Case "I:B"
            Open slPathFile For Input Lock Read Write As hlHandle
        Case "I:R"
            Open slPathFile For Input Lock Read As hlHandle
        Case "I:W"
            Open slPathFile For Input Lock Write As hlHandle

        Case "O:S"
            Open slPathFile For Output Shared As hlHandle
        Case "O:B"
            Open slPathFile For Output Lock Read Write As hlHandle
        Case "O:R"
            Open slPathFile For Output Lock Read As hlHandle
        Case "O:W"
            Open slPathFile For Output Lock Write As hlHandle

        Case "R:S"
            Open slPathFile For Random Shared As hlHandle
        Case "R:B"
            Open slPathFile For Random Lock Read Write As hlHandle
        Case "R:R"
            Open slPathFile For Random Lock Read As hlHandle
        Case "R:W"
            Open slPathFile For Random Lock Write As hlHandle
                        
            
        Case "A-B:S"
            Open slPathFile For Append Access Read Write Shared As hlHandle
        Case "A-B:B"
            Open slPathFile For Append Access Read Write Lock Read Write As hlHandle
        Case "A-B:R"
            Open slPathFile For Append Access Read Write Lock Read As hlHandle
        Case "A-B:W"
            Open slPathFile For Append Access Read Write Lock Write As hlHandle
            
        Case "A-W:S"
            Open slPathFile For Append Access Write Shared As hlHandle
        Case "A-W:B"
            Open slPathFile For Append Access Write Lock Read Write As hlHandle
        Case "A-W:R"
            Open slPathFile For Append Access Write Lock Read As hlHandle
        Case "A-W:W"
            Open slPathFile For Append Access Write Lock Write As hlHandle
            
        Case "B-B:S"
            Open slPathFile For Binary Access Read Write Shared As hlHandle
        Case "B-B:B"
            Open slPathFile For Binary Access Read Write Lock Read Write As hlHandle
        Case "B-B:R"
            Open slPathFile For Binary Access Read Write Lock Read As hlHandle
        Case "B-B:W"
            Open slPathFile For Binary Access Read Write Lock Write As hlHandle
            
        Case "B-R:S"
            Open slPathFile For Binary Access Read Shared As hlHandle
        Case "B-R:B"
            Open slPathFile For Binary Access Read Lock Read Write As hlHandle
        Case "B-R:R"
            Open slPathFile For Binary Access Read Lock Read As hlHandle
        Case "B-R:W"
            Open slPathFile For Binary Access Read Lock Write As hlHandle
            
        Case "B-W:S"
            Open slPathFile For Binary Access Write Shared As hlHandle
        Case "B-W:B"
            Open slPathFile For Binary Access Write Lock Read Write As hlHandle
        Case "B-W:R"
            Open slPathFile For Binary Access Write Lock Read As hlHandle
        Case "B-W:W"
            Open slPathFile For Binary Access Write Lock Write As hlHandle
            
        Case "I-R:S"
            Open slPathFile For Input Access Read Shared As hlHandle
        Case "I-R:B"
            Open slPathFile For Input Access Read Lock Read Write As hlHandle
        Case "I-R:R"
            Open slPathFile For Input Access Read Lock Read As hlHandle
        Case "I-R:W"
            Open slPathFile For Input Access Read Lock Write As hlHandle
            
        Case "O-W:S"
            Open slPathFile For Output Access Write Shared As hlHandle
        Case "O-W:B"
            Open slPathFile For Output Access Write Lock Read Write As hlHandle
        Case "O-W:R"
            Open slPathFile For Output Access Write Lock Read As hlHandle
        Case "O-W:W"
            Open slPathFile For Output Access Write Lock Write As hlHandle
            
        Case "R-B:S"
            Open slPathFile For Random Access Read Write Shared As hlHandle
        Case "R-B:B"
            Open slPathFile For Random Access Read Write Lock Read Write As hlHandle
        Case "R-B:R"
            Open slPathFile For Random Access Read Write Lock Read As hlHandle
        Case "R-B:W"
            Open slPathFile For Random Access Read Write Lock Write As hlHandle
            
        Case "R-R:S"
            Open slPathFile For Random Access Read Shared As hlHandle
        Case "R-R:B"
            Open slPathFile For Random Access Read Lock Read Write As hlHandle
        Case "R-R:R"
            Open slPathFile For Random Access Read Lock Read As hlHandle
        Case "R-R:W"
            Open slPathFile For Random Access Read Lock Write As hlHandle
            
        Case "R-W:S"
            Open slPathFile For Random Access Write Shared As hlHandle
        Case "R-W:B"
            Open slPathFile For Random Access Write Lock Read Write As hlHandle
        Case "R-W:R"
            Open slPathFile For Random Access Write Lock Read As hlHandle
        Case "R-W:W"
            Open slPathFile For Random Access Write Lock Write As hlHandle
            
        Case Else
            gLogMsg "Unknown file open type: " & slForClause, "AffErrorLog.Txt", False
            gFileOpen = -1
    End Select
    On Error GoTo 0
    Exit Function
gFileOpenErr:
    'On Error GoTo 0 -- removed 11/3/20 - this sets err.Number to 0
    If err.Number <> 0 Then
        gFileOpen = err.Number
    Else
        gFileOpen = -1
    End If
    On Error GoTo 0
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gGetCSIStamp                    *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get stamp for file from csi_get *
'*                                                     *
'*******************************************************
Function gGetCSIName(slID As String) As String
    Dim slName As String
    Dim ilRet As Integer

    slName = Space$(101)
    ilRet = csiGetName(slID, slName)
    If ilRet = 0 Then
        gGetCSIName = Trim$(Left$(slName, InStr(slName, Chr$(0)) - 1))
    Else
        gGetCSIName = ""
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gIMoveListBox                   *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate master list box with   *
'*                     name and code and populate local*
'*                     list box with name.  The code   *
'*                     number must be the first two    *
'*                     bytes within the record.        *
'*                                                     *
'*******************************************************
Function gIMoveListBox(frm As Form, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, slFileName As String, ilNameOffset As Integer, ilNameLength As Integer, ilFilterType() As Integer, slFilter() As String, ilFilterOffset() As Integer) As Integer
'
'   ilFilter(0) = CHARFILTER
'   sFilter(0) = "S"
'   iFilterOffset(0) = 2
'   ilRet = gIMoveListBox (MainForm, cbcLocal, tlSortCode(), slSortCodeTag, sFileName, iNameOffset, iNameLength, ilFilter(), sFilter(), iFilterOffset())
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       cbcLocal (I/O)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       sFileName (I)- File name
'       iNameLength (I)- Number of characters in the name field to be used to populate list box
'       iNameOffset (I)- Starting offset location of the name (first location in the file is 0)
'       ilFilter() (I)- NOFILTER = no filter required
'                         CHARFILTER = one char filter
'                         INTEGERFILTER = integer filter
'                         STRINGFILTER = string filter
'                         CHARFILTERNOT = one char filter (all records except character specified)
'                         INTEGERFILTERNOT = integer filter  (all records except integer specified)
'                         STRINGFILTERNOT = string filter (all records except string specified)
'       sFilter() (I)- NULL string("") or selective record character (if                                        '       CHARFILTER or NOT) or Integer as string (if INTEGERFILTERor                     '       NOT) or String (if STRINGFILTER or NOT, length is obtained from the string)
'       iFilterOffset() (I)- 0 or offset in the record from 0 to character location
'       ilRet (O)- Error code (0 if no error)
'

    Dim illoop As Integer   'For loop control variable
    Dim slNameCode As String    'Name and code values from master list box
    Dim ilRet As Integer    'Return value from parse
    Dim slName As String    'Name value from master list box
    Dim llLen As Long

    On Error GoTo IMoveErr
    ilRet = gIPopListBox(frm, tlSortCode(), slSortCodeTag, slFileName, ilNameOffset, ilNameLength, ilFilterType(), slFilter(), ilFilterOffset())    'Repopulate Traffic list box if required
    If Not ilRet Then
        If lbcLocal.ListCount > 0 Then
            gIMoveListBox = CP_MSG_NOPOPREQ
            Exit Function
        End If
    End If
    On Error GoTo 0
    lbcLocal.Clear
    llLen = 0
    For illoop = 0 To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gIMoveListBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    gIMoveListBox = CP_MSG_POPREQ
    Exit Function
IMoveErr:
    gDbg_HandleError "PopSubs: gIMoveListBox"
'    gIMoveListBox = CP_MSG_NOSHOW
'    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gIPopListBox                    *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with name and *
'*                     code number.  The code number   *
'*                     must be the first two bytes     *
'*                     within the record.              *
'*                                                     *
'*******************************************************
Function gIPopListBox(frm As Form, tlSortCode() As SORTCODE, slSortCodeTag As String, slFileName As String, ilNameOffset As Integer, ilNameLength As Integer, ilFilterType() As Integer, slFilter() As String, ilFilterOffset() As Integer) As Integer
'
'   ilFilter(0) = CHARFILTER
'   sFilter(0) = "S"
'   iFilterOffset(0) = 2
'   ilRet = gIPopListBox (MainForm, tlSortCode(), slSortCodeTag, sFileName, iNameOffset, iNameLength, ilFilter(), sFilter(), iFilterOffset())
'   Where:
'       MainForm (I)- Name of Form to unload if error
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       sFileName (I)- File name
'       iNameLength (I)- Number of characters in the name field to be used to populate list box
'       iNameOffset (I)- Starting offset location of the name (first location in the file is 0)
'       ilFilter() (I)- NOFILTER = no filter required
'                          CHARFILTER = one char filter
'                          INTEGERFILTER = integer filter
'                          STRINGFILTER = string filter
'                         CHARFILTERNOT = one char filter (all records except character specified)
'                         INTEGERFILTERNOT = integer filter  (all records except integer specified)
'                         STRINGFILTERNOT = string filter (all records except string specified)
'       sFilter() (I)- NULL string("") or selective record character (if                                        '       CHARFILTER or NOT) or Integer as string (if INTEGERFILTERor                     '       NOT) or String (if STRINGFILTER or NOT, length is obtained from the string)
'       iFilterOffset() (I)- 0 or offset in the record from 0 to character location
'       ilRet (O)- True=list was either populated or repolpulated
'                  False=List was OK- it didn't require populating
'

    Dim slStamp As String   'Date/Time stamp for file
    Dim hlFile As Integer
    Dim ilRecLen As Integer        'Record length
    Dim slName As String   'Record Name field
    Dim ilRet As Integer 'btrieve status
    Dim llNoRec As Long 'Number of records
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlStrTypeBuff As POPSTRINGTYPE   'Type field record
    Dim tlCodeNameBuff As POPICODENAME   'Code/Name fields record
    Dim ilLen As Integer   'record length
    Dim llRecPos As Long    'Physical rec location
    Dim illoop As Integer
    Dim slFilterStamp As String
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim slStr As String
    Dim blFilterMissedReason As Boolean
    Dim ilLowLimit As Integer

    llLen = 0
    slStamp = gFileDateTime(sgDBPath & slFileName)
    slFilterStamp = ""
    blFilterMissedReason = False
    For illoop = LBound(ilFilterType) To UBound(ilFilterType) Step 1
        Select Case ilFilterType(illoop) And &HFF
            Case CHARFILTER
                slFilterStamp = slFilterStamp & " C" & slFilter(illoop) & Trim$(str$(ilFilterOffset(illoop)))
            Case INTEGERFILTER
                slFilterStamp = slFilterStamp & " I" & slFilter(illoop) & Trim$(str$(ilFilterOffset(illoop)))
            Case STRINGFILTER
                slFilterStamp = slFilterStamp & " S" & slFilter(illoop) & Trim$(str$(ilFilterOffset(illoop)))
            Case CHARFILTERNOT
                slFilterStamp = slFilterStamp & "NC" & slFilter(illoop) & Trim$(str$(ilFilterOffset(illoop)))
                If illoop = LBound(ilFilterType) + 1 Then
                    If (slFilter(LBound(ilFilterType)) = "M") And (slFilter(LBound(ilFilterType) + 1) = "A") Then
                        blFilterMissedReason = True
                    End If
                End If
            Case INTEGERFILTERNOT
                slFilterStamp = slFilterStamp & "NI" & slFilter(illoop) & Trim$(str$(ilFilterOffset(illoop)))
            Case STRINGFILTERNOT
                slFilterStamp = slFilterStamp & "NS" & slFilter(illoop) & Trim$(str$(ilFilterOffset(illoop)))
        End Select
    Next illoop
    slStamp = slStamp & slFilterStamp


    'On Error GoTo gIPopListBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If slStamp = slSortCodeTag Then
            gIPopListBox = False
            Exit Function
        End If
    End If
    gIPopListBox = True
    ilSortCode = 0
    ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    slSortCodeTag = slStamp
    'Potential
    If (slFilter(LBound(slFilter)) = "P") And (StrComp(Trim$(slFileName), "Mnf.btr", 1) = 0) Then
        ilRet = gObtainMnfForType("P", sgPotMnfStamp, tgPotMnf())
        If ilRet = False Then
            Exit Function
        End If
        ReDim tlSortCode(0 To UBound(tgPotMnf) - LBound(tgPotMnf)) As SORTCODE
        For illoop = LBound(tgPotMnf) To UBound(tgPotMnf) - 1 Step 1
            slName = tgPotMnf(illoop).sName & "\" & Trim$(str$(tgPotMnf(illoop).iCode))
            tlSortCode(ilSortCode).sKey = slName
            ilSortCode = ilSortCode + 1
        Next illoop
        If UBound(tlSortCode) - 1 > 0 Then
            ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
        End If
        Exit Function
    End If
    'Business Category
    If (slFilter(LBound(slFilter)) = "B") And (StrComp(Trim$(slFileName), "Mnf.btr", 1) = 0) Then
        ilRet = gObtainMnfForType("B", sgBusCatMnfStamp, tgBusCatMnf())
        If ilRet = False Then
            Exit Function
        End If
        ReDim tlSortCode(0 To UBound(tgBusCatMnf) - LBound(tgBusCatMnf)) As SORTCODE
        For illoop = LBound(tgBusCatMnf) To UBound(tgBusCatMnf) - 1 Step 1
            slName = tgBusCatMnf(illoop).sName & "\" & Trim$(str$(tgBusCatMnf(illoop).iCode))
            tlSortCode(ilSortCode).sKey = slName
            ilSortCode = ilSortCode + 1
        Next illoop
        If UBound(tlSortCode) - 1 > 0 Then
            ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
        End If
        Exit Function
    End If
    'Exclude
    If (slFilter(LBound(slFilter)) = "X") And (StrComp(Trim$(slFileName), "Mnf.btr", 1) = 0) Then
        ilRet = gObtainMnfForType("X", sgExclMnfStamp, tgExclMnf())
        If ilRet = False Then
            Exit Function
        End If
        ReDim tlSortCode(0 To UBound(tgExclMnf) - LBound(tgExclMnf)) As SORTCODE
        For illoop = LBound(tgExclMnf) To UBound(tgExclMnf) - 1 Step 1
            slName = tgExclMnf(illoop).sName & "\" & Trim$(str$(tgExclMnf(illoop).iCode))
            tlSortCode(ilSortCode).sKey = slName
            ilSortCode = ilSortCode + 1
        Next illoop
        If UBound(tlSortCode) - 1 > 0 Then
            ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
        End If
        Exit Function
    End If
    'Missed Reason excluding Station reasons
    If (slFilter(LBound(slFilter)) = "M") And (blFilterMissedReason) And (StrComp(Trim$(slFileName), "Mnf.btr", 1) = 0) Then
        slStr = ""
        'ReDim tgMRMnf(1 To 1) As MNF
        ReDim tgMRMnf(0 To 0) As MNF
        ilRet = gObtainMnfForType("M", slStr, tgMRMnf())
        If ilRet = False Then
            Exit Function
        End If
        ReDim tlSortCode(0 To UBound(tgMRMnf) - LBound(tgMRMnf)) As SORTCODE
        For illoop = LBound(tgMRMnf) To UBound(tgMRMnf) - 1 Step 1
            If (Trim$(tgMRMnf(illoop).sCodeStn) <> "A") And (Trim$(tgMRMnf(illoop).sCodeStn) <> "R") Then
                slName = tgMRMnf(illoop).sName & "\" & Trim$(str$(tgMRMnf(illoop).iCode))
                tlSortCode(ilSortCode).sKey = slName
                ilSortCode = ilSortCode + 1
            End If
        Next illoop
        ReDim Preserve tlSortCode(0 To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - 1 > 0 Then
            ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
        End If
        Exit Function
    End If
    'Competitives
    If (slFilter(LBound(slFilter)) = "C") And (StrComp(Trim$(slFileName), "Mnf.btr", 1) = 0) Then
        ilRet = gObtainComp()
        If ilRet = False Then
            Exit Function
        End If
        ReDim tlSortCode(0 To UBound(tgCompMnf) - LBound(tgCompMnf)) As SORTCODE
        For illoop = LBound(tgCompMnf) To UBound(tgCompMnf) - 1 Step 1
            slName = tgCompMnf(illoop).sName & "\" & Trim$(str$(tgCompMnf(illoop).iCode))
            tlSortCode(ilSortCode).sKey = slName
            ilSortCode = ilSortCode + 1
        Next illoop
        If UBound(tlSortCode) - 1 > 0 Then
            ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
        End If
        Exit Function
    End If
    'Avails
    If (StrComp(Trim$(slFileName), "Anf.btr", 1) = 0) And (Trim$(slFilter(LBound(slFilter))) = "") Then
        ilRet = gObtainAvail()
        If ilRet = False Then
            Exit Function
        End If
        ReDim tlSortCode(0 To UBound(tgAvailAnf) - LBound(tgAvailAnf)) As SORTCODE
        For illoop = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1 Step 1
            slName = tgAvailAnf(illoop).sName & "\" & Trim$(str$(tgAvailAnf(illoop).iCode))
            tlSortCode(ilSortCode).sKey = slName
            ilSortCode = ilSortCode + 1
        Next illoop
        If UBound(tlSortCode) - 1 > 0 Then
            ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
        End If
        Exit Function
    End If
    hlFile = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo IPopListErr
    gBtrvErrorMsg ilRet, "gIPopListBox (btrOpen):" & slFileName, frm
    On Error GoTo 0
    ilRecLen = btrRecordLength(hlFile)  'Get and save record length
    If ilRecLen > Len(tmIRecBuff) Then
        ilRecLen = Len(tmIRecBuff)
    End If
    ilLen = 2 + ilNameLength  'Extract operation record size
    llNoRec = gExtNoRec(ilLen) 'btrRecords(hlFile) 'Obtain number of records
    btrExtClear hlFile   'Clear any previous extend operation
    ilRet = btrGetFirst(hlFile, tmIRecBuff, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlFile)
        On Error GoTo IPopListErr
        gBtrvErrorMsg ilRet, "gIPopListBox (btrReset):" & slFileName, frm
        On Error GoTo 0
        btrDestroy hlFile
        Exit Function
    Else
        On Error GoTo IPopListErr
        gBtrvErrorMsg ilRet, "gIPopListBox (btrGetFirst):" & slFileName, frm
        On Error GoTo 0
    End If
    'Extract information from first record if required as the extend operation will bypass this record
    Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "POPICODENAMEPK", POPICODENAMEPK) 'Set extract limits (all records)
    For illoop = LBound(ilFilterType) To UBound(ilFilterType) Step 1
        Select Case ilFilterType(illoop) And &HFF
            Case CHARFILTER
                tlCharTypeBuff.sType = slFilter(illoop)    'Extract all matching records
                If illoop = UBound(ilFilterType) Then
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
                Else
                    If ilFilterType(illoop) >= 256 Then
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), 1, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlCharTypeBuff, 1)
                    Else
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                    End If
                End If
            Case INTEGERFILTER
                tlIntTypeBuff.iType = CInt(slFilter(illoop))
                If illoop = UBound(ilFilterType) Then
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilFilterOffset(illoop), 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
                Else
                    If ilFilterType(illoop) >= 256 Then
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilFilterOffset(illoop), 2, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlIntTypeBuff, 2)
                    Else
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilFilterOffset(illoop), 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                    End If
                End If
            Case STRINGFILTER
                tlStrTypeBuff.sType = slFilter(illoop)    'Extract all matching records
                If illoop = UBound(ilFilterType) Then
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), Len(slFilter(illoop)), BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlStrTypeBuff, Len(slFilter(illoop)))
                Else
                    If ilFilterType(illoop) >= 256 Then
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), Len(slFilter(illoop)), BTRV_EXT_EQUAL, BTRV_EXT_OR, tlStrTypeBuff, Len(slFilter(illoop)))
                    Else
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), Len(slFilter(illoop)), BTRV_EXT_EQUAL, BTRV_EXT_AND, tlStrTypeBuff, Len(slFilter(illoop)))
                    End If
                End If
            Case CHARFILTERNOT
                tlCharTypeBuff.sType = slFilter(illoop)    'Extract all matching records
                If illoop = UBound(ilFilterType) Then
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
                Else
                    If ilFilterType(illoop) >= 256 Then
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_OR, tlCharTypeBuff, 1)
                    Else
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                    End If
                End If
            Case INTEGERFILTERNOT
                tlIntTypeBuff.iType = CInt(slFilter(illoop))
                If illoop = UBound(ilFilterType) Then
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilFilterOffset(illoop), 2, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
                Else
                    If ilFilterType(illoop) >= 256 Then
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilFilterOffset(illoop), 2, BTRV_EXT_NOT_EQUAL, BTRV_EXT_OR, tlIntTypeBuff, 2)
                    Else
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilFilterOffset(illoop), 2, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                    End If
                End If
            Case STRINGFILTERNOT
                tlStrTypeBuff.sType = slFilter(illoop)    'Extract all matching records
                If illoop = UBound(ilFilterType) Then
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), Len(slFilter(illoop)), BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlStrTypeBuff, Len(slFilter(illoop)))
                Else
                    If ilFilterType(illoop) >= 256 Then
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), Len(slFilter(illoop)), BTRV_EXT_NOT_EQUAL, BTRV_EXT_OR, tlStrTypeBuff, Len(slFilter(illoop)))
                    Else
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), Len(slFilter(illoop)), BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlStrTypeBuff, Len(slFilter(illoop)))
                    End If
                End If
        End Select
        If ilFilterType(illoop) <> NOFILTER Then
            On Error GoTo IPopListErr
            gBtrvErrorMsg ilRet, "gIPopListBox (btrExtAddLogicConst):" & slFileName, frm
            On Error GoTo 0
        End If
    Next illoop
    ilRet = btrExtAddField(hlFile, 0, 2)  'Extract iCode field
    On Error GoTo IPopListErr
    gBtrvErrorMsg ilRet, "gIPopListBox (btrExtAddField):" & slFileName, frm
    On Error GoTo 0
    ilRet = btrExtAddField(hlFile, ilNameOffset, ilNameLength) 'Extract Name field
    On Error GoTo IPopListErr
    gBtrvErrorMsg ilRet, "gIPopListBox (btrExtAddField):" & slFileName, frm
    On Error GoTo 0
    'ilRet = btrExtGetNextExt(hlFile)  'Extract records
    ilRet = btrExtGetNext(hlFile, tlCodeNameBuff, ilLen, llRecPos)    'Extract record
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo IPopListErr
        gBtrvErrorMsg ilRet, "gIPopListBox (btrExtGetNextExt):" & slFileName, frm
        On Error GoTo 0
        ilLen = 2 + ilNameLength  'Extract operation record size
        'ilRet = btrExtGetFirst(hlFile, tlCodeNameBuff, ilLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlFile, tlCodeNameBuff, ilLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            slName = Mid$(tlCodeNameBuff.sChar, 1, ilNameLength)
            slName = slName & "\" & Trim$(str$(tlCodeNameBuff.iCode))
            'If Not gOkAddStrToListBox(slName, llLen, True) Then
            '    Exit Do
            'End If
            'lbcCtrl.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
            tlSortCode(ilSortCode).sKey = slName
            If ilSortCode >= UBound(tlSortCode) Then
                ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 100) As SORTCODE
            End If
            ilSortCode = ilSortCode + 1
            ilRet = btrExtGetNext(hlFile, tlCodeNameBuff, ilLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlFile, tlCodeNameBuff, ilLen, llRecPos)
            Loop
        Loop
    End If
    'Sort then output new headers and lines
    ReDim Preserve tlSortCode(0 To ilSortCode) As SORTCODE
    If UBound(tlSortCode) - 1 > 0 Then
        ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
    End If
    ilRet = btrClose(hlFile)
    On Error GoTo IPopListErr
    gBtrvErrorMsg ilRet, "gIPopListBox (btrReset):" & slFileName, frm
    On Error GoTo 0
    btrDestroy hlFile
    Exit Function
IPopListErr:
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
    gDbg_HandleError "PopSubs: gIPopListBox"
gIPopListBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gLMoveListBox                   *
'*                                                     *
'*             Created:4/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate master list box with   *
'*                     name and code and populate local*
'*                     list box with name.  The code   *
'*                     number must be the first four   *
'*                     bytes within the record.        *
'*                                                     *
'*******************************************************
Function gLMoveListBox(frm As Form, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, slFileName As String, ilNameOffset As Integer, ilNameLength As Integer, ilFilterType() As Integer, slFilter() As String, ilFilterOffset() As Integer) As Integer
'
'   ilFilter(0) = CHARFILTER
'   sFilter(0) = "S"
'   iFilterOffset(0) = 2
'   ilRet = gLMoveListBox (MainForm, cbcLocal, tlSortCode(), slSortCodeTag, sFileName, iNameOffset, iNameLength, ilFilter(), sFilter(), iFilterOffset())
'   Where:
'       MainForm (I)- Name of Form to unload if error
'       cbcLocal (I/O)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       sFileName (I)- File name
'       iNameLength (I)- Number of characters in the name field to be used to populate list box
'       iNameOffset (I)- Starting offset location of the name (first location in the file is 0)
'       ilFilter() (I)- NOFILTER = no filter required
'                         CHARFILTER = one char filter
'                         INTEGERFILTER = integer filter
'                         STRINGFILTER = string filter
'                         CHARFILTERNOT = one char filter (all records except character specified)
'                         INTEGERFILTERNOT = integer filter  (all records except integer specified)
'                         STRINGFILTERNOT = string filter (all records except string specified)
'       sFilter() (I)- NULL string("") or selective record character (if                                        '       CHARFILTER or NOT) or Integer as string (if INTEGERFILTERor                     '       NOT) or String (if STRINGFILTER or NOT, length is obtained from the string)
'       iFilterOffset() (I)- 0 or offset in the record from 0 to character location
'       ilRet (O)- Error code (0 if no error)
'

    Dim illoop As Integer   'For loop control variable
    Dim slNameCode As String    'Name and code values from master list box
    Dim ilRet As Integer    'Return value from parse
    Dim slName As String    'Name value from master list box
    Dim llLen As Long

    On Error GoTo LMoveErr
    ilRet = gLPopListBox(frm, tlSortCode(), slSortCodeTag, slFileName, ilNameOffset, ilNameLength, ilFilterType(), slFilter(), ilFilterOffset())    'Repopulate Traffic list box if required
    If Not ilRet Then
        If lbcLocal.ListCount > 0 Then
            gLMoveListBox = CP_MSG_NOPOPREQ
            Exit Function
        End If
    End If
    On Error GoTo 0
    lbcLocal.Clear
    llLen = 0
    For illoop = 0 To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gLMoveListBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    gLMoveListBox = CP_MSG_POPREQ
    Exit Function
LMoveErr:
    gDbg_HandleError "PopSubs: gLMoveListBox"
'    gLMoveListBox = CP_MSG_NOSHOW
'    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gLPopListBox                    *
'*                                                     *
'*             Created:4/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with name and *
'*                     code number.  The code number   *
'*                     must be the first four bytes    *
'*                     within the record.              *
'*                                                     *
'*******************************************************
Function gLPopListBox(frm As Form, tlSortCode() As SORTCODE, slSortCodeTag As String, slFileName As String, ilNameOffset As Integer, ilNameLength As Integer, ilFilterType() As Integer, slFilter() As String, ilFilterOffset() As Integer)
'
'   ilFilter(0) = CHARFILTER
'   sFilter(0) = "S"
'   iFilterOffset(0) = 2
'   ilRet = gLPopListBox (MainForm, tlSortCode(), slSortCodeTag, sFileName, iNameOffset, iNameLength, ilFilter(), sFilter(), iFilterOffset())
'   Where:
'       MainForm (I)- Name of Form to unload if error
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       sFileName (I)- File name
'       iNameLength (I)- Number of characters in the name field to be used to populate list box
'       iNameOffset (I)- Starting offset location of the name (first location in the file is 0)
'       ilFilter (I)- NOFILTER = no filter required
'                        CHARFILTER = one char filter
'                        INTEGERFILTER = integer filter
'                        STRINGFILTER = string filter
'                        CHARFILTERNOT = one char filter (all records except character specified)
'                        INTEGERFILTERNOT = integer filter  (all records except integer specified)
'                        STRINGFILTERNOT = string filter (all records except string specified)
'       sFilter() (I)- NULL string("") or selective record character (if                                        '       CHARFILTER or NOT) or Integer as string (if INTEGERFILTERor                     '       NOT) or String (if STRINGFILTER or NOT, length is obtained from the string)
'       iFilterOffset (I)- 0 or offset in the record from 0 to character location
'       ilRet (O)- True=list was either populated or repolpulated
'                       False=List was OK- it didn't require populating
'
    Dim slStamp As String   'Date/Time stamp for file
    Dim hlFile As Integer
    Dim ilRecLen As Integer        'Record length
    Dim slName As String   'Record Name field
    Dim ilRet As Integer 'btrieve status
    Dim llNoRec As Long 'Number of records
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlStrTypeBuff As POPSTRINGTYPE   'Type field record
    Dim tlCodeNameBuff As POPLCODENAME
    Dim ilLen As Integer   'record length
    Dim llRecPos As Long    'Physical rec location
    Dim illoop As Integer
    Dim slFilterStamp As String
    Dim llLen As Long
    'Dim ilSortCode As Integer
    Dim llSortCode As Long
    Dim ilLowLimit As Integer

    llLen = 0
    slStamp = gFileDateTime(sgDBPath & slFileName)
    slFilterStamp = ""
    For illoop = LBound(ilFilterType) To UBound(ilFilterType) Step 1
        Select Case ilFilterType(illoop) And &HFF
            Case CHARFILTER
                slFilterStamp = slFilterStamp & " C" & slFilter(illoop) & Trim$(str$(ilFilterOffset(illoop)))
            Case INTEGERFILTER
                slFilterStamp = slFilterStamp & " I" & slFilter(illoop) & Trim$(str$(ilFilterOffset(illoop)))
            Case STRINGFILTER
                slFilterStamp = slFilterStamp & " S" & slFilter(illoop) & Trim$(str$(ilFilterOffset(illoop)))
            Case CHARFILTERNOT
                slFilterStamp = slFilterStamp & "NC" & slFilter(illoop) & Trim$(str$(ilFilterOffset(illoop)))
            Case INTEGERFILTERNOT
                slFilterStamp = slFilterStamp & "NI" & slFilter(illoop) & Trim$(str$(ilFilterOffset(illoop)))
            Case STRINGFILTERNOT
                slFilterStamp = slFilterStamp & "NS" & slFilter(illoop) & Trim$(str$(ilFilterOffset(illoop)))
        End Select
    Next illoop
    slStamp = slStamp & slFilterStamp

    'On Error GoTo gLPopListBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If slStamp = slSortCodeTag Then
            gLPopListBox = False
            Exit Function
        End If
    End If
    gLPopListBox = True
    hlFile = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlFile, "", sgDBPath & slFileName, BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo LPopListErr
    gBtrvErrorMsg ilRet, "gLPopListBox (btrOpen)" & slFileName, frm
    On Error GoTo 0
    ilRecLen = btrRecordLength(hlFile)  'Get and save record length
    'ilSortCode = 0
    llSortCode = 0
    ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    slSortCodeTag = slStamp
    ilLen = 4 + ilNameLength  'Extract operation record size
    llNoRec = gExtNoRec(ilLen) 'btrRecords(hlFile) 'Obtain number of records
    'Variable length records- must read in the whole record
    If (StrComp(slFileName, "Cmf.Btr", 1) = 0) Or (StrComp(slFileName, "Cef.Btr", 1) = 0) Or (StrComp(slFileName, "Csf.Btr", 1) = 0) Then
        ilRecLen = ilRecLen + 5004
    End If
    If ilRecLen > Len(tmLRecBuff) Then
        ilRecLen = Len(tmLRecBuff)
    End If
    btrExtClear hlFile   'Clear any previous extend operation
    ilRet = btrGetFirst(hlFile, tmLRecBuff, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlFile)
        On Error GoTo LPopListErr
        gBtrvErrorMsg ilRet, "gLPopListBox (btrReset):" & slFileName, frm
        On Error GoTo 0
        btrDestroy hlFile
        Exit Function
    Else
        On Error GoTo LPopListErr
        gBtrvErrorMsg ilRet, "gLPopListBox (btrGetFirst):" & slFileName, frm
        On Error GoTo 0
    End If
    'Extract information from first record if required as the extend operation will bypass this record
    Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "POPLCODENAMEPK", POPLCODENAMEPK) 'Set extract limits (all records)
    For illoop = LBound(ilFilterType) To UBound(ilFilterType) Step 1
        Select Case ilFilterType(illoop) And &HFF
            Case CHARFILTER
                tlCharTypeBuff.sType = slFilter(illoop)    'Extract all matching records
                If illoop = UBound(ilFilterType) Then
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
                Else
                    If ilFilterType(illoop) >= 256 Then
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), 1, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlCharTypeBuff, 1)
                    Else
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                    End If
                End If
            Case INTEGERFILTER
                tlIntTypeBuff.iType = CInt(slFilter(illoop))
                If illoop = UBound(ilFilterType) Then
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilFilterOffset(illoop), 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
                Else
                    If ilFilterType(illoop) >= 256 Then
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilFilterOffset(illoop), 2, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlIntTypeBuff, 2)
                    Else
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilFilterOffset(illoop), 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                    End If
                End If
            Case STRINGFILTER
                tlStrTypeBuff.sType = slFilter(illoop)    'Extract all matching records
                If illoop = UBound(ilFilterType) Then
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), Len(slFilter(illoop)), BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlStrTypeBuff, Len(slFilter(illoop)))
                Else
                    If ilFilterType(illoop) >= 256 Then
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), Len(slFilter(illoop)), BTRV_EXT_EQUAL, BTRV_EXT_OR, tlStrTypeBuff, Len(slFilter(illoop)))
                    Else
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), Len(slFilter(illoop)), BTRV_EXT_EQUAL, BTRV_EXT_AND, tlStrTypeBuff, Len(slFilter(illoop)))
                    End If
                End If
            Case CHARFILTERNOT
                tlCharTypeBuff.sType = slFilter(illoop)    'Extract all matching records
                If illoop = UBound(ilFilterType) Then
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
                Else
                    If ilFilterType(illoop) >= 256 Then
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_OR, tlCharTypeBuff, 1)
                    Else
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
                    End If
                End If
            Case INTEGERFILTERNOT
                tlIntTypeBuff.iType = CInt(slFilter(illoop))
                If illoop = UBound(ilFilterType) Then
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilFilterOffset(illoop), 2, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
                Else
                    If ilFilterType(illoop) >= 256 Then
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilFilterOffset(illoop), 2, BTRV_EXT_NOT_EQUAL, BTRV_EXT_OR, tlIntTypeBuff, 2)
                    Else
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilFilterOffset(illoop), 2, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                    End If
                End If
            Case STRINGFILTERNOT
                tlStrTypeBuff.sType = slFilter(illoop)    'Extract all matching records
                If illoop = UBound(ilFilterType) Then
                    ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), Len(slFilter(illoop)), BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlStrTypeBuff, Len(slFilter(illoop)))
                Else
                    If ilFilterType(illoop) >= 256 Then
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), Len(slFilter(illoop)), BTRV_EXT_NOT_EQUAL, BTRV_EXT_OR, tlStrTypeBuff, Len(slFilter(illoop)))
                    Else
                        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilFilterOffset(illoop), Len(slFilter(illoop)), BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlStrTypeBuff, Len(slFilter(illoop)))
                    End If
                End If
        End Select
        If ilFilterType(illoop) <> NOFILTER Then
            On Error GoTo LPopListErr
            gBtrvErrorMsg ilRet, "gLPopListBox (btrExtAddLogicConst):" & slFileName, frm
            On Error GoTo 0
        End If
    Next illoop
    ilRet = btrExtAddField(hlFile, 0, 4)
    On Error GoTo LPopListErr
    gBtrvErrorMsg ilRet, "gLPopListBox (btrExtAddField)" & slFileName, frm
    On Error GoTo 0
    ilRet = btrExtAddField(hlFile, ilNameOffset, ilNameLength) 'Extract Name field
    On Error GoTo LPopListErr
    gBtrvErrorMsg ilRet, "gLPopListBox (btrExtAddField)" & slFileName, frm
    On Error GoTo 0
    'ilRet = btrExtGetNextExt(hlFile)  'Extract records
    ilRet = btrExtGetNext(hlFile, tlCodeNameBuff, ilLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo LPopListErr
        gBtrvErrorMsg ilRet, "gLPopListBox (btrExtGetNextExt)" & slFileName, frm
        On Error GoTo 0
        ilLen = 4 + ilNameLength  'Extract operation record size
        'ilRet = btrExtGetFirst(hlFile, tlCodeNameBuff, ilLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlFile, tlCodeNameBuff, ilLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            slName = Trim$(Mid$(tlCodeNameBuff.sChar, 1, ilNameLength))
            slName = slName & "\" & Trim$(str$(tlCodeNameBuff.lCode))
            'If Not gOkAddStrToListBox(slName, llLen, True) Then
            '    Exit Do
            'End If
            'lbcCtrl.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
            tlSortCode(llSortCode).sKey = slName
            If llSortCode >= UBound(tlSortCode) Then
                ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 1) As SORTCODE
            End If
            llSortCode = llSortCode + 1
            ilRet = btrExtGetNext(hlFile, tlCodeNameBuff, ilLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlFile, tlCodeNameBuff, ilLen, llRecPos)
            Loop
        Loop
    End If
    'Sort then output new headers and lines
    ReDim Preserve tlSortCode(0 To llSortCode) As SORTCODE
    If UBound(tlSortCode) - 1 > 0 Then
        ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
    End If
    ilRet = btrClose(hlFile)
    On Error GoTo LPopListErr
    gBtrvErrorMsg ilRet, "gLPopListBox (btrReset): " & slFileName, frm
    On Error GoTo 0
    btrDestroy hlFile
    Exit Function
LPopListErr:
    ilRet = btrClose(hlFile)
    btrDestroy hlFile
    gDbg_HandleError "PopSubs: gLPopListBox"
gLPopListBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainAdvt                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgCommAdf for          *
'*                     collection                      *
'*                                                     *
'*******************************************************
Function gObtainAdvt() As Integer
'
'   ilRet = gObtainAdvt ()
'   Where:
'       tgCommAdf() (I)- ADFEXT record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Slf date/time stamp
    Dim hlAdf As Integer        'Slf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlAdf As ADF
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim ilLowLimit As Integer

    slStamp = gFileDateTime(sgDBPath & "Adf.Btr")

    'On Error GoTo gObtainAdvtErr2
    'ilRet = 0
    'ilLowLimit = LBound(tgCommAdf)
    'If ilRet <> 0 Then
    '    sgCommAdfStamp = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tgCommAdf).Ptr <> 0 Then
        ilLowLimit = LBound(tgCommAdf)
    Else
        sgCommAdfStamp = ""
        ilLowLimit = 0
    End If

    hlAdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    If (sgCommAdfStamp <> "") And (sgCommAdfStamp <> "~") Then
        ''If defined, assume Ok as items added with gBinarySearchAdf
        ''If StrComp(slStamp, sgCommAdfStamp, 1) = 0 Then
            ''If UBound(tgCommAdf) > 1 Then
            'If UBound(tgCommAdf) - 1 = btrRecords(hlAdf) Then
            If UBound(tgCommAdf) - ilLowLimit = btrRecords(hlAdf) Then
                btrDestroy hlAdf
                gObtainAdvt = True
                Exit Function
            End If
            'End If
        'End If
    End If
    igInternalAdfCount = 0
    'ReDim tgCommAdf(1 To 1) As ADFEXT
    ReDim tgCommAdf(ilLowLimit To ilLowLimit) As ADFEXT
    ilRecLen = Len(tlAdf) 'btrRecordLength(hlAdf)  'Get and save record length
    sgCommAdfStamp = slStamp
    ilUpperBound = UBound(tgCommAdf)
    ilExtLen = Len(tgCommAdf(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hlAdf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlAdf, tlAdf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        gObtainAdvt = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainAdvt = False
            ilRet = btrClose(hlAdf)
            btrDestroy hlAdf
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlAdf, llNoRec, -1, "UC", "ADFEXTPK", ADFEXTPK) 'Set extract limits (all records)
    ilOffSet = gFieldOffset("Adf", "AdfCode")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 2)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Adf", "AdfName")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 30)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Adf", "AdfAbbr")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 7)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Adf", "AdfCreditRestr")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    'TTP 10460 - RAB export: add advertiser credit status and agency credit status
    ilOffSet = gFieldOffset("Adf", "AdfCrdApp")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    
    ilOffSet = gFieldOffset("Adf", "AdfMnfSort")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 2) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Adf", "AdfBilAgyDir")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Adf", "AdfState")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Adf", "AdfAllowRepMG")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Adf", "AdfBonusOnInv")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Adf", "AdfRepInvGen")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If

    'Get only the first byte of the address, used to check on Non-Direct to see if they were direct at one time in Collect
    ilOffSet = gFieldOffset("Adf", "AdfCntrAddr1")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 40) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If

    '5-26-06 Adv Political flag
    ilOffSet = gFieldOffset("Adf", "AdfPolitical")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If

    ilOffSet = gFieldOffset("Adf", "AdfAddrID")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 9)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Adf", "AdfTrfCode")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 2) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Adf", "AdfArfInvCode")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 2) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Adf", "AdfPnfBuyer")
    ilRet = btrExtAddField(hlAdf, ilOffSet, 2) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    
    ilOffSet = gFieldOffset("Adf", "AdfLastYearNew")     '9-6-11 Last Year Advt became New
    ilRet = btrExtAddField(hlAdf, ilOffSet, 2)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If
    
    ilOffSet = gFieldOffset("Adf", "AdfLastMonthNew")    '9-6-11 Last Year Month Advt became New
    ilRet = btrExtAddField(hlAdf, ilOffSet, 2)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If

    ilOffSet = gFieldOffset("Adf", "AdfMnfInvTerms")    '9-8-13 Inv terms
    ilRet = btrExtAddField(hlAdf, ilOffSet, 2)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If

    ilOffSet = gFieldOffset("Adf", "AdfCodeStn")    '1-9-18 Station Code
    ilRet = btrExtAddField(hlAdf, ilOffSet, 10)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If

    ilOffSet = gFieldOffset("Adf", "AdfPnfpay")    '10-14-19 pnf payables
    ilRet = btrExtAddField(hlAdf, ilOffSet, 2)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAdvt = False
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        Exit Function
    End If

    'ilRet = btrExtGetNextExt(hlAdf)    'Extract record
    ilUpperBound = UBound(tgCommAdf)
    ilRet = btrExtGetNext(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainAdvt = False
            ilRet = btrClose(hlAdf)
            btrDestroy hlAdf
            Exit Function
        End If
        ilUpperBound = UBound(tgCommAdf)
        ilExtLen = Len(tgCommAdf(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If tgCommAdf(ilUpperBound).sRepInvGen <> "E" Then
                igInternalAdfCount = igInternalAdfCount + 1
            End If
            ilUpperBound = ilUpperBound + 1
            'ReDim Preserve tgCommAdf(1 To ilUpperBound) As ADFEXT
            ReDim Preserve tgCommAdf(ilLowLimit To ilUpperBound) As ADFEXT
            ilRet = btrExtGetNext(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlAdf)
    btrDestroy hlAdf
    'Sort by code so that binary search can be used
    'If UBound(tgCommAdf) - 1 > 1 Then
    If UBound(tgCommAdf) - 1 > ilLowLimit Then
        'ArraySortTyp fnAV(tgCommAdf(), 1), UBound(tgCommAdf) - 1, 0, LenB(tgCommAdf(1)), 0, -1, 0
        ArraySortTyp fnAV(tgCommAdf(), ilLowLimit), UBound(tgCommAdf) - ilLowLimit, 0, LenB(tgCommAdf(ilLowLimit)), 0, -1, 0
    End If
    gObtainAdvt = True
    Exit Function
gObtainAdvtErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainAgency                   *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgCommAgf for          *
'*                     collection                      *
'*                                                     *
'*******************************************************
Function gObtainAgency() As Integer
'
'   ilRet = gObtainAgency ()
'   Where:
'       tgCommAgy() (I)- AGFEXT record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Slf date/time stamp
    Dim hlAgf As Integer        'Slf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlagf As AGF
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim ilLowLimit As Integer

    slStamp = gFileDateTime(sgDBPath & "Agf.Btr")

    'On Error GoTo gObtainAgencyErr2
    'ilRet = 0
    'ilLowLimit = LBound(tgCommAgf)
    'If ilRet <> 0 Then
    '    sgCommAgfStamp = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tgCommAgf).Ptr <> 0 Then
        ilLowLimit = LBound(tgCommAgf)
    Else
        sgCommAgfStamp = ""
        ilLowLimit = 0
    End If

    hlAgf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    If (sgCommAgfStamp <> "") And (sgCommAgfStamp <> "~") Then
        ''If defined, assume Ok as items added with gBinarySearchAdf
        ''If StrComp(slStamp, sgCommAgfStamp, 1) = 0 Then
            ''If UBound(tgCommAgf) > 1 Then
            'If UBound(tgCommAgf) - 1 = btrRecords(hlAgf) Then
            If UBound(tgCommAgf) - ilLowLimit = btrRecords(hlAgf) Then
                btrDestroy hlAgf
                gObtainAgency = True
                Exit Function
            End If
            'End If
        'End If
    End If
    'ReDim tgCommAgf(1 To 1) As AGFEXT
    ReDim tgCommAgf(ilLowLimit To ilLowLimit) As AGFEXT
    ilRecLen = Len(tlagf) 'btrRecordLength(hlAgf)  'Get and save record length
    sgCommAgfStamp = slStamp
    ilUpperBound = UBound(tgCommAgf)
    ilExtLen = Len(tgCommAgf(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAgf) 'Obtain number of records
    btrExtClear hlAgf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlAgf, tlagf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        gObtainAgency = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainAgency = False
            ilRet = btrClose(hlAgf)
            btrDestroy hlAgf
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlAgf, llNoRec, -1, "UC", "AGFEXTPK", AGFEXTPK) 'Set extract limits (all records)
    ilOffSet = gFieldOffset("Agf", "AgfCode")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 2)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Agf", "AgfName")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 40)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Agf", "AgfCity")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 5) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Agf", "AgfCreditRestr")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    'TTP 10460 - RAB export: add advertiser credit status and agency credit status
    ilOffSet = gFieldOffset("Agf", "AgfCrdApp")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    
    ilOffSet = gFieldOffset("Agf", "AgfMnfSort")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 2) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Agf", "AgfState")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Agf", "AgfTrfCode")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 2) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Agf", "AgfArfInvCode")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 2) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Agf", "AgfPnfBuyer")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 2) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    
    ilOffSet = gFieldOffset("Agf", "AgfComm")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 2) 'Extract Agy Comm
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    
    ilOffSet = gFieldOffset("Agf", "AgfMnfInvTerms")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 2) 'Extract Agy Comm
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    '10/28/14: get 1 or 2 Place Rating
    ilOffSet = gFieldOffset("Agf", "Agf1or2DigitRating")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 1) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    
    '1-9-18 Station Code
    ilOffSet = gFieldOffset("Agf", "AgfCodeStn")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 10) '
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    
    ilOffSet = gFieldOffset("Agf", "AgfPnfPay")     '10-14-19
    ilRet = btrExtAddField(hlAgf, ilOffSet, 2) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    
    ilOffSet = gFieldOffset("Agf", "AgfRefID")     '4-20-21
    ilRet = btrExtAddField(hlAgf, ilOffSet, 12)  'Extract RefID
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAgency = False
        ilRet = btrClose(hlAgf)
        btrDestroy hlAgf
        Exit Function
    End If
    
    'ilRet = btrExtGetNextExt(hlAgf)    'Extract record
    ilUpperBound = UBound(tgCommAgf)
    ilRet = btrExtGetNext(hlAgf, tgCommAgf(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainAgency = False
            ilRet = btrClose(hlAgf)
            btrDestroy hlAgf
            Exit Function
        End If
        ilUpperBound = UBound(tgCommAgf)
        ilExtLen = Len(tgCommAgf(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlAgf, tgCommAgf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlAgf, tgCommAgf(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilUpperBound = ilUpperBound + 1
            ReDim Preserve tgCommAgf(ilLowLimit To ilUpperBound) As AGFEXT
            ilRet = btrExtGetNext(hlAgf, tgCommAgf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlAgf, tgCommAgf(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlAgf)
    btrDestroy hlAgf
    'Sort by code so that binary search can be used
    'If UBound(tgCommAgf) - 1 > 1 Then
    If UBound(tgCommAgf) - 1 > ilLowLimit Then
        'ArraySortTyp fnAV(tgCommAgf(), 1), UBound(tgCommAgf) - 1, 0, LenB(tgCommAgf(1)), 0, -1, 0
        ArraySortTyp fnAV(tgCommAgf(), ilLowLimit), UBound(tgCommAgf) - ilLowLimit, 0, LenB(tgCommAgf(ilLowLimit)), 0, -1, 0
    End If
    gObtainAgency = True
    Exit Function
gObtainAgencyErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainAvail                    *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgAvailAnf             *
'*                                                     *
'*******************************************************
Function gObtainAvail() As Integer
'
'   ilRet = gObtainAvail ()
'   Where:
'       tgAvailAnf() (I)- MNFCOMPEXT record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlAnf As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlAnf As ANF
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim ilLowLimit As Integer

    slStamp = gFileDateTime(sgDBPath & "Anf.Btr")

    'On Error GoTo gObtainAvailErr2
    'ilRet = 0
    'ilLowLimit = LBound(tgAvailAnf)
    'If ilRet <> 0 Then
    '    sgAvailAnfStamp = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tgAvailAnf).Ptr <> 0 Then
        ilLowLimit = LBound(tgAvailAnf)
    Else
        sgAvailAnfStamp = ""
        ilLowLimit = 0
    End If

    If sgAvailAnfStamp <> "" Then
        If StrComp(slStamp, sgAvailAnfStamp, 1) = 0 Then
            'If UBound(tgAvailAnf) > 1 Then
                gObtainAvail = True
                Exit Function
            'End If
        End If
    End If
    ReDim tgAvailAnf(ilLowLimit To ilLowLimit) As ANF
    hlAnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAvail = False
        ilRet = btrClose(hlAnf)
        btrDestroy hlAnf
        Exit Function
    End If
    ilRecLen = Len(tlAnf) 'btrRecordLength(hlAnf)  'Get and save record length
    sgAvailAnfStamp = slStamp
    ilUpperBound = UBound(tgAvailAnf)
    ilExtLen = Len(tgAvailAnf(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAnf) 'Obtain number of records
    btrExtClear hlAnf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlAnf, tlAnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlAnf)
        btrDestroy hlAnf
        gObtainAvail = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainAvail = False
            ilRet = btrClose(hlAnf)
            btrDestroy hlAnf
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlAnf, llNoRec, -1, "UC", "ANF", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlAnf, ilOffSet, ilRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainAvail = False
        ilRet = btrClose(hlAnf)
        btrDestroy hlAnf
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlAnf)    'Extract record
    ilUpperBound = UBound(tgAvailAnf)
    ilRet = btrExtGetNext(hlAnf, tgAvailAnf(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainAvail = False
            ilRet = btrClose(hlAnf)
            btrDestroy hlAnf
            Exit Function
        End If
        ilUpperBound = UBound(tgAvailAnf)
        ilExtLen = Len(tgAvailAnf(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlAnf, tgAvailAnf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlAnf, tgAvailAnf(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilUpperBound = ilUpperBound + 1
            ReDim Preserve tgAvailAnf(ilLowLimit To ilUpperBound) As ANF
            ilRet = btrExtGetNext(hlAnf, tgAvailAnf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlAnf, tgAvailAnf(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlAnf)
    btrDestroy hlAnf
    gObtainAvail = True
    Exit Function
gObtainAvailErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainChfClf                   *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get Contract (Chf), Lines (Clf) *
'*                                                     *
'*******************************************************
Function gObtainChfClf(hlChf As Integer, hlClf As Integer, llChfCode As Long, ilIncludeHistory As Integer, tlChf As CHF, tlClf() As CLFLIST) As Integer
    Dim ilRet As Integer
    Dim tlTChf As CHF
    Dim tlChfSrchKey As LONGKEY0            'CHF record image
    Dim tlChfSrchKey1 As CHFKEY1    'CntrNo; CntRevNo; PropVer
    Dim tlClfSrchKey As CLFKEY0            'CLF record image
    Dim ilChfRecLen As Integer        'CHF record length
    Dim ilClfRecLen As Integer
    Dim ilLnUpper As Integer
    Dim llNoRec As Long
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim ilAddLn As Integer
    Dim ilTest As Integer
    Dim llCntrNo As Long
    Dim slGpSort As String
    Dim slVehName As String
    Dim slVehSort As String
    Dim slDPSort As String
    Dim slLine As String
    Dim slRevNo As String
    Dim slPropVer As String
    Dim ilVef As Integer
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim ilMnfRecLen As Integer
    Dim tlMnfSrchKey As INTKEY0
    Dim hlRdf As Integer
    Dim tlRdf As RDF
    Dim ilRdfRecLen As Integer
    Dim tlRdfSrchKey As INTKEY0
    Dim slOALineNo As String    'Line number of Parent + A or Line Number of Parent + C or ~~~~ if standard line)
    ReDim tlClf(0 To 0) As CLFLIST
    tlClf(0).iStatus = -1 'Not Used
    tlClf(0).lRecPos = 0
    tlClf(0).iFirstCff = -1
    tlClf(0).iFirstCgf = -1
    tlClf(0).lFirstAvail = -1
    ilChfRecLen = Len(tlChf)
    ilClfRecLen = Len(tlClf(0).ClfRec)
    tlChfSrchKey.lCode = llChfCode
    ilRet = btrGetEqual(hlChf, tlChf, ilChfRecLen, tlChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        gObtainChfClf = False
        Exit Function
    End If
    hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainChfClf = False
        btrDestroy hlMnf
        Exit Function
    End If
    ilMnfRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
    hlRdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainChfClf = False
        btrDestroy hlMnf
        btrDestroy hlRdf
        Exit Function
    End If
    ilRdfRecLen = Len(tlRdf) 'btrRecordLength(hlMnf)  'Get and save record length
    ilRet = gObtainVef()
    llCntrNo = tlChf.lCntrNo
    ilLnUpper = UBound(tlClf)
    tlChfSrchKey1.lCntrNo = llCntrNo
    If (tlChf.sStatus = "O") Or (tlChf.sStatus = "H") Then
        'Find latest contract because it contains all current line from different revisions
        tlChfSrchKey1.iCntRevNo = 32000
        tlChfSrchKey1.iPropVer = 32000
    Else
        tlChfSrchKey1.iCntRevNo = tlChf.iCntRevNo   '32000
        tlChfSrchKey1.iPropVer = tlChf.iPropVer '32000
    End If
    ilRet = btrGetGreaterOrEqual(hlChf, tlTChf, ilChfRecLen, tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tlTChf.lCntrNo = llCntrNo)
        btrExtClear hlClf   'Clear any previous extend operation
        tlClfSrchKey.lChfCode = tlTChf.lCode
        tlClfSrchKey.iLine = 0
        tlClfSrchKey.iCntRevNo = tlTChf.iCntRevNo ' 0 show latest version
        tlClfSrchKey.iPropVer = tlTChf.iPropVer ' 0 show latest version
        ilRet = btrGetGreaterOrEqual(hlClf, tlClf(ilLnUpper).ClfRec, ilClfRecLen, tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If (tlClf(ilLnUpper).ClfRec.lChfCode = tlTChf.lCode) Then
            ilExtLen = Len(tlClf(ilLnUpper).ClfRec)
            llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
            Call btrExtSetBounds(hlClf, llNoRec, -1, "UC", "CLF", "") 'Set extract limits (all records)
            ilOffSet = gFieldOffset("Clf", "ClfChfCode")
            ilRet = btrExtAddLogicConst(hlClf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlTChf.lCode, 4)
            If ilRet <> BTRV_ERR_NONE Then
                btrDestroy hlMnf
                btrDestroy hlRdf
                gObtainChfClf = False
                Exit Function
            End If
            ilRet = btrExtAddField(hlClf, 0, ilExtLen) 'Extract start/end time, and days
            If ilRet <> BTRV_ERR_NONE Then
                btrDestroy hlMnf
                btrDestroy hlRdf
                gObtainChfClf = False
                Exit Function
            End If
            ilRet = btrExtGetNext(hlClf, tlClf(ilLnUpper).ClfRec, ilExtLen, tlClf(ilLnUpper).lRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                'ilRet = btrExtGetFirst(hlClf, tlClfExt, ilExtLen, llRecPos)
                ilExtLen = Len(tlClf(ilLnUpper).ClfRec)
                'If ilRet = BTRV_ERR_REJECT_COUNT Then
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlClf, tlClf(ilLnUpper).ClfRec, ilExtLen, tlClf(ilLnUpper).lRecPos)
                Loop
                'End If
                Do While ilRet = BTRV_ERR_NONE
                    ilAddLn = True
                    If (tlClf(ilLnUpper).ClfRec.iCntRevNo = tlChf.iCntRevNo) And (tlClf(ilLnUpper).ClfRec.iPropVer <> tlChf.iPropVer) Then
                        ilAddLn = False
                    ElseIf (tlClf(ilLnUpper).ClfRec.iCntRevNo > tlChf.iCntRevNo) Then
                        ilAddLn = False
                    Else
                        If Not ilIncludeHistory Then
                            'Test if line added- if so ignore line
                            For ilTest = 0 To ilLnUpper - 1 Step 1
                                If tlClf(ilLnUpper).ClfRec.iLine = tlClf(ilTest).ClfRec.iLine Then
                                    ilAddLn = False
                                    Exit For
                                End If
                            Next ilTest
                        End If
                    End If
                    If ilAddLn Then
                        slOALineNo = "~~~~"
                        'Build sort key (Vehicle Group Sort #; Vehicle Sort #; Vehicle Name; Daypart Sort #; Line #
                        slVehSort = "999"
                        slVehName = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
                        'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If tlClf(ilLnUpper).ClfRec.iVefCode = tgMVef(ilVef).iCode Then
                            ilVef = gBinarySearchVef(tlClf(ilLnUpper).ClfRec.iVefCode)
                            If ilVef <> -1 Then
                                slVehName = tgMVef(ilVef).sName
                                slVehSort = Trim$(str$(tgMVef(ilVef).iSort))
                                tlMnfSrchKey.iCode = tgMVef(ilVef).iOwnerMnfCode
                                If tlMnf.iCode <> tlMnfSrchKey.iCode Then
                                    ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        tlMnf.iGroupNo = 999
                                        tlMnf.iCode = 0
                                    End If
                                End If
                                slGpSort = Trim$(str$(tlMnf.iGroupNo))
                                'Exit For
                            End If
                        'Next ilVef
                        tlRdfSrchKey.iCode = tlClf(ilLnUpper).ClfRec.iRdfCode
                        If tlRdf.iCode <> tlRdfSrchKey.iCode Then
                            ilRet = btrGetEqual(hlRdf, tlRdf, ilRdfRecLen, tlRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then
                                tlRdf.iSortCode = 999
                                tlRdf.iCode = 0
                            End If
                        End If
                        slDPSort = Trim$(str$(tlRdf.iSortCode))
                        Do While Len(slGpSort) < 3
                            slGpSort = "0" & slGpSort
                        Loop
                        Do While Len(slVehSort) < 3
                            slVehSort = "0" & slVehSort
                        Loop
                        Do While Len(slDPSort) < 3
                            slDPSort = "0" & slDPSort
                        Loop
                        slLine = Trim$(str$(tlClf(ilLnUpper).ClfRec.iLine))
                        Do While Len(slLine) < 4
                            slLine = "0" & slLine
                        Loop
                        slRevNo = Trim$(str$(999 - tlClf(ilLnUpper).ClfRec.iCntRevNo))
                        Do While Len(slRevNo) < 3
                            slRevNo = "0" & slRevNo
                        Loop
                        slPropVer = Trim$(str$(tlClf(ilLnUpper).ClfRec.iPropVer))
                        Do While Len(slPropVer) < 3
                            slPropVer = "0" & slPropVer
                        Loop
                        'tlClf(ilLnUpper).sKey = slGpSort & slVehSort & slVehName & slDPSort & slLine & slRevNo & slPropVer

                        tlClf(ilLnUpper).sKey = slOALineNo & "|" & slGpSort & slVehSort & slVehName & slDPSort & slLine & slRevNo & slPropVer
                        tlClf(ilLnUpper).iStatus = 1
                        ilLnUpper = ilLnUpper + 1
                        ReDim Preserve tlClf(0 To ilLnUpper) As CLFLIST
                        tlClf(ilLnUpper).iStatus = -1 'Not Used
                        tlClf(ilLnUpper).iFirstCff = -1
                        tlClf(ilLnUpper).iFirstCgf = -1
                        tlClf(ilLnUpper).lFirstAvail = -1
                        tlClf(ilLnUpper).lRecPos = 0
                    End If
                    ilRet = btrExtGetNext(hlClf, tlClf(ilLnUpper).ClfRec, ilExtLen, tlClf(ilLnUpper).lRecPos)
                    'If ilRet = BTRV_ERR_REJECT_COUNT Then
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlClf, tlClf(ilLnUpper).ClfRec, ilExtLen, tlClf(ilLnUpper).lRecPos)
                    Loop
                    'End If
                Loop
            End If
        End If
        'All current lines are associated with the fully scheduled header
        'so if fully schedule header processed and not getting history then
        'processing is finished
        If (Not ilIncludeHistory) And (tlTChf.sSchStatus = "F") And (tlTChf.lCode = tlChf.lCode) Then
            Exit Do
        End If
        tlChfSrchKey1.lCntrNo = llCntrNo
        tlChfSrchKey1.iCntRevNo = tlTChf.iCntRevNo - 1
        tlChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hlChf, tlTChf, ilChfRecLen, tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Loop
    'Sort then output new headers and lines
    If UBound(tlClf) - 1 > 0 Then
        ArraySortTyp fnAV(tlClf(), 0), UBound(tlClf), 0, LenB(tlClf(0)), 0, LenB(tlClf(0).sKey), 0
    End If
    btrDestroy hlMnf
    btrDestroy hlRdf
    gObtainChfClf = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainCntr                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get Contract (Chf), Lines (Clf),*
'*                     Flight (Cff)                    *
'*                                                     *
'*******************************************************
Function gObtainCntr(hlChf As Integer, hlClf As Integer, hlCff As Integer, llChfCode As Long, ilIncludeHistory As Integer, tlChf As CHF, tlClf() As CLFLIST, tlCff() As CFFLIST, Optional blSort As Boolean = True) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim ilRet As Integer
    Dim llRif As Long
    Dim tlTChf As CHF
    Dim tlChfSrchKey As LONGKEY0            'CHF record image
    Dim tlChfSrchKey1 As CHFKEY1    'CntrNo; CntRevNo; PropVer
    Dim tlClfSrchKey As CLFKEY0            'CLF record image
    Dim tlCffSrchKey As CFFKEY0            'CFF record image
    Dim ilChfRecLen As Integer        'CHF record length
    Dim ilClfRecLen As Integer
    Dim ilCffRecLen As Integer
    'Dim ilLnUpper As Integer
    'Dim ilFlUpper As Integer
    Dim llLnUpper As Long
    Dim llFlUpper As Long
    Dim llNoRec As Long
    Dim ilExtLen As Integer
    Dim llClfIndex As Long
    Dim llCffIndex As Long
    Dim ilOffSet As Integer
    Dim ilAddLn As Integer
    Dim llTest As Long
    Dim llCntrNo As Long
    Dim slGpSort As String
    Dim slVehName As String
    Dim slVehSort As String
    Dim slDPSort As String
    Dim slLine As String
    Dim slRevNo As String
    Dim slPropVer As String
    Dim ilVef As Integer
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim ilMnfRecLen As Integer
    Dim tlMnfSrchKey As INTKEY0
    Dim hlRdf As Integer
    Dim tlRdf As RDF
    Dim ilRdfRecLen As Integer
    Dim ilDay As Integer
    Dim tlRdfSrchKey As INTKEY0
    Dim slOALineNo As String    'Line number of Parent + A or Line Number of Parent + C or ~~~~ if standard line)
    Dim llEliminateFirst As Long
    Dim llMaxClfCode As Long
    Dim llNext As Long



    ReDim tlClf(0 To 0) As CLFLIST
    tlClf(0).iStatus = -1 'Not Used
    tlClf(0).lRecPos = 0
    tlClf(0).iFirstCff = -1
    tlClf(0).iFirstCgf = -1
    tlClf(0).lFirstCff = -1
    tlClf(0).lFirstCgf = -1
    tlClf(0).lFirstAvail = -1
    ReDim tlCff(0 To 0) As CFFLIST
    tlCff(0).iStatus = -1 'Not Used
    tlCff(0).lRecPos = 0
    tlCff(0).iNextCff = -1
    tlCff(0).lNextCff = -1
    tlClf(0).lCxfCode = 0 'TTP 10193 - ReRate Comment
    ilChfRecLen = Len(tlChf)
    ilClfRecLen = Len(tlClf(0).ClfRec)
    ilCffRecLen = Len(tlCff(0).CffRec)
    tlChfSrchKey.lCode = llChfCode
    ilRet = btrGetEqual(hlChf, tlChf, ilChfRecLen, tlChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = ilRet
        gObtainCntr = False
        Exit Function
    End If
    If blSort Then
        hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            gObtainCntr = False
            btrDestroy hlMnf
            Exit Function
        End If
        ilMnfRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
        hlRdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            gObtainCntr = False
            btrDestroy hlMnf
            btrDestroy hlRdf
            Exit Function
        End If
        ilRdfRecLen = Len(tlRdf) 'btrRecordLength(hlMnf)  'Get and save record length
    End If
    ilRet = gObtainVef()
    If Not ilRet Then
        gObtainCntr = False
        If blSort Then
            btrDestroy hlMnf
            btrDestroy hlRdf
        End If
        Exit Function
    End If
    ilRet = gObtainRcfRifRdf()
    If Not ilRet Then
        gObtainCntr = False
        If blSort Then
            btrDestroy hlMnf
            btrDestroy hlRdf
        End If
        Exit Function
    End If
    llCntrNo = tlChf.lCntrNo
    llLnUpper = UBound(tlClf)
    llFlUpper = UBound(tlCff)
    tlChfSrchKey1.lCntrNo = llCntrNo
    If (tlChf.sStatus = "O") Or (tlChf.sStatus = "H") Then
        'Find latest contract because it contains all current line from different revisions
        tlChfSrchKey1.iCntRevNo = 32000
        tlChfSrchKey1.iPropVer = 32000
    Else
        tlChfSrchKey1.iCntRevNo = tlChf.iCntRevNo   '32000
        tlChfSrchKey1.iPropVer = tlChf.iPropVer '32000
    End If
    ilRet = btrGetGreaterOrEqual(hlChf, tlTChf, ilChfRecLen, tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tlTChf.lCntrNo = llCntrNo)
        btrExtClear hlClf   'Clear any previous extend operation
        tlClfSrchKey.lChfCode = tlTChf.lCode
        tlClfSrchKey.iLine = 0
        tlClfSrchKey.iCntRevNo = tlTChf.iCntRevNo ' 0 show latest version
        tlClfSrchKey.iPropVer = tlTChf.iPropVer ' 0 show latest version
        ilRet = btrGetGreaterOrEqual(hlClf, tlClf(llLnUpper).ClfRec, ilClfRecLen, tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If (tlClf(llLnUpper).ClfRec.lChfCode = tlTChf.lCode) Then
            ilExtLen = Len(tlClf(llLnUpper).ClfRec)
            llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
            Call btrExtSetBounds(hlClf, llNoRec, -1, "UC", "CLF", "") 'Set extract limits (all records)
            ilOffSet = gFieldOffset("Clf", "ClfChfCode")
            ilRet = btrExtAddLogicConst(hlClf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlTChf.lCode, 4)
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                If blSort Then
                    btrDestroy hlMnf
                    btrDestroy hlRdf
                End If
                gObtainCntr = False
                Exit Function
            End If
            ilRet = btrExtAddField(hlClf, 0, ilExtLen) 'Extract start/end time, and days
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                If blSort Then
                    btrDestroy hlMnf
                    btrDestroy hlRdf
                End If
                gObtainCntr = False
                Exit Function
            End If
            ilRet = btrExtGetNext(hlClf, tlClf(llLnUpper).ClfRec, ilExtLen, tlClf(llLnUpper).lRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                'ilRet = btrExtGetFirst(hlClf, tlClfExt, ilExtLen, llRecPos)
                ilExtLen = Len(tlClf(llLnUpper).ClfRec)
                'If ilRet = BTRV_ERR_REJECT_COUNT Then
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlClf, tlClf(llLnUpper).ClfRec, ilExtLen, tlClf(llLnUpper).lRecPos)
                Loop
                'End If
                Do While ilRet = BTRV_ERR_NONE
                    ilAddLn = True
                    If (tlClf(llLnUpper).ClfRec.iCntRevNo = tlChf.iCntRevNo) And (tlClf(llLnUpper).ClfRec.iPropVer <> tlChf.iPropVer) Then
                        ilAddLn = False
                    ElseIf (tlClf(llLnUpper).ClfRec.iCntRevNo > tlChf.iCntRevNo) Then
                        ilAddLn = False
                    Else
                        If Not ilIncludeHistory Then
                            'Test if line added- if so ignore line
                            For llTest = 0 To llLnUpper - 1 Step 1
                                If tlClf(llLnUpper).ClfRec.iLine = tlClf(llTest).ClfRec.iLine Then
                                    ilAddLn = False
                                    Exit For
                                End If
                            Next llTest
                        End If
                    End If
                    If ilAddLn Then
                        slOALineNo = "~~~~"
                        'Build sort key (Vehicle Group Sort #; Vehicle Sort #; Vehicle Name; Daypart Sort #; Line #
                        slVehSort = "999"
                        slVehName = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
                        'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If tlClf(llLnUpper).ClfRec.iVefCode = tgMVef(ilVef).iCode Then
                            ilVef = gBinarySearchVef(tlClf(llLnUpper).ClfRec.iVefCode)
                            If ilVef <> -1 Then
                                slVehName = tgMVef(ilVef).sName
                                slVehSort = Trim$(str$(tgMVef(ilVef).iSort))
                                If blSort Then
                                    tlMnfSrchKey.iCode = tgMVef(ilVef).iOwnerMnfCode
                                    If tlMnf.iCode <> tlMnfSrchKey.iCode Then
                                        ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            tlMnf.iGroupNo = 999
                                            tlMnf.iCode = 0
                                        End If
                                    End If
                                    slGpSort = Trim$(str$(tlMnf.iGroupNo))
                                Else
                                    slGpSort = ""
                                End If
                                'Exit For
                            End If
                        'Next ilVef
                        slDPSort = "999"
                        If blSort Then
                            If (Asc(tgSpf.sUsingFeatures5) And CNTRINVSORTRC) = CNTRINVSORTRC Then     'use rate card items sort code
                                For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                                    If (tgMRif(llRif).iRcfCode = tlChf.iRcfCode) And (tgMRif(llRif).iVefCode = tlClf(llLnUpper).ClfRec.iVefCode) And (tgMRif(llRif).iRdfCode = tlClf(llLnUpper).ClfRec.iRdfCode) Then
                                        If (tgMRif(llRif).iSort > 0) Then
                                            slDPSort = Trim$(str$(tgMRif(llRif).iSort))
                                        End If
                                        Exit For
                                    End If
                                Next llRif
                            End If
                        End If
                        If (Asc(tgSpf.sUsingFeatures5) And CNTRINVSORTLN) = CNTRINVSORTLN Then
                            slDPSort = Trim$(str$(tlClf(llLnUpper).ClfRec.iLine))
                        Else
                            If slDPSort = "999" Then
                                If blSort Then
                                    tlRdfSrchKey.iCode = tlClf(llLnUpper).ClfRec.iRdfCode
                                    If tlRdf.iCode <> tlRdfSrchKey.iCode Then
                                        ilRet = btrGetEqual(hlRdf, tlRdf, ilRdfRecLen, tlRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet = BTRV_ERR_NONE Then
                                            If tlRdf.iSortCode > 0 Then
                                                slDPSort = Trim$(str$(tlRdf.iSortCode))
                                            Else
                                                tlRdf.iSortCode = 999
                                                tlRdf.iCode = 0
                                            End If
                                        End If
                                    End If
                                End If
                                If (slDPSort = "999") Then
                                    slDPSort = Trim$(str$(tlClf(llLnUpper).ClfRec.iLine))
                                End If
                            End If
                        End If
                        Do While Len(slGpSort) < 3
                            slGpSort = "0" & slGpSort
                        Loop
                        Do While Len(slVehSort) < 3
                            slVehSort = "0" & slVehSort
                        Loop
                        Do While Len(slDPSort) < 4
                            slDPSort = "0" & slDPSort
                        Loop
                        slLine = Trim$(str$(tlClf(llLnUpper).ClfRec.iLine))
                        Do While Len(slLine) < 4
                            slLine = "0" & slLine
                        Loop
                        slRevNo = Trim$(str$(999 - tlClf(llLnUpper).ClfRec.iCntRevNo))
                        Do While Len(slRevNo) < 3
                            slRevNo = "0" & slRevNo
                        Loop
                        slPropVer = Trim$(str$(tlClf(llLnUpper).ClfRec.iPropVer))
                        Do While Len(slPropVer) < 3
                            slPropVer = "0" & slPropVer
                        Loop
                        'tlClf(llLnUpper).sKey = slGpSort & slVehSort & slVehName & slDPSort & slLine & slRevNo & slPropVer

                        tlClf(llLnUpper).sKey = slOALineNo & "|" & slGpSort & slVehSort & slVehName & slDPSort & slLine & slRevNo & slPropVer
                        tlClf(llLnUpper).iStatus = 1
                        tlClf(llLnUpper).bLnPriceSet = True
                        tlClf(llLnUpper).lCxfCode = tlClf(llLnUpper).ClfRec.lCxfCode 'TTP 10193 - ReRate Comment
                        
                        llLnUpper = llLnUpper + 1
                        ReDim Preserve tlClf(0 To llLnUpper) As CLFLIST
                        tlClf(llLnUpper).iStatus = -1 'Not Used
                        tlClf(llLnUpper).lFirstCff = -1
                        tlClf(llLnUpper).lFirstCgf = -1
                        tlClf(llLnUpper).lFirstAvail = -1
                        tlClf(llLnUpper).lRecPos = 0
                        tlClf(llLnUpper).bLnPriceSet = False
                    End If
                    ilRet = btrExtGetNext(hlClf, tlClf(llLnUpper).ClfRec, ilExtLen, tlClf(llLnUpper).lRecPos)
                    'If ilRet = BTRV_ERR_REJECT_COUNT Then
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlClf, tlClf(llLnUpper).ClfRec, ilExtLen, tlClf(llLnUpper).lRecPos)
                    Loop
                    'End If
                Loop
            End If
        End If
        'All current lines are associated with the fully scheduled header
        'so if fully schedule header processed and not getting history then
        'processing is finished
        If (Not ilIncludeHistory) And (tlTChf.sSchStatus = "F") And (tlTChf.lCode = tlChf.lCode) Then
            Exit Do
        End If
        tlChfSrchKey1.lCntrNo = llCntrNo
        tlChfSrchKey1.iCntRevNo = tlTChf.iCntRevNo - 1
        tlChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hlChf, tlTChf, ilChfRecLen, tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Loop
    'Sort then output new headers and lines
    If UBound(tlClf) - 1 > 0 Then
        'Clear last record as it contains the previous record
        llLnUpper = UBound(tlClf)
        ReDim Preserve tlClf(0 To llLnUpper - 1) As CLFLIST
        ReDim Preserve tlClf(0 To llLnUpper) As CLFLIST
        'Moved to after flight read so that records above 32000 can be eliminated
        'tlClf(llLnUpper).iStatus = -1 'Not Used
        'tlClf(llLnUpper).lFirstCff = -1
        'tlClf(llLnUpper).lFirstCgf = -1
        'tlClf(llLnUpper).lRecPos = 0
        'ArraySortTyp fnAV(tlClf(), 0), UBound(tlClf), 0, LenB(tlClf(0)), 0, LenB(tlClf(0).sKey), 0
    End If
    'Obtain Flight records for line
    For llClfIndex = LBound(tlClf) To UBound(tlClf) - 1 Step 1
        tlClf(llClfIndex).iOrigAnySpots = -1
        btrExtClear hlCff   'Clear any previous extend operation
        ilExtLen = Len(tlCff(llFlUpper).CffRec)  'Extract operation record size
        tlCffSrchKey.lChfCode = tlClf(llClfIndex).ClfRec.lChfCode
        tlCffSrchKey.iClfLine = tlClf(llClfIndex).ClfRec.iLine
        tlCffSrchKey.iCntRevNo = tlClf(llClfIndex).ClfRec.iCntRevNo
        tlCffSrchKey.iPropVer = tlClf(llClfIndex).ClfRec.iPropVer
        tlCffSrchKey.iStartDate(0) = 0
        tlCffSrchKey.iStartDate(1) = 0
        ilRet = btrGetGreaterOrEqual(hlCff, tlCff(llFlUpper).CffRec, ilCffRecLen, tlCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If (tlCff(llFlUpper).CffRec.lChfCode = tlClf(llClfIndex).ClfRec.lChfCode) And (tlCff(llFlUpper).CffRec.iClfLine = tlClf(llClfIndex).ClfRec.iLine) Then
            llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
            Call btrExtSetBounds(hlCff, llNoRec, -1, "UC", "CFF", "") 'Set extract limits (all records)
            ilOffSet = gFieldOffset("Cff", "CffChfCode")
            ilRet = btrExtAddLogicConst(hlCff, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlClf(llClfIndex).ClfRec.lChfCode, 4)
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                If blSort Then
                    btrDestroy hlMnf
                    btrDestroy hlRdf
                End If
                gObtainCntr = False
                Exit Function
            End If
            ilOffSet = gFieldOffset("Cff", "CffClfLine")
            ilRet = btrExtAddLogicConst(hlCff, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlClf(llClfIndex).ClfRec.iLine, 2)
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                If blSort Then
                    btrDestroy hlMnf
                    btrDestroy hlRdf
                End If
                gObtainCntr = False
                Exit Function
            End If
            ilOffSet = gFieldOffset("Cff", "CffCntRevNo")
            ilRet = btrExtAddLogicConst(hlCff, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlClf(llClfIndex).ClfRec.iCntRevNo, 2)
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                If blSort Then
                    btrDestroy hlMnf
                    btrDestroy hlRdf
                End If
                gObtainCntr = False
                Exit Function
            End If
            ilOffSet = gFieldOffset("Cff", "CffPropVer")
            ilRet = btrExtAddLogicConst(hlCff, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlClf(llClfIndex).ClfRec.iPropVer, 2)
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                If blSort Then
                    btrDestroy hlMnf
                    btrDestroy hlRdf
                End If
                gObtainCntr = False
                Exit Function
            End If
            ilRet = btrExtAddField(hlCff, 0, ilExtLen)  'Extract start date
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                If blSort Then
                    btrDestroy hlMnf
                    btrDestroy hlRdf
                End If
                gObtainCntr = False
                Exit Function
            End If
            ilRet = btrExtGetNext(hlCff, tlCff(llFlUpper).CffRec, ilExtLen, tlCff(llFlUpper).lRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                ilExtLen = Len(tlCff(llFlUpper).CffRec)  'Extract operation record size
                'If ilRet = BTRV_ERR_REJECT_COUNT Then
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlCff, tlCff(llFlUpper).CffRec, ilExtLen, tlCff(llFlUpper).lRecPos)
                Loop
                'End If
                Do While ilRet = BTRV_ERR_NONE
                    If tlCff(llFlUpper).CffRec.sDyWk = "D" Then
                        For ilDay = 0 To 6 Step 1
                            If tlCff(llFlUpper).CffRec.iDay(ilDay) <> 0 Then
                                tlClf(llClfIndex).iOrigAnySpots = 1
                                Exit For
                            End If
                        Next ilDay
                    Else
                        If (tlCff(llFlUpper).CffRec.iSpotsWk <> 0) Or (tlCff(llFlUpper).CffRec.iXSpotsWk <> 0) Then
                            tlClf(llClfIndex).iOrigAnySpots = 1
                        End If
                    End If
                    If tlClf(llClfIndex).lFirstCff = -1 Then
                        tlClf(llClfIndex).lFirstCff = llFlUpper
                    Else
                        tlCff(llFlUpper - 1).lNextCff = llFlUpper
                    End If
                    tlCff(llFlUpper).lNextCff = -1
                    tlCff(llFlUpper).iStatus = 1 'Old and retain
                    gUnpackDateLong tlCff(llFlUpper).CffRec.iStartDate(0), tlCff(llFlUpper).CffRec.iStartDate(1), tlCff(llFlUpper).lStartDate    'Week Start date
                    gUnpackDateLong tlCff(llFlUpper).CffRec.iEndDate(0), tlCff(llFlUpper).CffRec.iEndDate(1), tlCff(llFlUpper).lEndDate    'Week End date
                    llFlUpper = llFlUpper + 1
                    ReDim Preserve tlCff(0 To llFlUpper) As CFFLIST
                    tlCff(llFlUpper).iStatus = -1 'Not Used
                    tlCff(llFlUpper).lNextCff = -1
                    tlCff(llFlUpper).lRecPos = 0
                    'If llFlUpper > 32700 Then
                    '    Exit For
                    'End If
                    ilRet = btrExtGetNext(hlCff, tlCff(llFlUpper).CffRec, ilExtLen, tlCff(llFlUpper).lRecPos)
                    'If ilRet = BTRV_ERR_REJECT_COUNT Then
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlCff, tlCff(llFlUpper).CffRec, ilExtLen, tlCff(llFlUpper).lRecPos)
                    Loop
                    'End If
                Loop
            End If
        End If
    Next llClfIndex
    llEliminateFirst = -1
    If (llFlUpper > 32000) Or (llLnUpper > 32000) Then
        'Too many flights, reduce lines
        'Find last contract to retain
        For llClfIndex = LBound(tlClf) To UBound(tlClf) - 1 Step 1
            llNext = tlClf(llClfIndex).lFirstCff
            Do While llNext <> -1
                If llNext > 32000 Then
                    Exit For
                End If
                llNext = tlCff(llNext).lNextCff
            Loop
            If llClfIndex > 32000 Then
                Exit For
            End If
            llEliminateFirst = tlClf(llClfIndex).lFirstCff
        Next llClfIndex
        If llEliminateFirst <> -1 Then
            llMaxClfCode = -1
            For llClfIndex = LBound(tlClf) To UBound(tlClf) - 1 Step 1
                If tlClf(llClfIndex).lFirstCff = llEliminateFirst Then
                    llMaxClfCode = llClfIndex
                    Exit For
                End If
            Next llClfIndex
            If llMaxClfCode <> -1 Then
                ReDim Preserve tlClf(0 To llMaxClfCode - 1) As CLFLIST
                ReDim Preserve tlClf(0 To llMaxClfCode) As CLFLIST
                tlClf(llMaxClfCode).iStatus = -1 'Not Used
                tlClf(llMaxClfCode).lFirstCff = -1
                tlClf(llMaxClfCode).lFirstCgf = -1
                tlClf(llMaxClfCode).lFirstAvail = -1
                tlClf(llMaxClfCode).lRecPos = 0
            End If
        End If
    End If
    For llClfIndex = LBound(tlClf) To UBound(tlClf) - 1 Step 1
        tlClf(llClfIndex).iFirstCff = tlClf(llClfIndex).lFirstCff
        llNext = tlClf(llClfIndex).lFirstCff
        Do While llNext <> -1
            tlCff(llNext).iNextCff = tlCff(llNext).lNextCff
            llNext = tlCff(llNext).lNextCff
        Loop
        tlClf(llClfIndex).iFirstCgf = tlClf(llClfIndex).lFirstCgf
    Next llClfIndex
    'Sort then output new headers and lines
    If UBound(tlClf) - 1 >= 0 Then
        'Clear last record as it contains the previous record
        tlClf(UBound(tlClf)).iStatus = -1 'Not Used
        tlClf(UBound(tlClf)).iFirstCff = -1
        tlClf(UBound(tlClf)).iFirstCgf = -1
        tlClf(UBound(tlClf)).lFirstCff = -1
        tlClf(UBound(tlClf)).lFirstCgf = -1
        tlClf(UBound(tlClf)).lFirstAvail = -1
        tlClf(UBound(tlClf)).lRecPos = 0
    End If
    If UBound(tlClf) - 1 > 0 Then
        ArraySortTyp fnAV(tlClf(), 0), UBound(tlClf), 0, LenB(tlClf(0)), 0, LenB(tlClf(0).sKey), 0
    End If
    igBtrError = 0
    If blSort Then
        btrDestroy hlMnf
        btrDestroy hlRdf
    End If
    gObtainCntr = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainCntrForDate              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain an array of contracts    *
'*                     salesperson code                *
'*                                                     *
'*******************************************************
Function gObtainCntrForDate(frm As Form, slStartDate As String, slEndDate As String, slStatus As String, slCntrType As String, ilHOType As Integer, tlChfAdvtExt() As CHFADVTEXT) As Integer
'
'   ilRet = gObtainCntrForDate (MainForm, slStartDate, slEndDate, slStatus, slCntrType, ilHOType, tlChfAdvtExt() )
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       slStartDate(I)- Start Date
'       slEndDate(I)- End Date
'       slStatus (I)- chfStatus value or blank
'                         W=Working; D=Rejected; C=Completed; I=Unapproved; H=Hold; O=Order
'                         Multiple status can be specified (WDI)
'       slCntrType (I)- chfType value or blank
'                       C=Standard; V=Reservation; T=Remnant; R=DR; Q=PI; S=PSA; M=Promo
'       ilHOType (I)-  1=H or O only; 2=H or O or G or N (if G or N exists show it over H or O);
'                      3=H or O or G or N or W or C or I (if G or N or W or C or I exists show it over H or O)
'                        Note: G or N can't exist at the same time as W or C or I for an order
'                              G or N or W or C or I CntrRev > 0
'                      4=Return all matching contracts not deleted
'       tlChfAdvtExt(O)- Array of contracts of the structure CHFADVTEXT which match selection
'       ilRet (O)- Error code (0 if no error)
'
    gObtainCntrForDate = gCntrForActiveOHD(frm, slStartDate, slEndDate, "", "", slStatus, slCntrType, ilHOType, tlChfAdvtExt())
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainCntrSpot                 *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain spots for                *
'*                     specified dates and contract    *
'*          7-8-03 dh Add parameter to indicate 2 ways *
'*          to sort the results: 0 = current way (line,*
'*          vehicle, date & time or 1 = vehicle, line, *
'*          date & time                                *
'*                                                     *
'*******************************************************
Function gObtainCntrSpot(llVefCode As Long, ilVirtOnly As Integer, llChfCode As Long, ilLineNo As Integer, slSODate As String, slStartDate As String, slEndDate As String, tlSdfExtSort() As SDFEXTSORT, tlSdfExt() As SDFEXT, ilSortBy As Integer, ilAllSchdInSelVef As Integer, Optional blUseSchVehForMG As Boolean = False) As Integer
'Debug.Print "gObtainCntrSpot:" & llChfCode & "," & ilLineNo
'
'   ilRet = gObtainCntrSpot (llVefCode, llChfCode, ilLineNo, slStartDate, slEndDate, tlSdfExtSort(), tlSdfExt())
'   Where:
'       llVefCode(I)- Vehicle code (-1=All Vehicles; < -1= Vsf from contract; > 0 = Vehicle)
'                     Contract multi-vehicle can't be -1 as a vehicle must be defined
'                     prior to contracts
'       ilVirtOnly(I)- True=If llVefCode is for a virtual vehicle get its spots only
'                      False=Get all conventional vehicles and virtual vehicle spots
'       llChfCode() (I)- Contract code
'       ilLineNo(I)- Line number (-1=All lines)
'       slSODate(I)- S=Scheduled date (Use Actual Date from SMF); O=Ordered Date (Use Missed date from SMF)
'       slStartDate(I)- Start date  ("" = Start at earliest date)
'       slEndDate(I)- End date  ("" = TFN)
'       tlSdfExtSort(O)- Sort array containing key: Line#|Vehicle|Date|Time and Index into tlSdfExt
'                        this replaces list box sort
'       tlSdfExt(O)- Record images
'       ilSortBy (I) - 0 = line, vehicle name, date & time, 1 = vehicle name, line, date & time
'       ilAllSchdInSelVef (I)- True=Get all spots regardless of original contract vehicle if ilVefCode specified; False= Only get spots that ordered specified vehicle
'       blUseSchVefForMG (I) -
'       ilRet (O)- True if spots obtained OK; False if error
'

    Dim slStamp As String    'date/time stamp
    Dim hlSdf As Integer        'Sdf handle
    Dim ilSdfRecLen As Integer     'Record length
    Dim tlSdf As SDF
    Dim hlChf As Integer        'Chf handle
    Dim ilChfRecLen As Integer     'Record length
    Dim tlChf As CHF
    Dim hlVef As Integer        'Vef handle
    Dim ilVefRecLen As Integer     'Record length
    Dim tlVef As VEF
    Dim tlVefSrchKey As INTKEY0
    Dim hlVsf As Integer        'Vsf handle
    Dim tlVVsf As VSF           'Virtual Vehicle
    Dim tlCVsf As VSF           'Contract Multi- Vehicle
    Dim ilVsfReclen As Integer     'Record length
    Dim tlVsfSrchKey As LONGKEY0
    Dim hlSmf As Integer        'Smf handle
    Dim ilSmfRecLen As Integer     'Record length
    Dim tlSmf As SMF
    Dim tlSmfSrchKey As SMFKEY0
    Dim tlSmfSrchKey2 As LONGKEY0
    Dim slStr As String
    Dim llRecPos As Long        'Record location
    Dim llNoRec As Long
    Dim ilRet As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLTypeBuff As POPLCODE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim slDate As String
    Dim llDate As Long
    Dim slTime As String
    Dim llTime As Long
    Dim slNameSort As String
    Dim slLine As String
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    'Dim ilUpper As Integer
    Dim llUpper As Long
    Dim illoop As Integer
    Dim ilVVsf As Integer
    Dim ilCVsf As Integer
    Dim ilFound As Integer
    'Dim ilTest As Integer
    Dim llTest As Long
    Dim ilIncludeSpot As Integer
    Dim tlSdfSrchKey As SDFKEY0
    Dim tlSdfSrchKey3 As LONGKEY0
    Dim tlSdfSrchKey5 As LONGKEY0
    Dim ilUseKey5 As Integer
    Dim llLen As Long
    Dim llLkVsfCode As Long
    Dim ilSortLowLimit As Integer
    Dim ilLowLimit As Integer
    Dim llSdfExtIndex As Integer
    
    ilUseKey5 = False
    slStr = Trim$(str$(llVefCode)) & Trim$(str$(llChfCode)) & Trim$(str$(ilLineNo))
    If slStartDate <> "" Then
        llDate = gDateValue(slStartDate)
        slDate = Trim$(str$(llDate))
        Do While Len(slDate) < 6
            slDate = "0" & slDate
        Loop
    Else
        slDate = "000000"
    End If
    slStr = slStr & slDate
    If slEndDate <> "" Then
        llDate = gDateValue(slEndDate)
        slDate = Trim$(str$(llDate))
        Do While Len(slDate) < 6
            slDate = "0" & slDate
        Loop
    Else
        slDate = "000000"
    End If
    slStr = slStr & slDate
    slStamp = slStr 'FileDateTime(sgDBPath & "Sdf.Btr") & slStr

    'On Error GoTo gObtainCntrSpotErr2
    'ilRet = 0
    'ilSortLowLimit = LBound(tlSdfExtSort)
    'If ilRet <> 0 Then
    '    smSdfExtStamp = ""
    '    ilSortLowLimit = 0
    'End If
    'ilRet = 0
    'ilLowLimit = LBound(tlSdfExt)
    'If ilRet <> 0 Then
    '    smSdfExtStamp = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSdfExtSort).Ptr <> 0 Then
        ilSortLowLimit = LBound(tlSdfExtSort)
    Else
        smSdfExtStamp = ""
        ilSortLowLimit = 0
    End If
    If PeekArray(tlSdfExt).Ptr <> 0 Then
        ilLowLimit = LBound(tlSdfExt)
    Else
        smSdfExtStamp = ""
        ilLowLimit = 0
    End If

    If smSdfExtStamp <> "" Then
        If StrComp(slStamp, smSdfExtStamp, 1) = 0 Then
            If UBound(tlSdfExtSort) > ilSortLowLimit Then
                gObtainCntrSpot = True
                Exit Function
            End If
        End If
    End If
    gObtainCntrSpot = True
    hlSdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainCntrSpot = False
        ilRet = btrClose(hlSdf)
        btrDestroy hlSdf
        Exit Function
    End If
    ilSdfRecLen = Len(tlSdf) 'btrRecordLength(hlSdf)  'Get and save record length
    hlChf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainCntrSpot = False
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlChf)
        btrDestroy hlSdf
        btrDestroy hlChf
        Exit Function
    End If
    ilChfRecLen = Len(tlChf) 'btrRecordLength(hlSdf)  'Get and save record length
    hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainCntrSpot = False
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlChf)
        ilRet = btrClose(hlVef)
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlVef
        Exit Function
    End If
    ilVefRecLen = Len(tlVef) 'btrRecordLength(hlSdf)  'Get and save record length
    hlVsf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainCntrSpot = False
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlChf)
        ilRet = btrClose(hlVef)
        ilRet = btrClose(hlVsf)
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlVef
        btrDestroy hlVsf
        Exit Function
    End If
    ilVsfReclen = Len(tlVVsf) 'btrRecordLength(hlSdf)  'Get and save record length
    If llVefCode > 0 Then
        'For ilLoop = LBound(tlCVsf.iFSCode) To UBound(tlCVsf.iFSCode) Step 1
        '    tlCVsf.iFSCode(ilLoop) = 0
        'Next ilLoop
        'tlCVsf.iFSCode(LBound(tlCVsf.iFSCode)) = llVefCode
        ReDim imCntrSpotsVefCode(0 To 1) As Integer
        imCntrSpotsVefCode(0) = llVefCode
    Else
        If (llVefCode = -1) Or (llVefCode = 0) Then
            ''For ilLoop = LBound(tlCVsf.iFSCode) To UBound(tlCVsf.iFSCode) Step 1
            ''    tlCVsf.iFSCode(ilLoop) = 0
            ''Next ilLoop
            ''tlCVsf.iFSCode(LBound(tlCVsf.iFSCode)) = -1
            'ReDim imCntrSpotsVefCode(0 To 0) As Integer
            'ilRet = gObtainVef()
            'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            '    If (tgMVef(ilLoop).sType = "C") Or (tgMVef(ilLoop).sType = "S") Or (tgMVef(ilLoop).sType = "G") Or (tgMVef(ilLoop).sType = "V") Then
            '        imCntrSpotsVefCode(UBound(imCntrSpotsVefCode)) = tgMVef(ilLoop).iCode
            '        ReDim Preserve imCntrSpotsVefCode(0 To UBound(imCntrSpotsVefCode) + 1) As Integer
            '    End If
            'Next ilLoop
            ilUseKey5 = True
            ReDim imCntrSpotsVefCode(0 To 1) As Integer
            imCntrSpotsVefCode(0) = 32000
        Else
            llLkVsfCode = -llVefCode
            ReDim imCntrSpotsVefCode(0 To 0) As Integer
            Do While llLkVsfCode > 0
                tlVsfSrchKey.lCode = llLkVsfCode
                ilRet = btrGetEqual(hlVsf, tlCVsf, ilVsfReclen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    gObtainCntrSpot = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlChf)
                    ilRet = btrClose(hlVef)
                    ilRet = btrClose(hlVsf)
                    btrDestroy hlSdf
                    btrDestroy hlChf
                    btrDestroy hlVef
                    btrDestroy hlVsf
                    Exit Function
                End If
                For ilCVsf = LBound(tlCVsf.iFSCode) To UBound(tlCVsf.iFSCode) Step 1
                    If tlCVsf.iFSCode(ilCVsf) <> 0 Then
                        imCntrSpotsVefCode(UBound(imCntrSpotsVefCode)) = tlCVsf.iFSCode(ilCVsf)
                        ReDim Preserve imCntrSpotsVefCode(0 To UBound(imCntrSpotsVefCode) + 1) As Integer
                    End If
                Next ilCVsf
                llLkVsfCode = tlCVsf.lLkVsfCode
            Loop

        End If
    End If
    hlSmf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Erase imCntrSpotsVefCode
        gObtainCntrSpot = False
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlChf)
        ilRet = btrClose(hlVef)
        ilRet = btrClose(hlVsf)
        ilRet = btrClose(hlSmf)
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlVef
        btrDestroy hlVsf
        btrDestroy hlSmf
        Exit Function
    End If
    ilSmfRecLen = Len(tlSmf) 'btrRecordLength(hlSmf)  'Get and save record length
    ReDim tlSdfExtSort(ilSortLowLimit To ilSortLowLimit) As SDFEXTSORT
    'ReDim tlSdfExt(1 To 1) As SDFEXT
    ReDim tlSdfExt(ilLowLimit To ilLowLimit) As SDFEXT
    smSdfExtStamp = slStamp
    llLen = 0
    'For ilCVsf = LBound(tlCVsf.iFSCode) To UBound(tlCVsf.iFSCode) Step 1
    '    If tlCVsf.iFSCode(ilCVsf) <> 0 Then
    '        If tlCVsf.iFSCode(ilCVsf) <> -1 Then
    '            tlVefSrchKey.iCode = tlCVsf.iFSCode(ilCVsf)
    For ilCVsf = LBound(imCntrSpotsVefCode) To UBound(imCntrSpotsVefCode) - 1 Step 1
        If imCntrSpotsVefCode(ilCVsf) > 0 Then
            If Not ilUseKey5 Then
                tlVefSrchKey.iCode = imCntrSpotsVefCode(ilCVsf)
                ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> BTRV_ERR_NONE Then
                    Erase imCntrSpotsVefCode
                    gObtainCntrSpot = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlChf)
                    ilRet = btrClose(hlVef)
                    ilRet = btrClose(hlVsf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlChf
                    btrDestroy hlVef
                    btrDestroy hlVsf
                    btrDestroy hlSmf
                    Exit Function
                End If
                'If tlVef.sType <> "V" Then
                If (tlVef.sType <> "V") Or ((tlVef.sType = "V") And (ilVirtOnly)) Then
                    For illoop = LBound(tlVVsf.iFSCode) To UBound(tlVVsf.iFSCode) Step 1
                        tlVVsf.iFSCode(illoop) = 0
                    Next illoop
                    tlVVsf.iFSCode(LBound(tlVVsf.iFSCode)) = imCntrSpotsVefCode(ilCVsf) 'tlCVsf.iFSCode(ilCVsf)
                Else
                    tlVsfSrchKey.lCode = tlVef.lVsfCode
                    ilRet = btrGetEqual(hlVsf, tlVVsf, ilVsfReclen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        Erase imCntrSpotsVefCode
                        gObtainCntrSpot = False
                        ilRet = btrClose(hlSdf)
                        ilRet = btrClose(hlChf)
                        ilRet = btrClose(hlVef)
                        ilRet = btrClose(hlVsf)
                        ilRet = btrClose(hlSmf)
                        btrDestroy hlSdf
                        btrDestroy hlChf
                        btrDestroy hlVef
                        btrDestroy hlVsf
                        btrDestroy hlSmf
                        Exit Function
                    End If
                    'Store the virtual vehicle as the first vehicle
                    For ilVVsf = UBound(tlVVsf.iFSCode) To LBound(tlVVsf.iFSCode) + 1 Step -1
                        tlVVsf.iFSCode(ilVVsf) = tlVVsf.iFSCode(ilVVsf - 1)
                        tlVVsf.iNoSpots(ilVVsf) = tlVVsf.iNoSpots(ilVVsf - 1)
                    Next ilVVsf
                    tlVVsf.iFSCode(LBound(tlVVsf.iFSCode)) = imCntrSpotsVefCode(ilCVsf) 'tlCVsf.iFSCode(ilCVsf)
                    tlVVsf.iNoSpots(LBound(tlVVsf.iFSCode)) = 1
                End If
                'Else
                '    For ilLoop = LBound(tlVVsf.iFSCode) To UBound(tlVVsf.iFSCode) Step 1
                '        tlVVsf.iFSCode(ilLoop) = 0
                '    Next ilLoop
                '    tlVVsf.iFSCode(LBound(tlVVsf.iFSCode)) = -1
                'End If
            Else
                For illoop = LBound(tlVVsf.iFSCode) To UBound(tlVVsf.iFSCode) Step 1
                    tlVVsf.iFSCode(illoop) = 0
                Next illoop
                tlVVsf.iFSCode(LBound(tlVVsf.iFSCode)) = imCntrSpotsVefCode(ilCVsf) 'tlCVsf.iFSCode(ilCVsf)
            End If
            For ilVVsf = LBound(tlVVsf.iFSCode) To UBound(tlVVsf.iFSCode) Step 1
                If tlVVsf.iFSCode(ilVVsf) <> 0 Then
                    ilExtLen = 2 + 4 + 2 + 4 + 4 + 1 + 1 + 2 + 1 + 1 + 1 + 2 + 2 + 1 + 4 'VefCode+ChfCode+LineNo+Date+Time+Status+Trace+Length+PriceType+SpotType+Bill+adf+GameNo+Midnight+code Len(tlSdfExt(1)) - 9'Extract operation record size
                    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
                    btrExtClear hlSdf   'Clear any previous extend operation
                    If Not ilUseKey5 Then
                        If tlVVsf.iFSCode(ilVVsf) <= 0 Then
                            tlSdfSrchKey.iVefCode = 1
                        Else
                            tlSdfSrchKey.iVefCode = tlVVsf.iFSCode(ilVVsf)
                        End If
                        tlSdfSrchKey.lChfCode = llChfCode
                        tlSdfSrchKey.iLineNo = 0
                        tlSdfSrchKey.lFsfCode = 0
                        tlSdfSrchKey.iDate(0) = 0
                        tlSdfSrchKey.iDate(1) = 0
                        tlSdfSrchKey.sSchStatus = " "
                        tlSdfSrchKey.iTime(0) = 0
                        tlSdfSrchKey.iTime(1) = 0
                        ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Else
                        tlSdfSrchKey5.lCode = llChfCode
                        ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey5, INDEXKEY5, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    End If
                    If ilRet <> BTRV_ERR_END_OF_FILE Then
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDFEXTPK", SDFEXTPK) 'Set extract limits (all records)
                        If Not ilUseKey5 Then
                            If tlVVsf.iFSCode(ilVVsf) > 0 Then
                                tlIntTypeBuff.iType = tlVVsf.iFSCode(ilVVsf)
                                ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
                                ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                            Else
                                'tlIntTypeBuff.iType = 1
                                'ilOffset = gFieldOffset("Sdf", "SdfVefCode")
                                'ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_GTE, BTRV_EXT_AND, tlIntTypeBuff, 2)
                                'tlIntTypeBuff.iType = 32000
                                'ilOffset = gFieldOffset("Sdf", "SdfVefCode")
                                'ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_LTE, BTRV_EXT_AND, tlIntTypeBuff, 2)
                            End If
                        End If
                        tlLTypeBuff.lCode = llChfCode
                        ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLTypeBuff, 4)
                        If ilLineNo > 0 Then
                            tlIntTypeBuff.iType = ilLineNo
                            ilOffSet = gFieldOffset("Sdf", "SdfLineNo")
                            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                        End If
                        If slStartDate = "" Then
                            slDate = "1/1/1970"
                        Else
                            slDate = slStartDate
                        End If
                        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                        ilOffSet = gFieldOffset("Sdf", "SdfDate")
                        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
                        If slEndDate = "" Then
                            slDate = "12/31/2069"
                        Else
                            slDate = slEndDate
                        End If
                        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                        ilOffSet = gFieldOffset("Sdf", "SdfDate")
                        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 2)  'Extract Name
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 4)  'Extract iCode field
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        ilOffSet = gFieldOffset("Sdf", "SdfLineNo")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 2)  'Extract Name
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        ilOffSet = gFieldOffset("Sdf", "SdfDate")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 4) 'Extract Variation
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        ilOffSet = gFieldOffset("Sdf", "SdfTime")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 4) 'Extract Variation
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        ilOffSet = gFieldOffset("Sdf", "SdfSchStatus")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 1) 'Extract Variation
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        ilOffSet = gFieldOffset("Sdf", "SdfTracer")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 1) 'Extract Variation
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        ilOffSet = gFieldOffset("Sdf", "SdfLen")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 2) 'Extract Variation
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        ilOffSet = gFieldOffset("Sdf", "SdfPriceType")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 1) 'Extract Variation
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        ilOffSet = gFieldOffset("Sdf", "SdfSpotType")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 1) 'Extract Variation
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        ilOffSet = gFieldOffset("Sdf", "SdfBill")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 1) 'Extract Variation
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        ilOffSet = gFieldOffset("Sdf", "SdfAdfCode")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 2) 'Extract Variation
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        ilOffSet = gFieldOffset("Sdf", "SdfGameNo")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 2) 'Extract Variation
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        
                        '5/22/14: Add cross midnight
                        ilOffSet = gFieldOffset("Sdf", "SdfXCrossMidnight")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 1) 'Extract Variation
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If

                        ilOffSet = gFieldOffset("Sdf", "SdfCode")
                        ilRet = btrExtAddField(hlSdf, ilOffSet, 4) 'Extract Variation
                        If ilRet <> BTRV_ERR_NONE Then
                            Erase imCntrSpotsVefCode
                            gObtainCntrSpot = False
                            ilRet = btrClose(hlVsf)
                            ilRet = btrClose(hlVef)
                            ilRet = btrClose(hlChf)
                            ilRet = btrClose(hlSdf)
                            ilRet = btrClose(hlSmf)
                            btrDestroy hlVsf
                            btrDestroy hlVef
                            btrDestroy hlChf
                            btrDestroy hlSdf
                            btrDestroy hlSmf
                            Exit Function
                        End If
                        'ilRet = btrExtGetNextExt(hlSdf)    'Extract record
                        llUpper = UBound(tlSdfExt)
                        ilRet = btrExtGetNext(hlSdf, tlSdfExt(llUpper), ilExtLen, llRecPos)
                        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                                Erase imCntrSpotsVefCode
                                gObtainCntrSpot = False
                                ilRet = btrClose(hlVsf)
                                ilRet = btrClose(hlVef)
                                ilRet = btrClose(hlChf)
                                ilRet = btrClose(hlSdf)
                                ilRet = btrClose(hlSmf)
                                btrDestroy hlVsf
                                btrDestroy hlVef
                                btrDestroy hlChf
                                btrDestroy hlSdf
                                btrDestroy hlSmf
                                Exit Function
                            End If
                            ilExtLen = 2 + 4 + 2 + 4 + 4 + 1 + 1 + 2 + 1 + 1 + 1 + 2 + 2 + 1 + 4 'VefCode+ChfCode+LineNo+Date+Time+Status+Tracer+Length+PriceType+SpotType+Bill+adf+GameNo+Midnight+code Len(tlSdfExt(1)) - 4'Extract operation record size
                            llUpper = UBound(tlSdfExt)
                            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hlSdf, tlSdfExt(llUpper), ilExtLen, llRecPos)
                            Loop
                            Do While ilRet = BTRV_ERR_NONE
                                'Test if spot previously added- if so ignore (virtual then one virtual vehicle for contract, and
                                'the virtual vehicles book into same vehicles)
                                ilFound = False
                                For llTest = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
                                    If tlSdfExt(llTest).lRecPos = llRecPos Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next llTest
                                If Not ilFound Then
                                    If (tlSdfExt(llUpper).sSchStatus = "O") Or (tlSdfExt(llUpper).sSchStatus = "G") Then
                                        If (slSODate = "O") Then
                                            'we will pickup the spot when we do the SMF (because it's AS-ORDERED, which the SMF has the as-ordered info)
                                            ilFound = True
                                            tlSdfExt(llUpper).lMdDate = 0
                                        Else
                                            tlSmfSrchKey2.lCode = tlSdfExt(llUpper).lCode
                                            ilRet = btrGetEqual(hlSmf, tlSmf, ilSmfRecLen, tlSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilFound = True
                                            Else
                                                gUnpackDateLong tlSmf.iMissedDate(0), tlSmf.iMissedDate(1), tlSdfExt(llUpper).lMdDate
                                            End If
                                        End If
                                    Else
                                        tlSdfExt(llUpper).lMdDate = 0
                                    End If
                                End If
                                
                                If Not ilFound Then
                                    'tlSdfExt(llUpper).lMdDate = 0
                                    tlSdfExt(llUpper).lRecPos = llRecPos
                                    gUnpackDateForSort tlSdfExt(llUpper).iDate(0), tlSdfExt(llUpper).iDate(1), slDate
                                    gUnpackTime tlSdfExt(llUpper).iTime(0), tlSdfExt(llUpper).iTime(1), "A", "1", slTime
                                    llTime = CLng(gTimeToCurrency(slTime, False))
                                    slTime = Trim$(str$(llTime))
                                    Do While Len(slTime) < 5
                                        slTime = "0" & slTime
                                    Loop
                                    slLine = Trim$(str$(tlSdfExt(llUpper).iLineNo))
                                    'If tlVef.iCode <> tlSdfExt(llUpper).iVefCode Then
                                    ''    tlVefSrchKey.iCode = tlSdfExt(ilUpper).iVefCode
                                    ''    ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    ''    If ilRet <> BTRV_ERR_NONE Then
                                    '    ilRet = gBinarySearchVef(tlSdfExt(llUpper).iVefCode)
                                    '    If ilRet <> -1 Then
                                    '        tlVef = tgMVef(ilRet)
                                    '    Else
                                    '        tlVef.sName = "Missing"
                                    '    End If
                                    'End If
                                    If (ilSortBy = 0) And ((tlSdfExt(llUpper).sSchStatus = "O") Or (tlSdfExt(llUpper).sSchStatus = "G")) Then             'sort by line, veh, date & time
                                        If blUseSchVehForMG Then            'sort with the scheduled vehicle name, not the ordered
                                            ilRet = gBinarySearchVef(tlSdfExt(llUpper).iVefCode)
                                            If ilRet <> -1 Then
                                                tlVef = tgMVef(ilRet)
                                            Else
                                                tlVef.sName = "Missing"
                                            End If
                                        Else
                                            ilRet = gBinarySearchVef(tlSmf.iOrigSchVef)
                                            If ilRet <> -1 Then
                                                tlVef = tgMVef(ilRet)
                                            Else
                                                tlVef.sName = "Missing"
                                            End If
                                        End If
                                    Else
                                        If tlVef.iCode <> tlSdfExt(llUpper).iVefCode Then
                                        '    tlVefSrchKey.iCode = tlSdfExt(ilUpper).iVefCode
                                        '    ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        '    If ilRet <> BTRV_ERR_NONE Then
                                            ilRet = gBinarySearchVef(tlSdfExt(llUpper).iVefCode)
                                            If ilRet <> -1 Then
                                                tlVef = tgMVef(ilRet)
                                            Else
                                                tlVef.sName = "Missing"
                                            End If
                                        End If
                                    End If

                                    '7-8-03
                                    If ilSortBy = 0 Then            'sort by line, veh, date & time
                                        slNameSort = slLine & "|" & tlVef.sName & "|" & slDate & "|" & slTime
                                    Else                            'sort by veh, line, date & time
                                        slNameSort = tlVef.sName & "|" & slLine & "|" & slDate & "|" & slTime
                                    End If
                                    tlSdfExtSort(UBound(tlSdfExtSort)).sKey = slNameSort
                                    tlSdfExtSort(UBound(tlSdfExtSort)).lSdfExtIndex = llUpper
                                    ReDim Preserve tlSdfExtSort(ilSortLowLimit To UBound(tlSdfExtSort) + 1) As SDFEXTSORT
                                    'ilUpper = ilUpper + 1
                                    llUpper = llUpper + 1
                                    'ReDim Preserve tlSdfExt(1 To llUpper) As SDFEXT
                                    ReDim Preserve tlSdfExt(ilLowLimit To llUpper) As SDFEXT
                                End If
                                ilRet = btrExtGetNext(hlSdf, tlSdfExt(llUpper), ilExtLen, llRecPos)
                                Do While ilRet = BTRV_ERR_REJECT_COUNT
                                    ilRet = btrExtGetNext(hlSdf, tlSdfExt(llUpper), ilExtLen, llRecPos)
                                Loop
                            Loop
                        End If
                    End If
                End If
            Next ilVVsf
        End If
    Next ilCVsf
    
    'Get any vehicle spots
    llUpper = UBound(tlSdfExt)
    ilExtLen = Len(tlSmf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hlSmf   'Clear any previous extend operation
    tlSmfSrchKey.lChfCode = llChfCode
    If ilLineNo > 0 Then
        tlSmfSrchKey.iLineNo = ilLineNo
    Else
        tlSmfSrchKey.iLineNo = 0
    End If
    tlSmfSrchKey.iMissedDate(0) = 0
    tlSmfSrchKey.iMissedDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hlSmf, tlSmf, ilSmfRecLen, tlSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Erase imCntrSpotsVefCode
            ilRet = btrClose(hlSmf)
            btrDestroy hlSmf
            ilRet = btrClose(hlVsf)
            btrDestroy hlVsf
            ilRet = btrClose(hlVef)
            btrDestroy hlVef
            ilRet = btrClose(hlChf)
            btrDestroy hlChf
            ilRet = btrClose(hlSdf)
            btrDestroy hlSdf
            gObtainCntrSpot = True
            Exit Function
        End If
        Call btrExtSetBounds(hlSmf, llNoRec, -1, "UC", "SMF", "") 'Set extract limits (all records)
        tlLTypeBuff.lCode = llChfCode
        ilOffSet = gFieldOffset("Smf", "SmfChfCode")
        ilRet = btrExtAddLogicConst(hlSmf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLTypeBuff, 4)
        If ilLineNo > 0 Then
            tlIntTypeBuff.iType = ilLineNo
            ilOffSet = gFieldOffset("Smf", "SmfLineNo")
            ilRet = btrExtAddLogicConst(hlSmf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        End If
        'tlCharTypeBuff.sType = "A"    'Extract all matching records
        'ilOffset = gFieldOffset("Smf", "SmfMGSource")
        'ilRet = btrExtAddLogicConst(hlSmf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
        If (slStartDate = "") Then
            slDate = "1/1/1970"
        Else
            slDate = slStartDate
        End If
        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        If (slSODate = "O") Or (ilUseKey5) Then
            ilOffSet = gFieldOffset("Smf", "SmfMissedDate")
        Else
            ilOffSet = gFieldOffset("Smf", "SmfActualDate")
        End If
        ilRet = btrExtAddLogicConst(hlSmf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        If (slEndDate = "") Then
            slDate = "12/31/2069"
        Else
            slDate = slEndDate
        End If

        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        If (slSODate = "O") Or (ilUseKey5) Then
            ilOffSet = gFieldOffset("Smf", "SmfMissedDate")
        Else
            ilOffSet = gFieldOffset("Smf", "SmfActualDate")
        End If
        ilRet = btrExtAddLogicConst(hlSmf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        ilRet = btrExtAddField(hlSmf, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Erase imCntrSpotsVefCode
            ilRet = btrClose(hlSmf)
            btrDestroy hlSmf
            ilRet = btrClose(hlVsf)
            btrDestroy hlVsf
            ilRet = btrClose(hlVef)
            btrDestroy hlVef
            ilRet = btrClose(hlChf)
            btrDestroy hlChf
            ilRet = btrClose(hlSdf)
            btrDestroy hlSdf
            gObtainCntrSpot = True
            Exit Function
        End If
        ilRet = btrExtGetNext(hlSmf, tlSmf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Erase imCntrSpotsVefCode
                ilRet = btrClose(hlSmf)
                btrDestroy hlSmf
                ilRet = btrClose(hlVsf)
                btrDestroy hlVsf
                ilRet = btrClose(hlVef)
                btrDestroy hlVef
                ilRet = btrClose(hlChf)
                btrDestroy hlChf
                ilRet = btrClose(hlSdf)
                btrDestroy hlSdf
                gObtainCntrSpot = True
                Exit Function
            End If
            ilExtLen = Len(tlSmf)
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSmf, tlSmf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If (tlSmf.sSchStatus = "O") Or (tlSmf.sSchStatus = "G") Then
                    ilIncludeSpot = True
                    tlSdfSrchKey3.lCode = tlSmf.lSdfCode
                    ilRet = btrGetEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        ilRet = btrGetPosition(hlSdf, llRecPos)
                        'If tlSdf.iVefCode = tlSmf.iOrigSchVef Then
                        '    ilIncludeSpot = False
                        'End If
                    Else
                        ilIncludeSpot = False
                    End If
                Else
                    ilIncludeSpot = False
                End If
                If ilIncludeSpot Then
                    If Not ilUseKey5 Then
                        ilIncludeSpot = False
                        'For ilCVsf = LBound(tlCVsf.iFSCode) To UBound(tlCVsf.iFSCode) Step 1
                        '    If tlCVsf.iFSCode(ilCVsf) <> 0 Then
                        '        If tlCVsf.iFSCode(ilCVsf) <> -1 Then
                        '            tlVefSrchKey.iCode = tlCVsf.iFSCode(ilCVsf)
                        For ilCVsf = LBound(imCntrSpotsVefCode) To UBound(imCntrSpotsVefCode) - 1 Step 1
                            If imCntrSpotsVefCode(ilCVsf) > 0 Then
                                'tlVefSrchKey.iCode = imCntrSpotsVefCode(ilCVsf)
                                'ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                'If ilRet <> BTRV_ERR_NONE Then
                                ilRet = gBinarySearchVef(imCntrSpotsVefCode(ilCVsf))
                                If ilRet <> -1 Then
                                    tlVef = tgMVef(ilRet)
                                Else
                                    Erase imCntrSpotsVefCode
                                    gObtainCntrSpot = False
                                    ilRet = btrClose(hlSdf)
                                    ilRet = btrClose(hlChf)
                                    ilRet = btrClose(hlVef)
                                    ilRet = btrClose(hlVsf)
                                    ilRet = btrClose(hlSmf)
                                    btrDestroy hlSdf
                                    btrDestroy hlChf
                                    btrDestroy hlVef
                                    btrDestroy hlVsf
                                    btrDestroy hlSmf
                                    Exit Function
                                End If
                                If tlVef.sType <> "V" Then
                                    For illoop = LBound(tlVVsf.iFSCode) To UBound(tlVVsf.iFSCode) Step 1
                                        tlVVsf.iFSCode(illoop) = 0
                                    Next illoop
                                    tlVVsf.iFSCode(LBound(tlVVsf.iFSCode)) = imCntrSpotsVefCode(ilCVsf) 'tlCVsf.iFSCode(ilCVsf)
                                Else
                                    tlVsfSrchKey.lCode = tlVef.lVsfCode
                                    ilRet = btrGetEqual(hlVsf, tlVVsf, ilVsfReclen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        Erase imCntrSpotsVefCode
                                        gObtainCntrSpot = False
                                        ilRet = btrClose(hlSdf)
                                        ilRet = btrClose(hlChf)
                                        ilRet = btrClose(hlVef)
                                        ilRet = btrClose(hlVsf)
                                        ilRet = btrClose(hlSmf)
                                        btrDestroy hlSdf
                                        btrDestroy hlChf
                                        btrDestroy hlVef
                                        btrDestroy hlVsf
                                        btrDestroy hlSmf
                                        Exit Function
                                    End If
                                    'Store the virtual vehicle as the first vehicle
                                    For ilVVsf = UBound(tlVVsf.iFSCode) To LBound(tlVVsf.iFSCode) + 1 Step -1
                                        tlVVsf.iFSCode(ilVVsf) = tlVVsf.iFSCode(ilVVsf - 1)
                                        tlVVsf.iNoSpots(ilVVsf) = tlVVsf.iNoSpots(ilVVsf - 1)
                                    Next ilVVsf
                                    tlVVsf.iFSCode(LBound(tlVVsf.iFSCode)) = imCntrSpotsVefCode(ilCVsf) 'tlCVsf.iFSCode(ilCVsf)
                                    tlVVsf.iNoSpots(LBound(tlVVsf.iFSCode)) = 1
                                End If
                                'Else
                                '    For ilLoop = LBound(tlVVsf.iFSCode) To UBound(tlVVsf.iFSCode) Step 1
                                '        tlVVsf.iFSCode(ilLoop) = 0
                                '    Next ilLoop
                                '    tlVVsf.iFSCode(LBound(tlVVsf.iFSCode)) = -1
                                'End If
                                If ilAllSchdInSelVef Then
                                    ilIncludeSpot = True
                                Else
                                    For ilVVsf = LBound(tlVVsf.iFSCode) To UBound(tlVVsf.iFSCode) Step 1
                                        If tlVVsf.iFSCode(ilVVsf) <> 0 Then
                                            If tlVVsf.iFSCode(ilVVsf) > 0 Then
                                                If tlVVsf.iFSCode(ilVVsf) = tlSmf.iOrigSchVef Then
                                                    ilIncludeSpot = True
                                                    Exit For
                                                End If
                                            Else
                                                ilIncludeSpot = True
                                                Exit For
                                            End If
                                        End If
                                    Next ilVVsf
                                End If
                                If ilIncludeSpot Then
                                    Exit For
                                End If
                            End If
                        Next ilCVsf
                    End If
                End If
                
                If ilIncludeSpot Then
                    'TTP 10766 - Spot Discrepancy Report and Spot Discrepancy Summary by Month & Revisit TTP 10723
                    ilFound = False
                    llSdfExtIndex = llUpper 'Index for Adding a new item to array
                    For llTest = LBound(tlSdfExt) To UBound(tlSdfExt) - 1 Step 1
                        'If tlSdfExt(ilTest).lRecPos = llRecPos Then
                        If tlSdfExt(llTest).lCode = tlSdf.lCode Then
                            ilFound = True
                            llSdfExtIndex = llTest 'Index for Updating existing item in array
                            Exit For
                        End If
                    Next llTest
                    
                    'JW 6/29/23 Revisit TTP 10723
                    tlSdfExt(llSdfExtIndex).iVefCode = tlSdf.iVefCode     'Vehicle Code (combos not allowed)
                    tlSdfExt(llSdfExtIndex).lChfCode = tlSdf.lChfCode       'Contract code
                    tlSdfExt(llSdfExtIndex).iLineNo = tlSdf.iLineNo       'Contract code
                    tlSdfExt(llSdfExtIndex).iDate(0) = tlSdf.iDate(0)    'Schedule or missed Date of spot
                    tlSdfExt(llSdfExtIndex).iDate(1) = tlSdf.iDate(1)    'Schedule or missed Date of spot
                    tlSdfExt(llSdfExtIndex).iTime(0) = tlSdf.iTime(0)    'Schedule or missed Date of spot
                    tlSdfExt(llSdfExtIndex).iTime(1) = tlSdf.iTime(1)    'Schedule or missed Date of spot
                    tlSdfExt(llSdfExtIndex).sSchStatus = tlSdf.sSchStatus    'S=Scheduled, M=Missed, R=Ready to schd MG, U=Unscheduled MG,
                    tlSdfExt(llSdfExtIndex).sTracer = tlSdf.sTracer
                    tlSdfExt(llSdfExtIndex).iLen = tlSdf.iLen         'Spot length
                    tlSdfExt(llSdfExtIndex).sPriceType = tlSdf.sPriceType
                    tlSdfExt(llSdfExtIndex).sSpotType = tlSdf.sSpotType
                    tlSdfExt(llSdfExtIndex).sBill = tlSdf.sBill
                    tlSdfExt(llSdfExtIndex).iAdfCode = tlSdf.iAdfCode
                    tlSdfExt(llSdfExtIndex).iGameNo = tlSdf.iGameNo
                    gUnpackDateLong tlSmf.iMissedDate(0), tlSmf.iMissedDate(1), tlSdfExt(llSdfExtIndex).lMdDate
                    tlSdfExt(llSdfExtIndex).iStatus = 0     '0=From Sdf; 1= from Smf
                    tlSdfExt(llSdfExtIndex).lRecPos = llRecPos
                    tlSdfExt(llSdfExtIndex).lCode = tlSdf.lCode
                    
                    '-------------------------------------
                    'If we didnt just 'update' tlSdfExt, Add it to tlSdfExt and tlSdfExtSort arrays
                    If Not ilFound Then
                        'Get Date/Time for tlSdfExtSort.sKey
                        gUnpackDateForSort tlSdfExt(llSdfExtIndex).iDate(0), tlSdfExt(llSdfExtIndex).iDate(1), slDate
                        gUnpackTime tlSdfExt(llSdfExtIndex).iTime(0), tlSdfExt(llSdfExtIndex).iTime(1), "A", "1", slTime
                        llTime = CLng(gTimeToCurrency(slTime, False))
                        slTime = Trim$(str$(llTime))
                        Do While Len(slTime) < 5
                            slTime = "0" & slTime
                        Loop
                        'Get Line # for tlSdfExtSort.sKey
                        slLine = Trim$(str$(tlSdfExt(llSdfExtIndex).iLineNo))
                        'Get Vehicle Name for tlSdfExtSort.sKey
                        If slSODate = "O" And tlSdf.iVefCode <> tlSmf.iOrigSchVef Then tlSdf.iVefCode = tlSmf.iOrigSchVef
                        ilRet = gBinarySearchVef(tlSdf.iVefCode)
                        If ilRet <> -1 Then
                            tlVef = tgMVef(ilRet)
                        Else
                            tlVef.sName = "Missing"
                        End If
                        'Make Sort Key String for tlSdfExtSort.sKey
                        If ilSortBy = 0 Then                'sort by line, vehicle, date & time
                            slNameSort = slLine & "|" & tlVef.sName & "|" & slDate & "|" & slTime
                        Else                            'mg  or out: sort by vehicle, line, date & time
                            'use the vehicle mg or out spot is schedule in
                            slNameSort = tlVef.sName & "|" & slLine & "|" & slDate & "|" & slTime
                        End If
                        'Add it to tlSdfExt and tlSdfExtSort arrays
                        tlSdfExtSort(UBound(tlSdfExtSort)).sKey = slNameSort
                        tlSdfExtSort(UBound(tlSdfExtSort)).lSdfExtIndex = llUpper
                        ReDim Preserve tlSdfExtSort(ilSortLowLimit To UBound(tlSdfExtSort) + 1) As SDFEXTSORT
                        llUpper = llUpper + 1
                        ReDim Preserve tlSdfExt(ilLowLimit To llUpper) As SDFEXT
                    End If
                End If
                
                ilRet = btrExtGetNext(hlSmf, tlSmf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSmf, tlSmf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    If UBound(tlSdfExtSort) - 1 > ilSortLowLimit Then
        ArraySortTyp fnAV(tlSdfExtSort(), ilSortLowLimit), UBound(tlSdfExtSort) - ilSortLowLimit, 0, LenB(tlSdfExtSort(ilSortLowLimit)), 0, LenB(tlSdfExtSort(ilSortLowLimit).sKey), 0
    End If
    Erase imCntrSpotsVefCode
    ilRet = btrClose(hlSmf)
    btrDestroy hlSmf
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlChf)
    btrDestroy hlChf
    ilRet = btrClose(hlSdf)
    btrDestroy hlSdf
    gObtainCntrSpot = True
    Exit Function
gObtainCntrSpotErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainComp                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgCompMnf for          *
'*                     collection                      *
'*                                                     *
'*******************************************************
Function gObtainComp() As Integer
'
'   ilRet = gObtainComp ()
'   Where:
'       tgCompMnf() (I)- MNFCOMPEXT record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlMnf As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlMnf As MNF
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim slGetStamp As String
    Dim ilLowLimit As Integer

    slStamp = gFileDateTime(sgDBPath & "Mnf.Btr")

    'On Error GoTo gObtainCompErr2
    'ilRet = 0
    'ilLowLimit = LBound(tgCompMnf)
    'If ilRet <> 0 Then
    '    sgCompMnfStamp = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tgCompMnf).Ptr <> 0 Then
        ilLowLimit = LBound(tgCompMnf)
    Else
        sgCompMnfStamp = ""
        ilLowLimit = 0
    End If

    If sgCompMnfStamp <> "" Then
        If StrComp(slStamp, sgCompMnfStamp, 1) = 0 Then
            'If UBound(tgCompMnf) > 1 Then
                gObtainComp = True
                Exit Function
            'End If
        End If
        slGetStamp = ""
    Else
        slGetStamp = "" 'gGetCSIStamp("COMPMNF")
    End If
    If slGetStamp <> "" Then
'        sgCompMnfStamp = slGetStamp
'        ilRet = csiGetAlloc("COMPMNF", ilStartIndex, ilEndIndex)
'        ReDim tgCompMnf(ilStartIndex To ilEndIndex) As MNFCOMPEXT
'        For ilLoop = LBound(tgCompMnf) To UBound(tgCompMnf) Step 1
'            ilRet = csiGetRec("COMPMNF", ilLoop, VarPtr(tgCompMnf(ilLoop)), LenB(tgCompMnf(ilLoop)))
'        Next ilLoop
        gObtainComp = True
    Else
        'ReDim tgCompMnf(1 To 1) As MNFCOMPEXT
        ReDim tgCompMnf(ilLowLimit To ilLowLimit) As MNFCOMPEXT
        hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            gObtainComp = False
            ilRet = btrClose(hlMnf)
            btrDestroy hlMnf
            Exit Function
        End If
        ilRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
        sgCompMnfStamp = slStamp
        ilUpperBound = UBound(tgCompMnf)
        ilExtLen = Len(tgCompMnf(ilUpperBound))  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlMnf) 'Obtain number of records
        btrExtClear hlMnf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlMnf, tlMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlMnf)
            btrDestroy hlMnf
            gObtainComp = True
            Exit Function
        Else
            If ilRet <> BTRV_ERR_NONE Then
                gObtainComp = False
                ilRet = btrClose(hlMnf)
                btrDestroy hlMnf
                Exit Function
            End If
        End If
        Call btrExtSetBounds(hlMnf, llNoRec, -1, "UC", "MNFCOMPEXTPK", MNFCOMPEXTPK) 'Set extract limits (all records)
        tlCharTypeBuff.sType = "C"
        ilOffSet = gFieldOffset("Mnf", "MnfType")
        ilRet = btrExtAddLogicConst(hlMnf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        ilOffSet = gFieldOffset("Mnf", "MnfCode")
        ilRet = btrExtAddField(hlMnf, ilOffSet, 2)  'Extract iCode field
        If ilRet <> BTRV_ERR_NONE Then
            gObtainComp = False
            ilRet = btrClose(hlMnf)
            btrDestroy hlMnf
            Exit Function
        End If
        ilOffSet = gFieldOffset("Mnf", "MnfName")
        ilRet = btrExtAddField(hlMnf, ilOffSet, 20)  'Extract First Name field
        If ilRet <> BTRV_ERR_NONE Then
            gObtainComp = False
            ilRet = btrClose(hlMnf)
            btrDestroy hlMnf
            Exit Function
        End If
        ilOffSet = gFieldOffset("Mnf", "MnfUnitType")
        ilRet = btrExtAddField(hlMnf, ilOffSet, 6)  'Extract First Name field
        If ilRet <> CP_MSG_NONE Then
            gObtainComp = False
            ilRet = btrClose(hlMnf)
            btrDestroy hlMnf
            Exit Function
        End If
        'ilRet = btrExtGetNextExt(hlMnf)    'Extract record
        ilUpperBound = UBound(tgCompMnf)
        ilRet = btrExtGetNext(hlMnf, tgCompMnf(ilUpperBound), ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                gObtainComp = False
                ilRet = btrClose(hlMnf)
                btrDestroy hlMnf
                Exit Function
            End If
            ilUpperBound = UBound(tgCompMnf)
            ilExtLen = Len(tgCompMnf(ilUpperBound))  'Extract operation record size
            'ilRet = btrExtGetFirst(hlMnf, tgCompMnf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlMnf, tgCompMnf(ilUpperBound), ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilUpperBound = ilUpperBound + 1
                'ReDim Preserve tgCompMnf(1 To ilUpperBound) As MNFCOMPEXT
                ReDim Preserve tgCompMnf(ilLowLimit To ilUpperBound) As MNFCOMPEXT
                ilRet = btrExtGetNext(hlMnf, tgCompMnf(ilUpperBound), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlMnf, tgCompMnf(ilUpperBound), ilExtLen, llRecPos)
                Loop
            Loop
        End If
        ilRet = btrClose(hlMnf)
        btrDestroy hlMnf
'        ilRet = csiSetStamp("COMPMNF", sgCompMnfStamp)
'        ilRet = csiSetAlloc("COMPMNF", LBound(tgCompMnf), UBound(tgCompMnf))
'        For ilLoop = LBound(tgCompMnf) To UBound(tgCompMnf) Step 1
'            ilRet = csiSetRec("COMPMNF", ilLoop, VarPtr(tgCompMnf(ilLoop)), LenB(tgCompMnf(ilLoop)))
'        Next ilLoop
        gObtainComp = True
    End If
    Exit Function
gObtainCompErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainCorpCal                  *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgMCC                  *
'*                                                     *
'*******************************************************
Function gObtainCorpCal() As Integer
'
'   ilRet = gObtainCorpCal ()
'   Where:
'       tgMCC() (I)- COF record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlCof As Integer        'Cof handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim slGetStamp As String
    Dim ilLowLimit As Integer

    slStamp = gFileDateTime(sgDBPath & "Cof.Btr")

    'On Error GoTo gObtainCorpCalErr2
    'ilRet = 0
    'ilLowLimit = LBound(tgMCof)
    'If ilRet <> 0 Then
    '    sgMCofStamp = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tgMCof).Ptr <> 0 Then
        ilLowLimit = LBound(tgMCof)
    Else
        sgMCofStamp = ""
        ilLowLimit = 0
    End If

    If sgMCofStamp <> "" Then
        If StrComp(slStamp, sgMCofStamp, 1) = 0 Then
            'If UBound(tgMCof) > 1 Then
                gObtainCorpCal = True
                Exit Function
            'End If
        End If
        slGetStamp = ""
    Else
        slGetStamp = "" 'gGetCSIStamp("COF")
    End If
    If slGetStamp <> "" Then
'        sgMCofStamp = slGetStamp
'        ilRet = csiGetAlloc("COF", ilStartIndex, ilEndIndex)
'        ReDim tgMCof(ilStartIndex To ilEndIndex) As COF
'        For ilLoop = LBound(tgMCof) To UBound(tgMCof) Step 1
'            ilRet = csiGetRec("COF", ilLoop, VarPtr(tgMCof(ilLoop)), LenB(tgMCof(ilLoop)))
'        Next ilLoop
        gObtainCorpCal = True
    Else
        'ReDim tgMCof(1 To 1) As COF
        ReDim tgMCof(ilLowLimit To ilLowLimit) As COF
        hlCof = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlCof, "", sgDBPath & "Cof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            gObtainCorpCal = False
            ilRet = btrClose(hlCof)
            btrDestroy hlCof
            Exit Function
        End If
        'ilRecLen = Len(tgMCof(1)) 'btrRecordLength(hlCof)  'Get and save record length
        ilRecLen = Len(tgMCof(ilLowLimit)) 'btrRecordLength(hlCof)  'Get and save record length
        sgMCofStamp = slStamp
        ilUpperBound = UBound(tgMCof)
        ilExtLen = Len(tgMCof(ilUpperBound))  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlCof) 'Obtain number of records
        btrExtClear hlCof   'Clear any previous extend operation
        ilRet = btrGetFirst(hlCof, tgMCof(ilUpperBound), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlCof)
            btrDestroy hlCof
            gObtainCorpCal = True
            Exit Function
        Else
            If ilRet <> BTRV_ERR_NONE Then
                gObtainCorpCal = False
                ilRet = btrClose(hlCof)
                btrDestroy hlCof
                Exit Function
            End If
        End If
        Call btrExtSetBounds(hlCof, llNoRec, -1, "UC", "COF", "") 'Set extract limits (all records)
        ilOffSet = 0
        ilRet = btrExtAddField(hlCof, ilOffSet, ilExtLen)  'Extract First Name field
        If ilRet <> BTRV_ERR_NONE Then
            gObtainCorpCal = False
            ilRet = btrClose(hlCof)
            btrDestroy hlCof
            Exit Function
        End If
        'ilRet = btrExtGetNextExt(hlCof)    'Extract record
        ilUpperBound = UBound(tgMCof)
        ilRet = btrExtGetNext(hlCof, tgMCof(ilUpperBound), ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                gObtainCorpCal = False
                ilRet = btrClose(hlCof)
                btrDestroy hlCof
                Exit Function
            End If
            ilUpperBound = UBound(tgMCof)
            ilExtLen = Len(tgMCof(ilUpperBound))  'Extract operation record size
            'ilRet = btrExtGetFirst(hlCof, tgCompMnf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlCof, tgMCof(ilUpperBound), ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilUpperBound = ilUpperBound + 1
                'ReDim Preserve tgMCof(1 To ilUpperBound) As COF
                ReDim Preserve tgMCof(ilLowLimit To ilUpperBound) As COF
                ilRet = btrExtGetNext(hlCof, tgMCof(ilUpperBound), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlCof, tgMCof(ilUpperBound), ilExtLen, llRecPos)
                Loop
            Loop
        End If
        ilRet = btrClose(hlCof)
        btrDestroy hlCof
'        ilRet = csiSetStamp("COF", sgMCofStamp)
'        ilRet = csiSetAlloc("COF", LBound(tgMCof), UBound(tgMCof))
'        For ilLoop = LBound(tgMCof) To UBound(tgMCof) Step 1
'            ilRet = csiSetRec("COF", ilLoop, VarPtr(tgMCof(ilLoop)), LenB(tgMCof(ilLoop)))
'        Next ilLoop
        gObtainCorpCal = True
    End If
    Exit Function
gObtainCorpCalErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainLatestRcf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgMRcf, tgMRif, tgMRdf *
'*                                                     *
'*******************************************************
Function gObtainLatestRcf(tlRcf As RCF) As Integer 'VBC NR
'
'   ilRet = gObtainLatestRcf (tlRcf)
'   Where:
'       tlMRcf (O)- RCF record structure to be created
'       ilRet (O)- True = Found; False = error
'
    Dim hlRcf As Integer        'Rcf handle 'VBC NR
    Dim tlTRcf As RCF           'Test image for rcf 'VBC NR
    Dim ilRecLen As Integer     'Record length 'VBC NR
    Dim llNoRec As Long         'Number of records in Sof 'VBC NR
    Dim ilExtLen As Integer 'VBC NR
    Dim llRecPos As Long        'Record location 'VBC NR
    Dim ilRet As Integer 'VBC NR
    Dim ilOffSet As Integer 'VBC NR
    Dim llDate As Long 'VBC NR
    Dim llTDate As Long 'VBC NR

    gObtainLatestRcf = False 'VBC NR
    hlRcf = CBtrvTable(ONEHANDLE) 'CBtrvTable() 'VBC NR
    ilRet = btrOpen(hlRcf, "", sgDBPath & "Rcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'VBC NR
    If ilRet <> BTRV_ERR_NONE Then 'VBC NR
        igBtrError = ilRet 'VBC NR
        ilRet = btrClose(hlRcf) 'VBC NR
        btrDestroy hlRcf 'VBC NR
        Exit Function 'VBC NR
    End If 'VBC NR
    tlRcf.iYear = 0 'VBC NR
    ilRecLen = Len(tlTRcf) 'btrRecordLength(hlAdf)  'Get and save record length 'VBC NR
    ilExtLen = Len(tlTRcf)  'Extract operation record size 'VBC NR
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records 'VBC NR
    btrExtClear hlRcf   'Clear any previous extend operation 'VBC NR
    ilRet = btrGetFirst(hlRcf, tlTRcf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation 'VBC NR
    If ilRet <> BTRV_ERR_END_OF_FILE Then 'VBC NR
        If ilRet <> BTRV_ERR_NONE Then 'VBC NR
            igBtrError = ilRet 'VBC NR
            ilRet = btrClose(hlRcf) 'VBC NR
            btrDestroy hlRcf 'VBC NR
            Exit Function 'VBC NR
        End If 'VBC NR
        Call btrExtSetBounds(hlRcf, llNoRec, -1, "UC", "RCF", "") 'Set extract limits (all records) 'VBC NR
        ilOffSet = 0 'VBC NR
        ilRet = btrExtAddField(hlRcf, ilOffSet, ilRecLen)  'Extract iCode field 'VBC NR
        If ilRet <> BTRV_ERR_NONE Then 'VBC NR
            igBtrError = ilRet 'VBC NR
            ilRet = btrClose(hlRcf) 'VBC NR
            btrDestroy hlRcf 'VBC NR
            Exit Function 'VBC NR
        End If 'VBC NR
        'ilRet = btrExtGetNextExt(hlAdf)    'Extract record
        ilRet = btrExtGetNext(hlRcf, tlTRcf, ilExtLen, llRecPos) 'VBC NR
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then 'VBC NR
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then 'VBC NR
                igBtrError = ilRet 'VBC NR
                ilRet = btrClose(hlRcf) 'VBC NR
                btrDestroy hlRcf 'VBC NR
                Exit Function 'VBC NR
            End If 'VBC NR
            ilExtLen = Len(tlTRcf)  'Extract operation record size 'VBC NR
            'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT 'VBC NR
                ilRet = btrExtGetNext(hlRcf, tlTRcf, ilExtLen, llRecPos) 'VBC NR
            Loop 'VBC NR
            Do While ilRet = BTRV_ERR_NONE 'VBC NR
                If Left$(tlTRcf.sName, 1) <> "~" Then 'VBC NR
                    If tlTRcf.iYear > tlRcf.iYear Then 'VBC NR
                        tlRcf = tlTRcf 'VBC NR
                        gUnpackDateLong tlTRcf.iStartDate(0), tlTRcf.iStartDate(1), llDate 'VBC NR
                        gObtainLatestRcf = True 'VBC NR
                    ElseIf tlTRcf.iYear = tlRcf.iYear Then 'VBC NR
                        gUnpackDateLong tlTRcf.iStartDate(0), tlTRcf.iStartDate(1), llTDate 'VBC NR
                        If llTDate > llDate Then 'VBC NR
                            tlRcf = tlTRcf 'VBC NR
                            llDate = llTDate 'VBC NR
                        End If 'VBC NR
                    End If 'VBC NR
                End If 'VBC NR
                ilRet = btrExtGetNext(hlRcf, tlTRcf, ilExtLen, llRecPos) 'VBC NR
                Do While ilRet = BTRV_ERR_REJECT_COUNT 'VBC NR
                    ilRet = btrExtGetNext(hlRcf, tlTRcf, ilExtLen, llRecPos) 'VBC NR
                Loop 'VBC NR
            Loop 'VBC NR
        End If 'VBC NR
    End If 'VBC NR
    ilRet = btrClose(hlRcf) 'VBC NR
    btrDestroy hlRcf 'VBC NR
    Exit Function 'VBC NR
End Function 'VBC NR

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainMissedSpot               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain missed spots for         *
'*                     specified dates and advertiser  *
'*                                                     *
'*******************************************************
Function gObtainMissedSpot(slType As String, ilInVefCode As Integer, ilAdfCode As Integer, ilGameNo As Integer, slStartDate As String, slEndDate As String, ilSortOrder As Integer, tlSdfMdExt() As SDFMDEXT, slSortCodeTag As String, Optional blCheckOtherVehicles As Boolean = False, Optional blIncludeMissedMG As Boolean = False) As Integer
'
'   ilRet = gObtainMissedSpot (slType, ilInVefCode, ilAdfCode, slStartDate, slEndDate, ilSortOrder, lbcSortCtrl, tlSdfMdExt())
'   Where:
'       slType(I)- "M"=Missed; "C"=Cancelled; "H"=Hidden
'       ilInVefCode(I)- Vehicle code
'       ilAdfCode() (I)- Advertiser code (-1=All advertiser)
'       slStartDate(I)- Start date
'       slEndDate(I)- End date
'       ilSortOrder(I)- 0=Date|Time|Advt|CntrNo|LineNo; 1=Date|Day|Advt|CntrNo|LineNo; 2=Advt|ContrNo|Date|Time; -1=Ignore sort
'                       Following fields not set if ilSortOrder = -1
'                       lCntrNo; sCntrType; sAdvtAbbr
'       tlSdfMdExt(I/O)- Record images in sorted order
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- True if missed obtained OK; False if error
'
    Dim slStamp As String    'date/time stamp
    Dim hlSdf As Integer        'Sdf handle
    Dim ilSdfRecLen As Integer     'Record length
    Dim tlSdf As SDF
    Dim hlChf As Integer        'Sdf handle
    Dim ilChfRecLen As Integer     'Record length
    Dim tlChf As CHF
    Dim tlChfSrchKey As LONGKEY0
    Dim hlClf As Integer        'Sdf handle
    Dim ilClfRecLen As Integer     'Record length
    Dim tlClf As CLF
    Dim tlClfSrchKey As CLFKEY0
    Dim hlRaf As Integer        'Sdf handle
    Dim ilRafRecLen As Integer     'Record length
    Dim tlRaf As RAF
    Dim tlRafSrchKey As LONGKEY0
    Dim tlSdfSrchKey2 As SDFKEY2
    Dim tlSdfSrchKey3 As LONGKEY0

    Dim hlSxf As Integer        'Sdf handle
    Dim ilSxfRecLen As Integer     'Record length
    Dim tlSxf As SXF
    Dim tlSxfSrchKey2 As SXFKEY2
    
    Dim slStr As String
    'Dim llNoRec As Long
    Dim ilRet As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim slTime As String
    Dim llTime As Long
    Dim slNameSort As String
    Dim slCnt As String
    Dim slLine As String
    'Dim ilExtLen As Integer
    Dim ilUpper As Integer
    Dim illoop As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slDay As String
    Dim llLen As Long
    Dim ilVef As Integer
    Dim ilVff As Integer
    Dim ilVefCode As Integer
    '8/29/14: Add Game # to sort # 2: TTP 6713
    Dim slGameNo As String
    Dim ilPass As Integer
    Dim ilLowLimit As Integer

    llLen = 0
    slStr = slType & Trim$(str$(ilInVefCode)) & Trim$(str$(ilAdfCode)) & Trim$(str$(ilGameNo))
    llDate = gDateValue(slStartDate)
    llStartDate = llDate
    slDate = Trim$(str$(llDate))
    Do While Len(slDate) < 6
        slDate = "0" & slDate
    Loop
    slStr = slStr & slDate
    llDate = gDateValue(slEndDate)
    llEndDate = llDate
    slDate = Trim$(str$(llDate))
    Do While Len(slDate) < 6
        slDate = "0" & slDate
    Loop
    slStr = slStr & slDate
    slStamp = slStr 'FileDateTime(sgDBPath & "Sdf.Btr") & slStr

    'On Error GoTo gObtainMissedSpotErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSdfMdExt)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSdfMdExt).Ptr <> 0 Then
        ilLowLimit = LBound(tlSdfMdExt)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            gObtainMissedSpot = True
            Exit Function
        End If
    End If
    If Not gObtainAdvt() Then
        gObtainMissedSpot = False
        Exit Function
    End If
    gObtainMissedSpot = True
    hlSdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainMissedSpot = False
        ilRet = btrClose(hlSdf)
        btrDestroy hlSdf
        Exit Function
    End If
    ilSdfRecLen = Len(tlSdf) 'btrRecordLength(hlSdf)  'Get and save record length
    hlChf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainMissedSpot = False
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlChf)
        btrDestroy hlSdf
        btrDestroy hlChf
        Exit Function
    End If
    ilChfRecLen = Len(tlChf) 'btrRecordLength(hlSdf)  'Get and save record length
    'If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
    hlClf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainMissedSpot = False
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlChf)
        ilRet = btrClose(hlClf)
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlClf
        Exit Function
    End If
    ilClfRecLen = Len(tlClf) 'btrRecordLength(hlSdf)  'Get and save record length
    
    hlSxf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlSxf, "", sgDBPath & "Sxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainMissedSpot = False
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlChf)
        ilRet = btrClose(hlClf)
        ilRet = btrClose(hlSxf)
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlClf
        btrDestroy hlSxf
        Exit Function
    End If
    ilSxfRecLen = Len(tlSxf) 'btrRecordLength(hlSdf)  'Get and save record length
    
    If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
        hlRaf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            gObtainMissedSpot = False
            ilRet = btrClose(hlSdf)
            ilRet = btrClose(hlChf)
            ilRet = btrClose(hlClf)
            ilRet = btrClose(hlSxf)
            ilRet = btrClose(hlRaf)
            btrDestroy hlSdf
            btrDestroy hlChf
            btrDestroy hlClf
            btrDestroy hlSxf
            btrDestroy hlRaf
            Exit Function
        End If
        ilRafRecLen = Len(tlRaf) 'btrRecordLength(hlSdf)  'Get and save record length
    End If
    'ReDim tlSdfMdExt(1 To 1) As SDFMDEXT
    'temporarily set ubound to 0, after sort, add one.
    ReDim tlSdfMdExt(ilLowLimit To ilLowLimit) As SDFMDEXT
    slSortCodeTag = slStamp
    'ilExtLen = Len(tlSdfMdExt(1)) - 4'Extract operation record size
    'llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hlSdf   'Clear any previous extend operation
    'ilRet = btrGetFirst(hlSdf, tlSdf, ilSdfRecLen, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    ReDim ilCheckVefCode(0 To 1) As Integer
    ilCheckVefCode(0) = ilInVefCode
    If blCheckOtherVehicles Then
        ilVef = gBinarySearchVef(ilInVefCode)
        If ilVef <> -1 Then
            If tgMVef(ilVef).sType = "G" Then
                For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    If tgMVef(ilVef).iCode <> ilInVefCode Then
                        If (tgMVef(ilVef).sType = "G") And (tgMVef(ilVef).sState <> "D") Then
                            For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                                If tgMVef(ilVef).iCode = tgVff(ilVff).iVefCode Then
                                    If tgVff(ilVff).sMoveSportToSport = "Y" Then
                                        ilCheckVefCode(UBound(ilCheckVefCode)) = tgMVef(ilVef).iCode
                                        ReDim Preserve ilCheckVefCode(0 To UBound(ilCheckVefCode) + 1) As Integer
                                    End If
                                    Exit For
                                End If
                            Next ilVff
                        ElseIf ((tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "S")) And (tgMVef(ilVef).sState <> "D") Then
                            For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                                If ilInVefCode = tgVff(ilVff).iVefCode Then
                                    If tgVff(ilVff).sMoveNonToSport = "Y" Then
                                        ilCheckVefCode(UBound(ilCheckVefCode)) = tgMVef(ilVef).iCode
                                        ReDim Preserve ilCheckVefCode(0 To UBound(ilCheckVefCode) + 1) As Integer
                                    End If
                                    Exit For
                                End If
                            Next ilVff
                        End If
                    End If
                Next ilVef
            Else
                For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    If tgMVef(ilVef).iCode <> ilInVefCode Then
                        If (tgMVef(ilVef).sType = "G") And (tgMVef(ilVef).sState <> "D") Then
                            For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                                If tgMVef(ilVef).iCode = tgVff(ilVff).iVefCode Then
                                    If tgVff(ilVff).sMoveSportToNon = "Y" Then
                                        ilCheckVefCode(UBound(ilCheckVefCode)) = tgMVef(ilVef).iCode
                                        ReDim Preserve ilCheckVefCode(0 To UBound(ilCheckVefCode) + 1) As Integer
                                    End If
                                    Exit For
                                End If
                            Next ilVff
                        End If
                    End If
                Next ilVef
            End If
        End If
    End If
    For ilPass = 0 To 1 Step 1
        For ilVef = 0 To UBound(ilCheckVefCode) - 1 Step 1
            ilVefCode = ilCheckVefCode(ilVef)
        
            If ilPass = 0 Then
                tlSdfSrchKey2.iVefCode = ilVefCode
                tlSdfSrchKey2.sSchStatus = slType
                If ilAdfCode > 0 Then
                    tlSdfSrchKey2.iAdfCode = ilAdfCode
                Else
                    tlSdfSrchKey2.iAdfCode = 0
                End If
                gPackDate slStartDate, tlSdfSrchKey2.iDate(0), tlSdfSrchKey2.iDate(1)
                tlSdfSrchKey2.iTime(0) = 0
                tlSdfSrchKey2.iTime(1) = 0
                ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
            Else
                tlSxfSrchKey2.sType = "G"
                gPackDate slStartDate, tlSxfSrchKey2.iMissedDate(0), tlSxfSrchKey2.iMissedDate(1)
                ilRet = btrGetGreaterOrEqual(hlSxf, tlSxf, ilSxfRecLen, tlSxfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
                If ilRet = BTRV_ERR_NONE Then
                    tlSdfSrchKey3.lCode = tlSxf.lSdfCode
                    ilRet = btrGetEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                End If
            End If
            'This code added as replacement for Ext operation
            ilUpper = UBound(tlSdfMdExt)
            'Do While (ilRet = BTRV_ERR_NONE) And (tlSdf.iVefCode = ilVefCode) And (tlSdf.sSchStatus = slType)
            Do While (ilRet = BTRV_ERR_NONE)
                If (ilPass = 0) And (tlSdf.iVefCode <> ilVefCode) Or (tlSdf.sSchStatus <> slType) Then
                    Exit Do
                End If
                If (ilPass = 1) And (tlSxf.sType <> "G") Then
                    Exit Do
                End If
                If ilPass = 0 Then
                    gUnpackDate tlSdf.iDate(0), tlSdf.iDate(1), slDate
                Else
                    gUnpackDate tlSxf.iMissedDate(0), tlSxf.iMissedDate(1), slDate
                End If
                If (ilAdfCode > 0) Then
                    If (tlSdf.iAdfCode <> ilAdfCode) Then
                        Exit Do
                    End If
                    If gDateValue(slDate) > llEndDate Then
                        Exit Do
                    End If
                End If
                ilRet = False
                If ilVef <= 0 Then
                    If ((gDateValue(slDate) >= llStartDate) And (gDateValue(slDate) <= llEndDate) And (((ilGameNo = 0) And (tlSdf.iGameNo = 0)) Or ((ilGameNo = -1) And (tlSdf.iGameNo > 0)))) Or ((ilGameNo > 0) And (tlSdf.iGameNo = ilGameNo)) Then
                        ilRet = True
                    End If
                Else
                    If (gDateValue(slDate) >= llStartDate) And (gDateValue(slDate) <= llEndDate) Then
                        ilRet = True
                    End If
                End If
                If (ilPass = 1) And (ilRet = True) Then
                    '9/24/15: Always include the PMG information even if on same vehicle
                    'If (tlSdf.sSchStatus <> "M") Or (tlSxf.iMissedVefCode <> ilVefCode) Then
                    If (tlSdf.sSchStatus <> "M") Then
                        ilRet = False
                    Else
                        If (tlSdf.iVefCode <> ilVefCode) And (tlSxf.iMissedVefCode <> ilVefCode) Then
                            ilRet = False
                        Else
                            For illoop = LBound(tlSdfMdExt) To ilUpper - 1 Step 1
                                If tlSdf.lCode = tlSdfMdExt(illoop).lSdfCode Then
                                    ilRet = False
                                    Exit For
                                End If
                            Next illoop
                        End If
                    End If
                End If
                If ilRet Then
                    ilRet = btrGetPosition(hlSdf, tlSdfMdExt(ilUpper).lRecPos)
                    tlSdfMdExt(ilUpper).lSdfCode = tlSdf.lCode
                    tlSdfMdExt(ilUpper).lChfCode = tlSdf.lChfCode
                    tlSdfMdExt(ilUpper).iLineNo = tlSdf.iLineNo
                    tlSdfMdExt(ilUpper).lFsfCode = tlSdf.lFsfCode
                    tlSdfMdExt(ilUpper).iAdfCode = tlSdf.iAdfCode
                    tlSdfMdExt(ilUpper).iDate(0) = tlSdf.iDate(0)
                    tlSdfMdExt(ilUpper).iDate(1) = tlSdf.iDate(1)
                    tlSdfMdExt(ilUpper).iTime(0) = tlSdf.iTime(0)
                    tlSdfMdExt(ilUpper).iTime(1) = tlSdf.iTime(1)
                    tlSdfMdExt(ilUpper).iMnfMissed = tlSdf.iMnfMissed
                    tlSdfMdExt(ilUpper).iLen = tlSdf.iLen
                    tlSdfMdExt(ilUpper).sShow = ""
                    tlSdfMdExt(ilUpper).iRotNo = tlSdf.iRotNo
                    tlSdfMdExt(ilUpper).sPtType = tlSdf.sPtType
                    tlSdfMdExt(ilUpper).lCopyCode = tlSdf.lCopyCode
                    tlSdfMdExt(ilUpper).lchfcxfCode = -1
                    tlSdfMdExt(ilUpper).lchfcxfInt = -1
                    tlSdfMdExt(ilUpper).iGameNo = tlSdf.iGameNo
                    tlSdfMdExt(ilUpper).lClfCxfCode = -1
                    tlSdfMdExt(ilUpper).sCITFlag = ""
                    tlSdfMdExt(ilUpper).sCopyCIT = ""
                    tlSdfMdExt(ilUpper).iSelected = 0
                    tlSdfMdExt(ilUpper).lSxfCode = 0
                    If ilPass = 1 Then
                        tlSdfMdExt(ilUpper).lSxfCode = tlSxf.lCode
                    End If
    
                    If ilSortOrder <> -1 Then
                        If (ilSortOrder = 0) Or (ilSortOrder = 2) Then
                            gUnpackDateForSort tlSdfMdExt(ilUpper).iDate(0), tlSdfMdExt(ilUpper).iDate(1), slDate
                        Else
                            gUnpackDate tlSdfMdExt(ilUpper).iDate(0), tlSdfMdExt(ilUpper).iDate(1), slDate
                            slDay = Trim$(str$(gWeekDayStr(slDate)))
                            gUnpackDateForSort tlSdfMdExt(ilUpper).iDate(0), tlSdfMdExt(ilUpper).iDate(1), slDate
                        End If
                        gUnpackTime tlSdfMdExt(ilUpper).iTime(0), tlSdfMdExt(ilUpper).iTime(1), "A", "1", slTime
                        llTime = CLng(gTimeToCurrency(slTime, False))
                        slTime = Trim$(str$(llTime))
                        Do While Len(slTime) < 7
                            slTime = "0" & slTime
                        Loop
                        If tlSdfMdExt(ilUpper).lChfCode > 0 Then
                            If tlChf.lCode <> tlSdfMdExt(ilUpper).lChfCode Then
                                tlChfSrchKey.lCode = tlSdfMdExt(ilUpper).lChfCode
                                ilRet = btrGetEqual(hlChf, tlChf, ilChfRecLen, tlChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet <> BTRV_ERR_NONE Then
                                    'gObtainMissedSpot = False
                                    'ilRet = btrClose(hlSdf)
                                    'ilRet = btrClose(hlChf)
                                    'btrDestroy hlSdf
                                    'btrDestroy hlChf
                                    'Exit Function
                                    tlChf.lCntrNo = 0
                                    tlChf.iAdfCode = 0
                                    tlSdfMdExt(ilUpper).iAdfCode = 0
                                End If
                            End If
                            tlSdfMdExt(ilUpper).lCntrNo = tlChf.lCntrNo
                            slCnt = Trim$(str$(tlChf.lCntrNo))
                            Do While Len(slCnt) < 7
                                slCnt = "0" & slCnt
                            Loop
                            slLine = Trim$(str$(tlSdfMdExt(ilUpper).iLineNo))
                            tlSdfMdExt(ilUpper).sCntrType = tlChf.sType
                            tlSdfMdExt(ilUpper).sAdvtAbbr = ""
                        Else
                            slCnt = "Feed   "
                            slLine = "   "
                            tlSdfMdExt(ilUpper).sCntrType = ""
                            tlSdfMdExt(ilUpper).sAdvtAbbr = ""
                            tlChf.iAdfCode = tlSdfMdExt(ilUpper).iAdfCode
                        End If
                        If tlChf.iAdfCode <> 0 Then
                            'For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                            '    If tlChf.iAdfCode = tgCommAdf(ilLoop).iCode Then
                                illoop = gBinarySearchAdf(tlChf.iAdfCode)
                                If illoop <> -1 Then
                                    tlSdfMdExt(ilUpper).sAdvtAbbr = tgCommAdf(illoop).sAbbr
                            '        Exit For
                                End If
                            'Next ilLoop
                        Else
                            tlSdfMdExt(ilUpper).sAdvtAbbr = "Missing"
                        End If
                        If ilSortOrder = 0 Then
                            slNameSort = slDate & "|" & slTime & "|" & tlSdfMdExt(ilUpper).sAdvtAbbr & "|" & slCnt & "|" & slLine
                        ElseIf ilSortOrder = 2 Then
                            '8/29/14: Add Game # to sort # 2: TTP 6713
                            'slNameSort = tlSdfMdExt(ilUpper).sAdvtAbbr & "|" & slCnt & "|" & slDate & "|" & slTime
                            If (ilGameNo = -1) And (tlSdfMdExt(ilUpper).iGameNo > 0) Then
                                slGameNo = tlSdfMdExt(ilUpper).iGameNo
                                Do While Len(slGameNo) < 4
                                    slGameNo = "0" & slGameNo
                                Loop
                                slNameSort = slGameNo & "|" & tlSdfMdExt(ilUpper).sAdvtAbbr & "|" & slCnt & "|" & slDate & "|" & slTime
                            Else
                                slNameSort = tlSdfMdExt(ilUpper).sAdvtAbbr & "|" & slCnt & "|" & slDate & "|" & slTime
                            End If
                        Else
                            slNameSort = slDate & "|" & slDay & "|" & tlSdfMdExt(ilUpper).sAdvtAbbr & "|" & slCnt & "|" & slLine
                        End If
                        tlSdfMdExt(ilUpper).sKey = slNameSort
                    Else
                        tlSdfMdExt(ilUpper).sKey = ""
                    End If
                    'Bypass spots missing contracts because they might cause other problems (i.e. right click on spot to see information)
                    tlSdfMdExt(ilUpper).sLiveCopy = ""
                    tlSdfMdExt(ilUpper).sNetRegionAbbr = ""
                    If tlSdfMdExt(ilUpper).iAdfCode <> 0 Then
                        tlSdfMdExt(ilUpper).sSplitNetwork = "N"
                        'If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
                            ilClfRecLen = Len(tlClf)
                            tlClfSrchKey.lChfCode = tlSdf.lChfCode
                            tlClfSrchKey.iLine = tlSdf.iLineNo
                            tlClfSrchKey.iCntRevNo = 32000 ' 0 show latest Revision
                            tlClfSrchKey.iPropVer = 32000 ' 0 show latest version
                            ilRet = btrGetGreaterOrEqual(hlClf, tlClf, ilClfRecLen, tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            Do While (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlSdf.lChfCode) And (tlClf.iLine = tlSdf.iLineNo) And ((tlClf.sSchStatus <> "M") And (tlClf.sSchStatus <> "F"))
                                ilRet = btrGetNext(hlClf, tlClf, ilClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                            If (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tlSdf.lChfCode) And (tlClf.iLine = tlSdf.iLineNo) Then
                                tlSdfMdExt(ilUpper).sLiveCopy = tlClf.sLiveCopy
                                If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
                                    If tlClf.lRafCode > 0 Then
                                        tlSdfMdExt(ilUpper).sSplitNetwork = "Y"
                                        tlRafSrchKey.lCode = tlClf.lRafCode
                                        ilRet = btrGetEqual(hlRaf, tlRaf, ilRafRecLen, tlRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet = BTRV_ERR_NONE Then
                                            If Trim$(tlRaf.sAbbr) <> "" Then
                                                tlSdfMdExt(ilUpper).sNetRegionAbbr = tlRaf.sAbbr
                                            Else
                                                tlSdfMdExt(ilUpper).sNetRegionAbbr = Left$(tlRaf.sName, 5)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        'End If
                        ilUpper = ilUpper + 1
                        'ReDim Preserve tlSdfMdExt(1 To ilUpper) As SDFMDEXT
                        ReDim Preserve tlSdfMdExt(ilLowLimit To ilUpper) As SDFMDEXT
                    End If
        '        ilRet = btrExtGetNext(hlSdf, tlSdfMdExt(ilUpper), ilExtLen, llRecPos)
                End If
                If ilPass = 0 Then
                    ilRet = btrGetNext(hlSdf, tlSdf, ilSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Else
                    ilRet = btrGetNext(hlSxf, tlSxf, ilSxfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        tlSdfSrchKey3.lCode = tlSxf.lSdfCode
                        ilRet = btrGetEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            tlSdf.iVefCode = tlSxf.iMissedVefCode
                            tlSdf.iGameNo = tlSxf.iMissedGameNo
                            'L.Bianchi 06/03/2021 TTP 10201
                            'tlSdf.iDate(0) = tlSxf.iMissedDate(0)
                            'tlSdf.iDate(1) = tlSxf.iMissedDate(1)
                            'tlSdf.iTime(0) = tlSxf.iMissedTime(0)
                            'tlSdf.iTime(1) = tlSxf.iMissedTime(1)
                        End If
                    End If
                End If
            Loop
        Next ilVef
        If Not blIncludeMissedMG Then
            Exit For
        End If
    Next ilPass
    ilRet = btrClose(hlSdf)
    btrDestroy hlSdf
    ilRet = btrClose(hlChf)
    btrDestroy hlChf
    'If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
    ilRet = btrClose(hlClf)
    btrDestroy hlClf
    ilRet = btrClose(hlSxf)
    btrDestroy hlSxf
    If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
        ilRet = btrClose(hlRaf)
        btrDestroy hlRaf
    End If
    gObtainMissedSpot = True
    If ilSortOrder <> -1 Then
        'If UBound(tlSdfMdExt) - 1 > 1 Then
        If UBound(tlSdfMdExt) - 1 > ilLowLimit Then
            'ArraySortTyp fnAV(tlSdfMdExt(), 1), UBound(tlSdfMdExt) - 1, 0, LenB(tlSdfMdExt(1)), 0, LenB(tlSdfMdExt(1).sKey), 0
            ArraySortTyp fnAV(tlSdfMdExt(), ilLowLimit), UBound(tlSdfMdExt) - ilLowLimit, 0, LenB(tlSdfMdExt(ilLowLimit)), 0, LenB(tlSdfMdExt(ilLowLimit).sKey), 0
        End If
        '8/29/14: Remove Game # from sort # 2: TTP 6713
        If (ilGameNo = -1) And (tlSdfMdExt(ilUpper).iGameNo > 0) And (ilSortOrder = 2) Then
            For ilUpper = LBound(tlSdfMdExt) To UBound(tlSdfMdExt) - 1 Step 1
                tlSdfMdExt(ilUpper).sKey = Mid$(tlSdfMdExt(ilUpper).sKey, 6)
            Next ilUpper
        End If
    End If
    'set arrary so that lower bound is one instead of zero so that Spots screen will work
    If ilLowLimit = 0 Then
        ReDim Preserve tlSdfMdExt(0 To UBound(tlSdfMdExt) + 1) As SDFMDEXT
        For illoop = UBound(tlSdfMdExt) - 1 To 0 Step -1
            tlSdfMdExt(illoop + 1) = tlSdfMdExt(illoop)
        Next illoop
        tlSdfMdExt(0).lChfCode = -1
        tlSdfMdExt(0).lCntrNo = -1
        tlSdfMdExt(0).lRecPos = -1
        tlSdfMdExt(0).iSelected = 0
        tlSdfMdExt(0).lSdfCode = -1
        tlSdfMdExt(0).sKey = ""
    End If
    Exit Function
gObtainMissedSpotErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainMnfForType               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate Dayparts               *
'*                                                     *
'*******************************************************
Function gObtainMnfForType(slMnfType As String, slMMnfStamp As String, tlMMnf() As MNF) As Integer
'
'   ilRet = gObtainMnfForType (slMnfType, slMMnfStamp, tlMMnf())
'   Where:
'       slMnfType(I)- Mnf type
'       slMMnfStamp(I)- Time and type stamp
'       tlMMnf() (I)- MNF record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlMnf As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim ilLowLimit As Integer

    slStamp = gFileDateTime(sgDBPath & "Mnf.Btr") & slMnfType

    'On Error GoTo gObtainMnfForTypeErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlMMnf)
    'If ilRet <> 0 Then
    '    slMMnfStamp = ""
    '    'ilLowLimit = 1
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlMMnf).Ptr <> 0 Then
        ilLowLimit = LBound(tlMMnf)
    Else
        slMMnfStamp = ""
        ilLowLimit = 0
    End If

    If slMMnfStamp <> "" Then
        If StrComp(slStamp, slMMnfStamp, 1) = 0 Then
            'If UBound(tlMMnf) > 1 Then
                gObtainMnfForType = True
                Exit Function
            'End If
        End If
    End If
    ReDim tlMMnf(ilLowLimit To ilLowLimit) As MNF
    hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainMnfForType = False
        ilRet = btrClose(hlMnf)
        btrDestroy hlMnf
        Exit Function
    End If
    ilRecLen = Len(tlMMnf(ilLowLimit)) 'btrRecordLength(hlMnf)  'Get and save record length
    slMMnfStamp = slStamp
    ilUpperBound = UBound(tlMMnf)
    ilExtLen = Len(tlMMnf(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlMnf) 'Obtain number of records
    btrExtClear hlMnf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlMnf, tlMMnf(ilUpperBound), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlMnf)
        btrDestroy hlMnf
        gObtainMnfForType = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainMnfForType = False
            ilRet = btrClose(hlMnf)
            btrDestroy hlMnf
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlMnf, llNoRec, -1, "UC", "MNF", "") 'Set extract limits (all records)
    tlCharTypeBuff.sType = Left$(slMnfType, 1)    'Get left most character only
    ilOffSet = gFieldOffset("Mnf", "MnfType")
    ilRet = btrExtAddLogicConst(hlMnf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
    ilOffSet = 0
    ilRet = btrExtAddField(hlMnf, ilOffSet, ilExtLen)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainMnfForType = False
        ilRet = btrClose(hlMnf)
        btrDestroy hlMnf
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlMnf)    'Extract record
    ilUpperBound = UBound(tlMMnf)
    ilRet = btrExtGetNext(hlMnf, tlMMnf(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainMnfForType = False
            ilRet = btrClose(hlMnf)
            btrDestroy hlMnf
            Exit Function
        End If
        ilUpperBound = UBound(tlMMnf)
        ilExtLen = Len(tlMMnf(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlMnf, tgCompMnf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlMnf, tlMMnf(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If (Left$(slMnfType, 1) = "H") And (Len(slMnfType) = 2) Then
                If right$(slMnfType, 1) = Trim$(tlMMnf(ilUpperBound).sUnitType) Then
                    ilUpperBound = ilUpperBound + 1
                    ReDim Preserve tlMMnf(ilLowLimit To ilUpperBound) As MNF
                End If
            ElseIf (Left$(slMnfType, 1) = "M") Then
                'Ignore Affiliate missed reasons
                If (Trim$(tlMMnf(ilUpperBound).sCodeStn) <> "A") And (Trim$(tlMMnf(ilUpperBound).sCodeStn) <> "A") Then
                    ilUpperBound = ilUpperBound + 1
                    ReDim Preserve tlMMnf(ilLowLimit To ilUpperBound) As MNF
                End If
            Else
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tlMMnf(ilLowLimit To ilUpperBound) As MNF
            End If
            ilRet = btrExtGetNext(hlMnf, tlMMnf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlMnf, tlMMnf(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlMnf)
    btrDestroy hlMnf
    gObtainMnfForType = True
    Exit Function
gObtainMnfForTypeErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainRcfRifRdf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgMRcf, tgMRif, tgMRdf *
'*                                                     *
'*******************************************************
Function gObtainRcfRifRdf() As Integer
'
'   ilRet = gObtainRcfRifRdf ()
'   Where:
'       tgMRcf() (O)- RCF record structure to be created
'       tgMRif() (O)- RCF record structure to be created
'       tgMRdf() (O)- RCF record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Rcf or Rif or Rdf date/time stamp
    Dim hlRcf As Integer        'Rcf handle
    Dim hlRif As Integer        'Rif handle
    Dim hlRdf As Integer        'Rdf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim llUpperBound As Long
    Dim ilRifLowLimit As Integer
    Dim llRifLowLimit As Long
    Dim ilRdfLowLimit As Integer

    slStamp = gFileDateTime(sgDBPath & "Rcf.Btr")

    'On Error GoTo gObtainRcfRdfErr2
    'ilRet = 0
    'ilRifLowLimit = LBound(tgMRcf)
    'If ilRet <> 0 Then
    '    sgMRcfStamp = ""
    '    ilRifLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tgMRcf).Ptr <> 0 Then
        ilRifLowLimit = LBound(tgMRcf)
    Else
        sgMRcfStamp = ""
        ilRifLowLimit = 0
    End If

    If (sgMRcfStamp = "") Or (slStamp <> sgMRcfStamp) Then
        'ReDim tgMRcf(1 To 1) As RCF
        ReDim tgMRcf(ilRifLowLimit To ilRifLowLimit) As RCF
        hlRcf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlRcf, "", sgDBPath & "Rcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            gObtainRcfRifRdf = False
            ilRet = btrClose(hlRcf)
            btrDestroy hlRcf
            Exit Function
        End If
        'ilRecLen = Len(tgMRcf(1)) 'btrRecordLength(hlAdf)  'Get and save record length
        ilRecLen = Len(tgMRcf(ilRifLowLimit)) 'btrRecordLength(hlAdf)  'Get and save record length
        sgMRcfStamp = slStamp
        ilUpperBound = UBound(tgMRcf)
        ilExtLen = Len(tgMRcf(ilUpperBound))  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        btrExtClear hlRcf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlRcf, tgMRcf(ilUpperBound), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                gObtainRcfRifRdf = False
                ilRet = btrClose(hlRcf)
                btrDestroy hlRcf
                Exit Function
            End If
            Call btrExtSetBounds(hlRcf, llNoRec, -1, "UC", "RCF", "") 'Set extract limits (all records)
            ilOffSet = 0
            ilRet = btrExtAddField(hlRcf, ilOffSet, ilRecLen)  'Extract iCode field
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                gObtainRcfRifRdf = False
                ilRet = btrClose(hlRcf)
                btrDestroy hlRcf
                Exit Function
            End If
            'ilRet = btrExtGetNextExt(hlAdf)    'Extract record
            ilUpperBound = UBound(tgMRcf)
            ilRet = btrExtGetNext(hlRcf, tgMRcf(ilUpperBound), ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                    igBtrError = ilRet
                    gObtainRcfRifRdf = False
                    ilRet = btrClose(hlRcf)
                    btrDestroy hlRcf
                    Exit Function
                End If
                ilUpperBound = UBound(tgMRcf)
                ilExtLen = Len(tgMRcf(ilUpperBound))  'Extract operation record size
                'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlRcf, tgMRcf(ilUpperBound), ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    ilUpperBound = ilUpperBound + 1
                    'ReDim Preserve tgMRcf(1 To ilUpperBound) As RCF
                    ReDim Preserve tgMRcf(ilRifLowLimit To ilUpperBound) As RCF
                    ilRet = btrExtGetNext(hlRcf, tgMRcf(ilUpperBound), ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlRcf, tgMRcf(ilUpperBound), ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        ilRet = btrClose(hlRcf)
        btrDestroy hlRcf
    End If
    slStamp = gFileDateTime(sgDBPath & "Rif.Btr")

    'On Error GoTo gObtainRcfRdfErr2
    'ilRet = 0
    'llRifLowLimit = LBound(tgMRif)
    'If ilRet <> 0 Then
    '    sgMRifStamp = ""
    '    llRifLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tgMRif).Ptr <> 0 Then
        ilRifLowLimit = LBound(tgMRif)
    Else
        sgMRifStamp = ""
        ilRifLowLimit = 0
    End If

    If (sgMRifStamp = "") Or (slStamp <> sgMRifStamp) Then
        'ReDim tgMRif(1 To 1) As RIF
        ReDim tgMRif(llRifLowLimit To llRifLowLimit) As RIF
        hlRif = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlRif, "", sgDBPath & "Rif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            gObtainRcfRifRdf = False
            ilRet = btrClose(hlRif)
            btrDestroy hlRif
            Exit Function
        End If
        'ilRecLen = Len(tgMRif(1)) 'btrRecordLength(hlAdf)  'Get and save record length
        ilRecLen = Len(tgMRif(llRifLowLimit)) 'btrRecordLength(hlAdf)  'Get and save record length
        sgMRifStamp = slStamp
        llUpperBound = UBound(tgMRif)
        ilExtLen = Len(tgMRif(llUpperBound))  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        btrExtClear hlRif   'Clear any previous extend operation
        ilRet = btrGetFirst(hlRif, tgMRif(llUpperBound), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                gObtainRcfRifRdf = False
                ilRet = btrClose(hlRif)
                btrDestroy hlRif
                Exit Function
            End If
            Call btrExtSetBounds(hlRif, llNoRec, -1, "UC", "RIF", "") 'Set extract limits (all records)
            ilOffSet = 0
            ilRet = btrExtAddField(hlRif, ilOffSet, ilRecLen)  'Extract iCode field
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                gObtainRcfRifRdf = False
                ilRet = btrClose(hlRif)
                btrDestroy hlRif
                Exit Function
            End If
            'ilRet = btrExtGetNextExt(hlAdf)    'Extract record
            llUpperBound = UBound(tgMRif)
            ilRet = btrExtGetNext(hlRif, tgMRif(llUpperBound), ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                    igBtrError = ilRet
                    gObtainRcfRifRdf = False
                    ilRet = btrClose(hlRif)
                    btrDestroy hlRif
                    Exit Function
                End If
                llUpperBound = UBound(tgMRif)
                ilExtLen = Len(tgMRif(llUpperBound))  'Extract operation record size
                'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlRif, tgMRif(llUpperBound), ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    llUpperBound = llUpperBound + 1
                    'ReDim Preserve tgMRif(1 To llUpperBound) As RIF
                    ReDim Preserve tgMRif(llRifLowLimit To llUpperBound) As RIF
                    ilRet = btrExtGetNext(hlRif, tgMRif(llUpperBound), ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlRif, tgMRif(llUpperBound), ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        ilRet = btrClose(hlRif)
        btrDestroy hlRif
    End If
    slStamp = gFileDateTime(sgDBPath & "Rdf.Btr")

    'On Error GoTo gObtainRcfRdfErr2
    'ilRet = 0
    'ilRdfLowLimit = LBound(tgMRdf)
    'If ilRet <> 0 Then
    '    sgMRdfStamp = ""
    '    ilRdfLowLimit = 1
    'End If
    'On Error GoTo 0
    If PeekArray(tgMRdf).Ptr <> 0 Then
        ilRdfLowLimit = LBound(tgMRdf)
    Else
        sgMRdfStamp = ""
        ilRdfLowLimit = 0
    End If

    'If (sgMRdfStamp = "") Or (slStamp <> sgMRdfStamp) Then
    If (sgMRdfStamp = "") Then
        'ReDim tgMRdf(1 To 1) As RDF
        ReDim tgMRdf(ilRdfLowLimit To ilRdfLowLimit) As RDF
        hlRdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            gObtainRcfRifRdf = False
            ilRet = btrClose(hlRdf)
            btrDestroy hlRdf
            Exit Function
        End If
        'ilRecLen = Len(tgMRdf(1)) 'btrRecordLength(hlAdf)  'Get and save record length
        ilRecLen = Len(tgMRdf(ilRdfLowLimit)) 'btrRecordLength(hlAdf)  'Get and save record length
        sgMRdfStamp = slStamp
        ilUpperBound = UBound(tgMRdf)
        ilExtLen = Len(tgMRdf(ilUpperBound))  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        btrExtClear hlRdf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlRdf, tgMRdf(ilUpperBound), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                gObtainRcfRifRdf = False
                ilRet = btrClose(hlRdf)
                btrDestroy hlRdf
                Exit Function
            End If
            Call btrExtSetBounds(hlRdf, llNoRec, -1, "UC", "RDF", "") 'Set extract limits (all records)
            ilOffSet = 0
            ilRet = btrExtAddField(hlRdf, ilOffSet, ilRecLen)  'Extract iCode field
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                gObtainRcfRifRdf = False
                ilRet = btrClose(hlRdf)
                btrDestroy hlRdf
                Exit Function
            End If
            'ilRet = btrExtGetNextExt(hlAdf)    'Extract record
            ilUpperBound = UBound(tgMRdf)
            ilRet = btrExtGetNext(hlRdf, tgMRdf(ilUpperBound), ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                    igBtrError = ilRet
                    gObtainRcfRifRdf = False
                    ilRet = btrClose(hlRdf)
                    btrDestroy hlRdf
                    Exit Function
                End If
                ilUpperBound = UBound(tgMRdf)
                ilExtLen = Len(tgMRdf(ilUpperBound))  'Extract operation record size
                'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlRdf, tgMRdf(ilUpperBound), ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    ilUpperBound = ilUpperBound + 1
                    'ReDim Preserve tgMRdf(1 To ilUpperBound) As RDF
                    ReDim Preserve tgMRdf(ilRdfLowLimit To ilUpperBound) As RDF
                    ilRet = btrExtGetNext(hlRdf, tgMRdf(ilUpperBound), ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlRdf, tgMRdf(ilUpperBound), ilExtLen, llRecPos)
                    Loop
                Loop
            End If
        End If
        ilRet = btrClose(hlRdf)
        btrDestroy hlRdf
        'Sort by code so that binary search can be used
        'If UBound(tgMRdf) - 1 > 1 Then
        If UBound(tgMRdf) > 0 Then
            'ArraySortTyp fnAV(tgMRdf(), 1), UBound(tgMRdf) - 1, 0, LenB(tgMRdf(1)), 0, -1, 0
            ArraySortTyp fnAV(tgMRdf(), ilRdfLowLimit), UBound(tgMRdf) - ilRdfLowLimit, 0, LenB(tgMRdf(ilRdfLowLimit)), 0, -1, 0
        End If
    End If
    gObtainRcfRifRdf = True
    Exit Function
gObtainRcfRdfErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainRdf                      *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate Dayparts               *
'*                                                     *
'*******************************************************
Function gObtainRdf(slMRdfStamp As String, tlMRdf() As RDF) As Integer
'
'   ilRet = gObtainRdf (slMRdfStamp, tlMRdf())
'   Where:
'       tlMRdf() (I)- RDF record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlRdf As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim ilLowLimit As Integer

    slStamp = gFileDateTime(sgDBPath & "Rdf.Btr")

    'On Error GoTo gObtainRdfErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlMRdf)
    'If ilRet <> 0 Then
    '    slMRdfStamp = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlMRdf).Ptr <> 0 Then
        ilLowLimit = LBound(tlMRdf)
    Else
        slMRdfStamp = ""
        ilLowLimit = 0
    End If

    If slMRdfStamp <> "" Then
        If StrComp(slStamp, slMRdfStamp, 1) = 0 Then
            'If UBound(tlMRdf) > 1 Then
                gObtainRdf = True
                Exit Function
            'End If
        End If
    End If
    'ReDim tlMRdf(1 To 1) As RDF
    ReDim tlMRdf(ilLowLimit To ilLowLimit) As RDF
    hlRdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainRdf = False
        ilRet = btrClose(hlRdf)
        btrDestroy hlRdf
        Exit Function
    End If
    'ilRecLen = Len(tlMRdf(1)) 'btrRecordLength(hlRdf)  'Get and save record length
    ilRecLen = Len(tlMRdf(ilLowLimit)) 'btrRecordLength(hlRdf)  'Get and save record length
    slMRdfStamp = slStamp
    ilUpperBound = UBound(tlMRdf)
    ilExtLen = Len(tlMRdf(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlRdf) 'Obtain number of records
    btrExtClear hlRdf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlRdf, tlMRdf(ilUpperBound), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlRdf)
        btrDestroy hlRdf
        gObtainRdf = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainRdf = False
            ilRet = btrClose(hlRdf)
            btrDestroy hlRdf
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlRdf, llNoRec, -1, "UC", "RDF", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlRdf, ilOffSet, ilExtLen)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainRdf = False
        ilRet = btrClose(hlRdf)
        btrDestroy hlRdf
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlRdf)    'Extract record
    ilUpperBound = UBound(tlMRdf)
    ilRet = btrExtGetNext(hlRdf, tlMRdf(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainRdf = False
            ilRet = btrClose(hlRdf)
            btrDestroy hlRdf
            Exit Function
        End If
        ilUpperBound = UBound(tlMRdf)
        ilExtLen = Len(tlMRdf(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlRdf, tgCompMnf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlRdf, tlMRdf(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilUpperBound = ilUpperBound + 1
            'ReDim Preserve tlMRdf(1 To ilUpperBound) As RDF
            ReDim Preserve tlMRdf(ilLowLimit To ilUpperBound) As RDF
            ilRet = btrExtGetNext(hlRdf, tlMRdf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlRdf, tlMRdf(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlRdf)
    btrDestroy hlRdf
    'If UBound(tlMRdf) - 1 > 1 Then
    If UBound(tlMRdf) - ilLowLimit > ilLowLimit Then
        'ArraySortTyp fnAV(tlMRdf(), 1), UBound(tlMRdf) - 1, 0, LenB(tlMRdf(1)), 0, -1, 0
        ArraySortTyp fnAV(tlMRdf(), ilLowLimit), UBound(tlMRdf) - ilLowLimit, 0, LenB(tlMRdf(ilLowLimit)), 0, -1, 0
    End If
    gObtainRdf = True
    Exit Function
gObtainRdfErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainSalesperson              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgMSlf                 *
'*                                                     *
'*******************************************************
Function gObtainSalesperson() As Integer
'
'   ilRet = gObtainSalesperson ()
'   Where:
'       tgMSlf() (I)- SLF record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Slf date/time stamp
    Dim hlSlf As Integer        'Slf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlSlf As SLF
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilUpperBound As Integer
    Dim ilLowLimit As Integer

    slStamp = gFileDateTime(sgDBPath & "Slf.Btr")

    'On Error GoTo gObtainSalespersonErr2
    'ilRet = 0
    'ilLowLimit = LBound(tgMSlf)
    'If ilRet <> 0 Then
    '    sgMSlfStamp = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tgMSlf).Ptr <> 0 Then
        ilLowLimit = LBound(tgMSlf)
    Else
        sgMSlfStamp = ""
        ilLowLimit = 0
    End If

    If sgMSlfStamp <> "" Then
        If StrComp(slStamp, sgMSlfStamp, 1) = 0 Then
            'If UBound(tgMSlf) > 1 Then
                gObtainSalesperson = True
                Exit Function
            'End If
        End If
    End If
    'ReDim tgMSlf(1 To 1) As SLF
    ReDim tgMSlf(ilLowLimit To ilLowLimit) As SLF
    hlSlf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainSalesperson = False
        ilRet = btrClose(hlSlf)
        btrDestroy hlSlf
        Exit Function
    End If
    ilRecLen = Len(tlSlf) 'btrRecordLength(hlSlf)  'Get and save record length
    sgMSlfStamp = slStamp
    ilUpperBound = UBound(tgMSlf)
    ilExtLen = Len(tgMSlf(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSlf) 'Obtain number of records
    btrExtClear hlSlf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlSlf, tlSlf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlSlf)
        btrDestroy hlSlf
        gObtainSalesperson = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainSalesperson = False
            ilRet = btrClose(hlSlf)
            btrDestroy hlSlf
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlSlf, llNoRec, -1, "UC", "SLF", "") 'Set extract limits (all records)
    ilRet = btrExtAddField(hlSlf, 0, ilExtLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainSalesperson = False
        ilRet = btrClose(hlSlf)
        btrDestroy hlSlf
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlSlf)    'Extract record
    ilUpperBound = UBound(tgMSlf)
    ilRet = btrExtGetNext(hlSlf, tgMSlf(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainSalesperson = False
            ilRet = btrClose(hlSlf)
            btrDestroy hlSlf
            Exit Function
        End If
        ilUpperBound = UBound(tgMSlf)
        ilExtLen = Len(tgMSlf(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlSlf, tgMSlf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlSlf, tgMSlf(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilUpperBound = ilUpperBound + 1
            'ReDim Preserve tgMSlf(1 To ilUpperBound) As SLF
            ReDim Preserve tgMSlf(ilLowLimit To ilUpperBound) As SLF
            ilRet = btrExtGetNext(hlSlf, tgMSlf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSlf, tgMSlf(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlSlf)
    btrDestroy hlSlf
    'Sort by code so that binary search can be used
    'If UBound(tgMSlf) - 1 > 1 Then
    If UBound(tgMSlf) - ilLowLimit > ilLowLimit Then
        'ArraySortTyp fnAV(tgMSlf(), 1), UBound(tgMSlf) - 1, 0, LenB(tgMSlf(1)), 0, -1, 0
        ArraySortTyp fnAV(tgMSlf(), ilLowLimit), UBound(tgMSlf) - ilLowLimit, 0, LenB(tgMSlf(ilLowLimit)), 0, -1, 0
    End If
    gObtainSalesperson = True
    Exit Function
gObtainSalespersonErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainUrf                      *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgPopUrf               *
'*                                                     *
'*******************************************************
Function gObtainUrf() As Integer
'
'   ilRet = gObtainUrf ()
'   Where:
'       tgPopUrf() (I)- URF record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlUrf As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlUrf As URF
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim ilTest As Integer
    Dim ilDuplicates As Integer
    Dim ilDelete As Integer
    Dim tlUrfSrchKey As INTKEY0
    Dim ilLowLimit As Integer

    'Ignore time of URF since it is always changing as each user signon
    'Assume if sgUrfStamp defined, then it is Ok
    slStamp = gFileDateTime(sgDBPath & "Urf.Btr")

    'On Error GoTo gObtainUrfErr2
    'ilRet = 0
    'ilLowLimit = LBound(tgPopUrf)
    'If ilRet <> 0 Then
    '    sgPopUrfTag = ""
    '    ilLowLimit = 0  'V7.0 default was 1
    'End If
    'On Error GoTo 0
    If PeekArray(tgPopUrf).Ptr <> 0 Then
        ilLowLimit = LBound(tgPopUrf)
    Else
        sgPopUrfTag = ""
        ilLowLimit = 0
    End If

    If (sgPopUrfTag <> "") And (sgPopUrfTag <> "~") Then
        'Ignore time of URF since it is always changing as each user signon
        'Assume if sgUrfStamp defined, then it is Ok
        'If StrComp(slStamp, sgPopUrfTag, 1) = 0 Then
            'If UBound(tgPopUrf) > 1 Then
                gObtainUrf = True
                Exit Function
            'End If
        'End If
    End If
    Do
        ilDuplicates = False
        'ReDim tgPopUrf(1 To 1) As URF
        ReDim tgPopUrf(ilLowLimit To ilLowLimit) As URF
        hlUrf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            gObtainUrf = False
            ilRet = btrClose(hlUrf)
            btrDestroy hlUrf
            Exit Function
        End If
        ilRecLen = Len(tlUrf) 'btrRecordLength(hlUrf)  'Get and save record length
        sgPopUrfTag = slStamp
        ilUpperBound = UBound(tgPopUrf)
        ilExtLen = Len(tgPopUrf(ilUpperBound))  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlUrf) 'Obtain number of records
        btrExtClear hlUrf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlUrf, tlUrf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlUrf)
            btrDestroy hlUrf
            gObtainUrf = True
            Exit Function
        Else
            If ilRet <> BTRV_ERR_NONE Then
                gObtainUrf = False
                ilRet = btrClose(hlUrf)
                btrDestroy hlUrf
                Exit Function
            End If
        End If
        Call btrExtSetBounds(hlUrf, llNoRec, -1, "UC", "URF", "") 'Set extract limits (all records)
        ilOffSet = 0
        ilRet = btrExtAddField(hlUrf, ilOffSet, ilRecLen)  'Extract iCode field
        If ilRet <> BTRV_ERR_NONE Then
            gObtainUrf = False
            ilRet = btrClose(hlUrf)
            btrDestroy hlUrf
            Exit Function
        End If
        'ilRet = btrExtGetNextExt(hlUrf)    'Extract record
        ilUpperBound = UBound(tgPopUrf)
        ilRet = btrExtGetNext(hlUrf, tgPopUrf(ilUpperBound), ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                gObtainUrf = False
                ilRet = btrClose(hlUrf)
                btrDestroy hlUrf
                Exit Function
            End If
            ilUpperBound = UBound(tgPopUrf)
            ilExtLen = Len(tgPopUrf(ilUpperBound))  'Extract operation record size
            'ilRet = btrExtGetFirst(hlUrf, tgPopUrf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlUrf, tgPopUrf(ilUpperBound), ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilUpperBound = ilUpperBound + 1
                'ReDim Preserve tgPopUrf(1 To ilUpperBound) As URF
                ReDim Preserve tgPopUrf(ilLowLimit To ilUpperBound) As URF
                ilRet = btrExtGetNext(hlUrf, tgPopUrf(ilUpperBound), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlUrf, tgPopUrf(ilUpperBound), ilExtLen, llRecPos)
                Loop
            Loop
        End If
        'For ilLoop = 1 To UBound(tgPopUrf) - 1 Step 1
        For illoop = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
            gUrfDecrypt tgPopUrf(illoop)
        Next illoop
        ''Correct Duplicate names
        'Delete Duplicate Names
        'ilExtLen = Len(tgPopUrf(1))
        ilExtLen = Len(tgPopUrf(LBound(tgPopUrf)))
        'For ilLoop = 1 To UBound(tgPopUrf) - 1 Step 1
        For illoop = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
            'If tgPopUrf(ilLoop).sDelete <> "Y" Then
            If (tgPopUrf(illoop).iCode > 0) Then
                'For ilTest = 1 To UBound(tgPopUrf) - 1 Step 1
                For ilTest = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
                    'If (tgPopUrf(ilTest).iVefCode = tgPopUrf(ilLoop).iVefCode) And (ilLoop <> ilTest) Then
                    If (illoop <> ilTest) And (tgPopUrf(ilTest).iCode > 0) Then
                        If StrComp(tgPopUrf(ilTest).sName, tgPopUrf(illoop).sName, vbTextCompare) = 0 Then
                            ilDuplicates = True
                            ''Change Name so that it is unique
                            'tgPopUrf(ilTest).sName = Trim$(tgPopUrf(ilTest).sName) & "-" & Trim$(Str$(tgPopUrf(ilTest).iCode))
                            'gUrfEncrypt tgPopUrf(ilTest)
                            If (tgPopUrf(illoop).sDelete = "Y") And (tgPopUrf(ilTest).sDelete = "Y") Then
                                ilDelete = ilTest
                            ElseIf (tgPopUrf(illoop).sDelete <> "Y") And (tgPopUrf(ilTest).sDelete = "Y") Then
                                ilDelete = ilTest
                            ElseIf (tgPopUrf(illoop).sDelete = "Y") And (tgPopUrf(ilTest).sDelete <> "Y") Then
                                ilDelete = illoop
                            ElseIf (tgPopUrf(illoop).iVefCode = 0) And (tgPopUrf(ilTest).iVefCode = 0) Then
                                ilDelete = ilTest
                            ElseIf (tgPopUrf(illoop).iVefCode = 0) And (tgPopUrf(ilTest).iVefCode <> 0) Then
                                ilDelete = ilTest
                            ElseIf (tgPopUrf(illoop).iVefCode <> 0) And (tgPopUrf(ilTest).iVefCode = 0) Then
                                ilDelete = illoop
                            Else
                                ilDelete = ilTest
                            End If
                            tlUrfSrchKey.iCode = tgPopUrf(ilDelete).iCode
                            ilRet = btrGetEqual(hlUrf, tlUrf, ilExtLen, tlUrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                ilRet = btrDelete(hlUrf)
                                tgPopUrf(ilDelete).iCode = -tgPopUrf(ilDelete).iCode
                                'If ilRet = BTRV_ERR_NONE Then
                                '    ilRet = btrInsert(hlUrf, tgPopUrf(ilTest), ilExtLen, INDEXKEY0)
                                'End If
                            End If
                            'gUrfDecrypt tgPopUrf(ilTest)
                            'Exit For
                            If ilDelete = illoop Then
                                Exit For
                            End If
                        End If
                    End If
                Next ilTest
            End If
        Next illoop
        ilRet = btrClose(hlUrf)
        btrDestroy hlUrf
    Loop While ilDuplicates
    gObtainUrf = True
    Exit Function
gObtainUrfErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainVef                      *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgMVef                 *
'*                                                     *
'*******************************************************
Function gObtainVef() As Integer
'
'   ilRet = gObtainVef ()
'   Where:
'       tgMVef() (I)- VEF record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlVef As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim slGetStamp As String
    Dim ilLowLimit As Integer
    Dim ilVef As Integer

    slStamp = gFileDateTime(sgDBPath & "Vef.Btr")

    '11/26/17: Check Changed date/time
    If Not gFileChgd("vef.btr") Then
        gObtainVef = True
        Exit Function
    End If
    
    'On Error GoTo gObtainVefErr2
    'ilRet = 0
    'ilLowLimit = LBound(tgMVef)
    'If ilRet <> 0 Then
    '    sgMVefStamp = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tgMVef).Ptr <> 0 Then
        ilLowLimit = LBound(tgMVef)
    Else
        sgMVefStamp = ""
        ilLowLimit = 0
    End If

    hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = ilRet
        gObtainVef = False
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        Exit Function
    End If
    '6/9/18: replaced with gFileChgd
    'If (sgMVefStamp <> "") And (sgMVefStamp <> "~") Then
    '    'If defined, assume Ok as items added with gBinarySearchAdf
    '    'If StrComp(slStamp, sgMVefStamp, 1) = 0 Then
    '        'If UBound(tgMVef) > 1 Then
    '        If UBound(tgMVef) - ilLowLimit = btrRecords(hlVef) Then
    '            btrDestroy hlVef
    '            gObtainVef = True
    '            Exit Function
    '        End If
    '    'End If
    '    slGetStamp = ""
    'Else
        slGetStamp = "" 'gGetCSIStamp("VEF")
    'End If
    If slGetStamp <> "" Then
'        sgMVefStamp = slGetStamp
'        ilRet = csiGetAlloc("VEF", ilStartIndex, ilEndIndex)
'        ReDim tgMVef(ilStartIndex To ilEndIndex) As VEF
'        For ilLoop = LBound(tgMVef) To UBound(tgMVef) Step 1
'            ilRet = csiGetRec("VEF", ilLoop, VarPtr(tgMVef(ilLoop)), LenB(tgMVef(ilLoop)))
'        Next ilLoop
        gObtainVef = True
    Else
        'ReDim tgMVef(1 To 1) As VEF
        ReDim tgMVef(ilLowLimit To ilLowLimit) As VEF
        'ilRecLen = Len(tgMVef(1)) 'btrRecordLength(hlVef)  'Get and save record length
        ilRecLen = Len(tgMVef(ilLowLimit)) 'btrRecordLength(hlVef)  'Get and save record length
        sgMVefStamp = slStamp
        ilUpperBound = UBound(tgMVef)
        ilExtLen = Len(tgMVef(ilUpperBound))  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlVef) 'Obtain number of records
        btrExtClear hlVef   'Clear any previous extend operation
        ilRet = btrGetFirst(hlVef, tgMVef(ilUpperBound), ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlVef)
            btrDestroy hlVef
            gObtainVef = True
            Exit Function
        Else
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = ilRet
                gObtainVef = False
                ilRet = btrClose(hlVef)
                btrDestroy hlVef
                Exit Function
            End If
        End If
        Call btrExtSetBounds(hlVef, llNoRec, -1, "UC", "VEF", "") 'Set extract limits (all records)
        ilOffSet = 0
        ilRet = btrExtAddField(hlVef, ilOffSet, ilExtLen)  'Extract First Name field
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            gObtainVef = False
            ilRet = btrClose(hlVef)
            btrDestroy hlVef
            Exit Function
        End If
        'ilRet = btrExtGetNextExt(hlVef)    'Extract record
        ilUpperBound = UBound(tgMVef)
        ilRet = btrExtGetNext(hlVef, tgMVef(ilUpperBound), ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                igBtrError = ilRet
                gObtainVef = False
                ilRet = btrClose(hlVef)
                btrDestroy hlVef
                Exit Function
            End If
            ilUpperBound = UBound(tgMVef)
            ilExtLen = Len(tgMVef(ilUpperBound))  'Extract operation record size
            'ilRet = btrExtGetFirst(hlVef, tgCompMnf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlVef, tgMVef(ilUpperBound), ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Or (tgUrf(0).iMnfHubCode <= 0) Or ((tgUrf(0).iMnfHubCode = tgMVef(ilUpperBound).iMnfHubCode) And ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB)) Or ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) <> USINGHUB) Then
                    ilUpperBound = ilUpperBound + 1
                    'ReDim Preserve tgMVef(1 To ilUpperBound) As VEF
                    ReDim Preserve tgMVef(ilLowLimit To ilUpperBound) As VEF
                End If
                ilRet = btrExtGetNext(hlVef, tgMVef(ilUpperBound), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlVef, tgMVef(ilUpperBound), ilExtLen, llRecPos)
                Loop
            Loop
        End If
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        'Sort by vefCode
        'If UBound(tgMVef) > 2 Then
        If UBound(tgMVef) - ilLowLimit > ilLowLimit Then
            'ArraySortTyp fnAV(tgMVef(), 1), UBound(tgMVef) - 1, 0, LenB(tgMVef(1)), 0, -1, 0
            ArraySortTyp fnAV(tgMVef(), ilLowLimit), UBound(tgMVef) - ilLowLimit, 0, LenB(tgMVef(ilLowLimit)), 0, -1, 0
        End If
        ReDim tgVefName(0 To UBound(tgMVef)) As VEFNAME
        For ilVef = 0 To UBound(tgMVef) Step 1
            tgVefName(ilVef).sName = UCase(tgMVef(ilVef).sName)
            tgVefName(ilVef).iCode = tgMVef(ilVef).iCode
        Next ilVef
        ArraySortTyp fnAV(tgVefName(), 0), UBound(tgVefName), 0, LenB(tgVefName(0)), 0, LenB(tgVefName(0).sName), 0 'VBC NR
        
'        ilRet = csiSetStamp("VEF", sgMVefStamp)
'        ilRet = csiSetAlloc("VEF", LBound(tgMVef), UBound(tgMVef))
'        For ilLoop = LBound(tgMVef) To UBound(tgMVef) Step 1
'            ilRet = csiSetRec("VEF", ilLoop, VarPtr(tgMVef(ilLoop)), LenB(tgMVef(ilLoop)))
'        Next ilLoop
        gObtainVef = True
    End If
    gBuildDormantVef
    Exit Function
gObtainVefErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gOkAddStrToListBox              *
'*                                                     *
'*             Created:8/05/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Check that string can to       *
'*                      into list box                  *
'*                                                     *
'*******************************************************
Function gOkAddStrToListBox(slInStr As String, llLen As Long, ilShowMsg As Integer) As Integer
'
'   ilRet = gOkAddStrToListBox(slStr, llLen)
'   Where
'       slStr(I)- String to be added
'       llLen(I)- Running Length of all strings previuosly added
'       ilRet(O)- True=Ok to add string; False- No room
'
    Dim ilRet As Integer

    If fgWinVersion <= 3.9 Then
        llLen = llLen + Len(slInStr)
        If llLen > 60000 Then
            If ilShowMsg Then
                ilRet = MsgBox("Out of Room in List Box, Tell Counterpoint", vbOKOnly + vbExclamation, "List Box Room")
            End If
            gOkAddStrToListBox = False
            Exit Function
        End If
    End If
    gOkAddStrToListBox = True
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAASBoxViaCntr               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     advertiser or agency or         *
'*                     salespersons names and create   *
'*                     array of contracts associated    *
'*                     with names                      *
'*                                                     *
'*******************************************************
Function gPopAASBoxViaCntr(frm As Form, ilAAS As Integer, slStatus As String, slCntrType As String, ilCurrent As Integer, ilHOType As Integer, ilDormant As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, tlChfAdvtExt() As CHFADVTEXT) As Integer 'VBC NR
'
'   ilRet = gPopAASBoxViaCntr (MainForm, ilAAS, slStatus, slCntrType, ilCurrent, ilHOType, ilDormant, lbcLocal, tlSortCode(), slSortCodeTag, tlChfAdvtExt())
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAAS(I)- 0=Populate list boxes with Advertiser names
'                 1=Populate list boxes with Agency names
'                 2=Populate list boxes with Salesperson Names
'       slStatus (I)- chfStatus value or blank
'                         W=Working; D=Rejected; C=Completed; I=Unapproved; H=Hold; O=Order
'                         Multiple status can be specified (WDI)
'       slCntrType (I)- chfType value or blank
'                       C=Standard; V=Reservation; T=Remnant; R=DR; Q=PI; S=PSA; M=Promo
'       ilCurrent (I)- 0=Current (Active) (chfDelete <> y); 1=Past and Current (chfDelete <> y); 2=Current(Active) plus all cancel before start (chfDelete <> y); 3=All plus history (any value for chfDelete)
'       ilHOType (I)-  1=H or O only; 2=H or O or G or N (if G or N exists show it over H or O);
'                      3=H or O or G or N or W or C or I (if G or N or W or C or I exists show it over H or O)
'                        Note: G or N can't exist at the same time as W or C or I for an order
'                              G or N or W or C or I CntrRev > 0
'       ilDormant(I)- True=Include dormant advertisers; False=Exclude dormant advertisers
'       lbcLocal (O)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       tlChfAdvtExt(O)- Array of contracts of the structure CHFADVTEXT which selection criteria
'                        used to populate the list box
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'CHF date/time stamp 'VBC NR
    Dim hlChf As Integer        'CHF handle 'VBC NR
    Dim ilRecLen As Integer     'Record length 'VBC NR
    Dim llNoRec As Long         'Number of records in Sof 'VBC NR
    Dim tlChf As CHF 'VBC NR
    Dim slName As String 'VBC NR
    Dim ilExtLen As Integer 'VBC NR
    Dim llRecPos As Long        'Record location 'VBC NR
    Dim ilRet As Integer 'VBC NR
    Dim ilFound As Integer 'VBC NR
    Dim illoop As Integer 'VBC NR
    Dim slNameCode As String 'VBC NR
    Dim llTodayDate As Long 'VBC NR
    Dim hlVef As Integer        'Vef handle 'VBC NR
    Dim tlVef As VEF 'VBC NR
    Dim ilVefRecLen As Integer     'Record length 'VBC NR
    Dim hlVsf As Integer        'Vsf handle 'VBC NR
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length 'VBC NR
    Dim hlFile As Integer        'Adf handle 'VBC NR
    Dim tlAdf As ADF 'VBC NR
    Dim ilAdfRecLen As Integer     'Record length 'VBC NR
    Dim tlAdfSrchKey As INTKEY0 'VBC NR
    Dim tlagf As AGF 'VBC NR
    Dim ilAgfRecLen As Integer     'Record length 'VBC NR
    Dim tlAgfSrchKey As INTKEY0 'VBC NR
    Dim tlSlf As SLF 'VBC NR
    Dim ilSlfRecLen As Integer     'Record length 'VBC NR
    Dim tlSlfSrchKey As INTKEY0 'VBC NR
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record 'VBC NR
    Dim ilOffSet As Integer 'VBC NR
    Dim llLen As Long 'VBC NR
    Dim ilOper As Integer 'VBC NR
    Dim ilUpper As Integer 'VBC NR
    Dim slStr As String 'VBC NR
    Dim ilCntUpper As Integer 'VBC NR
    Dim ilSlf As Integer 'VBC NR
    Dim ilSlfCode As Integer 'VBC NR
    Dim ilTestCntrNo As Integer 'VBC NR
    Dim slCntrStatus As String 'VBC NR
    Dim slHOStatus As String 'VBC NR
    Dim ilSortCode As Integer 'VBC NR
    Dim ilPop As Integer 'VBC NR
    Dim slAdvtName As String 'VBC NR
    Dim ilSortLowLimit As Integer
    Dim ilLowLimit As Integer

    If slStatus = "" Then 'VBC NR
        slCntrStatus = "WCIDHO" 'VBC NR
    Else 'VBC NR
        slCntrStatus = slStatus 'VBC NR
    End If 'VBC NR
    slHOStatus = "" 'VBC NR
    If ilHOType = 1 Then 'VBC NR
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then 'VBC NR
            slHOStatus = slHOStatus & "H" 'VBC NR
        End If 'VBC NR
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then 'VBC NR
            slHOStatus = slHOStatus & "O" 'VBC NR
        End If 'VBC NR
    ElseIf ilHOType = 2 Then 'VBC NR
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then 'VBC NR
            slHOStatus = slHOStatus & "GH" 'VBC NR
        End If 'VBC NR
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then 'VBC NR
            slHOStatus = slHOStatus & "NO" 'VBC NR
        End If 'VBC NR
    ElseIf ilHOType = 3 Then 'VBC NR
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then 'VBC NR
            slHOStatus = slHOStatus & "GH" 'VBC NR
        End If 'VBC NR
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then 'VBC NR
            slHOStatus = slHOStatus & "NO" 'VBC NR
        End If 'VBC NR
        If (InStr(1, slCntrStatus, "H", 1) <> 0) Or (InStr(1, slCntrStatus, "O", 1) <> 0) Then 'VBC NR
            slHOStatus = slHOStatus & "WCI" 'VBC NR
        End If 'VBC NR
    End If 'VBC NR
    ilPop = True 'VBC NR
    llLen = 0 'VBC NR
    slStamp = gFileDateTime(sgDBPath & "Chf.Btr") & Trim$(str$(ilAAS)) & Trim$(slCntrStatus) & Trim$(slCntrType) & Trim$(str$(ilCurrent)) & Trim$(str$(ilHOType)) & Trim$(str$(ilDormant)) 'VBC NR

    'On Error GoTo gPopAASBoxViaCntrErr2 'VBC NR
    'ilRet = 0 'VBC NR
    'ilSortLowLimit = LBound(tlSortCode) 'VBC NR
    'If ilRet <> 0 Then 'VBC NR
    '    slSortCodeTag = "" 'VBC NR
    '    ilSortLowLimit = 0
    'End If 'VBC NR
    'ilRet = 0
    'ilLowLimit = LBound(tlChfAdvtExt) 'VBC NR
    'If ilRet <> 0 Then 'VBC NR
    '    ilLowLimit = 0
    'End If 'VBC NR
    'On Error GoTo 0 'VBC NR
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilSortLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilSortLowLimit = 0
    End If
    If PeekArray(tlChfAdvtExt).Ptr <> 0 Then
        ilLowLimit = LBound(tlChfAdvtExt)
    Else
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then 'VBC NR
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then 'VBC NR
            If lbcLocal.ListCount > 0 Then 'VBC NR
                gPopAASBoxViaCntr = CP_MSG_NOPOPREQ 'VBC NR
                Exit Function 'VBC NR
            End If 'VBC NR
            ilPop = False 'VBC NR
        End If 'VBC NR
    End If 'VBC NR
    gPopAASBoxViaCntr = CP_MSG_POPREQ 'VBC NR
    lbcLocal.Clear 'VBC NR
    slSortCodeTag = slStamp 'VBC NR
    If ilPop Then 'VBC NR
        llTodayDate = gDateValue(gNow()) 'VBC NR
        'gObtainVehComboList
        hlChf = CBtrvTable(ONEHANDLE) 'CBtrvTable() 'VBC NR
        ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrOpen):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilRecLen = Len(tlChf) 'btrRecordLength(hlChf)  'Get and save record length 'VBC NR
        hlVsf = CBtrvTable(ONEHANDLE) 'CBtrvTable() 'VBC NR
        ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrOpen):" & "Vsf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length 'VBC NR
        hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable() 'VBC NR
        ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrOpen):" & "Vef.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilVefRecLen = Len(tlVef) 'btrRecordLength(hlSlf)  'Get and save record length 'VBC NR
        If ilAAS = 1 Then 'VBC NR
            hlFile = CBtrvTable(ONEHANDLE) 'CBtrvTable() 'VBC NR
            ilRet = btrOpen(hlFile, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'VBC NR
            On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
            gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrOpen):" & "Agf.Btr", frm 'VBC NR
            On Error GoTo 0 'VBC NR
            ilAgfRecLen = Len(tlagf) 'btrRecordLength(hlSlf)  'Get and save record length 'VBC NR
            tlagf.iCode = 0 'VBC NR
        ElseIf ilAAS = 2 Then 'VBC NR
            hlFile = CBtrvTable(ONEHANDLE) 'CBtrvTable() 'VBC NR
            ilRet = btrOpen(hlFile, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'VBC NR
            On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
            gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrOpen):" & "Slf.Btr", frm 'VBC NR
            On Error GoTo 0 'VBC NR
            ilSlfRecLen = Len(tlSlf) 'btrRecordLength(hlSlf)  'Get and save record length 'VBC NR
            tlSlf.iCode = 0 'VBC NR
        Else 'VBC NR
            hlFile = CBtrvTable(ONEHANDLE) 'CBtrvTable() 'VBC NR
            ilRet = btrOpen(hlFile, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'VBC NR
            On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
            gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrOpen):" & "Adf.Btr", frm 'VBC NR
            On Error GoTo 0 'VBC NR
            ilAdfRecLen = Len(tlAdf) 'btrRecordLength(hlSlf)  'Get and save record length 'VBC NR
            tlAdf.iCode = 0 'VBC NR
        End If 'VBC NR
        ilSortCode = ilSortLowLimit 'VBC NR
        ReDim tlSortCode(ilSortLowLimit To ilSortLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found) 'VBC NR
        'ReDim tlChfAdvtExt(1 To 1) As CHFADVTEXT 'VBC NR
        ReDim tlChfAdvtExt(ilLowLimit To ilLowLimit) As CHFADVTEXT 'VBC NR
        ilCntUpper = ilLowLimit
        ilExtLen = Len(tlChfAdvtExt(ilCntUpper))  'Extract operation record size 'VBC NR
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlChf) 'Obtain number of records 'VBC NR
        btrExtClear hlChf   'Clear any previous extend operation 'VBC NR
        ilRet = btrGetFirst(hlChf, tlChf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation 'VBC NR
        If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then 'VBC NR
            ilRet = btrClose(hlChf) 'VBC NR
            On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
            gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrReset):" & "Chf.Btr", frm 'VBC NR
            On Error GoTo 0 'VBC NR
            btrDestroy hlChf 'VBC NR
            ilRet = btrClose(hlVsf) 'VBC NR
            On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
            gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrReset):" & "Vsf.Btr", frm 'VBC NR
            On Error GoTo 0 'VBC NR
            btrDestroy hlVsf 'VBC NR
            Exit Function 'VBC NR
        Else 'VBC NR
            On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
            gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrGetFirst):" & "Chf.Btr", frm 'VBC NR
            On Error GoTo 0 'VBC NR
        End If 'VBC NR
        'ReDim imAdvtCode(1 To 1) As Integer 'VBC NR
        ReDim imAdvtCode(0 To 0) As Integer 'VBC NR
        'ilUpper = 1 'VBC NR
        ilUpper = 0 'VBC NR
        Call btrExtSetBounds(hlChf, llNoRec, -1, "UC", "CHFADVTEXTPK", CHFADVTEXTPK) 'Set extract limits (all records) 'VBC NR
        ilSlfCode = tgUrf(0).iSlfCode 'VBC NR
        If (tgUrf(0).iGroupNo > 0) Then 'And (tgUrf(0).iSlfCode <= 0) Then 'VBC NR
            ilRet = gObtainUrf() 'VBC NR
            ilRet = gObtainSalesperson() 'VBC NR
        End If 'VBC NR
        tlCharTypeBuff.sType = "Y" 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfDelete") 'VBC NR
        If (slCntrType = "") Then 'VBC NR
            If ilCurrent <> 3 Then 'VBC NR
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1) 'VBC NR
            End If 'VBC NR
        Else 'VBC NR
            If ilCurrent <> 3 Then 'VBC NR
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1) 'VBC NR
            End If 'VBC NR
            'If slCntrStatus <> "" Then
            '    ilOper = BTRV_EXT_OR
            '    slStr = slCntrStatus
            '    Do While slStr <> ""
            '        If Len(slStr) = 1 Then
            '            If slCntrType <> "" Then
            '                ilOper = BTRV_EXT_AND
            '            Else
            '                ilOper = BTRV_EXT_LAST_TERM
            '            End If
            '        End If
            '        tlCharTypeBuff.sType = Left$(slStr, 1)
            '        ilOffset = gFieldOffset("Chf", "ChfStatus")
            '        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
            '        slStr = Mid$(slStr, 2)
            '    Loop
            'End If
            If slCntrType <> "" Then 'VBC NR
                ilOper = BTRV_EXT_OR 'VBC NR
                slStr = slCntrType 'VBC NR
                Do While slStr <> "" 'VBC NR
                    If Len(slStr) = 1 Then 'VBC NR
                        ilOper = BTRV_EXT_LAST_TERM 'VBC NR
                    End If 'VBC NR
                    tlCharTypeBuff.sType = Left$(slStr, 1) 'VBC NR
                    ilOffSet = gFieldOffset("Chf", "ChfType") 'VBC NR
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1) 'VBC NR
                    slStr = Mid$(slStr, 2) 'VBC NR
                Loop 'VBC NR
            End If 'VBC NR
        End If 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfCode") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract iCode field 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfCntrNo") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract Contract number 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfExtRevNo") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfCntRevNo") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfType") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfAdfCode") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfProduct") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 35) 'Extract Product 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfAgfCode") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfSlfCode1") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 20) 'Extract salesperson code 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfMnfDemo1") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract salesperson code 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfCxfInt") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfPropVer") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract end date 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfStatus") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfMnfPotnType") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract SellNet 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfStartDate") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfEndDate") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract end date 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfVefCode") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        ilOffSet = gFieldOffset("Chf", "ChfSifCode") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR

         '8-21-05 add pct of trade to array
        ilOffSet = gFieldOffset("Chf", "ChfPctTrade") 'VBC NR
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'pct trade 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        
        '7/12/10
        ilOffSet = gFieldOffset("Chf", "ChfCBSOrder")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopAASBoxViaCntrErr
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0

        '2/24/12
        ilOffSet = gFieldOffset("Chf", "ChfBillCycle")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopAASBoxViaCntrErr
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0

        '2/17/18
        ilOffSet = gFieldOffset("Chf", "ChfSource")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopAASBoxViaCntrErr
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
    
        'ilRet = btrExtGetNextExt(hlChf)    'Extract record
        ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos) 'VBC NR
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then 'VBC NR
            On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
            gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrExtGetNextExt):" & "Chf.Btr", frm 'VBC NR
            On Error GoTo 0 'VBC NR
            ilExtLen = Len(tlChfAdvtExt(ilCntUpper))  'Extract operation record size 'VBC NR
            'ilRet = btrExtGetFirst(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT 'VBC NR
                ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos) 'VBC NR
            Loop 'VBC NR
            Do While ilRet = BTRV_ERR_NONE 'VBC NR
                ilFound = True 'VBC NR
                ilTestCntrNo = False 'VBC NR
                'For Proposal CntRevNo = 0; For Orders CntRevNo >= 0 (for W, C, I CntRevNo > 0)
                If (tlChfAdvtExt(ilCntUpper).iCntRevNo = 0) And ((tlChfAdvtExt(ilCntUpper).sStatus <> "H") And (tlChfAdvtExt(ilCntUpper).sStatus <> "O") And (tlChfAdvtExt(ilCntUpper).sStatus <> "G") And (tlChfAdvtExt(ilCntUpper).sStatus <> "N")) Then 'Proposal 'VBC NR
                    If (InStr(1, slCntrStatus, tlChfAdvtExt(ilCntUpper).sStatus) = 0) Then 'VBC NR
                        ilFound = False 'VBC NR
                    End If 'VBC NR
                Else    'Order 'VBC NR
                    If (InStr(1, slHOStatus, tlChfAdvtExt(ilCntUpper).sStatus) <> 0) Then 'VBC NR
                        If (ilHOType = 2) Or (ilHOType = 3) Then 'VBC NR
                            ilTestCntrNo = True 'VBC NR
                        End If 'VBC NR
                    Else 'VBC NR
                        ilFound = False 'VBC NR
                    End If 'VBC NR
                End If 'VBC NR
                If ilFound Then 'VBC NR
                    '2/17/18: Allow Programmatic buys
                    tlChfAdvtExt(ilCntUpper).sSource = ""
                    ilFound = gTestChfAdvtExt(frm, ilSlfCode, tlChfAdvtExt(ilCntUpper), hlVsf, ilCurrent) 'VBC NR
                End If 'VBC NR
                If ilFound Then 'VBC NR
                    ilFound = False 'VBC NR
                    If ilTestCntrNo Then 'VBC NR
                        'For ilLoop = 1 To ilCntUpper - 1 Step 1 'VBC NR
                        For illoop = ilLowLimit To ilCntUpper - 1 Step 1 'VBC NR
                            If tlChfAdvtExt(illoop).lCntrNo = tlChfAdvtExt(ilCntUpper).lCntrNo Then 'VBC NR
                                If tlChfAdvtExt(illoop).iCntRevNo < tlChfAdvtExt(ilCntUpper).iCntRevNo Then 'VBC NR
                                    tlChfAdvtExt(illoop) = tlChfAdvtExt(ilCntUpper) 'VBC NR
                                End If 'VBC NR
                                ilFound = True 'VBC NR
                                Exit For 'VBC NR
                            End If 'VBC NR
                        Next illoop 'VBC NR
                    End If 'VBC NR
                    If Not ilFound Then 'VBC NR
                        ilFound = False 'VBC NR
                        If ilAAS = 1 Then 'VBC NR
                            If tlChfAdvtExt(ilCntUpper).iAgfCode <> 0 Then 'VBC NR
                                For illoop = LBound(imAdvtCode) To UBound(imAdvtCode) - 1 Step 1 'VBC NR
                                    If imAdvtCode(illoop) = tlChfAdvtExt(ilCntUpper).iAgfCode Then 'VBC NR
                                        ilFound = True 'VBC NR
                                        Exit For 'VBC NR
                                    End If 'VBC NR
                                Next illoop 'VBC NR
                            Else 'VBC NR
                                ilFound = True 'VBC NR
                            End If 'VBC NR
                            If Not ilFound Then 'VBC NR
                                imAdvtCode(ilUpper) = tlChfAdvtExt(ilCntUpper).iAgfCode 'VBC NR
                                ilUpper = ilUpper + 1 'VBC NR
                                'ReDim Preserve imAdvtCode(1 To ilUpper) As Integer 'VBC NR
                                ReDim Preserve imAdvtCode(0 To ilUpper) As Integer 'VBC NR
                            End If 'VBC NR
                        ElseIf ilAAS = 2 Then 'VBC NR
                            If ilSlfCode > 0 Then 'VBC NR
                                For illoop = LBound(imAdvtCode) To UBound(imAdvtCode) - 1 Step 1 'VBC NR
                                    If imAdvtCode(illoop) = ilSlfCode Then 'VBC NR
                                        ilFound = True 'VBC NR
                                        Exit For 'VBC NR
                                    End If 'VBC NR
                                Next illoop 'VBC NR
                                If Not ilFound Then 'VBC NR
                                    imAdvtCode(ilUpper) = ilSlfCode 'VBC NR
                                    ilUpper = ilUpper + 1 'VBC NR
                                    'ReDim Preserve imAdvtCode(1 To ilUpper) As Integer 'VBC NR
                                    ReDim Preserve imAdvtCode(0 To ilUpper) As Integer 'VBC NR
                                End If 'VBC NR
                            Else 'VBC NR
                                For ilSlf = 0 To 9 Step 1 'VBC NR
                                    If tlChfAdvtExt(ilCntUpper).iSlfCode(ilSlf) > 0 Then 'VBC NR
                                        ilFound = False 'VBC NR
                                        For illoop = LBound(imAdvtCode) To UBound(imAdvtCode) - 1 Step 1 'VBC NR
                                            If imAdvtCode(illoop) = tlChfAdvtExt(ilCntUpper).iSlfCode(ilSlf) Then 'VBC NR
                                                ilFound = True 'VBC NR
                                                Exit For 'VBC NR
                                            End If 'VBC NR
                                        Next illoop 'VBC NR
                                        If Not ilFound Then 'VBC NR
                                            imAdvtCode(ilUpper) = tlChfAdvtExt(ilCntUpper).iSlfCode(ilSlf) 'VBC NR
                                            ilUpper = ilUpper + 1 'VBC NR
                                            'ReDim Preserve imAdvtCode(1 To ilUpper) As Integer 'VBC NR
                                            ReDim Preserve imAdvtCode(0 To ilUpper) As Integer 'VBC NR
                                        End If 'VBC NR
                                    End If 'VBC NR
                                Next ilSlf 'VBC NR
                            End If 'VBC NR
                        Else 'VBC NR
                            For illoop = LBound(imAdvtCode) To UBound(imAdvtCode) - 1 Step 1 'VBC NR
                                If imAdvtCode(illoop) = tlChfAdvtExt(ilCntUpper).iAdfCode Then 'VBC NR
                                    ilFound = True 'VBC NR
                                    Exit For 'VBC NR
                                End If 'VBC NR
                            Next illoop 'VBC NR
                            If Not ilFound Then 'VBC NR
                                imAdvtCode(ilUpper) = tlChfAdvtExt(ilCntUpper).iAdfCode 'VBC NR
                                ilUpper = ilUpper + 1 'VBC NR
                                'ReDim Preserve imAdvtCode(1 To ilUpper) As Integer 'VBC NR
                                ReDim Preserve imAdvtCode(0 To ilUpper) As Integer 'VBC NR
                            End If 'VBC NR
                        End If 'VBC NR
                        ilCntUpper = ilCntUpper + 1 'VBC NR
                        'ReDim Preserve tlChfAdvtExt(1 To ilCntUpper) As CHFADVTEXT 'VBC NR
                        ReDim Preserve tlChfAdvtExt(ilLowLimit To ilCntUpper) As CHFADVTEXT 'VBC NR
                    End If 'VBC NR
                End If 'VBC NR
                ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos) 'VBC NR
                Do While ilRet = BTRV_ERR_REJECT_COUNT 'VBC NR
                    ilRet = btrExtGetNext(hlChf, tlChfAdvtExt(ilCntUpper), ilExtLen, llRecPos) 'VBC NR
                Loop 'VBC NR
            Loop 'VBC NR
            If ilAAS = 1 Then 'VBC NR
                For illoop = LBound(imAdvtCode) To UBound(imAdvtCode) - 1 Step 1 'VBC NR
                    tlAgfSrchKey.iCode = imAdvtCode(illoop) 'VBC NR
                    ilRet = btrGetEqual(hlFile, tlagf, ilAgfRecLen, tlAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'VBC NR
                    On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
                    gCPErrorMsg ilRet, "gPopAASBoxViaCntr (btrGetEqual: Agf)", frm 'VBC NR
                    On Error GoTo 0 'VBC NR
                    If (tlagf.sState = "A") Or ((tlagf.sState = "D") And (ilDormant)) Then 'VBC NR
                        slName = Trim$(tlagf.sName) & ", " & Trim$(tlagf.sCityID) 'VBC NR
                        Do While Len(slName) < Len(tlagf.sName) + Len(tlagf.sCityID) 'VBC NR
                            slName = slName & " " 'VBC NR
                        Loop 'VBC NR
                        slName = slName & "\" & Trim$(str$(tlagf.iCode)) 'VBC NR
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit For
                        'End If
                        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        tlSortCode(ilSortCode).sKey = slName 'VBC NR
                        If ilSortCode >= UBound(tlSortCode) Then 'VBC NR
                            ReDim Preserve tlSortCode(ilSortLowLimit To UBound(tlSortCode) + 100) As SORTCODE 'VBC NR
                        End If 'VBC NR
                        ilSortCode = ilSortCode + 1 'VBC NR
                    End If 'VBC NR
                Next illoop 'VBC NR
            ElseIf ilAAS = 2 Then 'VBC NR
                For illoop = LBound(imAdvtCode) To UBound(imAdvtCode) - 1 Step 1 'VBC NR
                    tlSlfSrchKey.iCode = imAdvtCode(illoop) 'VBC NR
                    ilRet = btrGetEqual(hlFile, tlSlf, ilSlfRecLen, tlSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'VBC NR
                    On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
                    gCPErrorMsg ilRet, "gPopAASBoxViaCntr (btrGetEqual: Slf)", frm 'VBC NR
                    On Error GoTo 0 'VBC NR
                    If (tlSlf.sState = "A") Or ((tlSlf.sState = "D") And (ilDormant)) Then 'VBC NR
                        If igSlfFirstNameFirst Then 'VBC NR
                            slName = Trim$(tlSlf.sFirstName) & " " & Trim$(tlSlf.sLastName) 'VBC NR
                        Else 'VBC NR
                            slName = Trim$(tlSlf.sLastName) & ", " & Trim$(tlSlf.sFirstName) 'VBC NR
                        End If 'VBC NR
                        Do While Len(slName) < Len(tlSlf.sFirstName) + Len(tlSlf.sLastName) 'VBC NR
                            slName = slName & " " 'VBC NR
                        Loop 'VBC NR
                        slName = slName & "\" & Trim$(str$(tlSlf.iCode)) 'VBC NR
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit For
                        'End If
                        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        tlSortCode(ilSortCode).sKey = slName 'VBC NR
                        If ilSortCode >= UBound(tlSortCode) Then 'VBC NR
                            ReDim Preserve tlSortCode(ilSortLowLimit To UBound(tlSortCode) + 100) As SORTCODE 'VBC NR
                        End If 'VBC NR
                        ilSortCode = ilSortCode + 1 'VBC NR
                    End If 'VBC NR
                Next illoop 'VBC NR
            Else 'VBC NR
                For illoop = LBound(imAdvtCode) To UBound(imAdvtCode) - 1 Step 1 'VBC NR
                    tlAdfSrchKey.iCode = imAdvtCode(illoop) 'VBC NR
                    ilRet = btrGetEqual(hlFile, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'VBC NR
                    On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
                    gCPErrorMsg ilRet, "gPopAASBoxViaCntr (btrGetEqual: Adf)", frm 'VBC NR
                    On Error GoTo 0 'VBC NR
                    If (tlAdf.sState = "A") Or ((tlAdf.sState = "D") And (ilDormant)) Then 'VBC NR
                        If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sAddrID) <> "") Then 'VBC NR
                            slAdvtName = Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID) 'VBC NR
                        Else 'VBC NR
                            slAdvtName = Trim$(tlAdf.sName) 'VBC NR
                        End If 'VBC NR
                        slName = slAdvtName & "\" & Trim$(str$(tlAdf.iCode)) 'VBC NR
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit For
                        'End If
                        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        tlSortCode(ilSortCode).sKey = slName 'VBC NR
                        If ilSortCode >= UBound(tlSortCode) Then 'VBC NR
                            ReDim Preserve tlSortCode(ilSortLowLimit To UBound(tlSortCode) + 100) As SORTCODE 'VBC NR
                        End If 'VBC NR
                        ilSortCode = ilSortCode + 1 'VBC NR
                    End If 'VBC NR
                Next illoop 'VBC NR
            End If 'VBC NR
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilSortLowLimit To ilSortCode) As SORTCODE 'VBC NR
            If UBound(tlSortCode) - ilSortLowLimit > ilSortLowLimit Then 'VBC NR
                ArraySortTyp fnAV(tlSortCode(), ilSortLowLimit), UBound(tlSortCode) - ilSortLowLimit, 0, LenB(tlSortCode(ilSortLowLimit)), 0, LenB(tlSortCode(ilSortLowLimit).sKey), 0 'VBC NR
            End If 'VBC NR
        End If 'VBC NR
        Erase imAdvtCode 'VBC NR
        ilRet = btrClose(hlFile) 'VBC NR
        btrDestroy hlFile 'VBC NR
        ilRet = btrClose(hlVsf) 'VBC NR
        btrDestroy hlVsf 'VBC NR
        ilRet = btrClose(hlVef) 'VBC NR
        btrDestroy hlVef 'VBC NR
        ilRet = btrClose(hlChf) 'VBC NR
        On Error GoTo gPopAASBoxViaCntrErr 'VBC NR
        gBtrvErrorMsg ilRet, "gPopAASBoxViaCntr (btrReset):" & "Chf.Btr", frm 'VBC NR
        On Error GoTo 0 'VBC NR
        btrDestroy hlChf 'VBC NR
    End If 'VBC NR
    llLen = 0 'VBC NR
    For illoop = ilSortLowLimit To UBound(tlSortCode) - 1 Step 1 'VBC NR
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop) 'VBC NR
        ilRet = gParseItem(slNameCode, 1, "\", slName) 'VBC NR
        slName = Trim$(slName) 'VBC NR
        If Not gOkAddStrToListBox(slName, llLen, True) Then 'VBC NR
            Exit For 'VBC NR
        End If 'VBC NR
        lbcLocal.AddItem slName  'Add ID to list box 'VBC NR
    Next illoop 'VBC NR
    Exit Function 'VBC NR
gPopAASBoxViaCntrErr: 'VBC NR
    ilRet = btrClose(hlFile) 'VBC NR
    btrDestroy hlFile 'VBC NR
    ilRet = btrClose(hlVsf) 'VBC NR
    btrDestroy hlVsf 'VBC NR
    ilRet = btrClose(hlVef) 'VBC NR
    btrDestroy hlVef 'VBC NR
    ilRet = btrClose(hlChf) 'VBC NR
    btrDestroy hlChf 'VBC NR
    gDbg_HandleError "PopSubs: gPopAASBoxViaCntr" 'VBC NR
'    gPopAASBoxViaCntr = CP_MSG_NOSHOW
'    Exit Function
gPopAASBoxViaCntrErr2: 'VBC NR
    ilRet = 1 'VBC NR
    Resume Next 'VBC NR
End Function 'VBC NR

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAdvtBox                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     advertiser names (if direct,    *
'*                     add \Direct to name)            *
'*                                                     *
'*******************************************************
Function gPopAdvtBox(frm As Form, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopAdvtBox (MainForm, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       lbcLocal (I/O)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       igPopAdfAgfDormant(I) - Include Dormant
'       igPopExternalAdvt(I) - Include external invoiced advertisers
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'Adf date/time stamp
    'Dim tlAdf As ADF
    Dim tlAdfExt As ADFEXT    'Advertiser extract record
    Dim slName As String
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilAdf As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Adf.Btr") & str$(igPopAdfAgfDormant)

    'On Error GoTo gPopAdvtBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopAdvtBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopAdvtBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilRet = gObtainAdvt()

            For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                tlAdfExt = tgCommAdf(ilAdf)
                If (igPopExternalAdvt = True) Or (tlAdfExt.sRepInvGen <> "E") Then
                    If (tlAdfExt.sState <> "D") Or (igPopAdfAgfDormant) Then
                        If Trim$(tlAdfExt.sBillAgyDir) = "D" Then
                            slName = Trim$(tlAdfExt.sName)
                            If Trim$(tlAdfExt.sAddrID) <> "" Then
                                slName = slName & ", " & Trim$(tlAdfExt.sAddrID) & "/Direct"
                            Else
                                slName = slName & "/Direct"
                            End If
                        Else
                            slName = Trim$(tlAdfExt.sName)
                        End If
                        Do While Len(slName) < Len(tlAdfExt.sName) + 21 'was 8, increased to 21 for /Non Direct added to Sales History and Collection
                            slName = slName & " "
                        Loop
                        slName = slName & "\" & Trim$(str$(tlAdfExt.iCode))
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit Do
                        'End If
                        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                    End If
                End If
            Next ilAdf
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopAdvtBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function

'    ilRet = btrClose(hlAdf)
'    btrDestroy hlAdf
    gDbg_HandleError "PopSubs: gPopAdvtBox"
'    gPopAdvtBox = CP_MSG_NOSHOW
'    Exit Function
gPopAdvtBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAdvtProdBox                 *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     advertiser Product names        *
'*                                                     *
'*******************************************************
Function gPopAdvtProdBox(frm As Form, ilAdvtCode As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopAdvtProdBox (MainForm, ilAdvt, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAdvt (I)- Advertise code value
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'Adf date/time stamp
    Dim hlPrf As Integer        'Adf handle
    Dim ilRecLen As Integer     'Record length
    Dim tlPrf As PRF
    Dim slName As String
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlSrchKey As PRFKEY1  'PrF key record image
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Prf.Btr") & Trim$(str$(ilAdvtCode))

    'On Error GoTo gPopAdvtProdBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopAdvtProdBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If

    gPopAdvtProdBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlPrf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopAdvtProdBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtProdBox (btrOpen): Prf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlPrf) 'btrRecordLength(hlPrf)  'Get and save record length
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlPrf)
        llNoRec = gExtNoRec(ilExtLen)
        btrExtClear hlPrf
        Call btrExtSetBounds(hlPrf, llNoRec, -1, "UC", "PRF", "") 'Set extract limits (all records)
        tlSrchKey.iAdfCode = ilAdvtCode
        ilRet = btrGetGreaterOrEqual(hlPrf, tlPrf, ilRecLen, tlSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
        If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
            ilRet = btrClose(hlPrf)
            On Error GoTo gPopAdvtProdBoxErr
            gBtrvErrorMsg ilRet, "gPopAdvtProdBox (btrReset):" & "Prf.Btr", frm
            On Error GoTo 0
            btrDestroy hlPrf
            Exit Function
        End If
        tlIntTypeBuff.iType = ilAdvtCode
        ilOffSet = gFieldOffset("Prf", "PrfAdfCode")
        ilRet = btrExtAddLogicConst(hlPrf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        Call btrExtSetBounds(hlPrf, llNoRec, -1, "UC", "PRF", "") 'Set extract limits (all records)
        ilOffSet = 0
        ilRet = btrExtAddField(hlPrf, ilOffSet, ilRecLen)  'Extract iCode field
        On Error GoTo gPopAdvtProdBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtProdBox (btrExtAddField):" & "Prf.Btr", frm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlPrf, tlPrf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopAdvtProdBoxErr
            gBtrvErrorMsg ilRet, "gPopAdvtProdBox (btrExtGetNextExt):" & "Prf.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlPrf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlPrf, tlPrf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slName = tlPrf.sName
                slName = slName & "\" & Trim$(str$(tlPrf.lCode))
                'If Not gOkAddStrToListBox(slName, llLen, True) Then
                '    Exit Do
                'End If
                'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                tlSortCode(ilSortCode).sKey = slName
                If ilSortCode >= UBound(tlSortCode) Then
                    ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                End If
                ilSortCode = ilSortCode + 1
                ilRet = btrExtGetNext(hlPrf, tlPrf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlPrf, tlPrf, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If
        ilRet = btrClose(hlPrf)
        On Error GoTo gPopAdvtProdBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtProdBox (btrReset):" & "Prf.Btr", frm
        On Error GoTo 0
        btrDestroy hlPrf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopAdvtProdBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop

    Exit Function
gPopAdvtProdBoxErr:
    ilRet = btrClose(hlPrf)
    btrDestroy hlPrf
    gDbg_HandleError "PopSubs: gPopAdvtProdBox"
'    gPopAdvtProdBox = CP_MSG_NOSHOW
'    Exit Function
gPopAdvtProdBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAgyBox                      *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     agency names/City ID            *
'*                                                     *
'*******************************************************
Function gPopAgyBox(frm As Form, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopAgyBox (MainForm, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'Agf date/time stamp
    Dim hlAgf As Integer        'Agf handle
    'Dim tlAgf As AGF
    Dim tlAgfExt As AGFEXT    'Agency extract record
    Dim slName As String
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilAgf As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Agf.Btr") & str$(igPopAdfAgfDormant)

    'On Error GoTo gPopAgyBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopAgyBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopAgyBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilRet = gObtainAgency()

            For ilAgf = LBound(tgCommAgf) To UBound(tgCommAgf) - 1 Step 1
                tlAgfExt = tgCommAgf(ilAgf)
                If (tlAgfExt.sState <> "D") Or (igPopAdfAgfDormant) Then
                    slName = Trim$(tlAgfExt.sName) & ", " & Trim$(tlAgfExt.sCityID)
                    slName = slName & "\" & Trim$(str$(tlAgfExt.iCode))
                    'If Not gOkAddStrToListBox(slName, llLen, True) Then
                    '    Exit Do
                    'End If
                    'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                    tlSortCode(ilSortCode).sKey = slName
                    If ilSortCode >= UBound(tlSortCode) Then
                        ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                    End If
                    ilSortCode = ilSortCode + 1
                End If
            Next ilAgf
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
    End If
    llLen = 0
    For illoop = 0 To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopAgyBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function

    ilRet = btrClose(hlAgf)
    btrDestroy hlAgf
    gDbg_HandleError "PopSubs: gPopAgyBox"
'    gPopAgyBox = CP_MSG_NOSHOW
'    Exit Function
gPopAgyBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAgyBoxNameCityBox           *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     agency, agency names and city   *
'*                                                     *
'*******************************************************
Function gPopAgyBoxNameCityBox(frm As Form, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, ilPopName As Integer, lbcName As Control, ilPopCity As Integer, lbcCity As Control) As Integer
'
'   ilRet = gPopAgyBoxNameCityBox (MainForm, lbcLocal, tlSortCode(), slSortCodeTag, ilPopName, lbcName, ilPopCity, lbcCity)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilPopName (I)- True=populate with unique names; False=Don't populate
'       lbcName (I)- List box to be populated with names
'       ilPopCity (I)- True=populate with unique city IDs; False=Don't populate
'       lbcCity (I)- List box to be populated with city IDs
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'Agf date/time stamp
    Dim hlAgf As Integer        'Agf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlagf As AGF
    Dim tlAgfExt As AGFEXT    'Agency extract record
    Dim slName As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilLowLimit As Integer
    'Dim ilPop As Integer

    'ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Agf.Btr") & Trim$(str$(ilPopName)) & Trim$(str$(ilPopCity))

    'On Error GoTo gPopAgyBoxNameCityBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopAgyBoxNameCityBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            'ilPop = False
        End If
    End If
    gPopAgyBoxNameCityBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp

    hlAgf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopAgyBoxNameCityBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrOpen):" & "Agf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlagf) 'btrRecordLength(hlAgf)  'Get and save record length
    ilSortCode = ilLowLimit
    ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    If ilPopName Then
        lbcName.Clear
    End If
    If ilPopCity Then
        lbcCity.Clear   'VB list box clear (list box used to retain code number so record can be found)
    End If
    ilExtLen = Len(tlAgfExt)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAgf) 'Obtain number of records
    btrExtClear hlAgf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlAgf, tlagf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlAgf)
        On Error GoTo gPopAgyBoxNameCityBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrReset):" & "Agf.Btr", frm
        On Error GoTo 0
        btrDestroy hlAgf
        Exit Function
    Else
        On Error GoTo gPopAgyBoxNameCityBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrGetFirst):" & "Agf.Btr", frm
        On Error GoTo 0
    End If
    Call btrExtSetBounds(hlAgf, llNoRec, -1, "UC", "AGFEXTPK", AGFEXTPK) 'Set extract limits (all records)
    ilOffSet = gFieldOffset("Agf", "AgfCode")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 2)  'Extract iCode field
    On Error GoTo gPopAgyBoxNameCityBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrExtAddField):" & "Agf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Agf", "AgfName")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 40)  'Extract Name field
    On Error GoTo gPopAgyBoxNameCityBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrExtAddField):" & "Agf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Agf", "AgfCity")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 5) 'Extract city ID
    On Error GoTo gPopAgyBoxNameCityBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrExtAddField):" & "Agf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Agf", "AgfCreditRestr")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 1) 'Extract city ID
    On Error GoTo gPopAgyBoxNameCityBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrExtAddField):" & "Agf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Agf", "AgfMnfSort")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 2) 'Extract city ID
    On Error GoTo gPopAgyBoxNameCityBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrExtAddField):" & "Agf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Agf", "AgfState")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 1) 'Extract city ID
    On Error GoTo gPopAgyBoxNameCityBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrExtAddField):" & "Agf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Agf", "AgfTrfCode")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 2) 'Extract city ID
    On Error GoTo gPopAgyBoxNameCityBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrExtAddField):" & "Agf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Agf", "AgfArfInvCode")
    ilRet = btrExtAddField(hlAgf, ilOffSet, 2) 'Extract city ID
    On Error GoTo gPopAgyBoxNameCityBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrExtAddField):" & "Agf.Btr", frm
    On Error GoTo 0
    'ilRet = btrExtGetNextExt(hlAgf)    'Extract record
    ilRet = btrExtGetNext(hlAgf, tlAgfExt, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo gPopAgyBoxNameCityBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrExtGetNextExt):" & "Agf.Btr", frm
        On Error GoTo 0
        ilExtLen = Len(tlAgfExt)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlAgf, tlAgfExt, ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlAgf, tlAgfExt, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            slName = Trim$(tlAgfExt.sName) & ", " & Trim$(tlAgfExt.sCityID)
            slName = slName & "\" & Trim$(str$(tlAgfExt.iCode))
            Do While Len(slName) < Len(tlagf.sName) + Len(tlagf.sCityID) + 2
                slName = slName & " "
            Loop
            'If Not gOkAddStrToListBox(slName, llLen, True) Then
            '    Exit Do
            'End If
            'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
            tlSortCode(ilSortCode).sKey = slName
            If ilSortCode >= UBound(tlSortCode) Then
                ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
            End If
            ilSortCode = ilSortCode + 1
            If ilPopName Then
                slName = Trim$(tlAgfExt.sName)
                gFindMatch slName, 0, lbcName
                If gLastFound(lbcName) < 0 Then
                    If Not gOkAddStrToListBox(slName, llLen, True) Then
                        Exit Do
                    End If
                    lbcName.AddItem slName
                End If
            End If
            If ilPopCity Then
                slName = Trim$(tlAgfExt.sCityID)
                gFindMatch slName, 0, lbcCity
                If gLastFound(lbcCity) < 0 Then
                    lbcCity.AddItem slName
                End If
            End If
            ilRet = btrExtGetNext(hlAgf, tlAgfExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlAgf, tlAgfExt, ilExtLen, llRecPos)
            Loop
        Loop
        'Sort then output new headers and lines
        ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
            ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
        End If
        llLen = 0
        For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
            slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            If ilRet <> CP_MSG_NONE Then
                gPopAgyBoxNameCityBox = CP_MSG_PARSE
                Exit Function
            End If
            slName = Trim$(slName)
            If Not gOkAddStrToListBox(slName, llLen, True) Then
                Exit For
            End If
            lbcLocal.AddItem slName  'Add ID to list box
        Next illoop
    End If
    ilRet = btrClose(hlAgf)
    On Error GoTo gPopAgyBoxNameCityBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyBoxNameCityBox (btrReset):" & "Agf.Btr", frm
    On Error GoTo 0
    btrDestroy hlAgf
    Exit Function
gPopAgyBoxNameCityBoxErr:
    ilRet = btrClose(hlAgf)
    btrDestroy hlAgf
    gDbg_HandleError "PopSubs: gPopAgyBoxNameCityBox"
'    gPopAgyBoxNameCityBox = CP_MSG_NOSHOW
'    Exit Function
gPopAgyBoxNameCityBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopBookNameBox                 *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     Book Name and Date              *
'*                                                     *
'*******************************************************
Function gPopBookNameBox(frm As Form, ilFlag As Integer, ilLnBookDnf As Integer, ilVefCode As Integer, ilSort As Integer, ilShow As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopBookNameBox (MainForm, ilFlag, ilLnBookDnf, ilVefCode, ilSort, ilShow, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilFlag(I)- 0=Obtain all bboks based on ilVefCode; 1=Only insert the default book for ilvefcode;
'                  2=only insert the line referenced book defined in ilLnBookDnf
'                  16=Like 0 but only for books with 16 buckets
'                  18=Like 0 but only for books with 18 buckets
'       ilLnBookDnf(I)- Line Book DNF (if zero, then use default book).  ilFlag = 2
'       ilVefCode(I)- Vehicle code (zero if ignored, If specified, then check if Drf defined for vehcile)
'       ilSort(I)- 0=Sort by Book Name only; 1= sort by date, then book name
'       ilShow(I)- 0=Book Name only; 1=Book Name followed by Date
'       lbcLocal (I/O)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'Dnf date/time stamp
    Dim hlDnf As Integer        'Dnf handle
    Dim ilDnfRecLen As Integer     'Record length
    Dim tlDnfSrchKey As INTKEY0
    Dim tlDnf As DNF
    Dim hlDrf As Integer        'Dnf handle
    Dim ilDrfRecLen As Integer     'Record length
    Dim tlDrf As DRF
    Dim tlDrfSrchKey As DRFKEY0
    Dim llNoRec As Long         'Number of records in Sof
    Dim slName As String
    Dim slDate As String
    Dim llDate As Long
    Dim slSortDate As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim llLen As Long
    Dim ilAdd As Integer
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    slStamp = gFileDateTime(sgDBPath & "Dnf.Btr") & gFileDateTime(sgDBPath & "Drf.Btr") & Trim$(str$(ilFlag)) & Trim$(str$(ilLnBookDnf)) & Trim$(str$(ilVefCode)) & Trim$(str$(ilSort)) & Trim$(str$(ilShow))

    'On Error GoTo gPopBookNameErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopBookNameBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopBookNameBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        hlDnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopBookNameBoxErr
        gBtrvErrorMsg ilRet, "gPopBookNameBox (btrOpen):" & "Dnf.Btr", frm
        On Error GoTo 0
        ilDnfRecLen = Len(tlDnf) 'btrRecordLength(hlDnf)  'Get and save record length
        If (ilFlag = 1) Or ((ilFlag = 2) And (ilLnBookDnf = 0)) Then  'Use default book
            'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
            '    If tgMVef(ilLoop).iCode = ilVefCode Then
                illoop = gBinarySearchVef(ilVefCode)
                If illoop <> -1 Then
                    tlDnfSrchKey.iCode = tgMVef(illoop).iDnfCode
                    ilRet = btrGetEqual(hlDnf, tlDnf, ilDnfRecLen, tlDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        '6/4/16: Replace GoSub
                        'GoSub CreateSort
                        mCreateSort ilSort, ilFlag, ilShow, tlDnf, slName, ilSortCode, tlSortCode()
                    End If
                    'Exit For
                End If
            'Next ilLoop
        ElseIf ilFlag = 2 Then
            'tlDnfSrchKey.iCode = ilLnBookDnf
            If ilLnBookDnf <= 0 Then
                tlDnfSrchKey.iCode = 0
            Else
                tlDnfSrchKey.iCode = ilLnBookDnf
            End If
            ilRet = btrGetEqual(hlDnf, tlDnf, ilDnfRecLen, tlDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                '6/4/16: Replace GoSub
                'GoSub CreateSort
                mCreateSort ilSort, ilFlag, ilShow, tlDnf, slName, ilSortCode, tlSortCode()
            Else
                'Use the default if line book not available
                'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                '    If tgMVef(ilLoop).iCode = ilVefCode Then
                    illoop = gBinarySearchVef(ilVefCode)
                    If illoop <> -1 Then
                        tlDnfSrchKey.iCode = tgMVef(illoop).iDnfCode
                        ilRet = btrGetEqual(hlDnf, tlDnf, ilDnfRecLen, tlDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            '6/4/16: Replace GoSub
                            'GoSub CreateSort
                            mCreateSort ilSort, ilFlag, ilShow, tlDnf, slName, ilSortCode, tlSortCode()
                        End If
                        'Exit For
                    End If
                'Next ilLoop
            End If
        Else
            hlDrf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
            ilRet = btrOpen(hlDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            On Error GoTo gPopBookNameBoxErr
            gBtrvErrorMsg ilRet, "gPopBookNameBox (btrOpen):" & "Drf.Btr", frm
            On Error GoTo 0
            ilDrfRecLen = Len(tlDrf) 'btrRecordLength(hlDrf)  'Get and save record length
            ilExtLen = Len(tlDnf)  'Extract operation record size
            llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlDnf) 'Obtain number of records
            btrExtClear hlDnf   'Clear any previous extend operation
            ilRet = btrGetFirst(hlDnf, tlDnf, ilDnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_END_OF_FILE Then
                ilRet = btrClose(hlDnf)
                btrDestroy hlDnf
                ilRet = btrClose(hlDrf)
                btrDestroy hlDrf
                Exit Function
            Else
                On Error GoTo gPopBookNameBoxErr
                gBtrvErrorMsg ilRet, "gPopBookNameBox (btrGetFirst):" & "Dnf.Btr", frm
                On Error GoTo 0
            End If
            Call btrExtSetBounds(hlDnf, llNoRec, -1, "UC", "DNF", "") 'Set extract limits (all records including first)
            ilOffSet = 0
            ilRet = btrExtAddField(hlDnf, ilOffSet, ilDnfRecLen)  'Extract iCode field
            On Error GoTo gPopBookNameBoxErr
            gBtrvErrorMsg ilRet, "gPopBookNameBox (btrExtAddField):" & "Dnf.Btr", frm
            On Error GoTo 0
            'ilRet = btrExtGetNextExt(hlDnf)    'Extract record
            ilRet = btrExtGetNext(hlDnf, tlDnf, ilExtLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo gPopBookNameBoxErr
                gBtrvErrorMsg ilRet, "gPopBookNameBox (btrExtGetNextExt):" & "Dnf.Btr", frm
                On Error GoTo 0
                ilExtLen = Len(tlDnf)  'Extract operation record size
                'ilRet = btrExtGetFirst(hlDnf, tlDnfExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlDnf, tlDnf, ilExtLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    ilAdd = False
                    If ((ilFlag <> 16) And (ilFlag <> 18)) Or ((ilFlag = 16) And (Trim$(tlDnf.sForm) <> "8")) Or ((ilFlag = 18) And (Trim$(tlDnf.sForm) = "8")) Then
                        If ilVefCode <= 0 Then
                            ilAdd = True
                        Else
                            'Test if Drf exist for vehicle specified
                            tlDrfSrchKey.iDnfCode = tlDnf.iCode
                            tlDrfSrchKey.sDemoDataType = "D"    'Demo Data
                            tlDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
                            tlDrfSrchKey.iVefCode = ilVefCode
                            tlDrfSrchKey.sInfoType = "D"        'Daypart
                            tlDrfSrchKey.iRdfCode = 0
                            ilRet = btrGetGreaterOrEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            If (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tlDnf.iCode) And (tlDrf.iVefCode = ilVefCode) Then
                                ilAdd = True
                            Else
                                tlDrfSrchKey.iDnfCode = tlDnf.iCode
                                tlDrfSrchKey.sDemoDataType = "D"    'Demo Data
                                tlDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
                                tlDrfSrchKey.iVefCode = ilVefCode
                                tlDrfSrchKey.sInfoType = "T"        'Daypart Or Time
                                tlDrfSrchKey.iRdfCode = 0
                                ilRet = btrGetGreaterOrEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                If (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tlDnf.iCode) And (tlDrf.iVefCode = ilVefCode) Then
                                    ilAdd = True
                                Else
                                    tlDrfSrchKey.iDnfCode = tlDnf.iCode
                                    tlDrfSrchKey.sDemoDataType = "D"    'Demo Data
                                    tlDrfSrchKey.iMnfSocEco = 0 'ilMnfSocEco
                                    tlDrfSrchKey.iVefCode = ilVefCode
                                    tlDrfSrchKey.sInfoType = "V"        'Daypart Or Time
                                    tlDrfSrchKey.iRdfCode = 0
                                    ilRet = btrGetGreaterOrEqual(hlDrf, tlDrf, ilDrfRecLen, tlDrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                    If (ilRet = BTRV_ERR_NONE) And (tlDrf.iDnfCode = tlDnf.iCode) And (tlDrf.iVefCode = ilVefCode) Then
                                        ilAdd = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If ilAdd Then
                        '6/4/16: Replace GoSub
                        'GoSub CreateSort
                        mCreateSort ilSort, ilFlag, ilShow, tlDnf, slName, ilSortCode, tlSortCode()
                        'If ilSort = 0 Then
                        '    gUnpackDate tlDnf.iBookDate(0), tlDnf.iBookDate(1), slDate
                        '    If ilShow = 0 Then
                        '        slName = Trim$(tlDnf.sBookName)
                        '    Else
                        '        slName = Trim$(tlDnf.sBookName) & ": " & slDate
                        '    End If
                        '    slName = slName & "\" & Trim$(Str$(tlDnf.iCode))
                        'Else
                        '    gUnpackDateLong tlDnf.iBookDate(0), tlDnf.iBookDate(1), llDate
                        '    llDate = 99999 - llDate
                        '    slSortDate = Trim$(Str$(llDate))
                        '    Do While Len(slSortDate) < 5
                        '        slSortDate = "0" & slSortDate
                        '    Loop
                        '    gUnpackDate tlDnf.iBookDate(0), tlDnf.iBookDate(1), slDate
                        '    If ilShow = 0 Then
                        '        slName = Trim$(tlDnf.sBookName)
                        '    Else
                        '        slName = Trim$(tlDnf.sBookName) & ": " & slDate
                        '    End If
                        '    slName = slSortDate & "|" & slName & "\" & Trim$(Str$(tlDnf.iCode))
                        'End If
                        ''If Not gOkAddStrToListBox(slName, llLen, True) Then
                        ''    Exit Do
                        ''End If
                        ''lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        'tlSortCode(ilSortCode).sKey = slName
                        'If ilSortCode >= UBound(tlSortCode) Then
                        '    ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 100) As SORTCODE
                        'End If
                        'ilSortCode = ilSortCode + 1
                    End If
                    ilRet = btrExtGetNext(hlDnf, tlDnf, ilExtLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlDnf, tlDnf, ilExtLen, llRecPos)
                    Loop
                Loop
            End If
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If
        On Error GoTo 0
        ilRet = btrClose(hlDnf)
        btrDestroy hlDnf
        ilRet = btrClose(hlDrf)
        btrDestroy hlDrf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        If ilSort = 0 Then
            ilRet = gParseItem(slNameCode, 1, "\", slName)
        Else
            ilRet = gParseItem(slNameCode, 2, "|", slName)
            slNameCode = slName
            ilRet = gParseItem(slNameCode, 1, "\", slName)
        End If
        If ilRet <> CP_MSG_NONE Then
            gPopBookNameBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopBookNameBoxErr:
    ilRet = btrClose(hlDnf)
    btrDestroy hlDnf
    ilRet = btrClose(hlDrf)
    btrDestroy hlDrf
    gDbg_HandleError "PopSubs: gPopBookNameBox"
'    gPopBookNameBox = CP_MSG_NOSHOW
'    Exit Function
gPopBookNameErr2:
    ilRet = 1
    Resume Next
'CreateSort:
'    If ilSort = 0 Then
'        gUnpackDate tlDnf.iBookDate(0), tlDnf.iBookDate(1), slDate
'        If ilShow = 0 Then
'            slName = Trim$(tlDnf.sBookName)
'        Else
'            slName = Trim$(tlDnf.sBookName) & ": " & slDate
'        End If
'        slName = slName & "\" & Trim$(Str$(tlDnf.iCode))
'    Else
'        gUnpackDateLong tlDnf.iBookDate(0), tlDnf.iBookDate(1), llDate
'        llDate = 99999 - llDate
'        slSortDate = Trim$(Str$(llDate))
'        Do While Len(slSortDate) < 5
'            slSortDate = "0" & slSortDate
'        Loop
'        gUnpackDate tlDnf.iBookDate(0), tlDnf.iBookDate(1), slDate
'        If ilShow = 0 Then
'            slName = Trim$(tlDnf.sBookName)
'        Else
'            slName = Trim$(tlDnf.sBookName) & ": " & slDate
'        End If
'        slName = slSortDate & "|" & slName & "\" & Trim$(Str$(tlDnf.iCode))
'    End If
'    'If Not gOkAddStrToListBox(slName, llLen, True) Then
'    '    Exit Do
'    'End If
'    'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
'    tlSortCode(ilSortCode).sKey = slName
'    If (ilFlag = 1) Or (ilFlag = 2) Then
'        ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 1) As SORTCODE
'    Else
'        If ilSortCode >= UBound(tlSortCode) Then
'            ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 100) As SORTCODE
'        End If
'    End If
'    ilSortCode = ilSortCode + 1
'    Return
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopCntrBoxViaArray             *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate List box with contract*
'*                      information                    *
'*                                                     *
'*******************************************************
Sub gPopCntrBoxViaArray(ilAAS As Integer, ilAASCodes() As Integer, tlChfAdvtExt() As CHFADVTEXT, ilShow As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String)
'
'   gPopCntrBoxViaArray ilAAS, ilAASCodes(), tlChfAdvtExt(), ilHOState, ilShow, lbcLocal, tlSortCode(), slSortCodeTag
'   Where:
'       ilAAS(I)- 0=ilAASCodes contains array of advertiser codes for which contract are to be shown
'                 1=ilAASCodes contains array of agency codes for which contract are to be showns
'                 2=ilAASCodes contains array of salesperson codes for which contract are to be shown
'       ilAASCodes()(I)- Array of advertiser or agency or salespersons code (dimension one greater then number to be checked)
'       tlChfAdvtExt(O)- Array of contracts of the structure CHFADVTEXT which selection criteria
'                        used to populate the list box (build from call to gPopAASBoxViaCntr)
'       ilShow(I)- 0=Only show numbers, 1= Show Number and advertiser (test site value) and product and internal comment,....
'                   2=Show Number, Dates, Product and vehicle
'                   3=Show Number, Advertiser, Dates
'                   4=Show Number, Dates
'                   5=Show Number, Advertiser
'       lbcLocal (O)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
    Dim illoop As Integer
    Dim slStr As String
    Dim slExtStr As String
    Dim slNameCode As String
    Dim slCode As String
    Dim slName As String
    Dim ilIndex As Integer
    Dim slShow As String
    Dim llCntrNo As Long
    Dim ilRevNo As Integer
    Dim ilVerNo As Integer
    Dim ilExtRevNo As Integer
    Dim ilError As Integer
    Dim ilOk As Integer
    Dim llLen As Long
    Dim ilRet As Integer
    Dim ilSlf As Integer
    Dim slStamp As String
    Dim hlAdf As Integer        'Adf handle
    Dim tlAdf As ADF
    Dim ilAdfRecLen As Integer     'Record length
    Dim tlAdfSrchKey As INTKEY0
    Dim hlVef As Integer        'Vef handle
    Dim tlVef As VEF
    Dim ilVefRecLen As Integer     'Record length
    Dim tlVefSrchKey As INTKEY0
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    slStamp = gFileDateTime(sgDBPath & "Chf.Btr") & Trim$(str$(ilAAS)) & Trim$(str$(ilShow))
    For ilIndex = LBound(ilAASCodes) To UBound(ilAASCodes) - 1 Step 1
        slStamp = slStamp & Trim$(str$(ilAASCodes(ilIndex)))
    Next ilIndex

    'On Error GoTo gPopCntrBoxViaArrayErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                Exit Sub
            End If
            ilPop = False
        End If
    End If
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hlVef
            Exit Sub
        End If
        ilVefRecLen = Len(tlVef) 'btrRecordLength(hlSlf)  'Get and save record length
        hlAdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hlVef
            btrDestroy hlAdf
            Exit Sub
        End If
        ilAdfRecLen = Len(tlAdf) 'btrRecordLength(hlSlf)  'Get and save record length
        tlAdf.iCode = 0
        hmCxf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hlVef
            btrDestroy hlAdf
            btrDestroy hmCxf
            Exit Sub
        End If
        ilError = False
        llLen = 0
        For ilIndex = LBound(ilAASCodes) To UBound(ilAASCodes) - 1 Step 1
            For illoop = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
                ilOk = False
                If ilAAS = 1 Then
                    If tlChfAdvtExt(illoop).iAgfCode = ilAASCodes(ilIndex) Then
                        ilOk = True
                    End If
                ElseIf ilAAS = 2 Then
                    For ilSlf = 0 To 9 Step 1
                        If tlChfAdvtExt(illoop).iSlfCode(ilSlf) = ilAASCodes(ilIndex) Then
                            ilOk = True
                            Exit For
                        End If
                    Next ilSlf
                Else
                    If tlChfAdvtExt(illoop).iAdfCode = ilAASCodes(ilIndex) Then
                        ilOk = True
                    End If
                End If
                If ilOk Then

                    slStr = Trim$(str$(99999999 - tlChfAdvtExt(illoop).lCntrNo))
                    Do While Len(slStr) < 8
                        slStr = "0" & slStr
                    Loop
                    slName = slStr
                    slStr = Trim$(str$(999 - tlChfAdvtExt(illoop).iCntRevNo))
                    Do While Len(slStr) < 3
                        slStr = "0" & slStr
                    Loop
                    slExtStr = Trim$(str$(999 - tlChfAdvtExt(illoop).iExtRevNo))
                    Do While Len(slExtStr) < 3
                        slExtStr = "0" & slExtStr
                    Loop
                    slName = slName & "|" & slStr & "-" & slExtStr & "|"
                    If (tlChfAdvtExt(illoop).sStatus = "W") Or (tlChfAdvtExt(illoop).sStatus = "C") Or (tlChfAdvtExt(illoop).sStatus = "I") Then
                        'Add Potential
                        If tlChfAdvtExt(illoop).iMnfPotnType > 0 Then
                            slStr = " "
                        Else
                            slStr = "~"
                        End If
                        slName = slName & slStr & "|"
                    Else
                        slName = slName & " |"
                    End If
                    slStr = Trim$(str$(999 - tlChfAdvtExt(illoop).iPropVer))
                    Do While Len(slStr) < 3
                        slStr = "0" & slStr
                    Loop
                    slName = slName & slStr & "|"
                    slName = slName & tlChfAdvtExt(illoop).sStatus & "|"
                    If ilShow = 0 Then
                        'slName = Trim$(Str$(tlChfAdvtExt(ilLoop).lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt(ilLoop).iCntRevNo)) & " V" & Trim$(Str$(tlChfAdvtExt(ilLoop).iPropVer))
                    ElseIf ilShow = 2 Then
                        'slName = Trim$(Str$(tlChfAdvtExt(ilLoop).lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt(ilLoop).iCntRevNo))
                        gUnpackDate tlChfAdvtExt(illoop).iStartDate(0), tlChfAdvtExt(illoop).iStartDate(1), slStartDate
                        gUnpackDate tlChfAdvtExt(illoop).iEndDate(0), tlChfAdvtExt(illoop).iEndDate(1), slEndDate
                        slName = slName & ": " & slStartDate & "-" & slEndDate
                        slName = slName & " " & Trim$(tlChfAdvtExt(illoop).sProduct)
                        If tlChfAdvtExt(illoop).lVefCode > 0 Then
                            If tlChfAdvtExt(illoop).lVefCode <> tlVef.iCode Then
                                tlVefSrchKey.iCode = tlChfAdvtExt(illoop).lVefCode
                                ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                'On Error GoTo gPopCntrBoxViaArrayErr
                                'gCPErrorMsg ilRet, "gPopCntrBoxViaArray (btrGetEqual: Adf)", Frm
                                'On Error GoTo 0
                            End If
                            slName = slName & " " & Trim$(tlVef.sName)
                        End If
                    ElseIf ilShow = 3 Then
                        'slName = Trim$(Str$(tlChfAdvtExt(ilLoop).lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt(ilLoop).iCntRevNo))
                        If tlChfAdvtExt(illoop).iAdfCode <> tlAdf.iCode Then
                            tlAdfSrchKey.iCode = tlChfAdvtExt(illoop).iAdfCode
                            ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            'On Error GoTo gPopCntrBoxViaArrayErr
                            'gCPErrorMsg ilRet, "gPopCntrBoxViaArray (btrGetEqual: Adf)", Frm
                            'On Error GoTo 0
                        End If
                        'slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt(ilLoop).sProduct)
                        If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sAddrID) <> "") Then
                            slName = slName & " " & Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID)
                        Else
                            slName = slName & " " & Trim$(tlAdf.sName)
                        End If
                        gUnpackDate tlChfAdvtExt(illoop).iStartDate(0), tlChfAdvtExt(illoop).iStartDate(1), slStartDate
                        gUnpackDate tlChfAdvtExt(illoop).iEndDate(0), tlChfAdvtExt(illoop).iEndDate(1), slEndDate
                        slName = slName & " " & slStartDate & "-" & slEndDate
                    ElseIf ilShow = 4 Then
                        'slName = Trim$(Str$(tlChfAdvtExt(ilLoop).lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt(ilLoop).iCntRevNo))
                        gUnpackDate tlChfAdvtExt(illoop).iStartDate(0), tlChfAdvtExt(illoop).iStartDate(1), slStartDate
                        gUnpackDate tlChfAdvtExt(illoop).iEndDate(0), tlChfAdvtExt(illoop).iEndDate(1), slEndDate
                        slName = slName & " " & slStartDate & "-" & slEndDate
                    ElseIf ilShow = 5 Then
                        'slName = Trim$(Str$(tlChfAdvtExt(ilLoop).lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt(ilLoop).iCntRevNo))
                        If tlChfAdvtExt(illoop).iAdfCode <> tlAdf.iCode Then
                            tlAdfSrchKey.iCode = tlChfAdvtExt(illoop).iAdfCode
                            ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            'On Error GoTo gPopCntrBoxViaArrayErr
                            'gCPErrorMsg ilRet, "gPopCntrBoxViaArray (btrGetEqual: Adf)", Frm
                            'On Error GoTo 0
                        End If
                        'slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt(ilLoop).sProduct)
                        If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sAddrID) <> "") Then
                            slName = slName & " " & Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID)
                        Else
                            slName = slName & " " & Trim$(tlAdf.sName)
                        End If
                    Else
                        'slName = Trim$(Str$(tlChfAdvtExt(ilLoop).lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt(ilLoop).iCntRevNo)) & " V" & Trim$(Str$(tlChfAdvtExt(ilLoop).iPropVer))
                        Select Case tlChfAdvtExt(illoop).sStatus
                            Case "W"
                                If tlChfAdvtExt(illoop).iCntRevNo > 0 Then
                                    slStr = "Rev Working"
                                Else
                                    slStr = "Working"
                                End If
                            Case "D"
                                slStr = "Rejected"
                            Case "C"
                                If tlChfAdvtExt(illoop).iCntRevNo > 0 Then
                                    slStr = "Rev Completed"
                                Else
                                    slStr = "Completed"
                                End If
                            Case "I"
                                If tlChfAdvtExt(illoop).iCntRevNo > 0 Then
                                    slStr = "Rev Unapproved"
                                Else
                                    slStr = "Unapproved"
                                End If
                            Case "G"
                                slStr = "Approved Hold"
                            Case "N"
                                slStr = "Approved Order"
                            Case "H"
                                slStr = "Hold"
                            Case "O"
                                slStr = "Order"
                        End Select
                        slName = slName & " " & slStr
                        slName = slName & " " & Trim$(tlChfAdvtExt(illoop).sProduct)
                        'Start Date Plus # weeks
                        gUnpackDate tlChfAdvtExt(illoop).iStartDate(0), tlChfAdvtExt(illoop).iStartDate(1), slStartDate
                        gUnpackDateLong tlChfAdvtExt(illoop).iEndDate(0), tlChfAdvtExt(illoop).iEndDate(1), llEndDate
                        slStr = str$((llEndDate - gDateValue(slStartDate)) \ 7 + 1)
                        slName = slName & " " & slStartDate & slStr
                        tmCxfSrchKey.lCode = tlChfAdvtExt(illoop).lCxfInt
                        If tmCxfSrchKey.lCode <> 0 Then
                            tmCxf.sComment = ""
                            imCxfRecLen = Len(tmCxf) '5027
                            ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                'If tmCxf.iStrLen > 0 Then
                                '    If tmCxf.iStrLen < 40 Then
                                '        slName = slName & " " & Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
                                '    Else
                                        'slName = slName & " " & Trim$(Left$(tmCxf.sComment, 40))
                                        slName = slName & " " & gStripChr0(Left$(tmCxf.sComment, 40))
                                '    End If
                                'End If
                            End If
                        End If
                    End If
                    slName = slName & "\" & Trim$(str$(tlChfAdvtExt(illoop).lCode))
                    'If Not gOkAddStrToListBox(slName, llLen, True) Then
                    '    ilError = True
                    '    Exit For
                    'End If
                    'lbcMster.AddItem slName
                    tlSortCode(ilSortCode).sKey = slName
                    If ilSortCode >= UBound(tlSortCode) Then
                        ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                    End If
                    ilSortCode = ilSortCode + 1
                End If
            Next illoop
            If ilError Then
                Exit For
            End If
        Next ilIndex
        'Sort then output new headers and lines
        ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - 1 > ilLowLimit Then
            ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
        End If
        btrDestroy hlAdf
        btrDestroy hlVef
        btrDestroy hmCxf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        'slNameCode = lbcMster.List(ilLoop)
        'ilRet = gParseItem(slNameCode, 1, "\", slName)
        'ilRet = gParseItem(slName, 1, "|", slCode)
        'llCntrNo = 99999999 - CLng(slCode)
        'slShow = Trim$(Str$(llCntrNo))
        'ilRet = gParseItem(slName, 2, "|", slCode)
        'ilRevNo = 999 - CLng(slCode)
        'slShow = slShow & " R" & Trim$(Str$(ilRevNo))
        'If ilShow = 1 Then
        '    'Potential
        '    ilRet = gParseItem(slName, 3, "|", slCode)
        '    If (Trim$(slCode) <> "") And (slCode <> "~") Then
        '        slShow = slShow & " " & slCode
        '    End If
        '    'Version
        '    ilRet = gParseItem(slName, 4, "|", slCode)
        '    If Trim$(slCode) <> "" Then
        '        ilVerNo = 999 - CLng(slCode)
        '        slShow = slShow & " V" & Trim$(Str$(ilVerNo))
        '    End If
        '    'Other fields
        '    ilRet = gParseItem(slName, 5, "|", slCode)
        '    slShow = slShow & " " & slCode
        'Else
        '    ilRet = gParseItem(slName, 3, "|", slCode)
        '    If Trim$(slCode) <> "" Then
        '        ilVerNo = 999 - CLng(slCode)
        '        slShow = slShow & " V" & Trim$(Str$(ilVerNo))
        '    End If
        'End If
        'lbcLocal.AddItem slShow
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 1, "|", slCode)
        llCntrNo = 99999999 - CLng(slCode)
        slShow = Trim$(str$(llCntrNo))
        ilRet = gParseItem(slName, 2, "|", slCode)
        ilRet = gParseItem(slCode, 1, "-", slExtStr)
        ilRevNo = 999 - CLng(slExtStr)
        ilRet = gParseItem(slCode, 2, "-", slExtStr)
        ilExtRevNo = 999 - CLng(slExtStr)
        ilRet = gParseItem(slName, 4, "|", slCode)
        ilVerNo = 999 - CLng(slCode)
        ilRet = gParseItem(slName, 5, "|", slCode)
        If (slCode = "W") Or (slCode = "C") Or (slCode = "I") Or (slCode = "D") Then
            If ilRevNo > 0 Then
                slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
            Else
                slShow = slShow & " V" & Trim$(str$(ilVerNo))
            End If
        Else
            slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
        End If
        If ilShow = 0 Then      'Number only
        ElseIf ilShow = 2 Then  'Number, Dates, Product, Vehicle
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 3 Then  'Number, Advertiser, Dates
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 4 Then  'Number, Dates
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 5 Then  'Number, Advertiser
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        Else                    'Number, Product, Internal comment
            'Potential
            ilRet = gParseItem(slName, 3, "|", slCode)
            If (Trim$(slCode) <> "") And (slCode <> "~") Then
                slShow = slShow & " " & slCode
            End If
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        End If
        If Not gOkAddStrToListBox(slShow, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slShow  'Add ID to list box
    Next illoop
    Exit Sub
gPopCntrBoxViaArrayErr2:
    ilRet = 1
    Resume Next
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopCntrForAASBox               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     contract number given adfcode,  *
'*                     or agency or user               *
'*                                                     *
'*******************************************************
Function gPopCntrForAASBox(frm As Form, ilAAS As Integer, ilAASCode As Integer, slStatus As String, slCntrType As String, ilCurrent As Integer, ilHOType As Integer, ilShow As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopCntrForAASBox (MainForm, ilAAS, ilAASCode, slStatus, slCntrType, ilCurrent, ilHOType, ilShow, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAAS(I)=0=Obtain Contracts for Specified Advertiser Code (ilAASCode) and User Salesperson (tgUrf(0).iSlfCode) defined as one of the contract salespersons
'                1=Obtain Contracts for Specified Agency Code (ilAASCode) and User Salesperson (tgUrf(0).iSlfCode) defined as one of the contract salespersons
'                2=Obtain Contracts for Specified Salesperson Code (ilAASCode) that matches one of the salespersons defined for the contract
'                3=Obtain Contracts for Specified Vehicle Code (ilAASCode)  and User Salesperson (tgUrf(0).iSlfCode) defined as one of the contract salespersons
'               -1=No selection by advertiser or agency or salesperson
'                Note if tgUrf(0).iSlfCode not specified, then salesperson test is bypassed
'       ilAASCode(I)- Advertiser or Agency code to obtain contracts for (-1 for all advertiser or agency)
'       slStatus (I)- chfStatus value or blank
'                         W=Working; D=Rejected; C=Completed; I=Unapproved; H=Hold; O=Order; G=Approved Hold; N=Approved Order
'                         Multiple status can be specified (WDI)
'                         If H or O (ilHOType indicates which H or O to show)
'       slCntrType (I)- chfType value or blank
'                       C=Standard; V=Reservation; T=Remnant; R=DR; Q=PI; S=PSA; M=Promo
'       ilCurrent (I)- 0=Current (Active) (chfDelete <> y); 1=Past and Current (chfDelete <> y); 2=Current(Active) plus all cancel before start (chfDelete <> y); 3=All plus history (any value for chfDelete)
'       ilHOType (I)-  1=H or O only; 2=H or O or G or N (if G or N exists show it over H or O);
'                      3=H or O or G or N or W or C or I (if G or N or W or C or I exists show it over H or O)
'                        Note: G or N can't exist at the same time as W or C or I for an order
'                              G or N or W or C or I CntrRev > 0
'                      4=H or O only and No W or C or I or N or G
'                      5=W or I only and No C or N or G (set slStatus = WC)
        'Old way
        '       ilHOType (I)- 1=Order only (ignore revision); 2=Order if no revision exist or revision only if it exist;
        '                      3=Include Revisision and Order; 4=Only Revision of orders
'       ilShow(I)-  0=Only show numbers,
'                   1=Show Number and advertiser (test site value) and product and internal comment,....
'                   2=Show Number, Dates, Product and vehicle
'                   3=Show Number, Advertiser, Dates
'                   4=Show Number, Dates
'                   5=Show Number, Advertiser
'                   6=Show Number, Status, Advertiser, Dates
'                   7=Show Number, Product
'       lbcLocal (O)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'CHF date/time stamp
    Dim hlChf As Integer        'CHF handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlChf As CHF
    Dim llCntrNo As Long
    Dim ilRevNo As Integer
    Dim ilVerNo As Integer
    Dim ilExtRevNo As Integer
    Dim slShow As String
    Dim slName As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim illoop As Integer
    Dim ilLoop1 As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim llTodayDate As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim slCode As String    'Sales source code number
    Dim hlVef As Integer        'Vef handle
    Dim tlVef As VEF
    Dim ilVefRecLen As Integer     'Record length
    Dim tlVefSrchKey As INTKEY0
    Dim hlVsf As Integer        'Vsf handle
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length
    Dim tlSrchKey As LONGKEY0
    Dim hlAdf As Integer        'Adf handle
    Dim tlAdf As ADF
    Dim ilAdfRecLen As Integer     'Record length
    Dim tlAdfSrchKey As INTKEY0
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim ilOffSet As Integer
    Dim tlVsf As VSF
    Dim llLen As Long
    Dim ilOper As Integer
    Dim slStr As String
    Dim slExtStr As String
    Dim ilSlfCode As Integer
    Dim ilVefCode As Integer
    Dim ilTestCntrNo As Integer
    Dim tlChfAdvtExt As CHFADVTEXT
    Dim slCntrStatus As String
    Dim slHOStatus As String
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    If ilCurrent = 4 Then
        lmInvStartDate = gDateValue(slSortCodeTag)
        lmInvEndDate = gDateValue(gObtainEndStd(slSortCodeTag))
    End If
    If slStatus = "" Then
        slCntrStatus = "WCIDHO"
    Else
        slCntrStatus = slStatus
    End If
    slHOStatus = ""
    If (ilHOType = 1) Or (ilHOType = 4) Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "H"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "O"
        End If
    ElseIf ilHOType = 2 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
    ElseIf ilHOType = 3 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
        If (InStr(1, slCntrStatus, "H", 1) <> 0) Or (InStr(1, slCntrStatus, "O", 1) <> 0) Then
            slHOStatus = slHOStatus & "WCI"
        End If
        If InStr(1, slCntrStatus, "H", 1) = 0 Then
            slHOStatus = slHOStatus & "G"
        End If
        If InStr(1, slCntrStatus, "O", 1) = 0 Then
            slHOStatus = slHOStatus & "N"
        End If
    ElseIf ilHOType = 5 Then
        slHOStatus = ""
        slCntrStatus = "WI"
    End If
    ilPop = True
    llLen = 0
    'ilRet = 0
    'On Error GoTo gPopCntrForAASBoxErr2
    'ilFound = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilFound = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilFound = 0
    End If
    
    slStamp = gFileDateTime(sgDBPath & "Chf.Btr") & Trim$(str$(ilAASCode)) & Trim$(slCntrStatus) & Trim$(slCntrType) & Trim$(str$(ilCurrent)) & Trim$(str$(ilHOType)) & Trim$(str$(ilShow))

    'On Error GoTo gPopCntrForAASBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopCntrForAASBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopCntrForAASBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        llTodayDate = gDateValue(gNow())
        'gObtainVehComboList
        hlChf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrOpen):" & "Chf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlChf) 'btrRecordLength(hlChf)  'Get and save record length
        hlVsf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrOpen):" & "Vsf.Btr", frm
        On Error GoTo 0
        ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length
        hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrOpen):" & "Vef.Btr", frm
        On Error GoTo 0
        ilVefRecLen = Len(tlVef) 'btrRecordLength(hlSlf)  'Get and save record length
        hlAdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrOpen):" & "Adf.Btr", frm
        On Error GoTo 0
        hmCxf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrOpen):" & "Adf.Btr", frm
        On Error GoTo 0
        ilAdfRecLen = Len(tlAdf) 'btrRecordLength(hlSlf)  'Get and save record length
        tlAdf.iCode = 0
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlChfAdvtExt)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlChf) 'Obtain number of records
        btrExtClear hlChf   'Clear any previous extend operation
        If (ilAAS = 0) And (ilAASCode > 0) Then
            tlIntTypeBuff.iType = ilAASCode
            ilRet = btrGetEqual(hlChf, tlChf, ilRecLen, tlIntTypeBuff, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Else
            ilRet = btrGetFirst(hlChf, tlChf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        End If
        If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
            ilRet = btrClose(hmCxf)
            btrDestroy hmCxf
            ilRet = btrClose(hlAdf)
            btrDestroy hlAdf
            ilRet = btrClose(hlVsf)
            btrDestroy hlVsf
            ilRet = btrClose(hlVef)
            btrDestroy hlVef
            ilRet = btrClose(hlChf)
            btrDestroy hlChf
            Exit Function
        Else
            On Error GoTo gPopCntrForAASBoxErr
            gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrGetFirst):" & "Chf.Btr", frm
            On Error GoTo 0
        End If
        Call btrExtSetBounds(hlChf, llNoRec, -1, "UC", "CHFADVTEXTPK", CHFADVTEXTPK) 'Set extract limits (all records)
        ilSlfCode = tgUrf(0).iSlfCode
        If (tgUrf(0).iGroupNo > 0) Then 'And (tgUrf(0).iSlfCode <= 0) Then
            ilRet = gObtainUrf()
            ilRet = gObtainSalesperson()
        End If
        If ilAAS = 1 Then
            If ilAASCode > 0 Then
                tlIntTypeBuff.iType = ilAASCode
                ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
                If (slCntrType = "") And (ilCurrent = 3) Then
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
                Else
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                End If
            End If
        ElseIf ilAAS = 2 Then
            If ilAASCode > 0 Then
                ilSlfCode = ilAASCode
            End If
        ElseIf ilAAS = 3 Then
            If ilAASCode > 0 Then
                ilVefCode = ilAASCode
            End If
        ElseIf ilAAS = 0 Then
            If ilAASCode > 0 Then
                tlIntTypeBuff.iType = ilAASCode
                ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
                If ilCurrent <> 3 Then
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                Else
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
                End If
            End If
        End If
        tlCharTypeBuff.sType = "Y"
        ilOffSet = gFieldOffset("Chf", "ChfDelete")
        'If selecting by advertiser- bypass slCntrType and slCntrStatus Test until get contract for speed
        'If ((slCntrStatus = "") And (slCntrType = "")) Or ((ilAAS = 0) And (ilAASCode > 0)) Then
        If (slCntrType = "") Or ((ilAAS = 0) And (ilAASCode > 0)) Then
            If ilCurrent <> 3 Then
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
            End If
        Else
            If ilCurrent <> 3 Then
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
            End If
            'If slCntrStatus <> "" Then
            '    ilOper = BTRV_EXT_OR
            '    slStr = slCntrStatus
            '    Do While slStr <> ""
            '        If Len(slStr) = 1 Then
            '            If slCntrType <> "" Then
            '                ilOper = BTRV_EXT_AND
            '            Else
            '                ilOper = BTRV_EXT_LAST_TERM
            '            End If
            '        End If
            '        tlCharTypeBuff.sType = Left$(slStr, 1)
            '        ilOffset = gFieldOffset("Chf", "ChfStatus")
            '        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
            '        slStr = Mid$(slStr, 2)
            '    Loop
            'End If
            If slCntrType <> "" Then
                ilOper = BTRV_EXT_OR
                slStr = slCntrType
                Do While slStr <> ""
                    If Len(slStr) = 1 Then
                        ilOper = BTRV_EXT_LAST_TERM
                    End If
                    tlCharTypeBuff.sType = Left$(slStr, 1)
                    ilOffSet = gFieldOffset("Chf", "ChfType")
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
                    slStr = Mid$(slStr, 2)
                Loop
            End If
        End If
        ilOffSet = gFieldOffset("Chf", "ChfCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract iCode field
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfCntrNo")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract Contract number
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfExtRevNo")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfCntRevNo")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfType")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfProduct")
        ilRet = btrExtAddField(hlChf, ilOffSet, 35) 'Extract Product
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfSlfCode1")
        ilRet = btrExtAddField(hlChf, ilOffSet, 20) 'Extract salesperson code
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfMnfDemo1")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract salesperson code
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfCxfInt")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfPropVer")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract end date
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfStatus")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfMnfPotnType")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract SellNet
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfStartDate")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfEndDate")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract end date
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfVefCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfSifCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0

         '8-21-05 add pct of trade to array
        ilOffSet = gFieldOffset("Chf", "ChfPctTrade")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'pct trade
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        
        '7/12/10
        ilOffSet = gFieldOffset("Chf", "ChfCBSOrder")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        
        '2/24/12
        ilOffSet = gFieldOffset("Chf", "ChfBillCycle")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0

        '2/17/18
        ilOffSet = gFieldOffset("Chf", "ChfSource")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0

        ReDim lmWCINGCntrNo(0 To 0) As Long     'Used to filter out Holds that have W or C or I or G or N
        'ilRet = btrExtGetNextExt(hlChf)    'Extract record
        ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopCntrForAASBoxErr
            gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrExtGetNextExt):" & "Chf.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlChfAdvtExt)  'Extract operation record size
            'ilRet = btrExtGetFirst(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilFound = True
                If (ilAAS = 3) And (ilVefCode <> -1) Then   '-1 = All Vehicles
                    ilFound = False
                    If tlChfAdvtExt.lVefCode > 0 Then
                        If tlChfAdvtExt.lVefCode = ilVefCode Then
                            ilFound = True
                        End If
                    ElseIf tlChfAdvtExt.lVefCode < 0 Then
                        tlSrchKey.lCode = -tlChfAdvtExt.lVefCode
                        ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        On Error GoTo gPopCntrForAASBoxErr
                        gBtrvErrorMsg ilRet, "gPopCntrForAASBox (btrGetEqual): Vsf.Btr", frm
                        On Error GoTo 0
                        Do While ilRet = BTRV_ERR_NONE
                            For illoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                                If tmVsf.iFSCode(illoop) > 0 Then
                                    If tmVsf.iFSCode(illoop) = ilVefCode Then
                                        ilFound = True
                                        Exit For
                                    End If
                                    If tmVsf.iFSCode(illoop) <> tlVef.iCode Then
                                        tlVefSrchKey.iCode = tmVsf.iFSCode(illoop)
                                        ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        On Error GoTo gPopCntrForAASBoxErr
                                        gCPErrorMsg ilRet, "gPopCntrForAASBox (btrGetEqual: Vef)", frm
                                        On Error GoTo 0
                                    End If
                                    If tlVef.sType = "V" Then
                                        If tlVef.iCode <> ilVefCode Then
                                            tlSrchKey.lCode = tlVef.lVsfCode
                                            ilRet = btrGetEqual(hlVsf, tlVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            For ilLoop1 = LBound(tlVsf.iFSCode) To UBound(tlVsf.iFSCode) Step 1
                                                If tlVsf.iFSCode(ilLoop1) > 0 Then
                                                    If tlVsf.iFSCode(ilLoop1) = ilVefCode Then
                                                        ilFound = True
                                                        Exit For
                                                    End If
                                                End If
                                            Next ilLoop1
                                        Else
                                            ilFound = True
                                        End If
                                    End If
                                End If
                            Next illoop
                            If ilFound Then
                                Exit Do
                            End If
                            If tmVsf.lLkVsfCode <= 0 Then
                                Exit Do
                            End If
                            tlSrchKey.lCode = tmVsf.lLkVsfCode
                            ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    Else
                        ilFound = True  'all vehicles
                    End If
                End If
                If ilHOType = 5 Then
                    ilTestCntrNo = True
                Else
                    ilTestCntrNo = False
                End If
                'For Proposal CntRevNo = 0; For Orders CntRevNo >= 0 (for W, C, I CntRevNo > 0)
                If (tlChfAdvtExt.iCntRevNo = 0) And ((tlChfAdvtExt.sStatus <> "H") And (tlChfAdvtExt.sStatus <> "O") And (tlChfAdvtExt.sStatus <> "G") And (tlChfAdvtExt.sStatus <> "N")) Then  'Proposal
                    'Proposal only
                    If (InStr(1, slCntrStatus, tlChfAdvtExt.sStatus) = 0) Then
                        ilFound = False
                    End If
                Else    'Order only
                    If ilHOType = 5 Then
                        If (InStr(1, slCntrStatus, tlChfAdvtExt.sStatus) = 0) Then
                            ilFound = False
                        End If
                    Else
                        If (InStr(1, slHOStatus, tlChfAdvtExt.sStatus) <> 0) Then
                            If (ilHOType = 2) Or (ilHOType = 3) Then
                                ilTestCntrNo = True
                            End If
                        Else
                            ilFound = False
                        End If
                    End If
                End If
                If slCntrType <> "" Then
                    If InStr(1, slCntrType, tlChfAdvtExt.sType) = 0 Then
                        ilFound = False
                    End If
                End If
                If ilFound Then
                    '2/17/18: Allow Programmatic buys
                    tlChfAdvtExt.sSource = ""
                    ilFound = gTestChfAdvtExt(frm, ilSlfCode, tlChfAdvtExt, hlVsf, ilCurrent)
                End If
                If (Not ilFound) And (ilHOType = 4) And (InStr(1, "WCING", tlChfAdvtExt.sStatus) <> 0) Then
                    lmWCINGCntrNo(UBound(lmWCINGCntrNo)) = tlChfAdvtExt.lCntrNo
                    ReDim Preserve lmWCINGCntrNo(0 To UBound(lmWCINGCntrNo) + 1) As Long   'Used to filter out Holds that have W or C or I or G or N
                End If
                If (Not ilFound) And (ilHOType = 5) And (InStr(1, "CNG", tlChfAdvtExt.sStatus) <> 0) Then
                    lmWCINGCntrNo(UBound(lmWCINGCntrNo)) = tlChfAdvtExt.lCntrNo
                    ReDim Preserve lmWCINGCntrNo(0 To UBound(lmWCINGCntrNo) + 1) As Long   'Used to filter out Holds that have W or C or I or G or N
                End If
                If ilFound Then
                    slStr = Trim$(str$(99999999 - tlChfAdvtExt.lCntrNo))
                    Do While Len(slStr) < 8
                        slStr = "0" & slStr
                    Loop
                    slName = slStr
                    slStr = Trim$(str$(999 - tlChfAdvtExt.iCntRevNo))
                    Do While Len(slStr) < 3
                        slStr = "0" & slStr
                    Loop
                    slExtStr = Trim$(str$(999 - tlChfAdvtExt.iExtRevNo))
                    Do While Len(slExtStr) < 3
                        slExtStr = "0" & slExtStr
                    Loop
                    slName = slName & "|" & slStr & "-" & slExtStr & "|"
                    If (tlChfAdvtExt.sStatus = "W") Or (tlChfAdvtExt.sStatus = "C") Or (tlChfAdvtExt.sStatus = "I") Then
                        'Add Potential
                        If tlChfAdvtExt.iMnfPotnType > 0 Then
                            slStr = " "
                        Else
                            slStr = "~"
                        End If
                        slName = slName & slStr & "|"
                    Else
                        slName = slName & " |"
                    End If
                    slStr = Trim$(str$(999 - tlChfAdvtExt.iPropVer))
                    Do While Len(slStr) < 3
                        slStr = "0" & slStr
                    Loop
                    slName = slName & slStr & "|"
                    slName = slName & tlChfAdvtExt.sStatus & "|"
                    If ilShow = 0 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo)) & " V" & Trim$(Str$(tlChfAdvtExt.iPropVer))
                    ElseIf ilShow = 2 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                        gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                        gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slEndDate
                        slName = slName & ": " & slStartDate & "-" & slEndDate
                        slName = slName & " " & Trim$(tlChfAdvtExt.sProduct)
                        If tlChfAdvtExt.lVefCode > 0 Then
                            If tlChfAdvtExt.lVefCode <> tlVef.iCode Then
                                tlVefSrchKey.iCode = tlChfAdvtExt.lVefCode
                                ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo gPopCntrForAASBoxErr
                                gCPErrorMsg ilRet, "gPopCntrForAASBox (btrGetEqual: Vef)", frm
                                On Error GoTo 0
                            End If
                            slName = slName & " " & Trim$(tlVef.sName)
                        End If
                    ElseIf ilShow = 3 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                        If tlChfAdvtExt.iAdfCode <> tlAdf.iCode Then
                            tlAdfSrchKey.iCode = tlChfAdvtExt.iAdfCode
                            ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo gPopCntrForAASBoxErr
                            gCPErrorMsg ilRet, "gPopCntrForAASBox (btrGetEqual: Adf)", frm
                            On Error GoTo 0
                        End If
                        'slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                        If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sAddrID) <> "") Then
                            slName = slName & " " & Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID)
                        Else
                            slName = slName & " " & Trim$(tlAdf.sName)
                        End If
                        gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                        gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slEndDate
                        slName = slName & " " & slStartDate & "-" & slEndDate
                    ElseIf ilShow = 4 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                        gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                        gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slEndDate
                        slName = slName & " " & slStartDate & "-" & slEndDate
                    ElseIf ilShow = 5 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                        If tlChfAdvtExt.iAdfCode <> tlAdf.iCode Then
                            tlAdfSrchKey.iCode = tlChfAdvtExt.iAdfCode
                            ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo gPopCntrForAASBoxErr
                            gCPErrorMsg ilRet, "gPopCntrForAASBox (btrGetEqual: Adf)", frm
                            On Error GoTo 0
                        End If
                        'slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                        If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sAddrID) <> "") Then
                            slName = slName & " " & Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID)
                        Else
                            slName = slName & " " & Trim$(tlAdf.sName)
                        End If
                    ElseIf ilShow = 7 Then
                        slName = slName & " " & Trim$(tlChfAdvtExt.sProduct)
                    Else
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo)) & " V" & Trim$(Str$(tlChfAdvtExt.iPropVer))
                        Select Case tlChfAdvtExt.sStatus
                            Case "W"
                                If tlChfAdvtExt.iCntRevNo > 0 Then
                                    slStr = "Rev Working"
                                Else
                                    slStr = "Working"
                                End If
                            Case "D"
                                slStr = "Rejected"
                            Case "C"
                                If tlChfAdvtExt.iCntRevNo > 0 Then
                                    slStr = "Rev Completed"
                                Else
                                    slStr = "Completed"
                                End If
                            Case "I"
                                If tlChfAdvtExt.iCntRevNo > 0 Then
                                    slStr = "Rev Unapproved"
                                Else
                                    slStr = "Unapproved"
                                End If
                            Case "G"
                                slStr = "Approved Hold"
                            Case "N"
                                slStr = "Approved Order"
                            Case "H"
                                slStr = "Hold"
                            Case "O"
                                slStr = "Order"
                        End Select
                        slName = slName & " " & slStr
                        If ilShow = 6 Then
                            If tlChfAdvtExt.iAdfCode <> tlAdf.iCode Then
                                tlAdfSrchKey.iCode = tlChfAdvtExt.iAdfCode
                                ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo gPopCntrForAASBoxErr
                                gCPErrorMsg ilRet, "gPopCntrForAASBox (btrGetEqual: Adf)", frm
                                On Error GoTo 0
                            End If
                            'slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                            If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sAddrID) <> "") Then
                                slName = slName & " " & Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID)
                            Else
                                slName = slName & " " & Trim$(tlAdf.sName)
                            End If
                        Else
                            slName = slName & " " & Trim$(tlChfAdvtExt.sProduct)
                        End If
                        'Start Date Plus # weeks
                        gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                        gUnpackDateLong tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), llEndDate
                        If ilShow = 6 Then
                            slName = slName & " " & slStartDate & "-" & Format$(llEndDate, "m/d/yy")
                        Else
                            slStr = str$((llEndDate - gDateValue(slStartDate)) \ 7 + 1)
                            slName = slName & " " & slStartDate & slStr
                            tmCxfSrchKey.lCode = tlChfAdvtExt.lCxfInt
                            If tmCxfSrchKey.lCode <> 0 Then
                                tmCxf.sComment = ""
                                imCxfRecLen = Len(tmCxf) '5027
                                ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    'If tmCxf.iStrLen > 0 Then
                                    '    If tmCxf.iStrLen < 40 Then
                                    '        slName = slName & " " & Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
                                    '    Else
                                            'slName = slName & " " & Trim$(Left$(tmCxf.sComment, 40))
                                            slName = slName & " " & gStripChr0(Left$(tmCxf.sComment, 40))
                                    '    End If
                                    'End If
                                End If
                            End If
                        End If
                    End If
                    ilFound = False
                    If ilTestCntrNo Then
                        For illoop = ilLowLimit To ilSortCode - 1 Step 1
                            slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 1, "\", slNameCode)
                            ilRet = gParseItem(slNameCode, 1, "|", slCode)
                            llCntrNo = 99999999 - CLng(slCode)
                            If llCntrNo = tlChfAdvtExt.lCntrNo Then
                                ilRet = gParseItem(slNameCode, 2, "|", slCode)
                                ilRet = gParseItem(slCode, 1, "-", slExtStr)
                                ilRevNo = 999 - CLng(slExtStr)
                                If tlChfAdvtExt.iCntRevNo > ilRevNo Then
                                    'Replace
                                    'lbcMster.RemoveItem ilLoop
                                    'llLen = llLen - Len(slNameCode)
                                    ilFound = True
                                    slName = slName & "\" & Trim$(str$(tlChfAdvtExt.lCode))
                                    tlSortCode(illoop).sKey = slName
                                Else
                                    'Leave
                                    ilFound = True
                                End If
                                Exit For
                            End If
                        Next illoop
                    End If
                    If Not ilFound Then
                        slName = slName & "\" & Trim$(str$(tlChfAdvtExt.lCode))
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit Do
                        'End If
                        'lbcMster.AddItem slName
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                    End If
                End If
                ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If (ilHOType = 4) Or (ilHOType = 5) Then
                For ilLoop1 = 0 To UBound(lmWCINGCntrNo) - 1 Step 1
                    For illoop = UBound(tlSortCode) - 1 To ilLowLimit Step -1
                        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 1, "\", slName)
                        ilRet = gParseItem(slName, 1, "|", slCode)
                        llCntrNo = 99999999 - CLng(slCode)
                        If lmWCINGCntrNo(ilLoop1) = llCntrNo Then
                            For ilIndex = illoop To UBound(tlSortCode) - 1 Step 1
                                tlSortCode(ilIndex) = tlSortCode(ilIndex + 1)
                            Next ilIndex
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) - 1) As SORTCODE
                        End If
                    Next illoop
                Next ilLoop1
            End If
            Erase lmWCINGCntrNo
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If
        ilRet = btrClose(hmCxf)
        btrDestroy hmCxf
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        ilRet = btrClose(hlVsf)
        btrDestroy hlVsf
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        ilRet = btrClose(hlChf)
        btrDestroy hlChf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 1, "|", slCode)
        llCntrNo = 99999999 - CLng(slCode)
        slShow = Trim$(str$(llCntrNo))
        ilRet = gParseItem(slName, 2, "|", slCode)
        ilRet = gParseItem(slCode, 1, "-", slExtStr)
        ilRevNo = 999 - CLng(slExtStr)
        ilRet = gParseItem(slCode, 2, "-", slExtStr)
        ilExtRevNo = 999 - CLng(slExtStr)
        ilRet = gParseItem(slName, 4, "|", slCode)
        ilVerNo = 999 - CLng(slCode)
        ilRet = gParseItem(slName, 5, "|", slCode)
        If (slCode = "W") Or (slCode = "C") Or (slCode = "I") Or (slCode = "D") Then
            If ilRevNo > 0 Then
                slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
            Else
                slShow = slShow & " V" & Trim$(str$(ilVerNo))
            End If
        Else
            slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
        End If
        If ilShow = 0 Then      'Number only
        ElseIf ilShow = 2 Then  'Number, Dates, Product, Vehicle
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 3 Then  'Number, Advertiser, Dates
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 4 Then  'Number, Dates
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 5 Then  'Number, Advertiser
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 7 Then  'Number, Product
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        Else                    'Number, Product, Internal comment
            'Potential
            ilRet = gParseItem(slName, 3, "|", slCode)
            If (Trim$(slCode) <> "") And (slCode <> "~") Then
                slShow = slShow & " " & slCode
            End If
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        End If
        If Not gOkAddStrToListBox(slShow, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slShow  'Add ID to list box
    Next illoop
    Exit Function
gPopCntrForAASBoxErr:
    ilRet = btrClose(hmCxf)
    btrDestroy hmCxf
    ilRet = btrClose(hlAdf)
    btrDestroy hlAdf
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlChf)
    btrDestroy hlChf
    gDbg_HandleError "PopSubs: gPopCntrForAASBox"
'    gPopCntrForAASBox = CP_MSG_NOSHOW
'    Exit Function
gPopCntrForAASBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopCntrForAASBoxByMkt          *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     contract number given adfcode,  *
'*                     or agency or user               *
'*                     Taken from gPopCntrForAASBox    *
'*                                                     *
'*******************************************************
Function gPopCntrForAASBoxByMkt(frm As Form, ilAAS As Integer, ilAASCode As Integer, slStatus As String, slCntrType As String, ilCurrent As Integer, ilHOType As Integer, ilShow As Integer, ilMktCode() As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopCntrForAASBoxByMkt (MainForm, ilAAS, ilAASCode, slStatus, slCntrType, ilCurrent, ilHOType, ilShow, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAAS(I)=0=Obtain Contracts for Specified Advertiser Code (ilAASCode) and User Salesperson (tgUrf(0).iSlfCode) defined as one of the contract salespersons
'                1=Obtain Contracts for Specified Agency Code (ilAASCode) and User Salesperson (tgUrf(0).iSlfCode) defined as one of the contract salespersons
'                2=Obtain Contracts for Specified Salesperson Code (ilAASCode) that matches one of the salespersons defined for the contract
'                3=Obtain Contracts for Specified Vehicle Code (ilAASCode)  and User Salesperson (tgUrf(0).iSlfCode) defined as one of the contract salespersons
'               -1=No selection by advertiser or agency or salesperson
'                Note if tgUrf(0).iSlfCode not specified, then salesperson test is bypassed
'       ilAASCode(I)- Advertiser or Agency code to obtain contracts for (-1 for all advertiser or agency)
'       slStatus (I)- chfStatus value or blank
'                         W=Working; D=Rejected; C=Completed; I=Unapproved; H=Hold; O=Order; G=Approved Hold; N=Approved Order
'                         Multiple status can be specified (WDI)
'                         If H or O (ilHOType indicates which H or O to show)
'       slCntrType (I)- chfType value or blank
'                       C=Standard; V=Reservation; T=Remnant; R=DR; Q=PI; S=PSA; M=Promo
'       ilCurrent (I)- 0=Current (Active) (chfDelete <> y); 1=Past and Current (chfDelete <> y); 2=Current(Active) plus all cancel before start (chfDelete <> y); 3=All plus history (any value for chfDelete)
'       ilHOType (I)-  1=H or O only; 2=H or O or G or N (if G or N exists show it over H or O);
'                      3=H or O or G or N or W or C or I (if G or N or W or C or I exists show it over H or O)
'                        Note: G or N can't exist at the same time as W or C or I for an order
'                              G or N or W or C or I CntrRev > 0
'                      4=H or O only and No W or C or I or N or G
'                      5=W or I only and No C or N or G (set slStatus = WC)
        'Old way
        '       ilHOType (I)- 1=Order only (ignore revision); 2=Order if no revision exist or revision only if it exist;
        '                      3=Include Revisision and Order; 4=Only Revision of orders
'       ilShow(I)-  0=Only show numbers,
'                   1=Show Number and advertiser (test site value) and product and internal comment,....
'                   2=Show Number, Dates, Product and vehicle
'                   3=Show Number, Advertiser, Dates
'                   4=Show Number, Dates
'                   5=Show Number, Advertiser
'                   6=Show Number, Status, Advertiser, Dates
'       lbcLocal (O)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'CHF date/time stamp
    Dim hlChf As Integer        'CHF handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlChf As CHF
    Dim llCntrNo As Long
    Dim ilRevNo As Integer
    Dim ilVerNo As Integer
    Dim ilExtRevNo As Integer
    Dim slShow As String
    Dim slName As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim illoop As Integer
    Dim ilLoop1 As Integer
    Dim ilIndex As Integer
    Dim slNameCode As String
    Dim llTodayDate As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim slCode As String    'Sales source code number
    Dim ilPos As Integer
    Dim hlVef As Integer        'Vef handle
    Dim tlVef As VEF
    Dim ilVefRecLen As Integer     'Record length
    Dim tlVefSrchKey As INTKEY0
    Dim hlVsf As Integer        'Vsf handle
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length
    Dim tlSrchKey As LONGKEY0
    Dim hlAdf As Integer        'Adf handle
    Dim tlAdf As ADF
    Dim ilAdfRecLen As Integer     'Record length
    Dim tlAdfSrchKey As INTKEY0
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim ilOffSet As Integer
    Dim tlVsf As VSF
    Dim llLen As Long
    Dim ilOper As Integer
    Dim slStr As String
    Dim slExtStr As String
    Dim ilSlfCode As Integer
    Dim ilVefCode As Integer
    Dim ilTestCntrNo As Integer
    Dim tlChfAdvtExt As CHFADVTEXT
    Dim slCntrStatus As String
    Dim slHOStatus As String
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    If slStatus = "" Then
        slCntrStatus = "WCIDHO"
    Else
        slCntrStatus = slStatus
    End If
    slHOStatus = ""
    If (ilHOType = 1) Or (ilHOType = 4) Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "H"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "O"
        End If
    ElseIf ilHOType = 2 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
    ElseIf ilHOType = 3 Then
        If InStr(1, slCntrStatus, "H", 1) <> 0 Then
            slHOStatus = slHOStatus & "GH"
        End If
        If InStr(1, slCntrStatus, "O", 1) <> 0 Then
            slHOStatus = slHOStatus & "NO"
        End If
        If (InStr(1, slCntrStatus, "H", 1) <> 0) Or (InStr(1, slCntrStatus, "O", 1) <> 0) Then
            slHOStatus = slHOStatus & "WCI"
        End If
        If InStr(1, slCntrStatus, "H", 1) = 0 Then
            slHOStatus = slHOStatus & "G"
        End If
        If InStr(1, slCntrStatus, "O", 1) = 0 Then
            slHOStatus = slHOStatus & "N"
        End If
    ElseIf ilHOType = 5 Then
        slHOStatus = ""
        slCntrStatus = "WI"
    End If
    ilPop = True
    llLen = 0
    'ilRet = 0
    'On Error GoTo gPopCntrForAASBoxByMktErr2
    'ilFound = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilFound = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilFound = 0
    End If

    
    slStamp = gFileDateTime(sgDBPath & "Chf.Btr") & Trim$(str$(ilAASCode)) & Trim$(slCntrStatus) & Trim$(slCntrType) & Trim$(str$(ilCurrent)) & Trim$(str$(ilHOType)) & Trim$(str$(ilShow))

    'On Error GoTo gPopCntrForAASBoxByMktErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopCntrForAASBoxByMkt = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopCntrForAASBoxByMkt = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        llTodayDate = gDateValue(gNow())
        'gObtainVehComboList
        hlChf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrOpen):" & "Chf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlChf) 'btrRecordLength(hlChf)  'Get and save record length
        hlVsf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrOpen):" & "Vsf.Btr", frm
        On Error GoTo 0
        ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length
        hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrOpen):" & "Vef.Btr", frm
        On Error GoTo 0
        ilVefRecLen = Len(tlVef) 'btrRecordLength(hlSlf)  'Get and save record length
        hlAdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrOpen):" & "Adf.Btr", frm
        On Error GoTo 0
        hmCxf = CBtrvTable(ONEHANDLE)
        ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrOpen):" & "Adf.Btr", frm
        On Error GoTo 0
        ilAdfRecLen = Len(tlAdf) 'btrRecordLength(hlSlf)  'Get and save record length
        tlAdf.iCode = 0
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlChfAdvtExt)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlChf) 'Obtain number of records
        btrExtClear hlChf   'Clear any previous extend operation
        If (ilAAS = 0) And (ilAASCode > 0) Then
            tlIntTypeBuff.iType = ilAASCode
            ilRet = btrGetEqual(hlChf, tlChf, ilRecLen, tlIntTypeBuff, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Else
            ilRet = btrGetFirst(hlChf, tlChf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        End If
        If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
            ilRet = btrClose(hmCxf)
            btrDestroy hmCxf
            ilRet = btrClose(hlAdf)
            btrDestroy hlAdf
            ilRet = btrClose(hlVsf)
            btrDestroy hlVsf
            ilRet = btrClose(hlVef)
            btrDestroy hlVef
            ilRet = btrClose(hlChf)
            btrDestroy hlChf
            Exit Function
        Else
            On Error GoTo gPopCntrForAASBoxByMktErr
            gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrGetFirst):" & "Chf.Btr", frm
            On Error GoTo 0
        End If
        Call btrExtSetBounds(hlChf, llNoRec, -1, "UC", "CHFADVTEXTPK", CHFADVTEXTPK) 'Set extract limits (all records)
        ilSlfCode = tgUrf(0).iSlfCode
        If (tgUrf(0).iGroupNo > 0) Then 'And (tgUrf(0).iSlfCode <= 0) Then
            ilRet = gObtainUrf()
            ilRet = gObtainSalesperson()
        End If
        If ilAAS = 1 Then
            If ilAASCode > 0 Then
                tlIntTypeBuff.iType = ilAASCode
                ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
                If (slCntrType = "") And (ilCurrent = 3) Then
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
                Else
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                End If
            End If
        ElseIf ilAAS = 2 Then
            If ilAASCode > 0 Then
                ilSlfCode = ilAASCode
            End If
        ElseIf ilAAS = 3 Then
            If ilAASCode > 0 Then
                ilVefCode = ilAASCode
            End If
        ElseIf ilAAS = 0 Then
            If ilAASCode > 0 Then
                tlIntTypeBuff.iType = ilAASCode
                ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
                If ilCurrent <> 3 Then
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                Else
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
                End If
            End If
        End If
        tlCharTypeBuff.sType = "Y"
        ilOffSet = gFieldOffset("Chf", "ChfDelete")
        'If selecting by advertiser- bypass slCntrType and slCntrStatus Test until get contract for speed
        'If ((slCntrStatus = "") And (slCntrType = "")) Or ((ilAAS = 0) And (ilAASCode > 0)) Then
        If (slCntrType = "") Or ((ilAAS = 0) And (ilAASCode > 0)) Then
            If ilCurrent <> 3 Then
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
            End If
        Else
            If ilCurrent <> 3 Then
                ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
            End If
            'If slCntrStatus <> "" Then
            '    ilOper = BTRV_EXT_OR
            '    slStr = slCntrStatus
            '    Do While slStr <> ""
            '        If Len(slStr) = 1 Then
            '            If slCntrType <> "" Then
            '                ilOper = BTRV_EXT_AND
            '            Else
            '                ilOper = BTRV_EXT_LAST_TERM
            '            End If
            '        End If
            '        tlCharTypeBuff.sType = Left$(slStr, 1)
            '        ilOffset = gFieldOffset("Chf", "ChfStatus")
            '        ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
            '        slStr = Mid$(slStr, 2)
            '    Loop
            'End If
            If slCntrType <> "" Then
                ilOper = BTRV_EXT_OR
                slStr = slCntrType
                Do While slStr <> ""
                    If Len(slStr) = 1 Then
                        ilOper = BTRV_EXT_LAST_TERM
                    End If
                    tlCharTypeBuff.sType = Left$(slStr, 1)
                    ilOffSet = gFieldOffset("Chf", "ChfType")
                    ilRet = btrExtAddLogicConst(hlChf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, ilOper, tlCharTypeBuff, 1)
                    slStr = Mid$(slStr, 2)
                Loop
            End If
        End If
        ilOffSet = gFieldOffset("Chf", "ChfCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract iCode field
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfCntrNo")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4)  'Extract Contract number
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfExtRevNo")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfCntRevNo")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract start date
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfType")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfAdfCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfProduct")
        ilRet = btrExtAddField(hlChf, ilOffSet, 35) 'Extract Product
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfAgfCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract advertiser code
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfSlfCode1")
        ilRet = btrExtAddField(hlChf, ilOffSet, 20) 'Extract salesperson code
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfMnfDemo1")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract salesperson code
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfCxfInt")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfPropVer")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract end date
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfStatus")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfMnfPotnType")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'Extract SellNet
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfStartDate")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract start date
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfEndDate")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract end date
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfVefCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Chf", "ChfSifCode")
        ilRet = btrExtAddField(hlChf, ilOffSet, 4) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0

         '8-21-05 add pct of trade to array
        ilOffSet = gFieldOffset("Chf", "ChfPctTrade")
        ilRet = btrExtAddField(hlChf, ilOffSet, 2) 'pct trade
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        
        '7/12/10
        ilOffSet = gFieldOffset("Chf", "ChfCBSOrder")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0

        '2/24/12
        ilOffSet = gFieldOffset("Chf", "ChfBillCycle")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0

        '2/17/18
        ilOffSet = gFieldOffset("Chf", "ChfSource")
        ilRet = btrExtAddField(hlChf, ilOffSet, 1) 'Extract Vehicle
        On Error GoTo gPopCntrForAASBoxByMktErr
        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtAddField):" & "Chf.Btr", frm
        On Error GoTo 0
        
        ReDim lmWCINGCntrNo(0 To 0) As Long     'Used to filter out Holds that have W or C or I or G or N
        'ilRet = btrExtGetNextExt(hlChf)    'Extract record
        ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopCntrForAASBoxByMktErr
            gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrExtGetNextExt):" & "Chf.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlChfAdvtExt)  'Extract operation record size
            'ilRet = btrExtGetFirst(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilFound = True
                If (ilAAS = 3) And (ilVefCode <> -1) Then   '-1 = All Vehicles
                    ilFound = False
                    If tlChfAdvtExt.lVefCode > 0 Then
                        If tlChfAdvtExt.lVefCode = ilVefCode Then
                            ilFound = True
                        End If
                    ElseIf tlChfAdvtExt.lVefCode < 0 Then
                        tlSrchKey.lCode = -tlChfAdvtExt.lVefCode
                        ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        On Error GoTo gPopCntrForAASBoxByMktErr
                        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrGetEqual): Vsf.Btr", frm
                        On Error GoTo 0
                        Do While ilRet = BTRV_ERR_NONE
                            For illoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                                If tmVsf.iFSCode(illoop) > 0 Then
                                    If tmVsf.iFSCode(illoop) = ilVefCode Then
                                        ilFound = True
                                        Exit For
                                    End If
                                    If tmVsf.iFSCode(illoop) <> tlVef.iCode Then
                                        tlVefSrchKey.iCode = tmVsf.iFSCode(illoop)
                                        ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        On Error GoTo gPopCntrForAASBoxByMktErr
                                        gCPErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrGetEqual: Vef)", frm
                                        On Error GoTo 0
                                    End If
                                    If tlVef.sType = "V" Then
                                        If tlVef.iCode <> ilVefCode Then
                                            tlSrchKey.lCode = tlVef.lVsfCode
                                            ilRet = btrGetEqual(hlVsf, tlVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                            For ilLoop1 = LBound(tlVsf.iFSCode) To UBound(tlVsf.iFSCode) Step 1
                                                If tlVsf.iFSCode(ilLoop1) > 0 Then
                                                    If tlVsf.iFSCode(ilLoop1) = ilVefCode Then
                                                        ilFound = True
                                                        Exit For
                                                    End If
                                                End If
                                            Next ilLoop1
                                        Else
                                            ilFound = True
                                        End If
                                    End If
                                End If
                            Next illoop
                            If ilFound Then
                                Exit Do
                            End If
                            If tmVsf.lLkVsfCode <= 0 Then
                                Exit Do
                            End If
                            tlSrchKey.lCode = tmVsf.lLkVsfCode
                            ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    Else
                        ilFound = True  'all vehicles
                    End If
                End If
                If (ilFound) And (tgSpf.sMktBase = "Y") Then
                    ilFound = False
                    If tlChfAdvtExt.lVefCode > 0 Then
                        'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If tlChfAdvtExt.lVefCode = tgMVef(ilIndex).iCode Then
                            ilIndex = gBinarySearchVef(CInt(tlChfAdvtExt.lVefCode))
                            If ilIndex <> -1 Then
                                For ilLoop1 = 0 To UBound(ilMktCode) - 1 Step 1
                                    If tgMVef(ilIndex).iMnfVehGp3Mkt = ilMktCode(ilLoop1) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilLoop1
                                'Exit For
                            End If
                        'Next ilIndex
                    ElseIf tlChfAdvtExt.lVefCode < 0 Then
                        tlSrchKey.lCode = -tlChfAdvtExt.lVefCode
                        ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        On Error GoTo gPopCntrForAASBoxByMktErr
                        gBtrvErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrGetEqual): Vsf.Btr", frm
                        On Error GoTo 0
                        Do While ilRet = BTRV_ERR_NONE
                            For illoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                                If tmVsf.iFSCode(illoop) > 0 Then
                                    'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                    '    If tmVsf.iFSCode(ilLoop) = tgMVef(ilIndex).iCode Then
                                        ilIndex = gBinarySearchVef(tmVsf.iFSCode(illoop))
                                        If ilIndex <> -1 Then
                                            For ilLoop1 = 0 To UBound(ilMktCode) - 1 Step 1
                                                If tgMVef(ilIndex).iMnfVehGp3Mkt = ilMktCode(ilLoop1) Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilLoop1
                                            'Exit For
                                        End If
                                    'Next ilIndex
                                    If ilFound Then
                                        Exit For
                                    End If
                                End If
                            Next illoop
                            If ilFound Then
                                Exit Do
                            End If
                            If tmVsf.lLkVsfCode <= 0 Then
                                Exit Do
                            End If
                            tlSrchKey.lCode = tmVsf.lLkVsfCode
                            ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    Else
                        ilFound = True  'all vehicles
                    End If
                End If
                If ilHOType = 5 Then
                    ilTestCntrNo = True
                Else
                    ilTestCntrNo = False
                End If
                'For Proposal CntRevNo = 0; For Orders CntRevNo >= 0 (for W, C, I CntRevNo > 0)
                If (tlChfAdvtExt.iCntRevNo = 0) And ((tlChfAdvtExt.sStatus <> "H") And (tlChfAdvtExt.sStatus <> "O") And (tlChfAdvtExt.sStatus <> "G") And (tlChfAdvtExt.sStatus <> "N")) Then  'Proposal
                    'Proposal only
                    If (InStr(1, slCntrStatus, tlChfAdvtExt.sStatus) = 0) Then
                        ilFound = False
                    End If
                Else    'Order only
                    If ilHOType = 5 Then
                        If (InStr(1, slCntrStatus, tlChfAdvtExt.sStatus) = 0) Then
                            ilFound = False
                        End If
                    Else
                        If (InStr(1, slHOStatus, tlChfAdvtExt.sStatus) <> 0) Then
                            If (ilHOType = 2) Or (ilHOType = 3) Then
                                ilTestCntrNo = True
                            End If
                        Else
                            ilFound = False
                        End If
                    End If
                End If
                If slCntrType <> "" Then
                    If InStr(1, slCntrType, tlChfAdvtExt.sType) = 0 Then
                        ilFound = False
                    End If
                End If
                If ilFound Then
                    ilFound = gTestChfAdvtExt(frm, ilSlfCode, tlChfAdvtExt, hlVsf, ilCurrent)
                End If
                If (Not ilFound) And (ilHOType = 4) And (InStr(1, "WCING", tlChfAdvtExt.sStatus) <> 0) Then
                    lmWCINGCntrNo(UBound(lmWCINGCntrNo)) = tlChfAdvtExt.lCntrNo
                    ReDim Preserve lmWCINGCntrNo(0 To UBound(lmWCINGCntrNo) + 1) As Long   'Used to filter out Holds that have W or C or I or G or N
                End If
                If (Not ilFound) And (ilHOType = 5) And (InStr(1, "CNG", tlChfAdvtExt.sStatus) <> 0) Then
                    lmWCINGCntrNo(UBound(lmWCINGCntrNo)) = tlChfAdvtExt.lCntrNo
                    ReDim Preserve lmWCINGCntrNo(0 To UBound(lmWCINGCntrNo) + 1) As Long   'Used to filter out Holds that have W or C or I or G or N
                End If
                If ilFound Then
                    slStr = Trim$(str$(99999999 - tlChfAdvtExt.lCntrNo))
                    Do While Len(slStr) < 8
                        slStr = "0" & slStr
                    Loop
                    slName = slStr
                    slStr = Trim$(str$(999 - tlChfAdvtExt.iCntRevNo))
                    Do While Len(slStr) < 3
                        slStr = "0" & slStr
                    Loop
                    slExtStr = Trim$(str$(999 - tlChfAdvtExt.iExtRevNo))
                    Do While Len(slExtStr) < 3
                        slExtStr = "0" & slExtStr
                    Loop
                    slName = slName & "|" & slStr & "-" & slExtStr & "|"
                    If (tlChfAdvtExt.sStatus = "W") Or (tlChfAdvtExt.sStatus = "C") Or (tlChfAdvtExt.sStatus = "I") Then
                        'Add Potential
                        If tlChfAdvtExt.iMnfPotnType > 0 Then
                            slStr = " "
                        Else
                            slStr = "~"
                        End If
                        slName = slName & slStr & "|"
                    Else
                        slName = slName & " |"
                    End If
                    slStr = Trim$(str$(999 - tlChfAdvtExt.iPropVer))
                    Do While Len(slStr) < 3
                        slStr = "0" & slStr
                    Loop
                    slName = slName & slStr & "|"
                    slName = slName & tlChfAdvtExt.sStatus & "|"
                    If ilShow = 0 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo)) & " V" & Trim$(Str$(tlChfAdvtExt.iPropVer))
                    ElseIf ilShow = 2 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                        gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                        gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slEndDate
                        slName = slName & ": " & slStartDate & "-" & slEndDate
                        slName = slName & " " & Trim$(tlChfAdvtExt.sProduct)
                        If tlChfAdvtExt.lVefCode > 0 Then
                            If tlChfAdvtExt.lVefCode <> tlVef.iCode Then
                                tlVefSrchKey.iCode = tlChfAdvtExt.lVefCode
                                ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo gPopCntrForAASBoxByMktErr
                                gCPErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrGetEqual: Vef)", frm
                                On Error GoTo 0
                            End If
                            slName = slName & " " & Trim$(tlVef.sName)
                        End If
                    ElseIf ilShow = 3 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                        If tlChfAdvtExt.iAdfCode <> tlAdf.iCode Then
                            tlAdfSrchKey.iCode = tlChfAdvtExt.iAdfCode
                            ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo gPopCntrForAASBoxByMktErr
                            gCPErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrGetEqual: Adf)", frm
                            On Error GoTo 0
                        End If
                        'slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                        If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sAddrID) <> "") Then
                            slName = slName & " " & Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID)
                        Else
                            slName = slName & " " & Trim$(tlAdf.sName)
                        End If
                        gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                        gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slEndDate
                        slName = slName & " " & slStartDate & "-" & slEndDate
                    ElseIf ilShow = 4 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                        gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                        gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slEndDate
                        slName = slName & " " & slStartDate & "-" & slEndDate
                    ElseIf ilShow = 5 Then
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo))
                        If tlChfAdvtExt.iAdfCode <> tlAdf.iCode Then
                            tlAdfSrchKey.iCode = tlChfAdvtExt.iAdfCode
                            ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo gPopCntrForAASBoxByMktErr
                            gCPErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrGetEqual: Adf)", frm
                            On Error GoTo 0
                        End If
                        'slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                        If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sAddrID) <> "") Then
                            slName = slName & " " & Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID)
                        Else
                            slName = slName & " " & Trim$(tlAdf.sName)
                        End If
                    Else
                        'slName = Trim$(Str$(tlChfAdvtExt.lCntrNo)) & " R" & Trim$(Str$(tlChfAdvtExt.iCntRevNo)) & " V" & Trim$(Str$(tlChfAdvtExt.iPropVer))
                        Select Case tlChfAdvtExt.sStatus
                            Case "W"
                                If tlChfAdvtExt.iCntRevNo > 0 Then
                                    slStr = "Rev Working"
                                Else
                                    slStr = "Working"
                                End If
                            Case "D"
                                slStr = "Rejected"
                            Case "C"
                                If tlChfAdvtExt.iCntRevNo > 0 Then
                                    slStr = "Rev Completed"
                                Else
                                    slStr = "Completed"
                                End If
                            Case "I"
                                If tlChfAdvtExt.iCntRevNo > 0 Then
                                    slStr = "Rev Unapproved"
                                Else
                                    slStr = "Unapproved"
                                End If
                            Case "G"
                                slStr = "Approved Hold"
                            Case "N"
                                slStr = "Approved Order"
                            Case "H"
                                slStr = "Hold"
                            Case "O"
                                slStr = "Order"
                        End Select
                        slName = slName & " " & slStr
                        If ilShow = 6 Then
                            If tlChfAdvtExt.iAdfCode <> tlAdf.iCode Then
                                tlAdfSrchKey.iCode = tlChfAdvtExt.iAdfCode
                                ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                On Error GoTo gPopCntrForAASBoxByMktErr
                                gCPErrorMsg ilRet, "gPopCntrForAASBoxByMkt (btrGetEqual: Adf)", frm
                                On Error GoTo 0
                            End If
                            'slName = slName & " " & Trim$(tlAdf.sName) '& "/" & Trim$(tlChfAdvtExt.sProduct)
                            If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sAddrID) <> "") Then
                                slName = slName & " " & Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID)
                            Else
                                slName = slName & " " & Trim$(tlAdf.sName)
                            End If
                        Else
                            slName = slName & " " & Trim$(tlChfAdvtExt.sProduct)
                        End If
                        'Start Date Plus # weeks
                        gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                        gUnpackDateLong tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), llEndDate
                        If ilShow = 6 Then
                            slName = slName & " " & slStartDate & "-" & Format$(llEndDate, "m/d/yy")
                        Else
                            slStr = str$((llEndDate - gDateValue(slStartDate)) \ 7 + 1)
                            slName = slName & " " & slStartDate & slStr
                            tmCxfSrchKey.lCode = tlChfAdvtExt.lCxfInt
                            If tmCxfSrchKey.lCode <> 0 Then
                                tmCxf.sComment = ""
                                imCxfRecLen = Len(tmCxf) '5027
                                ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    'If tmCxf.iStrLen > 0 Then
                                    '    If tmCxf.iStrLen < 40 Then
                                    '        slStr = Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
                                    '        'slName = slName & " " & Trim$(Left$(tmCxf.sComment, tmCxf.iStrLen))
                                    '    Else
                                            'slStr = Trim$(Left$(tmCxf.sComment, 40))
                                            slStr = gStripChr0(Left$(tmCxf.sComment, 40))
                                    '        'slName = slName & " " & Trim$(Left$(tmCxf.sComment, 40))
                                    '    End If
                                        Do While InStr(1, slStr, Chr(0), vbTextCompare) > 0
                                            ilPos = InStr(1, slStr, Chr(0), vbTextCompare)
                                            Mid$(slStr, ilPos, 1) = " "
                                        Loop
                                        slName = slName & " " & slStr
                                    'End If
                                End If
                            End If
                        End If
                    End If
                    ilFound = False
                    If ilTestCntrNo Then
                        For illoop = ilLowLimit To ilSortCode - 1 Step 1
                            slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
                            ilRet = gParseItem(slNameCode, 1, "\", slNameCode)
                            ilRet = gParseItem(slNameCode, 1, "|", slCode)
                            llCntrNo = 99999999 - CLng(slCode)
                            If llCntrNo = tlChfAdvtExt.lCntrNo Then
                                ilRet = gParseItem(slNameCode, 2, "|", slCode)
                                ilRet = gParseItem(slCode, 1, "-", slExtStr)
                                ilRevNo = 999 - CLng(slExtStr)
                                If tlChfAdvtExt.iCntRevNo > ilRevNo Then
                                    'Replace
                                    'lbcMster.RemoveItem ilLoop
                                    'llLen = llLen - Len(slNameCode)
                                    ilFound = True
                                    slName = slName & "\" & Trim$(str$(tlChfAdvtExt.lCode))
                                    tlSortCode(illoop).sKey = slName
                                Else
                                    'Leave
                                    ilFound = True
                                End If
                                Exit For
                            End If
                        Next illoop
                    End If
                    If Not ilFound Then
                        slName = slName & "\" & Trim$(str$(tlChfAdvtExt.lCode))
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit Do
                        'End If
                        'lbcMster.AddItem slName
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                    End If
                End If
                ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlChf, tlChfAdvtExt, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If (ilHOType = 4) Or (ilHOType = 5) Then
                For ilLoop1 = 0 To UBound(lmWCINGCntrNo) - 1 Step 1
                    For illoop = UBound(tlSortCode) - 1 To ilLowLimit Step -1
                        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
                        ilRet = gParseItem(slNameCode, 1, "\", slName)
                        ilRet = gParseItem(slName, 1, "|", slCode)
                        llCntrNo = 99999999 - CLng(slCode)
                        If lmWCINGCntrNo(ilLoop1) = llCntrNo Then
                            For ilIndex = illoop To UBound(tlSortCode) - 1 Step 1
                                tlSortCode(ilIndex) = tlSortCode(ilIndex + 1)
                            Next ilIndex
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) - 1) As SORTCODE
                        End If
                    Next illoop
                Next ilLoop1
            End If
            Erase lmWCINGCntrNo
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If
        ilRet = btrClose(hmCxf)
        btrDestroy hmCxf
        ilRet = btrClose(hlAdf)
        btrDestroy hlAdf
        ilRet = btrClose(hlVsf)
        btrDestroy hlVsf
        ilRet = btrClose(hlVef)
        btrDestroy hlVef
        ilRet = btrClose(hlChf)
        btrDestroy hlChf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        ilRet = gParseItem(slName, 1, "|", slCode)
        llCntrNo = 99999999 - CLng(slCode)
        slShow = Trim$(str$(llCntrNo))
        ilRet = gParseItem(slName, 2, "|", slCode)
        ilRet = gParseItem(slCode, 1, "-", slExtStr)
        ilRevNo = 999 - CLng(slExtStr)
        ilRet = gParseItem(slCode, 2, "-", slExtStr)
        ilExtRevNo = 999 - CLng(slExtStr)
        ilRet = gParseItem(slName, 4, "|", slCode)
        ilVerNo = 999 - CLng(slCode)
        ilRet = gParseItem(slName, 5, "|", slCode)
        If (slCode = "W") Or (slCode = "C") Or (slCode = "I") Or (slCode = "D") Then
            If ilRevNo > 0 Then
                slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
            Else
                slShow = slShow & " V" & Trim$(str$(ilVerNo))
            End If
        Else
            slShow = slShow & " R" & Trim$(str$(ilRevNo)) & "-" & Trim$(str$(ilExtRevNo))
        End If
        If ilShow = 0 Then      'Number only
        ElseIf ilShow = 2 Then  'Number, Dates, Product, Vehicle
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 3 Then  'Number, Advertiser, Dates
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 4 Then  'Number, Dates
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        ElseIf ilShow = 5 Then  'Number, Advertiser
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        Else                    'Number, Product, Internal comment
            'Potential
            ilRet = gParseItem(slName, 3, "|", slCode)
            If (Trim$(slCode) <> "") And (slCode <> "~") Then
                slShow = slShow & " " & slCode
            End If
            'Other fields
            ilRet = gParseItem(slName, 6, "|", slCode)
            slShow = slShow & " " & slCode
        End If
        If Not gOkAddStrToListBox(slShow, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slShow  'Add ID to list box
    Next illoop
    Exit Function
gPopCntrForAASBoxByMktErr:
    ilRet = btrClose(hmCxf)
    btrDestroy hmCxf
    ilRet = btrClose(hlAdf)
    btrDestroy hlAdf
    ilRet = btrClose(hlVsf)
    btrDestroy hlVsf
    ilRet = btrClose(hlVef)
    btrDestroy hlVef
    ilRet = btrClose(hlChf)
    btrDestroy hlChf
    gDbg_HandleError "PopSubs: gPopCntrForAASBoxByMkt"
'    gPopCntrForAASBoxByMkt = CP_MSG_NOSHOW
'    Exit Function
gPopCntrForAASBoxByMktErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopCopyForAdvtBox              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     copy for specified advertiser   *
'*                                                     *
'*******************************************************
Function gPopCopyForAdvtBox(frm As Form, ilAdvtCode As Integer, ilShowOrder As Integer, ilInputInvType As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopCopyForAdvtBox (MainForm, ilAdvtCode, ilShowOrder, ilInvType, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAdvtCode (I)- Advertiser code (-1 = All)
'       ilShowOrder (I)- 0=Copy only; 1=Copy/ISCI; 2=Copy/ISCI/Product; 3=ISCI/Copy; 4=ISCI; 5=ISCI/Product
'                       +256 (&H100) indicates that length should be shown first
'       ilInputInvType (I)-0 = Active only (exclude Purged and History); 1=Purged and History and Active; 2=Purged and History only;
'                        3=Active and History; 4= Purged only; 5=History only; 6=Active and Purged
'                       +256 (&H100) indicates that (Purged) and (History) text should be shown
'                       +512 (&H200) indicates only include inventory with length less then 30 sec (Billboards)
'                       +1024 (&H400) requires UseCartNo and mif to match
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'
    gPopCopyForAdvtBox = gPopCopyForAdvtBoxWithTest(frm, ilAdvtCode, -1, -1, "", ilShowOrder, ilInputInvType, lbcLocal, tlSortCode, slSortCodeTag, 0, 0)
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopCopyForAdvtBoxWithTest      *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     copy for specified advertiser   *
'*                                                     *
'*******************************************************
Function gPopCopyForAdvtBoxWithTest(frm As Form, ilAdvtCode As Integer, ilVefCode As Integer, ilLength As Integer, slLive As String, ilShowOrder As Integer, ilInputInvType As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, llActiveStartDate As Long, llActiveEndDate As Long, Optional blIncludeExtraInfo As Boolean = False, Optional llInCntrNo As Long = -1) As Integer
'
'   ilRet = gPopCopyForAdvtBoxWithTest (MainForm, ilAdvtCode, ilShowOrder, ilInvType, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAdvtCode (I)- Advertiser code (-1 = All)
'       ilVefCode (I)- Vehicle code (-1 = all)
'       ilLength (I)- Match spot length (-1 = all)
'       slLive (I)-  L= Live (requires script); R=No script; B = Both
'       ilShowOrder (I)- 0=Copy only; 1=Copy/ISCI; 2=Copy/ISCI/Product; 3=ISCI/Copy; 4=ISCI; 5=ISCI/Product
'                       +256 (&H100) indicates that length should be shown first
'       ilInputInvType (I)-0 = Active only (exclude Purged and History); 1=Purged and History and Active; 2=Purged and History only;
'                        3=Active and History; 4= Purged only; 5=History only; 6=Active and Purged
'                       +256 (&H100) indicates that (Purged) and (History) text should be shown
'                       +512 (&H200) indicates only include inventory with length less then 30 sec (Billboards)
'                       +1024 (&H400) requires UseCartNo and mif to match
'                       +2048 (&H800) requires Copy Inventory to have dates
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       llActiveStartDate(I)- Start date of rotation.  Inventory must be active on or before this date.  Used with Blackout only
'       llActiveEndDate(I)- Start date of rotation.  Inventory must be active on or after this date.  Used with Blackout only
'       blIncludeExtraInfo(I)- Include Product and Last Used Date after cifCode (field 2) as Field 3
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'Cif date/time stamp
    Dim hlCif As Integer        'Cif handle
    Dim hlTCif As Integer        'Cif handle
    Dim hlMcf As Integer
    Dim hlCpf As Integer
    Dim ilCifRecLen As Integer     'Record length
    Dim ilMcfRecLen As Integer     'Record length
    Dim ilCpfRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlCif As CIF
    Dim tlCifExt As CIFEXT    'Copy extract record
    Dim tlMcf As MCF
    Dim tlCpf As CPF
    Dim slName As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim ilShowLen As Integer
    Dim ilShowPH As Integer  'Show Purge and History names
    Dim ilInvType As Integer
    Dim ilSortOrder As Integer
    Dim llNowDate As Long
    Dim slDate As String
    'Dim ilRecRP As Integer  '1=Reused; 2=Purged
    Dim tlMcfSrchKey As INTKEY0 'MCF key record image
    Dim tlCpfSrchKey As LONGKEY0 'CPF key record image
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilBBLenOnly As Integer
    Dim ilLiveOk As Integer
    Dim ilMatchMode As Integer
    Dim ilHubOk As Integer
    Dim ilVehOk As Integer
    Dim ilVef As Integer
    Dim ilRotDateOk As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilPos As Integer
    Dim llCntrNo As Long
    Dim slStr As String
    Dim ilLowLimit As Integer
    Dim blCopyDatesRequired As Boolean

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Cif.Btr") & Trim$(str$(ilAdvtCode)) & Trim$(str$(ilLength)) & Trim$(slLive) & Trim$(str$(ilShowOrder)) & Trim$(str$(ilInputInvType)) & Trim$(str$(llActiveStartDate)) & Trim$(str$(llActiveEndDate)) & Trim$(str$(llInCntrNo))

    'On Error GoTo gPopCopyForAdvtWithTestErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopCopyForAdvtBoxWithTest = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    ilSortOrder = ilShowOrder And &HFF
    ilShowLen = ilShowOrder And &H100
    ilInvType = ilInputInvType And &HFF
    ilShowPH = ilInputInvType And &H100
    ilBBLenOnly = ilInputInvType And &H200
    ilMatchMode = ilInputInvType And &H400
    blCopyDatesRequired = False
    If (ilInputInvType And &H800) = &H800 Then
        blCopyDatesRequired = True
    End If
    gPopCopyForAdvtBoxWithTest = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlCif = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrOpen):" & "Cif.Btr", frm
        On Error GoTo 0
        ilCifRecLen = Len(tlCif) 'btrRecordLength(hlAdf)  'Get and save record length
        hlTCif = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlTCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrOpen):" & "Cif.Btr", frm
        On Error GoTo 0
        hlMcf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrOpen):" & "Mcf.Btr", frm
        On Error GoTo 0
        ilMcfRecLen = Len(tlMcf) 'btrRecordLength(hlAdf)  'Get and save record length
        tlMcf.iCode = 0
        hlCpf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrOpen):" & "Cpf.Btr", frm
        On Error GoTo 0
        ilCpfRecLen = Len(tlCpf) 'btrRecordLength(hlAdf)  'Get and save record length
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlCifExt)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlCif) 'Obtain number of records
        btrExtClear hlCif   'Clear any previous extend operation
        ilRet = btrGetFirst(hlCif, tlCif, ilCifRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlMcf)
            ilRet = btrClose(hlCpf)
            ilRet = btrClose(hlTCif)
            ilRet = btrClose(hlCif)
            On Error GoTo gPopCopyForAdvtBoxWithTestErr
            gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrReset):" & "Adf.Btr", frm
            On Error GoTo 0
            btrDestroy hlCpf
            btrDestroy hlMcf
            btrDestroy hlTCif
            btrDestroy hlCif
            Exit Function
        Else
            On Error GoTo gPopCopyForAdvtBoxWithTestErr
            gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrGetFirst):" & "Cif.Btr", frm
            On Error GoTo 0
        End If
        'Extract information from first record
    'Use UC (get cuurent) in setbounds instead of EG (bypass current)
    '    If Trim$(tlAdf.sBillAgyDir) = "D" Then
    '        slName = Trim$(tlAdf.sName) & "/Direct"
    '    Else
    '        slName = Trim$(tlAdf.sName)
    '    End If
    '    slName = slName & "\" & Trim$(Str$(tlAdf.iCode))
    '    lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
        slDate = Format$(gNow(), "m/d/yy")
        llNowDate = gDateValue(slDate)
        Call btrExtSetBounds(hlCif, llNoRec, -1, "UC", "CIFEXTPK", CIFEXTPK) 'Set extract limits (all records including first)
        If ilAdvtCode > 0 Then
            tlIntTypeBuff.iType = ilAdvtCode
            ilOffSet = gFieldOffset("Cif", "CifAdfCode")
            If ilInvType <> 1 Then
                ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
            Else
                ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
            End If
        End If
        Select Case ilInvType
            Case 0  'A
                tlCharTypeBuff.sType = "A"
                ilOffSet = gFieldOffset("Cif", "CifPurged")
                ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
            Case 2  'P or H
                tlCharTypeBuff.sType = "A"
                ilOffSet = gFieldOffset("Cif", "CifPurged")
                ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
            Case 3  'A or H
                tlCharTypeBuff.sType = "P"
                ilOffSet = gFieldOffset("Cif", "CifPurged")
                ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
            Case 4  'P
                tlCharTypeBuff.sType = "P"
                ilOffSet = gFieldOffset("Cif", "CifPurged")
                ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
            Case 5  'H
                tlCharTypeBuff.sType = "H"
                ilOffSet = gFieldOffset("Cif", "CifPurged")
                ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
            Case 6  'A or P
                tlCharTypeBuff.sType = "H"
                ilOffSet = gFieldOffset("Cif", "CifPurged")
                ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        End Select
        ilOffSet = gFieldOffset("Cif", "CifCode")
        ilRet = btrExtAddField(hlCif, ilOffSet, 4)  'Extract lCode field
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifMcfCode")
        ilRet = btrExtAddField(hlCif, ilOffSet, 2)  'Extract lCode field
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifName")
        ilRet = btrExtAddField(hlCif, ilOffSet, 5)  'Extract Name field
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifCut")
        ilRet = btrExtAddField(hlCif, ilOffSet, 1) 'Extract direct flag
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifLen")
        ilRet = btrExtAddField(hlCif, ilOffSet, 2) 'Extract direct flag
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifCpfCode")
        ilRet = btrExtAddField(hlCif, ilOffSet, 4)  'Extract lCode field
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifPurged")
        ilRet = btrExtAddField(hlCif, ilOffSet, 1)  'Extract Name field
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifRotStartDate")
        ilRet = btrExtAddField(hlCif, ilOffSet, 4)  'Extract lCode field
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifRotEndDate")
        ilRet = btrExtAddField(hlCif, ilOffSet, 4)  'Extract lCode field
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifCsfCode")
        ilRet = btrExtAddField(hlCif, ilOffSet, 4)  'Extract lCode field
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifReel")
        ilRet = btrExtAddField(hlCif, ilOffSet, 10)  'Extract Name field
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hlCif)    'Extract record
        ilRet = btrExtGetNext(hlCif, tlCifExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopCopyForAdvtBoxWithTestErr
            gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrExtGetNextExt):" & "Cif.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlCifExt)  'Extract operation record size
            'ilRet = btrExtGetFirst(hlCif, tlCifExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlCif, tlCifExt, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If llActiveStartDate > 0 Then
                    gUnpackDateLong tlCifExt.iRotStartDate(0), tlCifExt.iRotStartDate(1), llStartDate
                    If (llStartDate <= llActiveStartDate) Or ((llStartDate = 0) And (blCopyDatesRequired = False)) Then
                        gUnpackDateLong tlCifExt.iRotEndDate(0), tlCifExt.iRotEndDate(1), llEndDate
                        If (llEndDate >= llActiveEndDate) Or ((llEndDate = 0) And (blCopyDatesRequired = False)) Then
                            ilRotDateOk = True
                        Else
                            ilRotDateOk = False
                        End If
                    Else
                        ilRotDateOk = False
                    End If
                Else
                    ilRotDateOk = True
                    If blCopyDatesRequired Then
                        gUnpackDateLong tlCifExt.iRotStartDate(0), tlCifExt.iRotStartDate(1), llStartDate
                        If llStartDate = 0 Then
                            ilRotDateOk = False
                        Else
                            gUnpackDateLong tlCifExt.iRotEndDate(0), tlCifExt.iRotEndDate(1), llEndDate
                            If llEndDate = 0 Then
                                ilRotDateOk = False
                            End If
                        End If
                    End If
                End If
                If ((Asc(tgSpf.sUsingFeatures10) And VCREATIVEEXPORT) = VCREATIVEEXPORT) And (ilRotDateOk) And (llInCntrNo <> -1) Then
                    slStr = tlCifExt.sReel
                    If Trim$(slStr) <> "" Then
                        ilPos = InStr(1, slStr, "-", vbTextCompare)
                        If ilPos > 0 Then
                            llCntrNo = Val(Left(slStr, ilPos - 1))
                        Else
                            ilPos = InStr(1, slStr, " ", vbTextCompare)
                            If ilPos > 0 Then
                                llCntrNo = Val(Left(slStr, ilPos - 1))
                            Else
                                llCntrNo = Val(slStr)
                            End If
                        End If
                        If (llInCntrNo <> -1) And (llCntrNo <> llInCntrNo) Then
                            ilRotDateOk = False
                        End If
                    End If
                End If
                If ((ilLength = -1) Or (ilLength = tlCifExt.iLen)) And (ilRotDateOk) Then
                    ilLiveOk = False
                    ilHubOk = True
                    ilVehOk = True
                    If (tgSpf.sUseCartNo <> "N") And (tlCifExt.iMcfCode <> 0) Then
                        If tlMcf.iCode <> tlCifExt.iMcfCode Then
                            tlMcfSrchKey.iCode = tlCifExt.iMcfCode
                            ilRet = btrGetEqual(hlMcf, tlMcf, ilMcfRecLen, tlMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            On Error GoTo gPopCopyForAdvtBoxWithTestErr
                            gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrGetEqual):" & "Mcf.Btr", frm
                            On Error GoTo 0
                        End If
                        If (slLive = "") Or (slLive = "B") Or (((slLive = "L") Or (slLive = "M")) And (tlCifExt.lCsfCode > 0) And (tlMcf.sScript = "Y")) Or ((slLive = "R") And ((tlMcf.sScript <> "Y") Or (tlCifExt.lCsfCode <= 0) Or ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) <> LIVECOPY))) Then
                            ilLiveOk = True
                        Else
                            If ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) <> LIVECOPY) Then
                                ilLiveOk = True
                            End If
                        End If
                        If ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB) And (tgUrf(0).iCode > 2) And (tgUrf(0).iMnfHubCode > 0) Then
                            If tlMcf.iVefCode > 0 Then
                                ilVef = gBinarySearchVef(tlMcf.iVefCode)
                                If ilVef <> -1 Then
                                    If tgMVef(ilVef).iMnfHubCode <> tgUrf(0).iMnfHubCode Then
                                        ilHubOk = False
                                    End If
                                End If
                            End If
                        End If
                        If (Asc(tgSpf.sUsingFeatures3) And MEDIACODEBYVEH) = MEDIACODEBYVEH Then
                            If ilVefCode <> -1 Then
                                If (tlMcf.iVefCode <> 0) And (tlMcf.iVefCode <> ilVefCode) Then
                                    ilVehOk = False
                                End If
                            End If
                        End If
                    Else
                        If (slLive = "") Or (slLive = "B") Or (((slLive = "L") Or (slLive = "M")) And (tlCifExt.lCsfCode > 0)) Or ((slLive = "R") And ((tlCifExt.lCsfCode <= 0) Or ((Asc(tgSpf.sUsingFeatures) And LIVECOPY) <> LIVECOPY))) Then
                            ilLiveOk = True
                        End If
                    End If
                    If (ilMatchMode <> 0) And (ilLiveOk) And (ilHubOk) And (ilVehOk) Then
                        If (tgSpf.sUseCartNo = "N") And (tlCifExt.iMcfCode > 0) Then
                            ilLiveOk = False
                        End If
                        If (tgSpf.sUseCartNo <> "N") And (tlCifExt.iMcfCode <= 0) Then
                            ilLiveOk = False
                        End If
                    End If
                    If ilLiveOk Then
                        If (ilBBLenOnly = 0) Or ((ilBBLenOnly <> 0) And (tlCifExt.iLen < 30)) Then
                            If (tgSpf.sUseCartNo <> "N") And (tlCifExt.iMcfCode <> 0) Then
                                If tlMcf.iCode <> tlCifExt.iMcfCode Then
                                    tlMcfSrchKey.iCode = tlCifExt.iMcfCode
                                    ilRet = btrGetEqual(hlMcf, tlMcf, ilMcfRecLen, tlMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    On Error GoTo gPopCopyForAdvtBoxWithTestErr
                                    gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrGetEqual):" & "Mcf.Btr", frm
                                    On Error GoTo 0
                                End If
                                If ilSortOrder <> 4 Then
                                    slName = Trim$(tlMcf.sName) & Trim$(tlCifExt.sName)
                                    If (Len(Trim$(tlCifExt.sCut)) <> 0) Then
                                        slName = slName & "-" & tlCifExt.sCut
                                    End If
                                Else
                                    slName = ""
                                End If
                                If ilSortOrder > 0 Then
                                    If tlCifExt.lcpfCode > 0 Then
                                        tlCpfSrchKey.lCode = tlCifExt.lcpfCode
                                        ilRet = btrGetEqual(hlCpf, tlCpf, ilCpfRecLen, tlCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        On Error GoTo gPopCopyForAdvtBoxWithTestErr
                                        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrGetEqual):" & "Cpf.Btr", frm
                                        On Error GoTo 0
                                        If Trim$(tlCpf.sISCI) <> "" Then
                                            If (ilSortOrder = 1) Or (ilShowOrder = 2) Then
                                                slName = slName & " " & Trim$(tlCpf.sISCI)
                                            ElseIf ilSortOrder = 3 Then
                                                slName = Trim$(tlCpf.sISCI) & " " & slName
                                            ElseIf ilSortOrder = 4 Then
                                                slName = Trim$(tlCpf.sISCI)
                                            End If
                                        End If
                                        If ilSortOrder = 2 Then
                                            If Trim$(tlCpf.sName) <> "" Then
                                                slName = slName & " " & Trim$(tlCpf.sName)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If tlCifExt.lcpfCode > 0 Then
                                    tlCpfSrchKey.lCode = tlCifExt.lcpfCode
                                    ilRet = btrGetEqual(hlCpf, tlCpf, ilCpfRecLen, tlCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    On Error GoTo gPopCopyForAdvtBoxWithTestErr
                                    gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrGetEqual):" & "Cpf.Btr", frm
                                    On Error GoTo 0
                                    If Trim$(tlCpf.sISCI) <> "" Then
                                        slName = Trim$(tlCpf.sISCI)
                                    Else
                                        slName = "ISCI Missing"
                                    End If
                                    If (ilSortOrder = 2) Or (ilSortOrder = 5) Then
                                        If Trim$(tlCpf.sName) <> "" Then
                                            slName = slName & " " & Trim$(tlCpf.sName)
                                        End If
                                    End If
                                Else
                                    slName = "ISCI Missing"
                                End If
                            End If
                            If ilShowLen > 0 Then
                                slName = Trim$(str(tlCifExt.iLen)) & " " & slName
                            End If
                            'Ignore extra message to Non-purged/non reused only and for All (post Log, inventory by advertiser)
                            If ilShowPH > 0 Then
                                If tlCifExt.sPurged = "P" Then
                                    slName = slName & " (Purged)"    'If name chaned, then change in mTestForPurge within Copy.Frm
                                ElseIf tlCifExt.sPurged = "H" Then
                                    slName = slName & " (History)"    'If name chaned, then change in mTestForPurge within Copy.Frm
                                End If
                            End If
                            slName = slName & "\" & Trim$(str$(tlCifExt.lCode))
                            If blIncludeExtraInfo Then
                                gUnpackDateLong tlCifExt.iRotEndDate(0), tlCifExt.iRotEndDate(1), llEndDate
                                If tlCifExt.lcpfCode > 0 Then
                                    slName = slName & "\" & Trim$(tlCpf.sName) & "\" & llEndDate
                                Else
                                    slName = slName & "\" & "\" & llEndDate
                                End If
                            End If
                            'If Not gOkAddStrToListBox(slName, llLen, True) Then
                            '    Exit Do
                            'End If
                            'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                            tlSortCode(ilSortCode).sKey = slName
                            If ilSortCode >= UBound(tlSortCode) Then
                                ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                            End If
                            ilSortCode = ilSortCode + 1
                        End If
                    End If
                End If
                'End If
                ilRet = btrExtGetNext(hlCif, tlCifExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlCif, tlCifExt, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If
        ilRet = btrClose(hlCpf)
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrReset):" & "Cpf.Btr", frm
        On Error GoTo 0
        btrDestroy hlCpf
        ilRet = btrClose(hlMcf)
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrReset):" & "Mcf.Btr", frm
        On Error GoTo 0
        btrDestroy hlMcf
        ilRet = btrClose(hlTCif)
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrReset):" & "Cif.Btr", frm
        On Error GoTo 0
        btrDestroy hlTCif
        ilRet = btrClose(hlCif)
        On Error GoTo gPopCopyForAdvtBoxWithTestErr
        gBtrvErrorMsg ilRet, "gPopCopyForAdvtBoxWithTest (btrReset):" & "Cif.Btr", frm
        On Error GoTo 0
        btrDestroy hlCif
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopCopyForAdvtBoxWithTest = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopCopyForAdvtBoxWithTestErr:
    ilRet = btrClose(hlMcf)
    ilRet = btrClose(hlCpf)
    ilRet = btrClose(hlTCif)
    ilRet = btrClose(hlCif)
    btrDestroy hlCpf
    btrDestroy hlMcf
    btrDestroy hlTCif
    btrDestroy hlCif
    gDbg_HandleError "PopSubs: gPopCopyForAdvtBoxWithTest"
'    gPopCopyForAdvtBoxWithTest = CP_MSG_NOSHOW
'    Exit Function
gPopCopyForAdvtWithTestErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopCopyForMediaBox              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     copy for media by reverse date  *
'*                                                     *
'*******************************************************
Function gPopCopyForMediaBox(frm As Form, ilMcfCode As Integer, slPriorDate As String, ilShowRotDates As Integer, ilPurgedOnly As Integer, ilSortCart As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopCopyForMediaBox (MainForm, ilMcfCode, slPriorDate, ilShowRotDates, ilPurgedOnly, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilMediaCode (I)- Media code
'       slPriodDate (I)- inventory prior to specified date (if "", then all dates; if -1, then Never Activated)
'       ilShowRotDates(I) - True=Show rotation dates; False=Don't show
'       ilPurgedOnly(I)- True=Show only inventory marked as Purged; False=Show all Inventory
'       ilSortCart(I)- 0=Purge Date, thern cart #, 1=Cart #
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'
'       History inventory is excluded
'

    Dim slStamp As String    'Cif date/time stamp
    Dim hlCif As Integer        'Cif handle
    Dim ilCifRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlCif As CIF
    Dim tlCifExt As CIFINVEXT    'Copy extract record
    Dim tlInputExt As CIFINVEXT
    Dim slName As String
    Dim slInvName As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    '1/17/14 TTP 6600
    Dim llLoop As Long
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim slRotDate As String
    Dim slPurgeDate As String
    Dim llPriorDate As Long
    Dim ilAddRec As Integer
    Dim slDate1 As String
    Dim slDate2 As String
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim llLen As Long
    '1/17/14 TTP 6600
    Dim llSortCode As Long
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Cif.Btr") & Trim$(str$(ilMcfCode)) & slPriorDate & Trim$(str$(ilShowRotDates)) & Trim$(str$(ilPurgedOnly)) & Trim$(str$(ilSortCart))

    'On Error GoTo gPopCopyForMediaErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopCopyForMediaBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopCopyForMediaBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlCif = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCopyForMediaBoxErr
        gBtrvErrorMsg ilRet, "gPopCopyForMediaBox (btrOpen):" & "Cif.Btr", frm
        On Error GoTo 0
        ilCifRecLen = Len(tlCif) 'btrRecordLength(hlAdf)  'Get and save record length
        llSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        If slPriorDate = "-1" Then
            llPriorDate = -1
        Else
            llPriorDate = gDateValue(slPriorDate)
        End If
        ilExtLen = Len(tlCifExt)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlCif) 'Obtain number of records
        btrExtClear hlCif   'Clear any previous extend operation
        'Use key 1 to get duplicate copy numbers together
        ilRet = btrGetFirst(hlCif, tlCif, ilCifRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlCif)
            On Error GoTo gPopCopyForMediaBoxErr
            gBtrvErrorMsg ilRet, "gPopCopyForMediaBox (btrReset):" & "Adf.Btr", frm
            On Error GoTo 0
            btrDestroy hlCif
            Exit Function
        Else
            On Error GoTo gPopCopyForMediaBoxErr
            gBtrvErrorMsg ilRet, "gPopCopyForMediaBox (btrGetFirst):" & "Cif.Btr", frm
            On Error GoTo 0
        End If
        'Extract information from first record
    'Use UC (get cuurent) in setbounds instead of EG (bypass current)
    '    If Trim$(tlAdf.sBillAgyDir) = "D" Then
    '        slName = Trim$(tlAdf.sName) & "/Direct"
    '    Else
    '        slName = Trim$(tlAdf.sName)
    '    End If
    '    slName = slName & "\" & Trim$(Str$(tlAdf.iCode))
    '    lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
        Call btrExtSetBounds(hlCif, llNoRec, -1, "UC", "CIFINVEXTPK", CIFINVEXTPK) 'Set extract limits (all records including first)
        tlCharTypeBuff.sType = "H"
        ilOffSet = gFieldOffset("Cif", "CifPurged")
        ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
        tlIntTypeBuff.iType = ilMcfCode
        ilOffSet = gFieldOffset("Cif", "CifMcfCode")
        ilRet = btrExtAddLogicConst(hlCif, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        ilOffSet = gFieldOffset("Cif", "CifCode")
        ilRet = btrExtAddField(hlCif, ilOffSet, 4)  'Extract lCode field
        On Error GoTo gPopCopyForMediaBoxErr
        gBtrvErrorMsg ilRet, "gPopCopyForMediaBox (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifName")
        ilRet = btrExtAddField(hlCif, ilOffSet, 5)  'Extract Name field
        On Error GoTo gPopCopyForMediaBoxErr
        gBtrvErrorMsg ilRet, "gPopCopyForMediaBox (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifCut")
        ilRet = btrExtAddField(hlCif, ilOffSet, 1) 'Extract direct flag
        On Error GoTo gPopCopyForMediaBoxErr
        gBtrvErrorMsg ilRet, "gPopCopyForMediaBox (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifPurged")
        ilRet = btrExtAddField(hlCif, ilOffSet, 1)  'Extract Name field
        On Error GoTo gPopCopyForMediaBoxErr
        gBtrvErrorMsg ilRet, "gPopCopyForMediaBox (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifRotEndDate")
        ilRet = btrExtAddField(hlCif, ilOffSet, 4)  'Extract lCode field
        On Error GoTo gPopCopyForMediaBoxErr
        gBtrvErrorMsg ilRet, "gPopCopyForMediaBox (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Cif", "CifPurgeDate")
        ilRet = btrExtAddField(hlCif, ilOffSet, 4)  'Extract lCode field
        On Error GoTo gPopCopyForMediaBoxErr
        gBtrvErrorMsg ilRet, "gPopCopyForMediaBox (btrExtAddField):" & "Cif.Btr", frm
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hlCif)    'Extract record
        tlCifExt.sName = ""
        ilRet = btrExtGetNext(hlCif, tlInputExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopCopyForMediaBoxErr
            gBtrvErrorMsg ilRet, "gPopCopyForMediaBox (btrExtGetNextExt):" & "Cif.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlCifExt)  'Extract operation record size
            'ilRet = btrExtGetFirst(hlCif, tlCifExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlCif, tlInputExt, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If Trim$(tlCifExt.sName) = "" Then 'First record
                    tlCifExt = tlInputExt
                    ilAddRec = False
                Else
                    If (StrComp(tlInputExt.sName, tlCifExt.sName, 1) = 0) And (StrComp(tlInputExt.sCut, tlCifExt.sCut, 1) = 0) Then
                        'Save latest rotation end date
                        If (tlCifExt.sPurged = "P") And (tlInputExt.sPurged <> "P") Then
                            tlCifExt = tlInputExt
                        ElseIf (tlCifExt.sPurged <> "P") And (tlInputExt.sPurged = "P") Then
                        Else
                            gUnpackDate tlCifExt.iRotEndDate(0), tlCifExt.iRotEndDate(1), slDate1
                            gUnpackDate tlInputExt.iRotEndDate(0), tlInputExt.iRotEndDate(1), slDate2
                            If (slDate1 <> "") And (slDate2 <> "") Then
                                If gDateValue(slDate2) > gDateValue(slDate1) Then
                                    tlCifExt = tlInputExt
                                End If
                            ElseIf slDate2 <> "" Then
                                tlCifExt = tlInputExt
                            End If
                        End If
                        ilAddRec = False
                    Else
                        ilAddRec = True
                    End If
                End If
                If ilAddRec Then
                    If (Not ilPurgedOnly) Or ((ilPurgedOnly) And (tlCifExt.sPurged = "P")) Then
                        gUnpackDateForSort tlCifExt.iRotEndDate(0), tlCifExt.iRotEndDate(1), slRotDate
                        If slRotDate = "" Then
                            slRotDate = "000000"
                        End If
                        If (llPriorDate = 0) Or ((Val(slRotDate) < llPriorDate) And (llPriorDate <> -1)) Or ((slRotDate = "000000") And (llPriorDate = -1)) Then
                            gUnpackDateForSort tlCifExt.iPurgeDate(0), tlCifExt.iPurgeDate(1), slPurgeDate
                            If slPurgeDate = "" Then
                                slPurgeDate = "000000"
                            End If
                            If (llPriorDate <> -1) Or ((slPurgeDate = "000000") And (llPriorDate = -1)) Then
                                If tlCifExt.sPurged = "P" Then
                                    slName = slPurgeDate    'slRotDate 'slPurgeDate
                                    slName = "A" & slName
                                Else
                                    If slRotDate <> "000000" Then
                                        slName = slRotDate
                                        slName = "Z" & slName
                                    Else
                                        slName = "999999"
                                        slName = "Z" & slName
                                    End If
                                End If
                                If Trim$(tlCifExt.sCut) = "" Then
                                    slInvName = Trim$(tlCifExt.sName)
                                Else
                                    slInvName = Trim$(tlCifExt.sName) & "-" & Trim$(tlCifExt.sCut)
                                End If
                                If tlCifExt.sPurged = "P" Then
                                    slInvName = slInvName & ": Purged"
                                ElseIf (ilShowRotDates) Then
                                    If slRotDate <> "000000" Then
                                        gUnpackDate tlCifExt.iRotEndDate(0), tlCifExt.iRotEndDate(1), slRotDate
                                        slInvName = slInvName & ": " & slRotDate
                                    Else
                                        slInvName = slInvName & ": Unused"
                                    End If
                                End If
                                If ilSortCart <> 1 Then
                                    slName = slName & "|" & slInvName & "\" & Trim$(str$(tlCifExt.lCode))
                                Else
                                    slName = slInvName & "|" & slInvName & "\" & Trim$(str$(tlCifExt.lCode))
                                End If
                                'If Not gOkAddStrToListBox(slName, llLen, True) Then
                                '    Exit Do
                                'End If
                                'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                                tlSortCode(llSortCode).sKey = slName
                                If llSortCode >= UBound(tlSortCode) Then
                                    ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                                End If
                                llSortCode = llSortCode + 1
                            End If
                        End If
                    End If
                    tlCifExt = tlInputExt
                End If
                ilRet = btrExtGetNext(hlCif, tlInputExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlCif, tlInputExt, ilExtLen, llRecPos)
                Loop
            Loop
            'Add last obtained record
            If tlCifExt.sName <> "" Then
                If (Not ilPurgedOnly) Or ((ilPurgedOnly) And (tlCifExt.sPurged = "P")) Then
                    gUnpackDateForSort tlCifExt.iRotEndDate(0), tlCifExt.iRotEndDate(1), slRotDate
                    If slRotDate = "" Then
                        slRotDate = "000000"
                    End If
                    'If (llPriorDate = 0) Or (Val(slRotDate) < llPriorDate) Then
                    If (llPriorDate = 0) Or ((Val(slRotDate) < llPriorDate) And (llPriorDate <> -1)) Or ((slRotDate = "000000") And (llPriorDate = -1)) Then
                        'gUnpackDateForSort tlCifExt.iPurgeDate(0), tlCifExt.iPurgeDate(1), slPurgeDate
                        'If slPurgeDate = "" Then
                        '    slPurgeDate = "000000"
                        'End If
                        If tlCifExt.sPurged = "P" Then
                            slName = slRotDate 'slPurgeDate
                            slName = "A" & slName
                        Else
                            If slRotDate <> "000000" Then
                                slName = slRotDate
                                slName = "Z" & slName
                            Else
                                slName = "999999"
                                slName = "Z" & slName
                            End If
                        End If
                        If Trim$(tlCifExt.sCut) = "" Then
                            slInvName = Trim$(tlCifExt.sName)
                        Else
                            slInvName = Trim$(tlCifExt.sName) & "-" & Trim$(tlCifExt.sCut)
                        End If
                        If tlCifExt.sPurged = "P" Then
                            slInvName = slInvName & ": Purged"
                        ElseIf ilShowRotDates Then
                            If slRotDate <> "000000" Then
                                gUnpackDate tlCifExt.iRotEndDate(0), tlCifExt.iRotEndDate(1), slRotDate
                                slInvName = slInvName & ": " & slRotDate
                            Else
                                slInvName = slInvName & ": Unused"
                            End If
                        End If
                        If ilSortCart <> 1 Then
                            slName = slName & "|" & slInvName & "\" & Trim$(str$(tlCifExt.lCode))
                        Else
                            slName = slInvName & "|" & slInvName & "\" & Trim$(str$(tlCifExt.lCode))
                        End If
                        'If gOkAddStrToListBox(slName, llLen, True) Then
                        '    lbcMster.AddItem slName
                        'End If
                        tlSortCode(llSortCode).sKey = slName
                        If llSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        llSortCode = llSortCode + 1
                    End If
                End If
                tlCifExt = tlInputExt
            End If
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To llSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If
        ilRet = btrClose(hlCif)
        On Error GoTo gPopCopyForMediaBoxErr
        gBtrvErrorMsg ilRet, "gPopCopyForMediaBox (btrReset):" & "Cif.Btr", frm
        On Error GoTo 0
        btrDestroy hlCif
    End If
    llLen = 0
    For llLoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slName = tlSortCode(llLoop).sKey    'lbcMster.List(llLoop)
        ilRet = gParseItem(slName, 2, "|", slNameCode)
        If ilRet <> CP_MSG_NONE Then
            gPopCopyForMediaBox = CP_MSG_PARSE
            Exit Function
        End If
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopCopyForMediaBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next llLoop
    Exit Function
gPopCopyForMediaBoxErr:
    ilRet = btrClose(hlCif)
    btrDestroy hlCif
    gDbg_HandleError "PopSubs: gPopCopyForMediaBox"
'    gPopCopyForMediaBox = CP_MSG_NOSHOW
'    Exit Function
gPopCopyForMediaErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopEvtNmByTypeBox              *
'*                                                     *
'*             Created:7/29/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with event    *
'*                     type in event record order.     *
'*                                                     *
'*******************************************************
Function gPopEvtNmByTypeBox(frm As Form, ilInUseOnly As Integer, ilSpaceAlso As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopEvtNmByTypeBox (MainForm, ilInUse, ilSpace, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilInUse (I)- True= In use only; False = In use and not in use
'       ilSpace (I)- True=Include Line spaces and page eject; False=bypase line spaces and page eject
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'ETF date/time stamp
    Dim hlEtf As Integer        'ETF handle
    Dim ilRecLen As Integer     'Record length
    Dim tlEtf As ETF
    Dim slName As String
    Dim ilRet As Integer
    Dim slNameCode As String
    Dim illoop As Integer
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Etf.Btr")

    'On Error GoTo gPopEvtNmByTypeErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopEvtNmByTypeBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopEvtNmByTypeBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlEtf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlEtf, "", sgDBPath & "Etf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopEvtNmByTypeBoxErr
        gBtrvErrorMsg ilRet, "gPopEvtNmByTypeBox (btrOpen):" & "Etf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlEtf) 'btrRecordLength(hlEtf)  'Get and save record length
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilRet = btrGetFirst(hlEtf, tlEtf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            If tlEtf.iCode <> 4 Then    'Floating Avail
                If (tlEtf.sInUse = "Y") Or ((tlEtf.sInUse = "N") And (Not ilInUseOnly)) Then
                    If (ilSpaceAlso <> False) Or (Asc(tlEtf.sType) < Asc("A")) Or (Asc(tlEtf.sType) > Asc("D")) Then
                        slName = tlEtf.sType
                        slName = slName & "\" & Trim$(tlEtf.sName)
                        slName = slName & "\" & Trim$(str$(tlEtf.iCode))
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit Do
                        'End If
                        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                    End If
                End If
            End If
            ilRet = btrGetNext(hlEtf, tlEtf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Loop
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            On Error GoTo gPopEvtNmByTypeBoxErr
            gBtrvErrorMsg ilRet, "gPopEvtNmByTypeBox (btrGetFirst):" & "Etf.Btr", frm
            On Error GoTo 0
        End If
        'Sort then output new headers and lines
        ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
            ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
        End If
        ilRet = btrClose(hlEtf)
        On Error GoTo gPopEvtNmByTypeBoxErr
        gBtrvErrorMsg ilRet, "gPopEvtNmByTypeBox (btrReset):" & "Etf.Btr", frm
        On Error GoTo 0
        btrDestroy hlEtf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopEvtNmByTypeBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopEvtNmByTypeBoxErr:
    ilRet = btrClose(hlEtf)
    btrDestroy hlEtf
    gDbg_HandleError "PopSubs: gPopEvtNmByTypeBox"
'    gPopEvtNmByTypeBox = CP_MSG_NOSHOW
'    Exit Function
gPopEvtNmByTypeErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopLibByDateBox                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     Library names that exist for    *
'*                     specified dates (On air/Current)*
'*                                                     *
'*******************************************************
Function gPopLibByDateBox(frm As Form, ilVerType As Integer, slLibType As String, llStartDate As Long, llEndDate As Long, ilVefCode As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopLibByDateBox (MainForm, ilVer, slLibType, ilVEFCode, lStartDate, lEndDate, llStartDate, llEndDate, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilVerType (I)- 0=All; 1= Latest (newest)
'       slLibType (I)- 0=Library type ("R" = Regular; "S" = Special; "P"= Sports; "A" = All types)
'       ilVEFCode (I)- Vehicle code (negative if combo)
'       llStartDate (I)- Start Date to search for library names (zero allowed)
'       llEndDate (I) - End date of search (zero allow->TFN)
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'LVF date/time stamp
    Dim hlLvf As Integer        'LVF handle
    Dim tlLvf As LVF
    Dim ilRecLen As Integer     'Record length
    Dim tlSrchKey As LONGKEY0
    Dim hlLtf As Integer    'LTF handle
    Dim tlLtf As LTF        'Library Title
    Dim ilLtfRecLen As Integer
    Dim tlLtfSrchKey As INTKEY0  'Ltf key record image
    Dim hlLcf As Integer    'LCF handle
    Dim tlLcf As LCF        'Library calendar
    Dim ilLcfRecLen As Integer
    Dim tlLcfSrchKey As LCFKEY0  'LcF key record image
    Dim hlVef As Integer        'VSF handle
    Dim tlVef As VEF
    Dim ilVefRecLen As Integer     'Record length
    Dim tlVefSrchKey As INTKEY0
    Dim hlVsf As Integer        'VSF handle
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length
    Dim slDate As String
    Dim llDate As Long
    Dim slName As String
    Dim slName2 As String
    Dim llVersion As Long
    Dim slVersion As String
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slLibName As String
    Dim slCode As String    'Sales source code number
    Dim ilPos As Integer
    Dim ilVeh As Integer
    Dim ilShift As Integer
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Lvf.Btr") & Trim$(str$(ilVefCode)) & slLibType & Trim$(str$(ilVerType)) & Trim$(str$(llStartDate)) & Trim$(str$(llEndDate))

    'On Error GoTo gPopLibByDateBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopLibByDateBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopLibByDateBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlLvf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopLibByDateBoxErr
        gBtrvErrorMsg ilRet, "gPopLibByDateBox (btrOpen):" & "Lvf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlLvf) 'btrRecordLength(hlLvf)  'Get and save record length
        hlLtf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlLtf, "", sgDBPath & "Ltf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopLibByDateBoxErr
        gBtrvErrorMsg ilRet, "gPopLibByDateBox (btrOpen):" & "Ltf.Btr", frm
        On Error GoTo 0
        ilLtfRecLen = Len(tlLtf) 'btrRecordLength(hlLtf)  'Get and save record length
        hlLcf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopLibByDateBoxErr
        gBtrvErrorMsg ilRet, "gPopLibByDateBox (btrOpen):" & "Lcf.Btr", frm
        On Error GoTo 0
        ilLcfRecLen = Len(tlLcf) 'btrRecordLength(hlLcf)  'Get and save record length
        hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopLibByDateBoxErr
        gBtrvErrorMsg ilRet, "gPopLibByDateBox (btrOpen):" & "Vef.Btr", frm
        On Error GoTo 0
        ilVefRecLen = Len(tlVef) 'btrRecordLength(hlVsf)  'Get and save record length
        hlVsf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopLibByDateBoxErr
        gBtrvErrorMsg ilRet, "gPopLibByDateBox (btrOpen):" & "Vsf.Btr", frm
        On Error GoTo 0
        ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlVsf)  'Get and save record length
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        tlVefSrchKey.iCode = ilVefCode
        ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        On Error GoTo gPopLibByDateBoxErr
        gBtrvErrorMsg ilRet, "mPopLibByDateBoxRec (btrGetEqual): Vef.Btr", frm
        On Error GoTo 0
        If tlVef.sType = "V" Then   'Ignore library name if virtual vehicle
            ilRet = btrClose(hlLvf)
            ilRet = btrClose(hlLtf)
            ilRet = btrClose(hlLcf)
            ilRet = btrClose(hlVsf)
            ilRet = btrClose(hlVef)
            btrDestroy hlLvf
            btrDestroy hlLtf
            btrDestroy hlLcf
            btrDestroy hlVsf
            btrDestroy hlVef
            Exit Function
        End If
        'If ilVefCode > 0 Then
            For illoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                If illoop = LBound(tmVsf.iFSCode) Then
                    tmVsf.iFSCode(illoop) = ilVefCode
                Else
                    tmVsf.iFSCode(illoop) = 0
                End If
            Next illoop
        'Else
        '    tlVsfSrchKey.iCode = -ilVefCode
        '    ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfRecLen, tlVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        '    On Error GoTo gPopLibByDateBoxErr
        '    gBtrvErrorMsg ilRet, "mPopLibByDateBoxRec (btrGetEqual): Vsf.Btr", Frm
        '    On Error GoTo 0
        'End If
        For ilVeh = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
            If tmVsf.iFSCode(ilVeh) > 0 Then
                tlLcfSrchKey.iType = 0    'On air only
                tlLcfSrchKey.sStatus = "C"  'Current only
                tlLcfSrchKey.iVefCode = tmVsf.iFSCode(ilVeh)
                slDate = Format$(llStartDate, "m/d/yy")
                gPackDate slDate, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1)
                tlLcfSrchKey.iSeqNo = 0
                ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
                'Build list of unique library code numbers to minimize number of library reads
                'List is build into the local list box
                Do While (ilRet = BTRV_ERR_NONE) And (tlLcf.iType = 0) And (tlLcf.sStatus = "C") And (tlLcf.iVefCode = tmVsf.iFSCode(ilVeh))
                    gUnpackDate tlLcf.iLogDate(0), tlLcf.iLogDate(1), slDate
                    If (tlLcf.iLogDate(0) <= 7) And (tlLcf.iLogDate(1) = 0) Then
                        llDate = llEndDate + 1
                    Else
                        llDate = gDateValue(slDate)
                    End If
                    If (llEndDate = 0) Or (llDate <= llEndDate) Then
                        For illoop = LBound(tlLcf.lLvfCode) To UBound(tlLcf.lLvfCode) Step 1
                            If tlLcf.lLvfCode(illoop) > 0 Then
                                slCode = Trim$(str$(tlLcf.lLvfCode(illoop)))
                                Do While Len(slCode) < 9
                                    slCode = "0" & slCode
                                Loop
                                gFindMatch slCode, 0, lbcLocal
                                If gLastFound(lbcLocal) < 0 Then
                                    lbcLocal.AddItem slCode
                                End If
                            End If
                        Next illoop
                    Else
                        Exit Do
                    End If
                    ilRet = btrGetNext(hlLcf, tlLcf, ilLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_KEY_NOT_FOUND) Then
                    On Error GoTo gPopLibByDateBoxErr
                    gBtrvErrorMsg ilRet, "mPopLibByDateBoxRec (btrGetNext): Lcf.Btr", frm
                    On Error GoTo 0
                End If
            End If
        Next ilVeh
        'Read in each library and check version numbers
        For illoop = 0 To lbcLocal.ListCount - 1 Step 1
            tlSrchKey.lCode = CLng(lbcLocal.List(illoop))
            ilRet = btrGetEqual(hlLvf, tlLvf, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo gPopLibByDateBoxErr
            gBtrvErrorMsg ilRet, "mPopLibByDateBoxRec (btrGetEqual): Lvf.Btr", frm
            On Error GoTo 0
            tlLtfSrchKey.iCode = tlLvf.iLtfCode
            ilRet = btrGetEqual(hlLtf, tlLtf, ilLtfRecLen, tlLtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            On Error GoTo gPopLibByDateBoxErr
            gBtrvErrorMsg ilRet, "mPopLibByDateBoxRec (btrGetEqual): Ltf.Btr", frm
            On Error GoTo 0
            If (tlLtf.sType = slLibType) Or (slLibType = "A") Then
                slName = Trim$(tlLtf.sName) & "|"
                llVersion = 999999 - tlLvf.iVersion 'Descending order
                slVersion = Trim$(str$(llVersion))
                Do While Len(slVersion) < 6
                    slVersion = "0" & slVersion
                Loop
                slName = slName & slVersion
                slName = slName & "\" & Trim$(str$(tlLtf.iCode)) 'Trim$(Str$(tlLvf.lCode))
                'If Not gOkAddStrToListBox(slName, llLen, True) Then
                '    Exit For
                'End If
                'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                tlSortCode(ilSortCode).sKey = slName
                If ilSortCode >= UBound(tlSortCode) Then
                    ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                End If
                ilSortCode = ilSortCode + 1
            End If
        Next illoop
        'Sort then output new headers and lines
        ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
            ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
        End If
        lbcLocal.Clear
        If ilVerType = 1 Then
            illoop = UBound(tlSortCode) - 1
            Do While illoop >= ilLowLimit
            'For IlLoop = UBound(tlSortCode) - 1 To 1 Step -1
                slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
                ilRet = gParseItem(slNameCode, 1, "|", slName)
                'TTP 10310 - Rate card screen, Dayparts: Subscript out of range when pressing the Daypart by "Libraries" button
                slNameCode = ""
                If illoop > 0 Then
                    slNameCode = tlSortCode(illoop - 1).sKey  'lbcMster.List(ilLoop - 1)
                End If
                ilRet = gParseItem(slNameCode, 1, "|", slName2)
                If StrComp(slName, slName2, 1) = 0 Then
                    'lbcMster.RemoveItem ilLoop
                    For ilShift = illoop To UBound(tlSortCode) - 1 Step 1
                        tlSortCode(ilShift).sKey = tlSortCode(ilShift + 1).sKey
                    Next ilShift
                    ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) - 1) As SORTCODE
                End If
                illoop = illoop - 1
            'Next IlLoop
            Loop
        End If
        ilRet = btrClose(hlVef)
        On Error GoTo gPopLibByDateBoxErr
        gBtrvErrorMsg ilRet, "gPopLibByDateBox (btrReset):" & "Vef.Btr", frm
        On Error GoTo 0
        btrDestroy hlVef
        ilRet = btrClose(hlVsf)
        On Error GoTo gPopLibByDateBoxErr
        gBtrvErrorMsg ilRet, "gPopLibByDateBox (btrReset):" & "Vsf.Btr", frm
        On Error GoTo 0
        btrDestroy hlVsf
        ilRet = btrClose(hlLcf)
        On Error GoTo gPopLibByDateBoxErr
        gBtrvErrorMsg ilRet, "gPopLibByDateBox (btrReset):" & "Lcf.Btr", frm
        On Error GoTo 0
        btrDestroy hlLcf
        ilRet = btrClose(hlLtf)
        On Error GoTo gPopLibByDateBoxErr
        gBtrvErrorMsg ilRet, "gPopLibByDateBox (btrReset):" & "Ltf.Btr", frm
        On Error GoTo 0
        btrDestroy hlLtf
        ilRet = btrClose(hlLvf)
        On Error GoTo gPopLibByDateBoxErr
        gBtrvErrorMsg ilRet, "gPopLibByDateBox (btrReset):" & "Lvf.Btr", frm
        On Error GoTo 0
        btrDestroy hlLvf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopLibByDateBox = CP_MSG_PARSE
            Exit Function
        End If
        ilPos = InStr(slName, "|")
        If ilPos = 0 Then
            slName = Trim$(slName)
        Else
            ilRet = gParseItem(slName, 1, "|", slLibName)
            If ilRet <> CP_MSG_NONE Then
                gPopLibByDateBox = CP_MSG_PARSE
                Exit Function
            End If
            If ilVerType = 1 Then
                slName = slLibName
            Else
                ilRet = gParseItem(slName, 2, "|", slCode)
                If ilRet <> CP_MSG_NONE Then
                    gPopLibByDateBox = CP_MSG_PARSE
                    Exit Function
                End If
                llVersion = 999999 - CLng(slCode)
                slName = slLibName & "/" & Trim$(str$(llVersion))
            End If
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopLibByDateBoxErr:
    ilRet = btrClose(hlLvf)
    ilRet = btrClose(hlLtf)
    ilRet = btrClose(hlLcf)
    ilRet = btrClose(hlVsf)
    ilRet = btrClose(hlVef)
    btrDestroy hlLvf
    btrDestroy hlLtf
    btrDestroy hlLcf
    btrDestroy hlVsf
    btrDestroy hlVef
    gDbg_HandleError "PopSubs: gPopLibByDateBox"
'    gPopLibByDateBox = CP_MSG_NOSHOW
'    Exit Function
gPopLibByDateBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopMnfPlusFieldsBox            *
'*                                                     *
'*             Created:7/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     names from specified multi-name *
'*                                                     *
'*******************************************************
Function gPopMnfPlusFieldsBox(frm As Form, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, slMnfType As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slHardCost                                                                            *
'******************************************************************************************

'
'   ilRet = gPopMnfPlusFieldsBox (MainForm, lbcLocal, tlSortCode(), slSortCodeTag, sType)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       sType (I)- Which Mnf field is to be populated (I, INH, A, B, S, T, R, M, C, N, NNS, NOS, Y, YW, D, DS, DC, F, FG, 1, 2)
'                  where:
'                         INH get NTR type, exclude Hard Cost
'                         NNS get feed types exclusing subfeeds
'                         NOS get feed types that are subfeeds only
'                         YW transaction types with W- as type (unittype)
'                         D  get standard and custom demo
'                         DS get standard demos only
'                         DC get custom demos only
'                         FNG Exclude group name (items sorted with group name)
'                         FG Group Name (Exclude description)
'                         Hn get Vehicle Group Set which match n (number)
'                            If H or H0, Set Name included
'                         1 get Event Subtotal 1
'                         2 get Event subtotal 2
'       ilRet (O)- Error code (0 if no error)
'
'       Field #  Description
'         1      Name
'         2      Code
'         3      RPU (I or A) or SSComm (S) or UnitType (C)
'         4      UnitType (I) or Group Number (A or S)
'         5      SSComm (I)
'         6      Taxable (I)
'                GroupNo (D)
'
    Dim slStamp As String    'MNF date/time stamp
    Dim hlMnf As Integer        'MNF handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlMnf As MNF
    Dim slName As String
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim slNameCode As String
    Dim illoop As Integer
    Dim slStr As String
    Dim slSlspComm As String
    Dim slTax As String
    Dim slAcqCost As String
    Dim ilOffSet As Integer
    Dim ilAddRec As Integer
    Dim llLen As Long
    Dim ilPos As Integer
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Mnf.Btr") & slMnfType

    'On Error GoTo gPopMnfPlusFieldsBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopMnfPlusFieldsBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopMnfPlusFieldsBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        tlCharTypeBuff.sType = slMnfType    'Get left most character only
        If tlCharTypeBuff.sType = "D" Then
            ilRet = gObtainMnfForType(tlCharTypeBuff.sType, sgDemoMnfStamp, tgDemoMnf())
            If ilRet = False Then
                Exit Function
            End If
            For illoop = LBound(tgDemoMnf) To UBound(tgDemoMnf) - 1 Step 1
                tlMnf = tgDemoMnf(illoop)
                '6/4/16: Replaced GoSub
                'GoSub AddRecSub
                mAddRecSub1 tlMnf, slMnfType, slName, ilSortCode, tlSortCode()
            Next illoop
        ElseIf tlCharTypeBuff.sType = "F" Then
            ilRet = gObtainMnfForType(tlCharTypeBuff.sType, sgSocEcoMnfStamp, tgSocEcoMnf())
            If ilRet = False Then
                Exit Function
            End If
            For illoop = LBound(tgSocEcoMnf) To UBound(tgSocEcoMnf) - 1 Step 1
                tlMnf = tgSocEcoMnf(illoop)
                '6/4/16: Replaced GoSub
                'GoSub AddRecSub
                mAddRecSub1 tlMnf, slMnfType, slName, ilSortCode, tlSortCode()
            Next illoop
        Else
            hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
            ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            On Error GoTo gPopMnfPlusFieldsBoxErr
            gBtrvErrorMsg ilRet, "gPopMnfPlusFieldsBox (btrOpen):" & "Mnf.Btr", frm
            On Error GoTo 0
            ilRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
            llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlMnf) 'Obtain number of records
            btrExtClear hlMnf   'Clear any previous extend operation
            ilRet = btrGetFirst(hlMnf, tlMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_END_OF_FILE Then
                ilRet = btrClose(hlMnf)
                On Error GoTo gPopMnfPlusFieldsBoxErr
                gBtrvErrorMsg ilRet, "gPopMnfPlusFieldsBox (btrReset):" & "Mnf.Btr", frm
                On Error GoTo 0
                btrDestroy hlMnf
                Exit Function
            Else
                On Error GoTo gPopMnfPlusFieldsBoxErr
                gBtrvErrorMsg ilRet, "gPopMnfPlusFieldsBox (btrGetFirst):" & "Mnf.Btr", frm
                On Error GoTo 0
            End If
            'Extract information from first record
            Call btrExtSetBounds(hlMnf, llNoRec, -1, "UC", "MNF", "") 'Set extract limits (all records)
            tlCharTypeBuff.sType = slMnfType    'Get left most character only
            ilOffSet = gFieldOffset("Mnf", "MnfType")
            ilRet = btrExtAddLogicConst(hlMnf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
            ilRet = btrExtAddField(hlMnf, 0, ilRecLen)  'Extract all fields
            On Error GoTo gPopMnfPlusFieldsBoxErr
            gBtrvErrorMsg ilRet, "gPopMnfPlusFieldsBox (btrExtAddField):" & "Mnf.Btr", frm
            On Error GoTo 0
            'ilRet = btrExtGetNextExt(hlMnf)    'Extract record
            ilRet = btrExtGetNext(hlMnf, tlMnf, ilRecLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo gPopMnfPlusFieldsBoxErr
                gBtrvErrorMsg ilRet, "gPopMnfPlusFieldsBox (btrExtGetNextExt):" & "Mnf.Btr", frm
                On Error GoTo 0
                'ilRet = btrExtGetFirst(hlMnf, tlMnf, ilRecLen, llRecPos)
                ilRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlMnf, tlMnf, ilRecLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    '6/4/16: Replaced GoSub
                    'GoSub AddRecSub
                    mAddRecSub1 tlMnf, slMnfType, slName, ilSortCode, tlSortCode()
                    ilRet = btrExtGetNext(hlMnf, tlMnf, ilRecLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlMnf, tlMnf, ilRecLen, llRecPos)
                    Loop
                Loop
            End If
            ilRet = btrClose(hlMnf)
            On Error GoTo gPopMnfPlusFieldsBoxErr
            gBtrvErrorMsg ilRet, "gPopMnfPlusFieldsBox (btrReset):" & "Mnf.Btr", frm
            On Error GoTo 0
            btrDestroy hlMnf
        End If
        'Sort then output new headers and lines
        ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
            ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
        End If
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopMnfPlusFieldsBox = CP_MSG_PARSE
            Exit Function
        End If
        '5/11/10:  Removed group number from sort
        '5/7/10:  Group number changed from 2 to 3
        If slMnfType = "D" Then 'Remove Group Sort Number
            ''slName = Mid$(slName, 3)
            'slName = Mid$(slName, 4)
            slName = Mid$(slName, 2)
        End If
        If slMnfType = "FNG" Then 'Remove Group Name
            ilPos = InStr(slName, " ")
            If ilPos > 0 Then
                slName = Mid$(slName, ilPos + 1)
            End If
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopMnfPlusFieldsBoxErr:
    ilRet = btrClose(hlMnf)
    btrDestroy hlMnf
    gDbg_HandleError "PopSubs: gPopMnfPlusFieldsBox"

gPopMnfPlusFieldsBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopMnfPlusFieldsBoxNoForm            *
'*                                                     *
'*             Created:7/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     names from specified multi-name *
'*                                                     *
'*******************************************************
Function gPopMnfPlusFieldsBoxNoForm(lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, slMnfType As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slHardCost                                                                            *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  gPopMnfPlusFieldsBoxNoFormErr                                                         *
'******************************************************************************************

'
'   ilRet = gPopMnfPlusFieldsBoxNoForm (MainForm, lbcLocal, tlSortCode(), slSortCodeTag, sType)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       sType (I)- Which Mnf field is to be populated (I, INH, A, B, S, T, R, M, C, N, NNS, NOS, Y, YW, D, DS, DC, F, FG, 1, 2)
'                  where:
'                         INH get NTR type, exclude Hard Cost
'                         NNS get feed types exclusing subfeeds
'                         NOS get feed types that are subfeeds only
'                         YW transaction types with W- as type (unittype)
'                         D  get standard and custom demo
'                         DS get standard demos only
'                         DC get custom demos only
'                         FNG Exclude group name (items sorted with group name)
'                         FG Group Name (Exclude description)
'                         Hn get Vehicle Group Set which match n (number)
'                            If H or H0, Set Name included
'                         1 get Event Subtotal 1
'                         2 get Event subtotal 2
'       ilRet (O)- Error code (0 if no error)
'
'       Field #  Description
'         1      Name
'         2      Code
'         3      RPU (I or A) or SSComm (S) or UnitType (C)
'         4      UnitType (I) or Group Number (A or S)
'         5      SSComm (I)
'         6      Taxable (I)
'                GroupNo (D)
'
    Dim slStamp As String    'MNF date/time stamp
    Dim hlMnf As Integer        'MNF handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlMnf As MNF
    Dim slName As String
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim slNameCode As String
    Dim illoop As Integer
    Dim slStr As String
    Dim slSlspComm As String
    Dim slTax As String
    Dim slAcqCost As String
    Dim ilOffSet As Integer
    Dim ilAddRec As Integer
    Dim llLen As Long
    Dim ilPos As Integer
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Mnf.Btr") & slMnfType

    'On Error GoTo gPopMnfPlusFieldsBoxNoFormErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopMnfPlusFieldsBoxNoForm = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopMnfPlusFieldsBoxNoForm = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        tlCharTypeBuff.sType = slMnfType    'Get left most character only
        If tlCharTypeBuff.sType = "D" Then
            ilRet = gObtainMnfForType(tlCharTypeBuff.sType, sgDemoMnfStamp, tgDemoMnf())
            If ilRet = False Then
                Exit Function
            End If
            For illoop = LBound(tgDemoMnf) To UBound(tgDemoMnf) - 1 Step 1
                tlMnf = tgDemoMnf(illoop)
                '6/4/16: Replace GoSub
                'GoSub AddRecSub
                mAddRecSub2 tlMnf, slMnfType, slName, ilSortCode, tlSortCode()
            Next illoop
        ElseIf tlCharTypeBuff.sType = "F" Then
            ilRet = gObtainMnfForType(tlCharTypeBuff.sType, sgSocEcoMnfStamp, tgSocEcoMnf())
            If ilRet = False Then
                Exit Function
            End If
            For illoop = LBound(tgSocEcoMnf) To UBound(tgSocEcoMnf) - 1 Step 1
                tlMnf = tgSocEcoMnf(illoop)
                '6/4/16: Replace GoSub
                'GoSub AddRecSub
                mAddRecSub2 tlMnf, slMnfType, slName, ilSortCode, tlSortCode()
            Next illoop
        Else
            hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
            ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            'On Error GoTo gPopMnfPlusFieldsBoxNoFormErr
            'gBtrvErrorMsg ilRet, "gPopMnfPlusFieldsBoxNoForm (btrOpen):" & "Mnf.Btr", Frm
            'On Error GoTo 0

            ilRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
            llNoRec = gExtNoRec(ilRecLen) 'btrRecords(hlMnf) 'Obtain number of records
            btrExtClear hlMnf   'Clear any previous extend operation
            ilRet = btrGetFirst(hlMnf, tlMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_END_OF_FILE Then
                ilRet = btrClose(hlMnf)
                'On Error GoTo gPopMnfPlusFieldsBoxNoFormErr
                'gBtrvErrorMsg ilRet, "gPopMnfPlusFieldsBoxNoForm (btrReset):" & "Mnf.Btr", Frm
                'On Error GoTo 0
                btrDestroy hlMnf
                Exit Function
            Else
                'On Error GoTo gPopMnfPlusFieldsBoxNoFormErr
                'gBtrvErrorMsg ilRet, "gPopMnfPlusFieldsBoxNoForm (btrGetFirst):" & "Mnf.Btr", Frm
                'On Error GoTo 0
            End If
            'Extract information from first record

            Call btrExtSetBounds(hlMnf, llNoRec, -1, "UC", "MNF", "") 'Set extract limits (all records)
            tlCharTypeBuff.sType = slMnfType    'Get left most character only
            ilOffSet = gFieldOffset("Mnf", "MnfType")
            ilRet = btrExtAddLogicConst(hlMnf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
            ilRet = btrExtAddField(hlMnf, 0, ilRecLen)  'Extract all fields
            'On Error GoTo gPopMnfPlusFieldsBoxNoFormErr
            'gBtrvErrorMsg ilRet, "gPopMnfPlusFieldsBoxNoForm (btrExtAddField):" & "Mnf.Btr", Frm
            'On Error GoTo 0
            ''ilRet = btrExtGetNextExt(hlMnf)    'Extract record
            ilRet = btrExtGetNext(hlMnf, tlMnf, ilRecLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                'On Error GoTo gPopMnfPlusFieldsBoxNoFormErr
                'gBtrvErrorMsg ilRet, "gPopMnfPlusFieldsBoxNoForm (btrExtGetNextExt):" & "Mnf.Btr", Frm
                'On Error GoTo 0
                ''ilRet = btrExtGetFirst(hlMnf, tlMnf, ilRecLen, llRecPos)
                ilRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlMnf, tlMnf, ilRecLen, llRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    '6/4/16: Replace GoSub
                    'GoSub AddRecSub
                    mAddRecSub2 tlMnf, slMnfType, slName, ilSortCode, tlSortCode()
                    ilRet = btrExtGetNext(hlMnf, tlMnf, ilRecLen, llRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlMnf, tlMnf, ilRecLen, llRecPos)
                    Loop
                Loop
            End If
            ilRet = btrClose(hlMnf)
            'On Error GoTo gPopMnfPlusFieldsBoxNoFormErr
            'gBtrvErrorMsg ilRet, "gPopMnfPlusFieldsBoxNoForm (btrReset):" & "Mnf.Btr", Frm
            'On Error GoTo 0
            btrDestroy hlMnf
        End If
        'Sort then output new headers and lines
        ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
            ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
        End If
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopMnfPlusFieldsBoxNoForm = CP_MSG_PARSE
            Exit Function
        End If
        '5/11/10:  Remove Group number from sort
        '5/7/10:  Group number changed from 2 to 3
        If slMnfType = "D" Then 'Remove Group Sort Number
            ''slName = Mid$(slName, 3)
            'slName = Mid$(slName, 4)
            slName = Mid$(slName, 2)
        End If
        If slMnfType = "FNG" Then 'Remove Group Name
            ilPos = InStr(slName, " ")
            If ilPos > 0 Then
                slName = Mid$(slName, ilPos + 1)
            End If
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopMnfPlusFieldsBoxNoFormErr: 'VBC NR
    ilRet = btrClose(hlMnf)
    btrDestroy hlMnf
    gDbg_HandleError "PopSubs: gPopMnfPlusFieldsBoxNoForm"

gPopMnfPlusFieldsBoxNoFormErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopOfficePlus  Box             *
'*                                                     *
'*             Created:6/03/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with sales    *
'*                     source and sales office names   *
'*                     Like gPopOfficeOfficeBox except *
'*                     Rank and State added            *
'*                                                     *
'*******************************************************
Function gPopOfficePlusBox(frm As Form, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopOfficePlusBox (MainForm, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       tlSortCode (I/O)- Sorted List containing Rank|Office name/Source Name|state\code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slSofStamp As String    'Sof date/time stamp
    Dim slMNmStamp As String    'Mnf date/time stamp
    Dim slStamp As String       'Concat. stamp
    Dim hlSof As Integer        'Sof handle
    Dim ilSofRecLen As Integer     'Record length
    Dim hlMnf As Integer        'Mnf handle
    Dim ilMnfRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlSof As SOF
    Dim tlMnf As MNF
    Dim slName As String
    Dim tlMnfSrchKey As INTKEY0
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim slMktRank As String
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slSofStamp = gFileDateTime(sgDBPath & "sof.btr")
    slMNmStamp = gFileDateTime(sgDBPath & "Mnf.btr")
    slStamp = slSofStamp & " \ " & slMNmStamp

    'On Error GoTo gPopOfficePlusErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            'If lbcLocal.ListCount > 0 Then
                gPopOfficePlusBox = CP_MSG_NOPOPREQ
                Exit Function
            'End If
            'ilPop = False
        End If
    End If
    gPopOfficePlusBox = CP_MSG_POPREQ
    'lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlSof = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopOfficePlusBoxErr
        gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrOpen):" & "Sof.Btr", frm
        On Error GoTo 0
        ilSofRecLen = Len(tlSof) 'btrRecordLength(hlSof)  'Get and save record length
        hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopOfficePlusBoxErr
        gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrOpen):" & "Sof.Btr", frm
        On Error GoTo 0
        ilMnfRecLen = Len(tlMnf)  'btrRecordLength(hlMnf)  'Get and save record length
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlSof)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSof) 'Obtain number of records
        btrExtClear hlSof   'Clear any previous extend operation
        ilRet = btrGetFirst(hlSof, tlSof, ilSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlSof)
            On Error GoTo gPopOfficePlusBoxErr
            gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrReset):" & "Sof.Btr", frm
            On Error GoTo 0
            btrDestroy hlSof
            ilRet = btrClose(hlMnf)
            On Error GoTo gPopOfficePlusBoxErr
            gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrReset):" & "Mnf.Btr", frm
            On Error GoTo 0
            btrDestroy hlMnf
            Exit Function
        Else
            On Error GoTo gPopOfficePlusBoxErr
            gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrGetFirst):" & "Sof.Btr", frm
            On Error GoTo 0
        End If
        'Extract information from first record
    'Use UC (get cuurent) in setbounds instead of EG (bypass current)
    '    tlMnfSrchKey.iCode = tlSof.iMnfSSCode
    '    ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '    On Error GoTo gPopOfficePlusBoxErr
    '    gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrGetEqual):" & "Mnf.Btr", Frm
    '    On Error GoTo 0
    '    slName = Trim$(tlSof.sName) & "/" & Trim$(tlMnf.sName)
    '    slName = slName & "\" & Trim$(Str$(tlSof.iCode))
    '    lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
        Call btrExtSetBounds(hlSof, llNoRec, -1, "UC", "SOF", "") 'Set extract limits (all records)
        'ilOffset = gFieldOffset("Sof", "SofCode")
        'ilRet = btrExtAddField(hlSof, ilOffset, 2)  'Extract iCode field
        'On Error GoTo gPopOfficePlusBoxErr
        'gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrExtAddField):" & "Sof.Btr", Frm
        'On Error GoTo 0
        'ilOffset = gFieldOffset("Sof", "SofMnfSSCode")
        'ilRet = btrExtAddField(hlSof, ilOffset, 2)  'Extract iMnfSSCode field
        'On Error GoTo gPopOfficePlusBoxErr
        'gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrExtAddField):" & "Sof.Btr", Frm
        'On Error GoTo 0
        'ilOffset = gFieldOffset("Sof", "SofName")
        'ilRet = btrExtAddField(hlSof, ilOffset, 20) 'Extract Name field
        'On Error GoTo gPopOfficePlusBoxErr
        'gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrExtAddField):" & "Sof.Btr", Frm
        'On Error GoTo 0
        ilOffSet = 0
        ilRet = btrExtAddField(hlSof, ilOffSet, ilExtLen)  'Extract iCode field
        On Error GoTo gPopOfficePlusBoxErr
        gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrExtAddField):" & "Sof.Btr", frm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlSof, tlSof, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopOfficePlusBoxErr
            gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrExtGetNextExt):" & "Sof.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlSof)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSof, tlSof, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tlMnfSrchKey.iCode = tlSof.iMnfSSCode
                ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                On Error GoTo gPopOfficePlusBoxErr
                gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrGetEqual):" & "Mnf.Btr", frm
                On Error GoTo 0
                slName = Trim$(tlSof.sName) & "/" & Trim$(tlMnf.sName)
                slMktRank = Trim$(str$(tlSof.iMktRank))
                Do While Len(slMktRank) < 5
                    slMktRank = "0" & slMktRank
                Loop
                If tlSof.sState = "D" Then
                    slName = slMktRank & "|" & slName & "|D\" & Trim$(str$(tlSof.iCode))
                Else
                    slName = slMktRank & "|" & slName & "|A\" & Trim$(str$(tlSof.iCode))
                End If
                'If Not gOkAddStrToListBox(slName, llLen, True) Then
                '    Exit Do
                'End If
                'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                tlSortCode(ilSortCode).sKey = slName
                If ilSortCode >= UBound(tlSortCode) Then
                    ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                End If
                ilSortCode = ilSortCode + 1
                ilRet = btrExtGetNext(hlSof, tlSof, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSof, tlSof, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If
        ilRet = btrClose(hlSof)
        On Error GoTo gPopOfficePlusBoxErr
        gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrReset):" & "Sof.Btr", frm
        On Error GoTo 0
        btrDestroy hlSof
        ilRet = btrClose(hlMnf)
        On Error GoTo gPopOfficePlusBoxErr
        gBtrvErrorMsg ilRet, "gPopOfficePlusBox (btrReset):" & "Mnf.Btr", frm
        On Error GoTo 0
        btrDestroy hlMnf
    End If
    'llLen = 0
    'For ilLoop = 0 To UBound(tlSortCode) - 1 Step 1
    '    slNameCode = tlSortCode(ilLoop).sKey    'lbcMster.List(ilLoop)
    '    ilRet = gParseItem(slNameCode, 1, "\", slStr)
    '    If ilRet <> CP_MSG_NONE Then
    '        gPopOfficePlusBox = CP_MSG_PARSE
    '        Exit Function
    '    End If
    '    ilRet = gParseItem(slStr, 2, "|", slName)
    '    slName = Trim$(slName)
    '    If Not gOkAddStrToListBox(slName, llLen, True) Then
    '        Exit For
    '    End If
    '    lbcLocal.AddItem slName  'Add ID to list box
    'Next ilLoop
    Exit Function
gPopOfficePlusBoxErr:
    ilRet = btrClose(hlSof)
    ilRet = btrClose(hlMnf)
    btrDestroy hlSof
    btrDestroy hlMnf
    gDbg_HandleError "PopSubs: gPopOfficePlusBox"
'    gPopOfficePlusBox = CP_MSG_NOSHOW
'    Exit Function
gPopOfficePlusErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopOfficeSourceBox             *
'*                                                     *
'*             Created:6/03/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with sales    *
'*                     source and sales office names   *
'*                                                     *
'*******************************************************
Function gPopOfficeSourceBox(frm As Form, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopOfficeSourceBox (MainForm, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slSofStamp As String    'Sof date/time stamp
    Dim slMNmStamp As String    'Mnf date/time stamp
    Dim slStamp As String       'Concat. stamp
    Dim hlSof As Integer        'Sof handle
    Dim ilSofRecLen As Integer     'Record length
    Dim hlMnf As Integer        'Mnf handle
    Dim ilMnfRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlSof As SOF
    Dim tlMnf As MNF
    Dim tlSofMnfExt As SOFMNFEXT    'Sales office and sales source code extract record
    Dim slName As String
    Dim tlMnfSrchKey As INTKEY0
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slSofStamp = gFileDateTime(sgDBPath & "sof.btr")
    slMNmStamp = gFileDateTime(sgDBPath & "Mnf.btr")
    slStamp = slSofStamp & " \ " & slMNmStamp

    'On Error GoTo gPopOfficeSourceErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopOfficeSourceBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopOfficeSourceBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlSof = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopOfficeSourceBoxErr
        gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrOpen):" & "Sof.Btr", frm
        On Error GoTo 0
        ilSofRecLen = Len(tlSof) 'btrRecordLength(hlSof)  'Get and save record length
        hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopOfficeSourceBoxErr
        gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrOpen):" & "Sof.Btr", frm
        On Error GoTo 0
        ilMnfRecLen = Len(tlMnf)  'btrRecordLength(hlMnf)  'Get and save record length
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlSofMnfExt)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSof) 'Obtain number of records
        btrExtClear hlSof   'Clear any previous extend operation
        ilRet = btrGetFirst(hlSof, tlSof, ilSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlSof)
            On Error GoTo gPopOfficeSourceBoxErr
            gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrReset):" & "Sof.Btr", frm
            On Error GoTo 0
            btrDestroy hlSof
            ilRet = btrClose(hlMnf)
            On Error GoTo gPopOfficeSourceBoxErr
            gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrReset):" & "Mnf.Btr", frm
            On Error GoTo 0
            btrDestroy hlMnf
            Exit Function
        Else
            On Error GoTo gPopOfficeSourceBoxErr
            gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrGetFirst):" & "Sof.Btr", frm
            On Error GoTo 0
        End If
        'Extract information from first record
    'Use UC (get cuurent) in setbounds instead of EG (bypass current)
    '    tlMnfSrchKey.iCode = tlSof.iMnfSSCode
    '    ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '    On Error GoTo gPopOfficeSourceBoxErr
    '    gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrGetEqual):" & "Mnf.Btr", Frm
    '    On Error GoTo 0
    '    slName = Trim$(tlSof.sName) & "/" & Trim$(tlMnf.sName)
    '    slName = slName & "\" & Trim$(Str$(tlSof.iCode))
    '    lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
        Call btrExtSetBounds(hlSof, llNoRec, -1, "UC", "SOFMNFEXTPK", SOFMNFEXTPK) 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Sof", "SofCode")
        ilRet = btrExtAddField(hlSof, ilOffSet, 2)  'Extract iCode field
        On Error GoTo gPopOfficeSourceBoxErr
        gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrExtAddField):" & "Sof.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Sof", "SofMnfSSCode")
        ilRet = btrExtAddField(hlSof, ilOffSet, 2)  'Extract iMnfSSCode field
        On Error GoTo gPopOfficeSourceBoxErr
        gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrExtAddField):" & "Sof.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Sof", "SofName")
        ilRet = btrExtAddField(hlSof, ilOffSet, 20) 'Extract Name field
        On Error GoTo gPopOfficeSourceBoxErr
        gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrExtAddField):" & "Sof.Btr", frm
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hlSof)    'Extract record
        ilRet = btrExtGetNext(hlSof, tlSofMnfExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopOfficeSourceBoxErr
            gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrExtGetNextExt):" & "Sof.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlSofMnfExt)  'Extract operation record size
            'ilRet = btrExtGetFirst(hlSof, tlSofMnfExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSof, tlSofMnfExt, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tlMnfSrchKey.iCode = tlSofMnfExt.iMnfSSCode
                ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                On Error GoTo gPopOfficeSourceBoxErr
                gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrGetEqual):" & "Mnf.Btr", frm
                On Error GoTo 0
                slName = Trim$(tlSofMnfExt.sSOFName) & "/" & Trim$(tlMnf.sName)
                slName = slName & "\" & Trim$(str$(tlSofMnfExt.iCode))
                'If Not gOkAddStrToListBox(slName, llLen, True) Then
                '    Exit Do
                'End If
                'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                tlSortCode(ilSortCode).sKey = slName
                If ilSortCode >= UBound(tlSortCode) Then
                    ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                End If
                ilSortCode = ilSortCode + 1
                ilRet = btrExtGetNext(hlSof, tlSofMnfExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSof, tlSofMnfExt, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If
        ilRet = btrClose(hlSof)
        On Error GoTo gPopOfficeSourceBoxErr
        gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrReset):" & "Sof.Btr", frm
        On Error GoTo 0
        btrDestroy hlSof
        ilRet = btrClose(hlMnf)
        On Error GoTo gPopOfficeSourceBoxErr
        gBtrvErrorMsg ilRet, "gPopOfficeSourceBox (btrReset):" & "Mnf.Btr", frm
        On Error GoTo 0
        btrDestroy hlMnf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopOfficeSourceBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopOfficeSourceBoxErr:
    ilRet = btrClose(hlSof)
    ilRet = btrClose(hlMnf)
    btrDestroy hlSof
    btrDestroy hlMnf
    gDbg_HandleError "PopSubs: gPopOfficeSourceBox"
'    gPopOfficeSourceBox = CP_MSG_NOSHOW
'    Exit Function
gPopOfficeSourceErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopPersonnelBox                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     Personnel names                 *
'*                                                     *
'*******************************************************
Function gPopPersonnelBox(frm As Form, ilAdvtOrAgy As Integer, ilCode As Integer, slType As String, ilIncludeDormant As Integer, ilShow As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopPersonnelBox (MainForm, ilAdvtOrAgy, ilCode, slType, ilIncludeDormant, ilShow, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAdvtOrAgy(I)- 0=retrieve Advertiser; 1=retrieve Agency
'       ilCode (I)- Advertise or Agency code value (-1 for all)
'       slType(I)- B=Buyer Names; P=Payable Contact Names
'       ilIncludeDormant(I)- Include Dormant
'       ilShow(I)- 0=Name only; 1=Name/Phone #, 2=Name/Phone #/Fax #
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'Adf date/time stamp
    Dim hlPnf As Integer        'Adf handle
    Dim ilRecLen As Integer     'Record length
    Dim tlPnf As PNF
    Dim slName As String
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim slStr As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlSrchKey1 As PNFKEY1  'Pnf key record image
    Dim tlSrchKey2 As PNFKEY2  'Pnf key record image
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Pnf.Btr") & Trim$(str$(ilAdvtOrAgy)) & Trim$(str$(ilCode)) & slType

    'On Error GoTo gPopPersonnelBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopPersonnelBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopPersonnelBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlPnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlPnf, "", sgDBPath & "Pnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopPersonnelBoxErr
        gBtrvErrorMsg ilRet, "gPopPersonnelBox (btrOpen): Pnf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlPnf) 'btrRecordLength(hlPnf)  'Get and save record length
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlPnf)
        llNoRec = gExtNoRec(ilExtLen)
        btrExtClear hlPnf
        Call btrExtSetBounds(hlPnf, llNoRec, -1, "UC", "PNF", "") 'Set extract limits (all records)
        If ilAdvtOrAgy = 0 Then
            If ilCode <= 0 Then
                tlSrchKey2.iAdfCode = 0
            Else
                tlSrchKey2.iAdfCode = ilCode
            End If
            ilRet = btrGetGreaterOrEqual(hlPnf, tlPnf, ilRecLen, tlSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
            If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
                ilRet = btrClose(hlPnf)
                On Error GoTo gPopPersonnelBoxErr
                gBtrvErrorMsg ilRet, "gPopPersonnelBox (btrGetReaterOrEqual):" & "Pnf.Btr", frm
                On Error GoTo 0
                btrDestroy hlPnf
                Exit Function
            End If
            If ilCode >= 0 Then
                tlIntTypeBuff.iType = ilCode
                ilOffSet = gFieldOffset("Pnf", "PnfAdfCode")
                ilRet = btrExtAddLogicConst(hlPnf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
            End If
        Else
            If ilCode <= 0 Then
                tlSrchKey1.iAgfCode = 0
            Else
                tlSrchKey1.iAgfCode = ilCode
            End If
            ilRet = btrGetGreaterOrEqual(hlPnf, tlPnf, ilRecLen, tlSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
            If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
                ilRet = btrClose(hlPnf)
                On Error GoTo gPopPersonnelBoxErr
                gBtrvErrorMsg ilRet, "gPopPersonnelBox (btrGetGreaterOrEqual):" & "Pnf.Btr", frm
                On Error GoTo 0
                btrDestroy hlPnf
                Exit Function
            End If
            If ilCode >= 0 Then
                tlIntTypeBuff.iType = ilCode
                ilOffSet = gFieldOffset("Pnf", "PnfAgfCode")
                ilRet = btrExtAddLogicConst(hlPnf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
            End If
        End If
        Call btrExtSetBounds(hlPnf, llNoRec, -1, "UC", "PNF", "") 'Set extract limits (all records)
        ilOffSet = 0
        ilRet = btrExtAddField(hlPnf, ilOffSet, ilRecLen)  'Extract iCode field
        On Error GoTo gPopPersonnelBoxErr
        gBtrvErrorMsg ilRet, "gPopPersonnelBox (btrExtAddField):" & "Pnf.Btr", frm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlPnf, tlPnf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopPersonnelBoxErr
            gBtrvErrorMsg ilRet, "gPopPersonnelBox (btrExtGetNextExt):" & "Pnf.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlPnf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlPnf, tlPnf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If (tlPnf.sState = "A") Or ((tlPnf.sState = "D") And (ilIncludeDormant)) Then
                    If tlPnf.sType = slType Then
                        If slType = "B" Then
                            slName = Left$(tlPnf.sName, 20)
                        Else
                            slName = tlPnf.sName
                        End If
                        If (ilShow = 1) Or (ilShow = 2) Then
                            slStr = gFormatPhoneNo(tlPnf.sPhone)
                            slName = slName & " " & slStr
                            If (ilShow = 2) Then
                                slStr = gFormatPhoneNo(tlPnf.sFax)
                                If slStr <> "" Then
                                    slName = slName & " Fax " & slStr
                                End If
                            End If
                        End If
                        slName = slName & "\" & Trim$(str$(tlPnf.iCode))
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit Do
                        'End If
                        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                    End If
                End If
                ilRet = btrExtGetNext(hlPnf, tlPnf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlPnf, tlPnf, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If
        ilRet = btrClose(hlPnf)
        On Error GoTo gPopPersonnelBoxErr
        gBtrvErrorMsg ilRet, "gPopPersonnelBox (btrReset):" & "Pnf.Btr", frm
        On Error GoTo 0
        btrDestroy hlPnf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopPersonnelBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopPersonnelBoxErr:
    ilRet = btrClose(hlPnf)
    btrDestroy hlPnf
    gDbg_HandleError "PopSubs: gPopPersonnelBox"
'    gPopPersonnelBox = CP_MSG_NOSHOW
'    Exit Function
gPopPersonnelBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopProgLibBox                  *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     Library names                   *
'*                                                     *
'*******************************************************
Function gPopProgLibBox(frm As Form, ilVerType As Integer, slLibType As String, ilVefCode As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopProgLibBox (MainForm, ilVer, slLibType, ilVEFCode, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilVerType (I)- 0=All (ALLLIB) Version # in back; 1= Latest (newest) (LATESTLIB); 2=All (ALLLIBFRONT) version # in front
'       slLibType (I)- 0=Library type ("R" = Regular; "S" = Special; "P"= Sports; "A" = All types)
'       ilVEFCode (I)- Vehicle code
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'LTF date/time stamp
    Dim hlLtf As Integer        'LTF handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Lvf
    Dim tlLtf As LTF
    Dim tlLtfExt As LTFEXT    'Extract record
    Dim hlLvf As Integer    'LVF handle
    Dim tlLvf As LVF        'Library version
    Dim ilLvfRecLen As Integer
    Dim tlLvfSrchKey As LVFKEY1  'LcF key record image
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim slName As String
    Dim slName2 As String
    Dim llVersion As Long
    Dim slVersion As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slLibName As String
    Dim slCode As String    'Sales source code number
    Dim ilPos As Integer
    Dim ilOffSet As Integer
    Dim llLen As Long
    Dim ilError As Integer
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    ilError = False
    slStamp = gFileDateTime(sgDBPath & "Lvf.Btr") & slLibType & Trim$(str$(ilVerType)) & Trim$(str$(ilVefCode))

    'On Error GoTo gPopProgLibBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopProgLibBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopProgLibBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlLtf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlLtf, "", sgDBPath & "Ltf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopProgLibBoxErr
        gBtrvErrorMsg ilRet, "gPopProgLibBox (btrOpen):" & "Ltf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlLtf) 'btrRecordLength(hlLtf)  'Get and save record length
        hlLvf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopProgLibBoxErr
        gBtrvErrorMsg ilRet, "gPopProgLibBox (btrOpen):" & "Lvf.Btr", frm
        On Error GoTo 0
        ilLvfRecLen = Len(tlLvf) 'btrRecordLength(hlLvf)  'Get and save record length
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlLtfExt)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlLtf) 'Obtain number of records
        btrExtClear hlLtf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlLtf, tlLtf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlLtf)
            ilRet = btrClose(hlLvf)
            On Error GoTo gPopProgLibBoxErr
            gBtrvErrorMsg ilRet, "gPopProgLibBox (btrReset):" & "Ltf.Btr", frm
            On Error GoTo 0
            btrDestroy hlLvf
            btrDestroy hlLtf
            Exit Function
        Else
            On Error GoTo gPopProgLibBoxErr
            gBtrvErrorMsg ilRet, "gPopProgLibBox (btrGetFirst):" & "Ltf.Btr", frm
            On Error GoTo 0
        End If
        Call btrExtSetBounds(hlLtf, llNoRec, -1, "UC", "LTFEXTPK", LTFEXTPK) 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Ltf", "LtfVefCode")
        If slLibType <> "A" Then
            tlIntTypeBuff.iType = ilVefCode
            ilRet = btrExtAddLogicConst(hlLtf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
            tlCharTypeBuff.sType = slLibType
            ilOffSet = gFieldOffset("Ltf", "LtfType")
            ilRet = btrExtAddLogicConst(hlLtf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        Else
            tlIntTypeBuff.iType = ilVefCode
            ilRet = btrExtAddLogicConst(hlLtf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        End If
        ilOffSet = gFieldOffset("Ltf", "LtfCode")
        ilRet = btrExtAddField(hlLtf, ilOffSet, 2)  'Extract iCode field
        On Error GoTo gPopProgLibBoxErr
        gBtrvErrorMsg ilRet, "gPopProgLibBox (btrExtAddField):" & "Ltf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Ltf", "LtfName")
        ilRet = btrExtAddField(hlLtf, ilOffSet, 20)  'Extract Name
        On Error GoTo gPopProgLibBoxErr
        gBtrvErrorMsg ilRet, "gPopProgLibBox (btrExtAddField):" & "Ltf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Ltf", "LtfVar")
        ilRet = btrExtAddField(hlLtf, ilOffSet, 2) 'Extract Variation
        On Error GoTo gPopProgLibBoxErr
        gBtrvErrorMsg ilRet, "gPopProgLibBox (btrExtAddField):" & "Ltf.Btr", frm
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hlLtf)    'Extract record
        ilRet = btrExtGetNext(hlLtf, tlLtfExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopProgLibBoxErr
            gBtrvErrorMsg ilRet, "gPopProgLibBox (btrExtGetNextExt):" & "Chf.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlLtfExt)  'Extract operation record size
            'ilRet = btrExtGetFirst(hlLtf, tlLtfExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlLtf, tlLtfExt, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If tlLtfExt.iVar <> 0 Then
                    slName = Trim$(tlLtfExt.sName) & "-" & Trim$(str$(tlLtfExt.iVar)) & "|"
                Else
                    slName = Trim$(tlLtfExt.sName) & "|"
                End If
                If ilVerType = LATESTLIB Then
                    tlLvfSrchKey.iLtfCode = tlLtfExt.iCode
                    tlLvfSrchKey.iVersion = 32000
                    ilRet = btrGetGreaterOrEqual(hlLvf, tlLvf, ilLvfRecLen, tlLvfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
                    If (ilRet = BTRV_ERR_NONE) And (tlLvf.iLtfCode = tlLtfExt.iCode) And (tlLvf.iVersion > 0) Then
                        llVersion = 999999 - tlLvf.iVersion
                        slVersion = Trim$(str$(llVersion))
                        Do While Len(slVersion) < 6
                            slVersion = "0" & slVersion
                        Loop
                        slName = slName & slVersion
                        slName = slName & "\" & Trim$(str$(tlLvf.lCode))
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit Do
                        'End If
                        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                    End If
                Else
                    slName2 = slName
                    tlLvfSrchKey.iLtfCode = tlLtfExt.iCode
                    tlLvfSrchKey.iVersion = 32000
                    ilRet = btrGetGreaterOrEqual(hlLvf, tlLvf, ilLvfRecLen, tlLvfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
                    Do While (ilRet = BTRV_ERR_NONE) And (tlLvf.iLtfCode = tlLtfExt.iCode)
                        llVersion = 999999 - tlLvf.iVersion
                        slVersion = Trim$(str$(llVersion))
                        Do While Len(slVersion) < 6
                            slVersion = "0" & slVersion
                        Loop
                        slName = slName & slVersion
                        slName = slName & "\" & Trim$(str$(tlLvf.lCode))
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    ilError = True
                        '    Exit Do
                        'End If
                        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                        slName = slName2
                        ilRet = btrGetNext(hlLvf, tlLvf, ilLvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If ilError Then
                        Exit Do
                    End If
                End If
                ilRet = btrExtGetNext(hlLtf, tlLtfExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlLtf, tlLtfExt, ilExtLen, llRecPos)
                Loop
            Loop
        '    If ilVerType = LATESTLIB Then
        '        For ilLoop = lbcMster.ListCount - 1 To 1 Step -1
        '            slNameCode = lbcMster.List(ilLoop)
        '            ilRet = gParseItem(slNameCode, 1, "|", slName)
        '            slNameCode = lbcMster.List(ilLoop - 1)
        '            ilRet = gParseItem(slNameCode, 1, "|", slName2)
        '            If StrComp(slName, slName2, 1) = 0 Then
        '                lbcMster.RemoveItem ilLoop
        '            End If
        '        Next ilLoop
        '    End If
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If
        ilRet = btrClose(hlLvf)
        On Error GoTo gPopProgLibBoxErr
        gBtrvErrorMsg ilRet, "gPopProgLibBox (btrReset):" & "Lvf.Btr", frm
        On Error GoTo 0
        btrDestroy hlLvf
        ilRet = btrClose(hlLtf)
        On Error GoTo gPopProgLibBoxErr
        gBtrvErrorMsg ilRet, "gPopProgLibBox (btrReset):" & "Ltf.Btr", frm
        On Error GoTo 0
        btrDestroy hlLtf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopProgLibBox = CP_MSG_PARSE
            Exit Function
        End If
        ilPos = InStr(slName, "|")
        If ilPos = 0 Then
            slName = Trim$(slName)
        Else
            ilRet = gParseItem(slName, 1, "|", slLibName)
            If ilRet <> CP_MSG_NONE Then
                gPopProgLibBox = CP_MSG_PARSE
                Exit Function
            End If
            If ilVerType = LATESTLIB Then
                slName = slLibName
            Else
                ilRet = gParseItem(slName, 2, "|", slCode)
                If ilRet <> CP_MSG_NONE Then
                    gPopProgLibBox = CP_MSG_PARSE
                    Exit Function
                End If
                llVersion = 999999 - CLng(slCode)
                If ilVerType = ALLLIBFRONT Then
                    slName = Trim$(str$(llVersion)) & "/" & slLibName
                Else
                    slName = slLibName & "/" & Trim$(str$(llVersion))
                End If
            End If
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopProgLibBoxErr:
    ilRet = btrClose(hlLtf)
    ilRet = btrClose(hlLvf)
    btrDestroy hlLvf
    btrDestroy hlLtf
    gDbg_HandleError "PopSubs: gPopProgLibBox"
'    gPopProgLibBox = CP_MSG_NOSHOW
'    Exit Function
gPopProgLibBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopRateCardBox                 *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with Rate     *
'*                     Card names active as specified  *
'*                     date Vehicles.                  *
'*                                                     *
'*******************************************************
Function gPopRateCardBox(frm As Form, llActiveDate As Long, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, ilVehCode As Integer) As Integer
'
'   ilRet = gPopRateCardBox (MainForm, llDate, lbcLocal, tlSortCode(), slSortCodeTag, ilVehCode)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       llDate (I)- Rate cards that are active after or equal to date (zero=all rate cards)
'       lbcLocal (I)- List box to be populated from the master list box (sorted in descending order)
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilVehCode (I)- Vehicle code to get rate card for or -1 if all Vehicles
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'RCF date/time stamp
    'Dim hlRcf As Integer        'RCF handle
    'Dim llNoRec As Long         'Number of records in Sof
    'Dim tlRcf As RCF
    'Dim tlRcfExt As RCFEXT    'Advertiser extract record
    Dim slName As String
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slVehName As String
    Dim slVehType As String
    Dim slDate As String
    Dim llLen As Long
    Dim ilRcf As Integer
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilActiveYear As Integer
    Dim ilMonth As Integer
    Dim ilVef As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Rcf.Btr") & gFileDateTime(sgDBPath & "Urf.btr") & Trim$(str$(llActiveDate)) & Trim$(str$(ilVehCode))

    'On Error GoTo gPopRateCardBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopRateCardBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopRateCardBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If llActiveDate > 0 Then
        slDate = Format$(llActiveDate - 32, "m/d/yy")
        gObtainMonthYear 0, slDate, ilMonth, ilActiveYear
    Else
        ilActiveYear = 0
    End If
    If ilPop Then
        ilRet = gObtainRcfRifRdf()
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)

            'Do While ilRet = BTRV_ERR_NONE
            For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                ilFound = False
                If (ilVehCode = -1) Or (ilVehCode = tgMRcf(ilRcf).iVefCode) Then
                    If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Then
                        ilFound = True
                    Else
                        'If tgRcf(ilRcf).iVefCode >= 0 Then
                            If tgMRcf(ilRcf).iVefCode > 0 Then
                                For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                                    If (tgUrf(illoop).iVefCode = tgMRcf(ilRcf).iVefCode) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next illoop
                            Else
                                ilFound = True
                            End If
                        'Else
                        '    For ilLoop = LBound(igVehComboCode) To UBound(igVehComboCode) - 1 Step 1
                        '        If igVehComboCode(ilLoop) = -tlRcfExt.iVefCode Then
                        '            ilFound = True
                        '            Exit For
                        '        End If
                        '    Next ilLoop
                        'End If
                    End If
                End If
                If ilFound Then
                    'If (llActiveDate > 0) And ((tgMRcf(ilRcf).iEndDate(0) > 0) Or (tgMRcf(ilRcf).iEndDate(1) > 0)) Then
                    '    gUnpackDate tgMRcf(ilRcf).iEndDate(0), tgMRcf(ilRcf).iEndDate(1), slDate
                    '    llDate = gDateValue(slDate)
                    '    If llDate < llActiveDate Then
                    '        ilFound = False
                    '    End If
                    'End If
                    If tgMRcf(ilRcf).iYear < ilActiveYear Then
                        ilFound = False
                    End If
                    'If ilYear > 0 Then
                    '    If tgRcf(ilRcf).iYear < ilYear Then
                    '        ilFound = False
                    '    End If
                    'End If
                End If
                If ilFound Then
                    'Salesperson can't view any rate card that starts with a "~"
                    '2/3/98
                    'If tgUrf(0).iSlfCode > 0 Then
                    If (tgUrf(0).sWorkToHold <> "Y") And (tgUrf(0).sWorkToOrder <> "Y") And (tgUrf(0).sCompToHold <> "Y") And (tgUrf(0).sCompToOrder <> "Y") And (tgUrf(0).sIncompToHold <> "Y") And (tgUrf(0).sIncompToOrder <> "Y") Then
                        If Left$(tgMRcf(ilRcf).sName, 1) = "~" Then
                            ilFound = False
                        End If
                    ElseIf (tgUrf(0).iSlfCode > 0) And (tgSpf.sGUsePropSys = "N") Then
                        If Left$(tgMRcf(ilRcf).sName, 1) = "~" Then
                            ilFound = False
                        End If
                    End If
                End If
                If ilFound Then
                    If tgMRcf(ilRcf).iVefCode = 0 Then
                        slVehName = "[All Vehicles]"
                        slVehType = ""
                    Else
                        gObtainVehicleName tgMRcf(ilRcf).iVefCode, slVehName, slVehType
                    End If
        '            If tgSpf.sSSellNet = "Y" Then
        '                If slVehType <> "S" Then
        '                    ilFound = False
        '                End If
        '            Else
        '                If slVehType <> "N" Then
        '                    ilFound = False
        '                End If
        '            End If
                End If
                If ilFound And (tgMRcf(ilRcf).iVefCode > 0) And (tgUrf(0).iCode > 2) And (tgUrf(0).iMnfHubCode > 0) Then
                    If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
                        ilVef = gBinarySearchVef(tgMRcf(ilRcf).iVefCode)
                        If ilVef <> -1 Then
                            If tgMVef(ilVef).iMnfHubCode <> tgUrf(0).iMnfHubCode Then
                                ilFound = False
                            End If
                        End If
                    End If
                End If
                'If ilFound And (tgMRcf(ilRcf).iVefCode > 0) And (tgSpf.sMktBase = "Y") Then
                '    ilFound = False
                '    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                '    '    If tgMRif(ilPrg).iVefCode = tgMVef(ilVef).iCode Then
                '        ilVef = gBinarySearchVef(tgMRcf(ilRcf).iVefCode)
                '        If ilVef <> -1 Then
                '            For ilLoop = 0 To UBound(ilMktCode) - 1 Step 1
                '                If tgMVef(ilVef).iMnfVehGp3Mkt = ilMktCode(ilLoop) Then
                '                    ilFound = True
                '                    Exit For
                '                End If
                '            Next ilLoop
                '            'Exit For
                '        End If
                '    'Next ilVef
                'End If
                If ilFound Then
                    gUnpackDateForSort tgMRcf(ilRcf).iStartDate(0), tgMRcf(ilRcf).iStartDate(1), slName
                    'slName = Trim$(Str$(9999 - tgRcf(ilRcf).iYear))
'Changed to descending order 2/16/04
'                    slName = gSubStr("99999", slName)
                    If ilVehCode = -1 Then
                        slName = slName & "\" & Trim$(tgMRcf(ilRcf).sName) & str$(tgMRcf(ilRcf).iYear) & "/" & slVehName
                    Else
                        slName = slName & "\" & Trim$(tgMRcf(ilRcf).sName) & str$(tgMRcf(ilRcf).iYear)
                    End If
                    slName = slName & "\" & Trim$(str$(tgMRcf(ilRcf).iCode))
                    'If Not gOkAddStrToListBox(slName, llLen, True) Then
                    '    Exit For
                    'End If
                    'lbcMster.AddItem slName    'Add Name (retain matching sorted order) and Code number to list box
                    tlSortCode(ilSortCode).sKey = slName
                    If ilSortCode >= UBound(tlSortCode) Then
                        ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                    End If
                    ilSortCode = ilSortCode + 1
                End If
                'ilRet = btrExtGetNext(hlRcf, tgMRcf(ilRcf), ilExtLen, llRecPos)
                'Do While ilRet = BTRV_ERR_REJECT_COUNT
                '    ilRet = btrExtGetNext(hlRcf, tgMRcf(ilRcf), ilExtLen, llRecPos)
                'Loop
            Next ilRcf
            'Populate in newest date is firstFRcf.lDyRate
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                'ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
'Change to descending (removed date subtraction, 2/16/04
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 1, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        'End If
        'ilRet = btrClose(hlRcf)
        'On Error GoTo gPopRateCardBoxErr
        'gBtrvErrorMsg ilRet, "gPopRateCardBox (btrReset):" & "Rcf.Btr", Frm
        'On Error GoTo 0
        'btrDestroy hlRcf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1 '-1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopRateCardBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function

    'ilRet = btrClose(hlRcf)
    'btrDestroy hlRcf
    gDbg_HandleError "PopSubs: gPopRateCardBox"
'    gPopRateCardBox = CP_MSG_NOSHOW
'    Exit Function
gPopRateCardBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopSalespersonBox              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                   Salesperson names (first and last)*
'*                                                     *
'*******************************************************
Function gPopSalespersonBox(frm As Form, ilInType As Integer, ilInChkUrfSlfCode As Integer, ilDormant As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, ilFirstNameFirst As Integer) As Integer
'
'   ilRet = gPopSalespersonBox (MainForm, ilType, ilChkUrfSlfCode, ilDormant, lbcLocal, tlSortCode(), slSortCodeTag, ilFirstName)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilType(I)- 0=All; 1=Salespersons and Negotiator only; 2=Planner only, 3=Manager only; 4=Salespersons, Negotiator and Planner; 5= Matching group
'       ilChkUrfSlfCode(I)- True=Check If User is salesperson and if so only populate with his/her name.  If user
'                                is not salesperson, then populate with all salesperson matching ilType
'                           False=ignore User salesperson code
'       ilDormant(I)- True=Include dormant salespersons; False=Exclude dormant salespersons
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilFirstNameFirst (I)- True=First name first
'                             False = Last name first
'       ilRet (O)- Error code (0 if no error)
'
'
    Dim slStamp As String    'Slf date/time stamp
    Dim hlSlf As Integer        'Slf handle
    'Dim llNoRec As Long         'Number of records in Sof
    'Dim tlSlf As SLF
    'Dim tlSlfExt As SLFEXT    'Salesperson extract record
    Dim slName As String
    'Dim ilExtLen As Integer
    'Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    'Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    'Dim ilOffset As Integer
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilSlf As Integer
    Dim ilFound As Integer
    Dim ilPop As Integer
    Dim ilType As Integer
    Dim ilChkUrfSlfCode As Integer
    Dim blManager As Boolean
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    ilType = ilInType
    ilChkUrfSlfCode = ilInChkUrfSlfCode
    If (tgUrf(0).iGroupNo <= 0) And (ilType = 5) Then
        ilType = 0
    End If

    slStamp = gFileDateTime(sgDBPath & "Slf.Btr") & Trim$(str$(ilType)) & Trim$(str$(ilChkUrfSlfCode)) & Trim$(str$(ilDormant))

    'On Error GoTo gPopSalespersonBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopSalespersonBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    lbcLocal.Clear
    slSortCodeTag = slStamp
    gPopSalespersonBox = CP_MSG_POPREQ
    If (ilType = 5) Then
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilRet = gObtainSalesperson()
        If Not ilRet Then
            Exit Function
        End If
        blManager = True
        If tgUrf(0).iSlfCode > 0 Then
            ilSlf = gBinarySearchSlf(tgUrf(0).iSlfCode)
            If ilSlf = -1 Then
                Exit Function
            End If
            If (tgMSlf(ilSlf).sJobTitle = "S") Then
                blManager = False
            End If
        End If
        If Not blManager Then
            If ilFirstNameFirst Then
                slName = Trim$(tgMSlf(ilSlf).sFirstName) & " " & Trim$(tgMSlf(ilSlf).sLastName)
            Else
                slName = Trim$(tgMSlf(ilSlf).sLastName) & ", " & Trim$(tgMSlf(ilSlf).sFirstName)
            End If
            Do While Len(slName) < Len(tgMSlf(ilSlf).sFirstName) + Len(tgMSlf(ilSlf).sLastName) + 2
                slName = slName & " "
            Loop
            slName = slName & "\" & Trim$(str$(tgMSlf(ilSlf).iCode))
            tlSortCode(ilSortCode).sKey = slName
            If ilSortCode >= UBound(tlSortCode) Then
                ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
            End If
            ilSortCode = ilSortCode + 1
        Else
            For illoop = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
                If tgPopUrf(illoop).iGroupNo = tgUrf(0).iGroupNo Then
                    ilSlf = gBinarySearchSlf(tgPopUrf(illoop).iSlfCode)
                    If ilSlf <> -1 Then
                        If (tgMSlf(ilSlf).sState = "A") Or ((tgMSlf(ilSlf).sState = "D") And (ilDormant)) Then
                            If ilFirstNameFirst Then
                                slName = Trim$(tgMSlf(ilSlf).sFirstName) & " " & Trim$(tgMSlf(ilSlf).sLastName)
                            Else
                                slName = Trim$(tgMSlf(ilSlf).sLastName) & ", " & Trim$(tgMSlf(ilSlf).sFirstName)
                            End If
                            Do While Len(slName) < Len(tgMSlf(ilSlf).sFirstName) + Len(tgMSlf(ilSlf).sLastName) + 2
                                slName = slName & " "
                            Loop
                            slName = slName & "\" & Trim$(str$(tgMSlf(ilSlf).iCode))
                            tlSortCode(ilSortCode).sKey = slName
                            If ilSortCode >= UBound(tlSortCode) Then
                                ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                            End If
                            ilSortCode = ilSortCode + 1
                        End If
                    End If
                End If
            Next illoop
        End If
        ReDim Preserve tlSortCode(0 To ilSortCode) As SORTCODE
    Else
        If (ilChkUrfSlfCode) And (tgUrf(0).iSlfCode > 0) Then
            ReDim tlSortCode(ilLowLimit To ilLowLimit + 1) As SORTCODE
            slName = sgUserSalespersonName & "\" & Trim$(str$(tgUrf(0).iSlfCode))
            tlSortCode(ilLowLimit).sKey = slName
            ilSortCode = ilSortCode + 1
        Else
            If ilPop Then
                ilSortCode = ilLowLimit
                ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
                ilRet = gObtainSalesperson()
                If Not ilRet Then
                    Exit Function
                End If
                For ilSlf = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
                    ilFound = False
                    If (Not ilChkUrfSlfCode) Or (tgUrf(0).iSlfCode = 0) Then
                        If ilType = 0 Then
                            ilFound = True
                        ElseIf ilType = 1 Then
                            If (tgMSlf(ilSlf).sJobTitle = "S") Or (tgMSlf(ilSlf).sJobTitle = "N") Then
                                ilFound = True
                            End If
                        ElseIf ilType = 2 Then
                            If (tgMSlf(ilSlf).sJobTitle = "P") Then
                                ilFound = True
                            End If
                        ElseIf ilType = 3 Then
                            If (tgMSlf(ilSlf).sJobTitle = "M") Then
                                ilFound = True
                            End If
                        ElseIf ilType = 4 Then
                            If (tgMSlf(ilSlf).sJobTitle = "S") Or (tgMSlf(ilSlf).sJobTitle = "N") Or (tgMSlf(ilSlf).sJobTitle = "P") Then
                                ilFound = True
                            End If
                        End If
                        If ilFound Then
                            ilFound = False
                            If (tgMSlf(ilSlf).sState = "A") Or ((tgMSlf(ilSlf).sState = "D") And (ilDormant)) Then
                                ilFound = True
                            End If
                        End If
                    Else
                        If tgUrf(0).iSlfCode = tgMSlf(ilSlf).iCode Then
                            ilFound = True
                        End If
                    End If
                    If ilFound Then
                        If ilFirstNameFirst Then
                            slName = Trim$(tgMSlf(ilSlf).sFirstName) & " " & Trim$(tgMSlf(ilSlf).sLastName)
                        Else
                            slName = Trim$(tgMSlf(ilSlf).sLastName) & ", " & Trim$(tgMSlf(ilSlf).sFirstName)
                        End If
                        Do While Len(slName) < Len(tgMSlf(ilSlf).sFirstName) + Len(tgMSlf(ilSlf).sLastName) + 2
                            slName = slName & " "
                        Loop
                        slName = slName & "\" & Trim$(str$(tgMSlf(ilSlf).iCode))
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                    End If
                Next ilSlf
                    'Sort then output new headers and lines
                    ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
                    If UBound(tlSortCode) - ilLowLimit > 0 Then
                        ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
                    End If
            End If
        End If
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopSalespersonBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function

    ilRet = btrClose(hlSlf)
    btrDestroy hlSlf
    gDbg_HandleError "PopSubs: gPopSalespersonBox"
'    gPopSalespersonBox = CP_MSG_NOSHOW
'    Exit Function
gPopSalespersonBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopShortTitleBox               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     advertiser Short title names    *
'*                                                     *
'*******************************************************
Function gPopShortTitleBox(frm As Form, ilAdvtCode As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopShortTitleBox (MainForm, ilAdvt, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAdvt (I)- Advertise code value
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'Adf date/time stamp
    Dim hlSif As Integer        'Adf handle
    Dim ilRecLen As Integer     'Record length
    Dim tlSif As SIF
    Dim slName As String
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlSrchKey As SIFKEY1  'Sif key record image
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Sif.Btr") & Trim$(str$(ilAdvtCode))

    'On Error GoTo gPopShortTitleBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopShortTitleBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopShortTitleBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlSif = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopShortTitleBoxErr
        gBtrvErrorMsg ilRet, "gPopShortTitleBox (btrOpen): Sif.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlSif) 'btrRecordLength(hlSif)  'Get and save record length
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlSif)
        llNoRec = gExtNoRec(ilExtLen)
        btrExtClear hlSif
        Call btrExtSetBounds(hlSif, llNoRec, -1, "UC", "SIF", "") 'Set extract limits (all records)
        tlSrchKey.iAdfCode = ilAdvtCode
        ilRet = btrGetGreaterOrEqual(hlSif, tlSif, ilRecLen, tlSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
        If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
            ilRet = btrClose(hlSif)
            On Error GoTo gPopShortTitleBoxErr
            gBtrvErrorMsg ilRet, "gPopShortTitleBox (btrGetGreaterOrEqual):" & "Sif.Btr", frm
            On Error GoTo 0
            btrDestroy hlSif
            Exit Function
        End If
        tlIntTypeBuff.iType = ilAdvtCode
        ilOffSet = gFieldOffset("Sif", "SifAdfCode")
        ilRet = btrExtAddLogicConst(hlSif, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        Call btrExtSetBounds(hlSif, llNoRec, -1, "UC", "SIF", "") 'Set extract limits (all records)
        ilOffSet = 0
        ilRet = btrExtAddField(hlSif, ilOffSet, ilRecLen)  'Extract iCode field
        On Error GoTo gPopShortTitleBoxErr
        gBtrvErrorMsg ilRet, "gPopShortTitleBox (btrExtAddField):" & "Sif.Btr", frm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlSif, tlSif, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopShortTitleBoxErr
            gBtrvErrorMsg ilRet, "gPopShortTitleBox (btrExtGetNextExt):" & "Sif.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlSif)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSif, tlSif, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slName = tlSif.sName
                slName = slName & "\" & Trim$(str$(tlSif.lCode))
                'If Not gOkAddStrToListBox(slName, llLen, True) Then
                '    Exit Do
                'End If
                'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                tlSortCode(ilSortCode).sKey = slName
                If ilSortCode >= UBound(tlSortCode) Then
                    ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                End If
                ilSortCode = ilSortCode + 1
                ilRet = btrExtGetNext(hlSif, tlSif, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSif, tlSif, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If



        ilRet = btrClose(hlSif)
        On Error GoTo gPopShortTitleBoxErr
        gBtrvErrorMsg ilRet, "gPopShortTitleBox (btrReset):" & "Sif.Btr", frm
        On Error GoTo 0
        btrDestroy hlSif
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopShortTitleBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopShortTitleBoxErr:
    ilRet = btrClose(hlSif)
    btrDestroy hlSif
    gDbg_HandleError "PopSubs: gPopShortTitleBox"
'    gPopShortTitleBox = CP_MSG_NOSHOW
'    Exit Function
gPopShortTitleBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopSlspBudgetBox               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     unique years that budget created*
'*                     for salespeople                 *
'*                                                     *
'*******************************************************
Function gPopSlspBudgetBox(frm As Form, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopSlspBudgetBox (MainForm, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       lbcLocal (I)- List box to be populated from the master list box (sort = False)
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'Bsf date/time stamp
    Dim hlBsf As Integer        'Bsf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlBsf As BSF
    Dim tlBsfExt As POPINTEGERTYPE   'Type field record
    Dim slName As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim ilOffSet As Integer
    Dim slYear As String
    Dim llLen As Long
    Dim ilFound As Integer
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Bsf.Btr")

    'On Error GoTo gPopSlspBudgetBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopSlspBudgetBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopSlspBudgetBox = CP_MSG_POPREQ
    lbcLocal.Clear   'VB list box clear (list box used to retain code number so record can be found)
    slSortCodeTag = slStamp
    If ilPop Then
        hlBsf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlBsf, "", sgDBPath & "Bsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopSlspBudgetBoxErr
        gBtrvErrorMsg ilRet, "gPopSlspBudgetBox (btrOpen):" & "Bsf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlBsf) 'btrRecordLength(hlBsf)  'Get and save record length
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlBsfExt)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlBsf) 'Obtain number of records
        btrExtClear hlBsf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlBsf, tlBsf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlBsf)
            On Error GoTo gPopSlspBudgetBoxErr
            gBtrvErrorMsg ilRet, "gPopSlspBudgetBox (btrReset):" & "Bsf.Btr", frm
            On Error GoTo 0
            btrDestroy hlBsf
            Exit Function
        Else
            On Error GoTo gPopSlspBudgetBoxErr
            gBtrvErrorMsg ilRet, "gPopSlspBudgetBox (btrGetFirst):" & "Bsf.Btr", frm
            On Error GoTo 0
        End If
        Call btrExtSetBounds(hlBsf, llNoRec, -1, "UC", "POPINTEGERTYPEPK", POPINTEGERTYPEPK) 'Set extract limits (all records including first)
        slYear = Year(gNow())
        '1/15/10: Only allow 5 year to avoid changing integer to long with ilSortCode
        tlIntTypeBuff.iType = CInt(slYear) - 5  '20
        ilOffSet = gFieldOffset("Bsf", "BsfYear")
        ilRet = btrExtAddLogicConst(hlBsf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        ilOffSet = gFieldOffset("Bsf", "BsfYear")
        ilRet = btrExtAddField(hlBsf, ilOffSet, 2)  'Extract iCode field
        On Error GoTo gPopSlspBudgetBoxErr
        gBtrvErrorMsg ilRet, "gPopSlspBudgetBox (btrExtAddField):" & "Bsf.Btr", frm
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hlBsf)    'Extract record
        ilRet = btrExtGetNext(hlBsf, tlBsfExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopSlspBudgetBoxErr
            gBtrvErrorMsg ilRet, "gPopSlspBudgetBox (btrExtGetNextExt):" & "Bsf.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlBsfExt)  'Extract operation record size
            'ilRet = btrExtGetFirst(hlBsf, tlBsfExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlBsf, tlBsfExt, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slName = Trim$(str$(9999 - tlBsfExt.iType))
                'gFindMatch slName, 0, lbcMster
                'If lbcMster.LastFound < 0 Then
                ilFound = False
'                For ilLoop = 0 To ilSortCode - 1 Step 1
'                    If Trim$(tlSortCode(ilLoop).sKey) = slName Then
'                        ilFound = True
'                        Exit For
'                    End If
'                Next ilLoop
                If Not ilFound Then
                    'If Not gOkAddStrToListBox(slName, llLen, True) Then
                    '    Exit Do
                    'End If
                    'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                    tlSortCode(ilSortCode).sKey = slName
                    If ilSortCode >= UBound(tlSortCode) Then
                        ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                    End If
                    ilSortCode = ilSortCode + 1
                End If
                ilRet = btrExtGetNext(hlBsf, tlBsfExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlBsf, tlBsfExt, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
                ilSortCode = ilLowLimit
                For illoop = ilLowLimit + 1 To UBound(tlSortCode) - 1 Step 1
                    If Trim$(tlSortCode(ilSortCode).sKey) <> Trim$(tlSortCode(illoop).sKey) Then
                        ilSortCode = ilSortCode + 1
                        tlSortCode(ilSortCode) = tlSortCode(illoop)
                    End If
                Next illoop
                ilSortCode = ilSortCode + 1
                ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            End If
        End If
        ilRet = btrClose(hlBsf)
        On Error GoTo gPopSlspBudgetBoxErr
        gBtrvErrorMsg ilRet, "gPopSlspBudgetBox (btrReset):" & "Bsf.Btr", frm
        On Error GoTo 0
        btrDestroy hlBsf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slYear = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        slName = gSubStr("9999", slYear)
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopSlspBudgetBoxErr:
    ilRet = btrClose(hlBsf)
    btrDestroy hlBsf
    gDbg_HandleError "PopSubs: gPopSlspBudgetBox"
'    gPopSlspBudgetBox = CP_MSG_NOSHOW
'    Exit Function
gPopSlspBudgetBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopSSourceByGroupBox           *
'*                                                     *
'*             Created:7/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     advertiser names                *
'*                                                     *
'*******************************************************
Function gPopSSourceByGroupBox(frm As Form, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopSSourceByGroupBox (MainForm, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'MNF date/time stamp
    Dim hlMnf As Integer        'MNF handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlMnf As MNF
    Dim tlMnfExt As MNFEXT    'Advertiser extract record
    Dim slName As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim slNameCode As String
    Dim illoop As Integer
    Dim ilOffSet As Integer
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    'On Error GoTo gPopSSourceByGroupBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If
    
    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Mnf.Btr")
    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopSSourceByGroupBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopSSourceByGroupBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopSSourceByGroupBoxErr
        gBtrvErrorMsg ilRet, "gPopSSourceByGroupBox (btrOpen):" & "Mnf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tlMnfExt)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlMnf) 'Obtain number of records
        btrExtClear hlMnf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlMnf, tlMnf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_END_OF_FILE Then
            ilRet = btrClose(hlMnf)
            On Error GoTo gPopSSourceByGroupBoxErr
            gBtrvErrorMsg ilRet, "gPopSSourceByGroupBox (btrReset):" & "Mnf.Btr", frm
            On Error GoTo 0
            btrDestroy hlMnf
            Exit Function
        Else
            On Error GoTo gPopSSourceByGroupBoxErr
            gBtrvErrorMsg ilRet, "gPopSSourceByGroupBox (btrGetFirst):" & "Mnf.Btr", frm
            On Error GoTo 0
        End If
        'Extract information from first record
    'Use UC (get cuurent) in setbounds instead of EG (bypass current)
    '    If Trim$(tlMnf.sType) = "S" Then
    '        slName = Trim$(Str$(tlMnf.iGroupNo))
    '        Do While Len(slName) < 2
    '            slName = "0" & slName
    '        Loop
    '        slName = slName & "\" & Trim$(tlMnf.sName)
    '        slName = slName & "\" & Trim$(Str$(tlMnf.iCode))
    '        lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
    '    End If
        Call btrExtSetBounds(hlMnf, llNoRec, -1, "UC", "MNFEXTPK", MNFEXTPK) 'Set extract limits (all records)
        tlCharTypeBuff.sType = "S"
        ilOffSet = gFieldOffset("Mnf", "MnfType")
        ilRet = btrExtAddLogicConst(hlMnf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        ilOffSet = gFieldOffset("Mnf", "MnfCode")
        ilRet = btrExtAddField(hlMnf, ilOffSet, 2)  'Extract iCode field
        On Error GoTo gPopSSourceByGroupBoxErr
        gBtrvErrorMsg ilRet, "gPopSSourceByGroupBox (btrExtAddField):" & "Mnf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Mnf", "MnfName")
        ilRet = btrExtAddField(hlMnf, ilOffSet, 20)  'Extract Name field
        On Error GoTo gPopSSourceByGroupBoxErr
        gBtrvErrorMsg ilRet, "gPopSSourceByGroupBox (btrExtAddField):" & "Mnf.Btr", frm
        On Error GoTo 0
        ilOffSet = gFieldOffset("Mnf", "MnfGroupNo")
        ilRet = btrExtAddField(hlMnf, ilOffSet, 2) 'Extract part of group # field
        On Error GoTo gPopSSourceByGroupBoxErr
        gBtrvErrorMsg ilRet, "gPopSSourceByGroupBox (btrExtAddField):" & "Mnf.Btr", frm
        On Error GoTo 0
        'ilRet = btrExtGetNextExt(hlMnf)    'Extract record
        ilRet = btrExtGetNext(hlMnf, tlMnfExt, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopSSourceByGroupBoxErr
            gBtrvErrorMsg ilRet, "gPopSSourceByGroupBox (btrExtGetNextExt):" & "Mnf.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tlMnfExt)  'Extract operation record size
            'ilRet = btrExtGetFirst(hlMnf, tlMnfExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlMnf, tlMnfExt, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slName = Trim$(str$(tlMnfExt.iGroupNo))
                Do While Len(slName) < 2
                    slName = "0" & slName
                Loop
                slName = slName & "\" & Trim$(tlMnfExt.sName)
                slName = slName & "\" & Trim$(str$(tlMnfExt.iCode))
                'If Not gOkAddStrToListBox(slName, llLen, True) Then
                '    Exit Do
                'End If
                'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                tlSortCode(ilSortCode).sKey = slName
                If ilSortCode >= UBound(tlSortCode) Then
                    ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                End If
                ilSortCode = ilSortCode + 1
                ilRet = btrExtGetNext(hlMnf, tlMnfExt, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlMnf, tlMnfExt, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If
        ilRet = btrClose(hlMnf)
        On Error GoTo gPopSSourceByGroupBoxErr
        gBtrvErrorMsg ilRet, "gPopSSourceByGroupBox (btrReset):" & "Mnf.Btr", frm
        On Error GoTo 0
        btrDestroy hlMnf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopSSourceByGroupBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopSSourceByGroupBoxErr:
    ilRet = btrClose(hlMnf)
    btrDestroy hlMnf
    gDbg_HandleError "PopSubs: gPopSSourceByGroupBox"
'    gPopSSourceByGroupBox = CP_MSG_NOSHOW
'    Exit Function
gPopSSourceByGroupBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopUserListBox                 *
'*                                                     *
'*             Created:5/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*           Comments:Populate list Box with user names*
'*                                                     *
'*******************************************************
Function gPopUserListBox(frm As Form, lbcCtrl As Control, ilIncludeErased As Integer) As Integer
'
'   ilRet = gPopUserListBox (MainForm, cbcCtrl)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       cbcCtrl (I)- List box control that will be populated with names
'       ilRet (O)- True=list was either populated or repopulated
'                  False=List was OK- it didn't require populating
'

    Dim ilRecLen As Integer     'URF record length
    Dim hlUrf As Integer        'User Option file handle
    Dim tlUrf As URF
    Dim slStamp As String   'Current time stamp
    Dim llNoRec As Long     'Number of records (must be one)
    Dim ilRet As Integer
    Dim illoop As Integer
'    Dim ilVefRecLen As Integer
'    Dim hlVef As Integer
'    Dim tlVef As VEF
'    Dim tlVefSrchKey As INTKEY0    'key record image
    Dim ilVehOk As Integer
    Dim slSyncDate As String
    Dim slSyncTime As String
    Dim ilUrf As Integer

    ilRet = 0
    On Error GoTo gPopUserListBoxErr2
    slStamp = gFileDateTime(sgDBPath & "Urf.btr")
    If ilRet <> 0 Then
'        MsgBox "Unable to locate " & sgDBPath & "Urf.Btr" & " Error " & Str$(ilRet)
        'If igBkgdProg = 0 Then
        '    MsgBox "Unable to locate " & sgDBPath & "Urf.Btr" & " Error " & Str$(ilRet)
        'ElseIf igBkgdProg = 1 Then
        '    gLogMsg "Unable to locate " & sgDBPath & "Urf.Btr" & " Error " & Str$(ilRet), "Bkgd_Schd.Txt", False
        'ElseIf igBkgdProg = 2 Then
        '    gLogMsg "Unable to locate " & sgDBPath & "Urf.Btr" & " Error " & Str$(ilRet), "Set_Credit.Txt", False
        'Else
        '    gLogMsg "Unable to locate " & sgDBPath & "Urf.Btr" & " Error " & Str$(ilRet), "TrafficErrors.Txt", False
        'End If
        gMsgBox "Unable to locate " & sgDBPath & "Urf.Btr" & " Error " & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "User Path Error"
        gPopUserListBox = CP_MSG_NOPOPREQ
        Exit Function
    End If
    If lbcCtrl.Tag <> "" Then
        If StrComp(slStamp, lbcCtrl.Tag, 1) = 0 Then
            gPopUserListBox = CP_MSG_NOPOPREQ
            Exit Function
        End If
    End If
    gPopUserListBox = CP_MSG_POPREQ
    hlUrf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopUserListBoxErr
    gBtrvErrorMsg ilRet, "gPopUserListBox (btrOpen):" & "Urf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlUrf)  'btrRecordLength(hlUrf)  'Get and save record length
'    hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
'    ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    On Error GoTo gPopUserListBoxErr
'    gBtrvErrorMsg ilRet, "gPopUserListBoxRead (btrOpen):" & "Vef.btr", Frm
'    On Error GoTo 0
'    ilVefRecLen = Len(tlVef)  'btrRecordLength(hlVef)  'Get and save record length
    llNoRec = btrRecords(hlUrf) 'Obtain number of records
    If llNoRec = 0 Then
        tlUrf.sName = "Counterpoint"
        tlUrf.iVefCode = 0
        tlUrf.sPassword = "JFDL"
        tlUrf.sRept = "System"
        tlUrf.sPhoneNo = ""
        tlUrf.sCity = ""
        tlUrf.lEMailCefCode = 0
        tlUrf.iSlfCode = 0
        tlUrf.sClnMoYr = "M"    'Calendar by month
        tlUrf.sClnType = "S"    'Calendar type-standard
        tlUrf.sClnLayout = "A"  'Calendar- Across
        tlUrf.iClnLeft = 100    'Calendar left
        tlUrf.iClnTop = 200 'Calendar top
        tlUrf.iClcLeft = 2775   'Traffic.Width - 2775 - 100   '2625 width of calc form
        tlUrf.iClcTop = 200
        For illoop = 0 To 69 Step 1
            tlUrf.sWin(illoop) = "I"
        Next illoop
        tlUrf.iDefVeh = 0
        tlUrf.sGrid = "I"
        tlUrf.sPrice = "I"
        tlUrf.sCredit = "I"
        tlUrf.sPayRate = "I"
        tlUrf.sMerge = "I"
        tlUrf.sDelete = "N"
        tlUrf.sHideSpots = "I"
        tlUrf.sChgBilled = "I"
        tlUrf.sChgCntr = "I"
        tlUrf.sChgCrRt = "I"
        tlUrf.sBouChk = "I"
        tlUrf.sReprintLogAlert = "Y"
        tlUrf.sIncompAlert = "Y"
        tlUrf.sCompAlert = "Y"
        tlUrf.sSchAlert = "Y"
        tlUrf.sHoldAlert = "Y"
        tlUrf.sRateCardAlert = "Y"
        tlUrf.sResearchAlert = "Y"
        tlUrf.sAvailAlert = "Y"
        tlUrf.sCrdChkAlert = "Y"
        tlUrf.sDeniedAlert = "Y"
        tlUrf.sCrdLimitAlert = "Y"
        tlUrf.sMoveAlert = "Y"
        tlUrf.sPrgmmaticAlert = "V"
        tlUrf.sAdvanceAvails = "H"
        tlUrf.sShowNRMsg = "Y"
        tlUrf.sWorkToDead = "Y"
        tlUrf.sWorkToComp = "Y"
        tlUrf.sWorkToHold = "Y"
        tlUrf.sWorkToOrder = "Y"
        tlUrf.sCompToIncomp = "Y"
        tlUrf.sCompToDead = "Y"
        tlUrf.sCompToHold = "Y"
        tlUrf.sCompToOrder = "Y"
        tlUrf.sIncompToDead = "Y"
        tlUrf.sIncompToComp = "Y"
        tlUrf.sIncompToHold = "Y"
        tlUrf.sIncompToOrder = "Y"
        tlUrf.sDeadToWork = "Y"
        tlUrf.sHoldToOrder = "Y"
        tlUrf.sReviseCntr = "Y"
        tlUrf.sResvType = "I"
        tlUrf.sRemType = "I"
        tlUrf.sDRType = "I"
        tlUrf.sPIType = "I"
        tlUrf.sPSAType = "I"
        tlUrf.sPromoType = "I"
        tlUrf.sRefResvType = "I"
        tlUrf.sUseComputeCMC = "H"
        tlUrf.iGroupNo = 0
        tlUrf.sBlockRU = "N"
        tlUrf.sRCView = "R"
        tlUrf.sRegionCopy = "I"
        tlUrf.sChgPrices = "I"
        tlUrf.sChgLnBillPrice = "I"
        tlUrf.sActFlightButton = "I"
        tlUrf.sOldPassword(0) = ""
        tlUrf.sOldPassword(1) = ""
        tlUrf.sOldPassword(2) = ""
        tlUrf.sLiveLogPostOnly = "N"
        tlUrf.sSportPropOnly = "N"
        tlUrf.iRemoteUserID = 0
        tlUrf.iMnfHubCode = 0
        tlUrf.sShowNRMsg = "Y"
                ' Dan M new selected fields 4/10/09
        tlUrf.sChangeCSIDate = "I"
        tlUrf.sAllowInvDisplay = "I"
        tlUrf.sActivityLog = "V"

        'Add new value into gInitSuperUser (initsubs.bas)
        tlUrf.iCode = 0  'Autoincrement
        tlUrf.iRemoteID = 0
        tlUrf.iAutoCode = tlUrf.iCode
        gUrfEncrypt tlUrf
        ilRet = btrInsert(hlUrf, tlUrf, ilRecLen, INDEXKEY0)
        On Error GoTo gPopUserListBoxErr
        gBtrvErrorMsg ilRet, "gPopUserListBox (btrGetFirst):" & "Urf.Btr", frm
        On Error GoTo 0
        gUrfDecrypt tlUrf
        Do
            tlUrf.iRemoteID = 0
            tlUrf.iAutoCode = tlUrf.iCode
            gGetSyncDateTime slSyncDate, slSyncTime
            gPackDate slSyncDate, tlUrf.iSyncDate(0), tlUrf.iSyncDate(1)
            gPackTime slSyncTime, tlUrf.iSyncTime(0), tlUrf.iSyncTime(1)
            gUrfEncrypt tlUrf
            ilRet = btrUpdate(hlUrf, tlUrf, ilRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        gUrfDecrypt tlUrf
        tlUrf.sName = "Guide"
        tlUrf.iVefCode = 0
        tlUrf.sPassword = "Guide"
        tlUrf.iCode = 0  'Autoincrement
        tlUrf.iRemoteID = 0
        tlUrf.iAutoCode = tlUrf.iCode
        gUrfEncrypt tlUrf
        ilRet = btrInsert(hlUrf, tlUrf, ilRecLen, INDEXKEY0)
        On Error GoTo gPopUserListBoxErr
        gBtrvErrorMsg ilRet, "gPopUserListBox (btrInsert):" & "Urf.Btr", frm
        On Error GoTo 0
        gUrfDecrypt tlUrf
        Do
            tlUrf.iRemoteID = 0
            tlUrf.iAutoCode = tlUrf.iCode
            gPackDate slSyncDate, tlUrf.iSyncDate(0), tlUrf.iSyncDate(1)
            gPackTime slSyncTime, tlUrf.iSyncTime(0), tlUrf.iSyncTime(1)
            gUrfEncrypt tlUrf
            ilRet = btrUpdate(hlUrf, tlUrf, ilRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        gUrfDecrypt tlUrf
        sgPopUrfTag = ""
        ilRet = gObtainUrf()
    End If
    lbcCtrl.Clear
    lbcCtrl.Tag = slStamp
'    ilRet = btrGetFirst(hlUrf, tlUrf, ilRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    Do While ilRet = BTRV_ERR_NONE
'        gUrfDecrypt tlUrf
    ilRet = 0
    'On Error GoTo gPopUserListBoxErr2
    'ilUrf = LBound(tgPopUrf)
    If PeekArray(tgPopUrf).Ptr <> 0 Then
        ilUrf = LBound(tgPopUrf)
    Else
        ilRet = 1
        ilUrf = 0
    End If
    If ilRet <> 0 Then
        ilRet = gObtainUrf()
    End If
    For ilUrf = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
        tlUrf = tgPopUrf(ilUrf)
        If tlUrf.iCode = 1 Then
            sgCPName = Trim$(tlUrf.sName)
        End If
        If tlUrf.iCode = 2 Then
            sgSUName = Trim$(tlUrf.sName)
        End If
        If (tlUrf.sDelete <> "Y") Or (ilIncludeErased) Then
            gFindMatch tlUrf.sName, 0, lbcCtrl
            If gLastFound(lbcCtrl) < 0 Then
                If tlUrf.iCode <= 2 Then
                    ilVehOk = True
                Else
                    If tlUrf.iVefCode = 0 Then  'All vehicles
                        ilVehOk = True
                    Else
'                        ilVehOk = False
'                        tlVefSrchKey.iCode = tlUrf.iVefCode
'                        ilRet = btrGetEqual(hlVef, tlVef, ilVefRecLen, tlVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'                        If ilRet = BTRV_ERR_NONE Then
'                            tlVef = tgMVef(ilVef)
'                            If tgSpf.sSSellNet = "Y" Then
'                                If (tlVef.sType = "S") Or (tlVef.sType = "C") Then
'                                    ilVehOk = True
'                                End If
'                            Else
'                                If tlVef.sType = "C" Then
'                                    ilVehOk = True
'                                End If
'                            End If
'                        End If
                        ilVehOk = True
                    End If
                End If
                If ilVehOk Then
                    lbcCtrl.AddItem Trim$(tlUrf.sName)
                End If
            End If
        End If
'        ilRet = btrGetNext(hlUrf, tlUrf, ilRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    Loop
     Next ilUrf
'    If ilRet <> BTRV_ERR_END_OF_FILE Then
'        On Error GoTo gPopUserListBoxErr
'        gBtrvErrorMsg ilRet, "gPopUserListBox (btrGetFirst):" & "Urf.Btr", Frm
'        On Error GoTo 0
'    End If
'    ilRet = btrClose(hlVef)
'    btrDestroy hlVef
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    Exit Function
gPopUserListBoxErr:
'    ilRet = btrClose(hlVef)
    ilRet = btrClose(hlUrf)
'    btrDestroy hlVef
    btrDestroy hlUrf
    gDbg_HandleError "PopSubs: gPopUserListBox"
gPopUserListBoxErr2:
    ilRet = err.Number
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopUserVehicleBox             *
'*                                                     *
'*             Created:5/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with Vehicle  *
'*                     which are permitted for the user*
'*                                                     *
'*******************************************************
Function gPopUserVehicleBox(frm As Form, llVehType As Long, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, Optional ilIncludeVefCode As Integer = 0) As Integer
'
'   ilRet = gPopUserVehicleBox (MainForm, llVehType, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error
'       llVehType(I)- Right to Left (0 to 15) Bit On means include; Bit Off exclude
'           Conventional without Feed = VEHCONV_WO_FEED = &H1
'           Conventional with Feed = VEHCONV_W_FEED = &H2
'           Selling = VEHSELLING = &H4
'           Airing = VEHAIRING = &H8
'           Log = VEHLOG = &H10     Vehicle with Log Vehicle will be included if VEHLOGVEHICLE is also set on, or exclude if VEHLOGVEHICLE is not set on
'           Virtual = VEHVIRTUAL = &H20
'           Log Vehicle = VEHLOGVEHICLE = &H40  This is only used if VEHLOG set on
'           Package Vehicle = VEHPACKAGE = &H80   This is any kind of Package
'           Standard Package = VEHSTDPKG = &H100  Standard package vehicles only
'           Simulcast Vehicle = VEHSIMUL = &H200  Simulcast vehicle
'
'           Active Vehicles = ACTIVEVEH = &H2000
'           Dormant Vehicles = DORMANTVEF = &H4000
'
'           Old definition (Changed 12/21/95)
'           ilType(I)- 0=selling and conventional(Both); 1=Airing; 2=Unused; 3=All except Log and Dormant;
'                  4=Conventional(Both); 5= Selling; 6 = Airing and Conventional(with feed);
'                  7 = Airing and Conventional(Both); 8=All vehicles; 9 = Airing, Conventional without Log Vehicle and Log Vehicles;
'                  10 = Airing, Selling, Conventional without Log Vehicle and Log Vehicles;
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- True=list was either populated or repopulated
'                  False=List was OK- it didn't require populating
'

    'Dim hlVef As Integer        'Vehicle file handle
    'Dim tlVef As VEF
    Dim slStamp As String   'Current time stamp
    Dim ilRet As Integer
    Dim slName As String
    Dim illoop As Integer
    'Dim tlSrchKey As INTKEY0
    Dim slNameCode As String
    Dim ilOk As Integer 'Does Vehicle match type
    Dim ilAllVehicles As Integer
    Dim llLen As Long
    Dim ilIndex As Integer
    Dim slGpSort As String
    Dim slVehSort As String
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim ilMnfRecLen As Integer
    Dim tlMnfSrchKey As INTKEY0
    Dim ilSortVehbyGroup As Integer 'True=Sort vehicle by groups; False=Sort by name only
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilSortVehbyGroup = False
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Vef.btr") & Trim$(str$(llVehType))

    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    ilPop = True
    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopUserVehicleBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If

    gPopUserVehicleBox = CP_MSG_POPREQ
    lbcLocal.Clear
    lbcLocal.Tag = ""
    slSortCodeTag = slStamp
    If ilPop Then
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilRet = gObtainVef()
        'DoEvents    'This eliminated error in NWCalls
        For illoop = 0 To 20 Step 1
            ilOk = ilOk
        Next illoop
        If ilRet = False Then
            Exit Function
        End If
        hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hlMnf
            Exit Function
        End If
        ilMnfRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length

        ilAllVehicles = False
        If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Then
            ilAllVehicles = True
        Else
            For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                If tgUrf(illoop).iVefCode = 0 Then
                    ilAllVehicles = True
                    Exit For
                End If
            Next illoop
        End If
        'hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        'ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo gPopUserVehicleBoxErr
        'gBtrvErrorMsg ilRet, "gPopUserVehicleBox (btrOpen):" & "Vef.btr", Frm
        'On Error GoTo 0
        'ilRecLen = Len(tlVef)  'btrRecordLength(hlVef)  'Get and save record length
        If ilAllVehicles Then
            'ilExtLen = Len(tlVef)  'Extract operation record size
            'llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
            'btrExtClear hlVef   'Clear any previous extend operation
            'ilRet = btrGetFirst(hlVef, tlVef, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            'If ilRet = BTRV_ERR_END_OF_FILE Then
            '    ilRet = btrClose(hlVef)
            '    btrDestroy hlVef
            '    Exit Function
            'Else
            '    If ilRet <> BTRV_ERR_NONE Then
            '        ilRet = btrClose(hlVef)
            '        btrDestroy hlVef
            '        Exit Function
            '    End If
            'End If
            'Call btrExtSetBounds(hlVef, llNoRec, -1, "UC") 'Set extract limits (all records)
            'ilOffset = 0
            'ilRet = btrExtAddField(hlVef, 0, ilExtLen)  'Extract iCode field
            'If ilRet <> BTRV_ERR_NONE Then
            '    ilRet = btrClose(hlVef)
            '    btrDestroy hlVef
            '    Exit Function
            'End If
            ''ilRet = btrExtGetNextExt(hlSlf)    'Extract record
            'ilRet = btrExtGetNext(hlVef, tlVef, ilExtLen, llRecPos)
            'If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            '    If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            '        ilRet = btrClose(hlVef)
            '        btrDestroy hlVef
            '        Exit Function
            '    End If
            '    ilExtLen = Len(tlVef)  'Extract operation record size
            '    'ilRet = btrExtGetFirst(hlSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
            '    Do While ilRet = BTRV_ERR_REJECT_COUNT
            '        ilRet = btrExtGetNext(hlVef, tlVef, ilExtLen, llRecPos)
            '    Loop
                'Do While ilRet = BTRV_ERR_NONE
                For illoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    If mTestVehType(llVehType, tgMVef(illoop), ilIncludeVefCode) Then
                        If ilSortVehbyGroup Then
                            slVehSort = Trim$(str$(tgMVef(illoop).iSort))
                            tlMnfSrchKey.iCode = tgMVef(illoop).iOwnerMnfCode
                            If tlMnf.iCode <> tlMnfSrchKey.iCode Then
                                ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet <> BTRV_ERR_NONE Then
                                    tlMnf.iGroupNo = 999
                                End If
                            End If
                            slGpSort = Trim$(str$(tlMnf.iGroupNo))
                            Do While Len(slGpSort) < 3
                                slGpSort = "0" & slGpSort
                            Loop
                            Do While Len(slVehSort) < 3
                                slVehSort = "0" & slVehSort
                            Loop
                        Else
                            slGpSort = "000"
                            slVehSort = "000"
                        End If
                        slName = tgMVef(illoop).sName 'trim cause xxx and xxx y to sort in wrong order Trim$(tlVef.sName)
                        slName = slGpSort & "|" & slVehSort & "|" & slName & "\" & Trim$(str$(tgMVef(illoop).iCode))
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit For
                        'End If
                        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                    End If
                    'ilRet = btrExtGetNext(hlVef, tlVef, ilExtLen, llRecPos)
                    'Do While ilRet = BTRV_ERR_REJECT_COUNT
                    '    ilRet = btrExtGetNext(hlVef, tlVef, ilExtLen, llRecPos)
                    'Loop
                'Loop
                Next illoop
            'End If
        Else
            For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                'tlSrchKey.iCode = tgUrf(ilLoop).iVefCode
                'ilRet = btrGetEqual(hlVef, tlVef, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                'On Error GoTo gPopUserVehicleBoxErr
                'gBtrvErrorMsg ilRet, "gPopUserVehicleBox (btrGetEqual):" & "Vef.btr", Frm
                'On Error GoTo 0
                'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                '    If tgMVef(ilIndex).iCode = tgUrf(ilLoop).iVefCode Then
                    ilIndex = gBinarySearchVef(tgUrf(illoop).iVefCode)
                    If ilIndex <> -1 Then
                        If mTestVehType(llVehType, tgMVef(ilIndex), ilIncludeVefCode) Then
                            If ilSortVehbyGroup Then
                                slVehSort = Trim$(str$(tgMVef(ilIndex).iSort))
                                tlMnfSrchKey.iCode = tgMVef(ilIndex).iOwnerMnfCode
                                If tlMnf.iCode <> tlMnfSrchKey.iCode Then
                                    ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        tlMnf.iGroupNo = 999
                                    End If
                                End If
                                slGpSort = Trim$(str$(tlMnf.iGroupNo))
                                Do While Len(slGpSort) < 3
                                    slGpSort = "0" & slGpSort
                                Loop
                                Do While Len(slVehSort) < 3
                                    slVehSort = "0" & slVehSort
                                Loop
                            Else
                                slGpSort = "000"
                                slVehSort = "000"
                            End If
                            slName = tgMVef(ilIndex).sName    'Trim$(tlVef.sName)
                            slName = slGpSort & "|" & slVehSort & "|" & slName & "\" & Trim$(str$(tgMVef(ilIndex).iCode))
                            'If Not gOkAddStrToListBox(slName, llLen, True) Then
                            '    Exit For
                            'End If
                            'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                            tlSortCode(ilSortCode).sKey = slName
                            If ilSortCode >= UBound(tlSortCode) Then
                                ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                            End If
                            ilSortCode = ilSortCode + 1
                        End If
                        'Exit For
                    End If
                'Next ilIndex
            Next illoop
        End If
        'Sort then output new headers and lines
        ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
            ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
        End If
        btrDestroy hlMnf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopUserVehicleBox = CP_MSG_PARSE
            Exit Function
        End If
        ilRet = gParseItem(slName, 3, "|", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopUserVehicleBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box

    Next illoop
    'ilRet = btrClose(hlVef)
    'btrDestroy hlVef
    Exit Function

    btrDestroy hlMnf
    'ilRet = btrClose(hlVef)
    'btrDestroy hlVef
    gDbg_HandleError "PopSubs: gPopUserVehicleBox"
gPopUserVehicleBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopUserVehicleBoxNoForm             *
'*                                                     *
'*             Created:5/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with Vehicle  *
'*                     which are permitted for the user*
'*                                                     *
'*******************************************************
Function gPopUserVehicleBoxNoForm(llVehType As Long, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopUserVehicleBoxNoForm (MainForm, llVehType, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error
'       llVehType(I)- Right to Left (0 to 15) Bit On means include; Bit Off exclude
'           Conventional without Feed = VEHCONV_WO_FEED = &H1
'           Conventional with Feed = VEHCONV_W_FEED = &H2
'           Selling = VEHSELLING = &H4
'           Airing = VEHAIRING = &H8
'           Log = VEHLOG = &H10     Vehicle with Log Vehicle will be included if VEHLOGVEHICLE is also set on, or exclude if VEHLOGVEHICLE is not set on
'           Virtual = VEHVIRTUAL = &H20
'           Log Vehicle = VEHLOGVEHICLE = &H40  This is only used if VEHLOG set on
'           Package Vehicle = VEHPACKAGE = &H80   This is any kind of Package
'           Standard Package = VEHSTDPKG = &H100  Standard package vehicles only
'           Simulcast Vehicle = VEHSIMUL = &H200  Simulcast vehicle
'
'           Active Vehicles = ACTIVEVEH = &H2000
'           Dormant Vehicles = DORMANTVEF = &H4000
'
'           Old definition (Changed 12/21/95)
'           ilType(I)- 0=selling and conventional(Both); 1=Airing; 2=Unused; 3=All except Log and Dormant;
'                  4=Conventional(Both); 5= Selling; 6 = Airing and Conventional(with feed);
'                  7 = Airing and Conventional(Both); 8=All vehicles; 9 = Airing, Conventional without Log Vehicle and Log Vehicles;
'                  10 = Airing, Selling, Conventional without Log Vehicle and Log Vehicles;
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- True=list was either populated or repopulated
'                  False=List was OK- it didn't require populating
'

    'Dim hlVef As Integer        'Vehicle file handle
    'Dim tlVef As VEF
    Dim slStamp As String   'Current time stamp
    Dim ilRet As Integer
    Dim slName As String
    Dim illoop As Integer
    'Dim tlSrchKey As INTKEY0
    Dim slNameCode As String
    Dim ilOk As Integer 'Does Vehicle match type
    Dim ilAllVehicles As Integer
    Dim llLen As Long
    Dim ilIndex As Integer
    Dim slGpSort As String
    Dim slVehSort As String
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim ilMnfRecLen As Integer
    Dim tlMnfSrchKey As INTKEY0
    Dim ilSortVehbyGroup As Integer 'True=Sort vehicle by groups; False=Sort by name only
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer

    ilSortVehbyGroup = False
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Vef.btr") & Trim$(str$(llVehType))
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    ilPop = True
    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopUserVehicleBoxNoForm = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If

    gPopUserVehicleBoxNoForm = CP_MSG_POPREQ
    lbcLocal.Clear
    lbcLocal.Tag = ""
    slSortCodeTag = slStamp
    If ilPop Then
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilRet = gObtainVef()
        'DoEvents    'This eliminated error in NWCalls
        For illoop = 0 To 20 Step 1
            ilOk = ilOk
        Next illoop
        If ilRet = False Then
            Exit Function
        End If
        hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hlMnf
            Exit Function
        End If
        ilMnfRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length

        ilAllVehicles = False
        If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Then
            ilAllVehicles = True
        Else
            For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                If tgUrf(illoop).iVefCode = 0 Then
                    ilAllVehicles = True
                    Exit For
                End If
            Next illoop
        End If
        'hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        'ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo gPopUserVehicleBoxErr
        'gBtrvErrorMsg ilRet, "gPopUserVehicleBox (btrOpen):" & "Vef.btr", Frm
        'On Error GoTo 0
        'ilRecLen = Len(tlVef)  'btrRecordLength(hlVef)  'Get and save record length
        If ilAllVehicles Then
            'ilExtLen = Len(tlVef)  'Extract operation record size
            'llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
            'btrExtClear hlVef   'Clear any previous extend operation
            'ilRet = btrGetFirst(hlVef, tlVef, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            'If ilRet = BTRV_ERR_END_OF_FILE Then
            '    ilRet = btrClose(hlVef)
            '    btrDestroy hlVef
            '    Exit Function
            'Else
            '    If ilRet <> BTRV_ERR_NONE Then
            '        ilRet = btrClose(hlVef)
            '        btrDestroy hlVef
            '        Exit Function
            '    End If
            'End If
            'Call btrExtSetBounds(hlVef, llNoRec, -1, "UC") 'Set extract limits (all records)
            'ilOffset = 0
            'ilRet = btrExtAddField(hlVef, 0, ilExtLen)  'Extract iCode field
            'If ilRet <> BTRV_ERR_NONE Then
            '    ilRet = btrClose(hlVef)
            '    btrDestroy hlVef
            '    Exit Function
            'End If
            ''ilRet = btrExtGetNextExt(hlSlf)    'Extract record
            'ilRet = btrExtGetNext(hlVef, tlVef, ilExtLen, llRecPos)
            'If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            '    If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            '        ilRet = btrClose(hlVef)
            '        btrDestroy hlVef
            '        Exit Function
            '    End If
            '    ilExtLen = Len(tlVef)  'Extract operation record size
            '    'ilRet = btrExtGetFirst(hlSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
            '    Do While ilRet = BTRV_ERR_REJECT_COUNT
            '        ilRet = btrExtGetNext(hlVef, tlVef, ilExtLen, llRecPos)
            '    Loop
                'Do While ilRet = BTRV_ERR_NONE
                For illoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    If mTestVehType(llVehType, tgMVef(illoop)) Then
                        If ilSortVehbyGroup Then
                            slVehSort = Trim$(str$(tgMVef(illoop).iSort))
                            tlMnfSrchKey.iCode = tgMVef(illoop).iOwnerMnfCode
                            If tlMnf.iCode <> tlMnfSrchKey.iCode Then
                                ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet <> BTRV_ERR_NONE Then
                                    tlMnf.iGroupNo = 999
                                End If
                            End If
                            slGpSort = Trim$(str$(tlMnf.iGroupNo))
                            Do While Len(slGpSort) < 3
                                slGpSort = "0" & slGpSort
                            Loop
                            Do While Len(slVehSort) < 3
                                slVehSort = "0" & slVehSort
                            Loop
                        Else
                            slGpSort = "000"
                            slVehSort = "000"
                        End If
                        slName = tgMVef(illoop).sName 'trim cause xxx and xxx y to sort in wrong order Trim$(tlVef.sName)
                        slName = slGpSort & "|" & slVehSort & "|" & slName & "\" & Trim$(str$(tgMVef(illoop).iCode))
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit For
                        'End If
                        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                    End If
                    'ilRet = btrExtGetNext(hlVef, tlVef, ilExtLen, llRecPos)
                    'Do While ilRet = BTRV_ERR_REJECT_COUNT
                    '    ilRet = btrExtGetNext(hlVef, tlVef, ilExtLen, llRecPos)
                    'Loop
                'Loop
                Next illoop
            'End If
        Else
            For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                'tlSrchKey.iCode = tgUrf(ilLoop).iVefCode
                'ilRet = btrGetEqual(hlVef, tlVef, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                'On Error GoTo gPopUserVehicleBoxErr
                'gBtrvErrorMsg ilRet, "gPopUserVehicleBox (btrGetEqual):" & "Vef.btr", Frm
                'On Error GoTo 0
                'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                '    If tgMVef(ilIndex).iCode = tgUrf(ilLoop).iVefCode Then
                    ilIndex = gBinarySearchVef(tgUrf(illoop).iVefCode)
                    If ilIndex <> -1 Then
                        If mTestVehType(llVehType, tgMVef(ilIndex)) Then
                            If ilSortVehbyGroup Then
                                slVehSort = Trim$(str$(tgMVef(ilIndex).iSort))
                                tlMnfSrchKey.iCode = tgMVef(ilIndex).iOwnerMnfCode
                                If tlMnf.iCode <> tlMnfSrchKey.iCode Then
                                    ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        tlMnf.iGroupNo = 999
                                    End If
                                End If
                                slGpSort = Trim$(str$(tlMnf.iGroupNo))
                                Do While Len(slGpSort) < 3
                                    slGpSort = "0" & slGpSort
                                Loop
                                Do While Len(slVehSort) < 3
                                    slVehSort = "0" & slVehSort
                                Loop
                            Else
                                slGpSort = "000"
                                slVehSort = "000"
                            End If
                            slName = tgMVef(ilIndex).sName    'Trim$(tlVef.sName)
                            slName = slGpSort & "|" & slVehSort & "|" & slName & "\" & Trim$(str$(tgMVef(ilIndex).iCode))
                            'If Not gOkAddStrToListBox(slName, llLen, True) Then
                            '    Exit For
                            'End If
                            'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                            tlSortCode(ilSortCode).sKey = slName
                            If ilSortCode >= UBound(tlSortCode) Then
                                ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                            End If
                            ilSortCode = ilSortCode + 1
                        End If
                        'Exit For
                    End If
                'Next ilIndex
            Next illoop
        End If
        'Sort then output new headers and lines
        ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
            ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
        End If
        btrDestroy hlMnf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopUserVehicleBoxNoForm = CP_MSG_PARSE
            Exit Function
        End If
        ilRet = gParseItem(slName, 3, "|", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopUserVehicleBoxNoForm = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    'ilRet = btrClose(hlVef)
    'btrDestroy hlVef
    Exit Function

    btrDestroy hlMnf
    'ilRet = btrClose(hlVef)
    'btrDestroy hlVef
    gDbg_HandleError "PopSubs: gPopUserVehicleBoxNoForm"
gPopUserVehicleBoxNoFormErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopUserVehicleByMkt            *
'*                                                     *
'*             Created:5/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with Vehicle  *
'*                     which are permitted for the user*
'*                     Taken from gPopUserVehicleBox   *
'*                                                     *
'*******************************************************
Function gPopUserVehicleByMkt(frm As Form, llVehType As Long, ilMktCode() As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, Optional ilIncludeVefCode As Integer = 0) As Integer
'
'   ilRet = gPopUserVehicleByMkt (MainForm, llVehType, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error
'       llVehType(I)- Right to Left (0 to 15) Bit On means include; Bit Off exclude
'           Conventional without Feed = VEHCONV_WO_FEED = &H1
'           Conventional with Feed = VEHCONV_W_FEED = &H2
'           Selling = VEHSELLING = &H4
'           Airing = VEHAIRING = &H8
'           Log = VEHLOG = &H10     Vehicle with Log Vehicle will be included if VEHLOGVEHICLE is also set on, or exclude if VEHLOGVEHICLE is not set on
'           Virtual = VEHVIRTUAL = &H20
'           Log Vehicle = VEHLOGVEHICLE = &H40  This is only used if VEHLOG set on
'           Package Vehicle = VEHPACKAGE = &H80   This is any kind of Package
'           Standard Package = VEHSTDPKG = &H100  Standard package vehicles only
'           Simulcast Vehicle = VEHSIMUL = &H200  Simulcast vehicle
'
'           Active Vehicles = ACTIVEVEH = &H2000
'           Dormant Vehicles = DORMANTVEF = &H4000
'
'           Old definition (Changed 12/21/95)
'           ilType(I)- 0=selling and conventional(Both); 1=Airing; 2=Unused; 3=All except Log and Dormant;
'                  4=Conventional(Both); 5= Selling; 6 = Airing and Conventional(with feed);
'                  7 = Airing and Conventional(Both); 8=All vehicles; 9 = Airing, Conventional without Log Vehicle and Log Vehicles;
'                  10 = Airing, Selling, Conventional without Log Vehicle and Log Vehicles;
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- True=list was either populated or repopulated
'                  False=List was OK- it didn't require populating
'

    'Dim hlVef As Integer        'Vehicle file handle
    'Dim tlVef As VEF
    Dim slStamp As String   'Current time stamp
    Dim ilRet As Integer
    Dim slName As String
    Dim illoop As Integer
    'Dim tlSrchKey As INTKEY0
    Dim slNameCode As String
    Dim ilOk As Integer 'Does Vehicle match type
    Dim ilAllVehicles As Integer
    Dim llLen As Long
    Dim ilIndex As Integer
    Dim slGpSort As String
    Dim slVehSort As String
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim ilMnfRecLen As Integer
    Dim tlMnfSrchKey As INTKEY0
    Dim ilSortVehbyGroup As Integer 'True=Sort vehicle by groups; False=Sort by name only
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilMkt As Integer
    Dim ilMktOk As Integer
    Dim ilLowLimit As Integer

    ilSortVehbyGroup = False
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Vef.btr") & Trim$(str$(llVehType))

    If ((llVehType And VEHBYMKT) = VEHBYMKT) And (tgSpf.sMktBase = "Y") Then
        For ilMkt = 0 To UBound(ilMktCode) - 1 Step 1
            slStamp = slStamp & Trim$(str$(ilMktCode(ilMkt)))
        Next ilMkt
    End If
    'On Error GoTo gPopUserVehicleByMktErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    ilPop = True
    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopUserVehicleByMkt = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If

    gPopUserVehicleByMkt = CP_MSG_POPREQ
    lbcLocal.Clear
    lbcLocal.Tag = ""
    slSortCodeTag = slStamp
    If ilPop Then
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilRet = gObtainVef()
        'DoEvents    'This eliminated error in NWCalls
        For illoop = 0 To 20 Step 1
            ilOk = ilOk
        Next illoop
        If ilRet = False Then
            Exit Function
        End If
        hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hlMnf
            Exit Function
        End If
        ilMnfRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length

        ilAllVehicles = False
        If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Then
            ilAllVehicles = True
        Else
            For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                If tgUrf(illoop).iVefCode = 0 Then
                    ilAllVehicles = True
                    Exit For
                End If
            Next illoop
        End If
        'hlVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        'ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo gPopUserVehicleByMktErr
        'gBtrvErrorMsg ilRet, "gPopUserVehicleByMkt (btrOpen):" & "Vef.btr", Frm
        'On Error GoTo 0
        'ilRecLen = Len(tlVef)  'btrRecordLength(hlVef)  'Get and save record length
        If ilAllVehicles Then
            For illoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                'Market vehicle selection
                If ((llVehType And VEHBYMKT) = VEHBYMKT) And (tgSpf.sMktBase = "Y") Then
                    ilMktOk = False
                    For ilMkt = 0 To UBound(ilMktCode) - 1 Step 1
                        If tgMVef(illoop).iMnfVehGp3Mkt = ilMktCode(ilMkt) Then
                            ilMktOk = True
                            Exit For
                        End If
                    Next ilMkt
                Else
                    ilMktOk = True
                End If
                If ilMktOk Then
                    If mTestVehType(llVehType, tgMVef(illoop), ilIncludeVefCode) Then
                        If ilSortVehbyGroup Then
                            slVehSort = Trim$(str$(tgMVef(illoop).iSort))
                            tlMnfSrchKey.iCode = tgMVef(illoop).iOwnerMnfCode
                            If tlMnf.iCode <> tlMnfSrchKey.iCode Then
                                ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet <> BTRV_ERR_NONE Then
                                    tlMnf.iGroupNo = 999
                                End If
                            End If
                            slGpSort = Trim$(str$(tlMnf.iGroupNo))
                            Do While Len(slGpSort) < 3
                                slGpSort = "0" & slGpSort
                            Loop
                            Do While Len(slVehSort) < 3
                                slVehSort = "0" & slVehSort
                            Loop
                        Else
                            slGpSort = "000"
                            slVehSort = "000"
                        End If
                        slName = tgMVef(illoop).sName 'trim cause xxx and xxx y to sort in wrong order Trim$(tlVef.sName)
                        slName = slGpSort & "|" & slVehSort & "|" & slName & "\" & Trim$(str$(tgMVef(illoop).iCode))
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit For
                        'End If
                        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                    End If
                End If
            Next illoop
        Else
            For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                'tlSrchKey.iCode = tgUrf(ilLoop).iVefCode
                'ilRet = btrGetEqual(hlVef, tlVef, ilRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                'On Error GoTo gPopUserVehicleByMktErr
                'gBtrvErrorMsg ilRet, "gPopUserVehicleByMkt (btrGetEqual):" & "Vef.btr", Frm
                'On Error GoTo 0
                'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                '    If tgMVef(ilIndex).iCode = tgUrf(ilLoop).iVefCode Then
                        'Market vehicle selection
                    ilIndex = gBinarySearchVef(tgUrf(illoop).iVefCode)
                    If ilIndex <> -1 Then
                        If ((llVehType And VEHBYMKT) = VEHBYMKT) And (tgSpf.sMktBase = "Y") Then
                            ilMktOk = False
                            For ilMkt = 0 To UBound(ilMktCode) - 1 Step 1
                                If tgMVef(ilIndex).iMnfVehGp3Mkt = ilMktCode(ilMkt) Then
                                    ilMktOk = True
                                    Exit For
                                End If
                            Next ilMkt
                        Else
                            ilMktOk = True
                        End If
                        If ilMktOk Then
                            If mTestVehType(llVehType, tgMVef(ilIndex), ilIncludeVefCode) Then
                                If ilSortVehbyGroup Then
                                    slVehSort = Trim$(str$(tgMVef(ilIndex).iSort))
                                    tlMnfSrchKey.iCode = tgMVef(ilIndex).iOwnerMnfCode
                                    If tlMnf.iCode <> tlMnfSrchKey.iCode Then
                                        ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            tlMnf.iGroupNo = 999
                                        End If
                                    End If
                                    slGpSort = Trim$(str$(tlMnf.iGroupNo))
                                    Do While Len(slGpSort) < 3
                                        slGpSort = "0" & slGpSort
                                    Loop
                                    Do While Len(slVehSort) < 3
                                        slVehSort = "0" & slVehSort
                                    Loop
                                Else
                                    slGpSort = "000"
                                    slVehSort = "000"
                                End If
                                slName = tgMVef(ilIndex).sName    'Trim$(tlVef.sName)
                                slName = slGpSort & "|" & slVehSort & "|" & slName & "\" & Trim$(str$(tgMVef(ilIndex).iCode))
                                'If Not gOkAddStrToListBox(slName, llLen, True) Then
                                '    Exit For
                                'End If
                                'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                                tlSortCode(ilSortCode).sKey = slName
                                If ilSortCode >= UBound(tlSortCode) Then
                                    ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                                End If
                                ilSortCode = ilSortCode + 1
                            End If
                        End If
                        'Exit For
                    End If
                'Next ilIndex
            Next illoop
        End If
        'Sort then output new headers and lines
        ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
            ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
        End If
        btrDestroy hlMnf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopUserVehicleByMkt = CP_MSG_PARSE
            Exit Function
        End If
        ilRet = gParseItem(slName, 3, "|", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopUserVehicleByMkt = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    'ilRet = btrClose(hlVef)
    'btrDestroy hlVef
    Exit Function

    btrDestroy hlMnf
    'ilRet = btrClose(hlVef)
    'btrDestroy hlVef
    gDbg_HandleError "PopSubs: gPopUserVehicleByMkt"
gPopUserVehicleByMktErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopVehBudgetBox                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     advertiser names (if direct,    *
'*                     add \Direct to name)            *
'*                                                     *
'*******************************************************
Function gPopVehBudgetBox(frm As Form, ilType As Integer, ilSortOrder As Integer, ilShowOrder As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopVehBudgetBox (MainForm, ilSortBox, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilType(I)-0=Budget; 1=Actuals; 2=Budget
'       ilSortBox(I)-0=Sort Year/Budget Name; 1=Budget Name/Year
'       ilShowOrder(I)- 0=Year/Budget Name; 1=Budget Name/Year
'       lbcLocal (I)- List box to be populated from the master list box (sort = False)
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'
'       Note: In lbcCtrl the year is stored in reverse order (to get year => 9999-year)
'             In lbcLocal the year is store as true year (the sort property must be false)
'

    Dim slStamp As String    'Bvf date/time stamp
    Dim hlBvf As Integer        'Bvf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlBvf As BVF
    Dim tlBvfExt As BVFEXT    'Advertiser extract record
    '10/10/11:  Get names directly from NF and Year from BVF
'    Dim hlMnf As Integer
'    Dim tlMnf As MNF
'    Dim tlMnfSrchKey As INTKEY0
'    Dim ilMnfRecLen As Integer
    Dim slMnfStamp As String
    Dim tlBvfSrchKey2 As INTKEY0
    
    Dim slName As String
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim ilFound As Integer
    Dim slNameCode As String
    Dim slNameYear As String
    Dim slYear As String
    Dim ilOffSet As Integer
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Bvf.Btr") & str$(ilType) & str$(ilSortOrder) & str$(ilShowOrder)

    'On Error GoTo gPopVehBudgetBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopVehBudgetBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopVehBudgetBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then

        
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        hlBvf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopVehBudgetBoxErr
        gBtrvErrorMsg ilRet, "gPopVehBudgetBox (btrOpen):" & "Bvf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tlBvf) 'btrRecordLength(hlBvf)  'Get and save record length
        slMnfStamp = ""
        'ReDim tlMnf(1 To 1) As MNF
        ReDim tlMnf(0 To 0) As MNF
        ilRet = gObtainMnfForType("U", slMnfStamp, tlMnf())
        If ilRet = False Then
            Exit Function
        End If
        For illoop = LBound(tlMnf) To UBound(tlMnf) - 1 Step 1
            tlBvfSrchKey2.iCode = tlMnf(illoop).iCode
            ilRet = btrGetEqual(hlBvf, tlBvf, ilRecLen, tlBvfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slName = Trim$(str$(9999 - tlBvf.iYear))
                If ilSortOrder = 0 Then
                    slName = slName & "/" & Trim$(tlMnf(illoop).sName) & "\" & Trim$(str$(tlMnf(illoop).iCode))
                Else
                    slName = Trim$(tlMnf(illoop).sName) & "/" & slName & "\" & Trim$(str$(tlMnf(illoop).iCode))
                End If
                tlSortCode(ilSortCode).sKey = slName
                ilSortCode = ilSortCode + 1
                ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            End If
        Next illoop
        If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
            ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
       End If
    End If
    ilRet = btrClose(hlBvf)
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slNameYear)
        If ilSortOrder = 0 Then
            ilRet = gParseItem(slNameYear, 1, "/", slYear)
            ilRet = gParseItem(slNameYear, 2, "/", slName)
        Else
            ilRet = gParseItem(slNameYear, 1, "/", slName)
            ilRet = gParseItem(slNameYear, 2, "/", slYear)
        End If
        If ilShowOrder = 0 Then
            slName = gSubStr("9999", slYear) & "/" & slName
        Else
            slName = slName & "/" & gSubStr("9999", slYear)
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    Exit Function
gPopVehBudgetBoxErr:
    ilRet = btrClose(hlBvf)
    btrDestroy hlBvf
'    ilRet = btrClose(hlMnf)
'    btrDestroy hlMnf
    gDbg_HandleError "PopSubs: gPopVehBudgetBox"
'    gPopVehBudgetBox = CP_MSG_NOSHOW
'    Exit Function
gPopVehBudgetBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopVehViaRateCardBox           *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     Vehicle names from the rate     *
'*                     card programs which are allowed *
'*                     for a user                      *
'*                     (only N/A or selling allowed)   *
'*                                                     *
'*******************************************************
Function gPopVehViaRateCardBox(frm As Form, ilRcfCode As Integer, ilMktCode() As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, Optional slFilterType As String = "", Optional blRequireAdServer As Boolean = False, Optional blCheckLTF As Boolean = True) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilPrg                                                                                 *
'******************************************************************************************

'
'   ilRet = gPopVehViaRateCardBox (MainForm, ilRcfCode, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error
'       ilRcfCode (I)- Rcf Code to get Rif for
'       lbcLocal (I)- List box to be populated from the master list box (sorted in descending order)
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       slFilterType(I)- Obtain which vehicles (A=Air Time; P=Podcast Spots; M=Podcast CPM; B = Air Time and Podcast Spots; X = None
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'RCF date/time stamp
    Dim slName As String
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slVehName As String
    Dim slVehSort As String
    Dim llLen As Long
    Dim slGpSort As String
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim ilMnfRecLen As Integer
    Dim tlMnfSrchKey As INTKEY0
    Dim ilSortVehbyGroup As Integer 'True=Sort vehicle by groups; False=Sort by name only
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim slSortChar As String
    Dim llRif As Long
    Dim ilLowLimit As Integer
    Dim slSQLQuery As String
    Dim pvf_rst As ADODB.Recordset
    Dim ilVff As Integer
    ilPop = True
    ilSortVehbyGroup = False
    llLen = 0
    'TTP 10574 - Contract screen: new slowness when saving a large contract.
    'slStamp = gFileDateTime(sgDBPath & "Rcf.Btr") & Trim$(str$(ilRcfCode))
    slStamp = gFileDateTime(sgDBPath & "Rcf.Btr") & Trim$(str$(ilRcfCode)) & blCheckLTF

    'On Error GoTo gPopVehViaRateCardBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If
    ''TTP 10438 - Proposals/Contracts screen, NTR tab: podcast medium vehicle for use with ad server only doesn't appear in vehicle list WHEN the podcast Ad Server option is ON
    ''We need to reload the list so that we can check for Podcast vehicles with LTF on Air Tab and skip check on NTR tab
    'TTP 10574 - Contract screen: new slowness when saving a large contract.
    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopVehViaRateCardBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopVehViaRateCardBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            gPopVehViaRateCardBox = CP_MSG_NOPOPREQ
            btrDestroy hlMnf
            Exit Function
        End If
        ilMnfRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
        'gObtainVehComboList
        ilRet = gObtainVef()
        ilSortCode = 0
        If UBound(tgMRcf) > 1 Then
            ReDim tlSortCode(ilLowLimit To UBound(tgMRif) / (UBound(tgMRcf) - 1)) As SORTCODE 'VB list box clear (list box used to retain code number so record can be found)
        Else
            If UBound(tgMRif) > 10000 Then
                ReDim tlSortCode(ilLowLimit To 10000) As SORTCODE 'VB list box clear (list box used to retain code number so record can be found)
            Else
                ReDim tlSortCode(ilLowLimit To UBound(tgMRif)) As SORTCODE 'VB list box clear (list box used to retain code number so record can be found)
            End If
        End If
        For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
            If tgMRif(llRif).iRcfCode = ilRcfCode Then
                ilFound = False
                If tgUrf(0).iCode <= 2 Then
                    ilFound = True
                Else
                    'If tgRpf(ilPrg).iVefCode >= 0 Then
                        For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                            If (tgUrf(illoop).iVefCode = tgMRif(llRif).iVefCode) Then
                                ilFound = True
                                Exit For
                            End If
                        Next illoop
                    'Else
                    '    For ilLoop = LBound(igVehComboCode) To UBound(igVehComboCode) - 1 Step 1
                    '        If igVehComboCode(ilLoop) = -tgRpf(ilPrg).iVefCode Then
                    '            ilFound = True
                    '            Exit For
                    '        End If
                    '    Next ilLoop
                    'End If
                End If
                If Not ilFound Then
                    For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                        If (tgUrf(illoop).iVefCode = 0) Then
                            ilFound = True
                            Exit For
                        End If
                    Next illoop
                End If
                If (ilFound) And (tgUrf(0).iCode > 2) And (tgUrf(0).iMnfHubCode > 0) Then
                    If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB Then
                        ilVef = gBinarySearchVef(tgMRif(llRif).iVefCode)
                        If ilVef <> -1 Then
                            If tgMVef(ilVef).iMnfHubCode <> tgUrf(0).iMnfHubCode Then
                                ilFound = False
                            End If
                        End If
                    End If
                End If
                If ilFound And (tgSpf.sMktBase = "Y") Then
                    ilFound = False
                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If tgMRif(ilPrg).iVefCode = tgMVef(ilVef).iCode Then
                        ilVef = gBinarySearchVef(tgMRif(llRif).iVefCode)
                        If ilVef <> -1 Then
                            For illoop = 0 To UBound(ilMktCode) - 1 Step 1
                                If tgMVef(ilVef).iMnfVehGp3Mkt = ilMktCode(illoop) Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next illoop
                            'Exit For
                        End If
                    'Next ilVef
                End If
                If ilFound Then
                    ilVpf = gBinarySearchVpf(tgMRif(llRif).iVefCode)
                    If ilVpf <> -1 Then
                        'If slFilterType = "M" Then
                        '    If tgVpf(ilVpf).sGMedium <> "M" Then
                        '        ilFound = False
                        '    End If
                        'ElseIf slFilterType = "P" Then
                        If slFilterType = "P" Then
                            If (tgVpf(ilVpf).sGMedium <> "P") Then
                                ilFound = False
                            End If
                        ElseIf slFilterType = "A" Then
                            If (tgVpf(ilVpf).sGMedium = "P") Then   'Or (tgVpf(ilVpf).sGMedium = "M") Then
                                ilFound = False
                            End If
                        ElseIf slFilterType = "B" Then
                            'If (tgVpf(ilVpf).sGMedium = "M") Then
                            '    ilFound = False
                            'End If
                            If (tgVpf(ilVpf).sGMedium = "P") Then   'Or (tgVpf(ilVpf).sGMedium = "M") Then
                                ilVef = gBinarySearchVef(tgMRif(llRif).iVefCode)
                                If ilVef <> -1 Then
                                    If ((Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER) Then
                                        'TTP 10438 - Proposals/Contracts screen, NTR tab: podcast medium vehicle for use with ad server only doesn't appear in vehicle list WHEN the podcast Ad Server option is ON
                                        If blCheckLTF Then
                                            If gExistLtf(tgMRif(llRif).iVefCode) = False Then '5/20/2021 Bianchi - exclude podcast without programming spots
                                                ilFound = False
                                            End If
                                        End If
                                    End If
                                    
                                    If tgMVef(ilVef).sType = "P" Then
                                        'If CPM package, ignore package vehicle
                                        slSQLQuery = "Select pvfType from PVF_Package_Vehicle Where pvfCode = " & tgMVef(ilVef).lPvfCode
                                        Set pvf_rst = gSQLSelectCall(slSQLQuery)
                                        If Not pvf_rst.EOF Then
                                            If pvf_rst!pvfType = "C" Then
                                                ilFound = False
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf slFilterType = "X" Then
                            ilFound = False
                        End If
                    End If
                End If
                If ilFound Then
                    'gObtainVehicleName tgMRif(ilPrg).iVefCode, slVehName, slVehType
                    slVehSort = "999"
                    slVehName = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
                    slSortChar = "B"
                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If tgMRif(ilPrg).iVefCode = tgMVef(ilVef).iCode Then
                        ilVef = gBinarySearchVef(tgMRif(llRif).iVefCode)
                        If ilVef <> -1 Then
                            slVehName = tgMVef(ilVef).sName
                            If tgMVef(ilVef).sType = "P" Then
                                slSortChar = "A"
                            End If
                            If ilSortVehbyGroup Then
                                slVehSort = Trim$(str$(tgMVef(ilVef).iSort))
                                tlMnfSrchKey.iCode = tgMVef(ilVef).iOwnerMnfCode
                                If tlMnf.iCode <> tlMnfSrchKey.iCode Then
                                    ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        tlMnf.iGroupNo = 999
                                    End If
                                End If
                                slGpSort = Trim$(str$(tlMnf.iGroupNo))
                            End If
                            'Exit For
                        End If
                    'Next ilVef
                    If slVehName = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" Then
                        slVehName = "~~ Missing" & str$(tgMRif(llRif).iVefCode) & " ~~~"
                    End If
                    If ilSortVehbyGroup Then
                        Do While Len(slGpSort) < 3
                            slGpSort = "0" & slGpSort
                        Loop
                        Do While Len(slVehSort) < 3
                            slVehSort = "0" & slVehSort
                        Loop
                    Else
                        slGpSort = "000"
                        slVehSort = "000"
                    End If

                    slName = slSortChar & slGpSort & "|" & slVehSort & "|" & slVehName & "\" & Trim$(str$(Abs(tgMRif(llRif).iVefCode)))
                    'gFindMatch slName, 0, lbcLocal
                    'If lbcLocal.LastFound = -1 Then
                    ilFound = False
'                    For ilLoop = 0 To ilSortCode - 1 Step 1
'                        If Trim$(tlSortCode(ilLoop).sKey) = slName Then
'                            ilFound = True
'                            Exit For
'                        End If
'                    Next ilLoop

                    'TTP 10174 - 6/2/21 - Prevent Podcast vehicles with No AdServer
                    If blRequireAdServer Then
                        If tgVpf(ilVpf).sGMedium = "P" And tgMVef(ilVef).sType = "C" Then  'Podcast, but not Packages
                            ilVff = gBinarySearchVff(tgMRif(llRif).iVefCode)
                            If ilVff <> -1 Then
                                If tgVff(ilVff).iAvfCode = 0 Then
                                    ilFound = True 'Skip this one, it has No AdServer
                                End If
                            End If
                        End If
                    End If
                    
                    If Not ilFound Then
                        'If Not gOkAddStrToListBox(slName, llLen, True) Then
                        '    Exit For
                        'End If
                        'lbcLocal.AddItem slName
                        tlSortCode(ilSortCode).sKey = slName
                        If ilSortCode >= UBound(tlSortCode) Then
                            ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                        End If
                        ilSortCode = ilSortCode + 1
                        'If tgRpf(ilPrg).iVefCode >= 0 Then
                            'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                        'Else
                        '    lbcComboMster.AddItem slName
                        'End If
                    End If
                End If
            End If
        Next llRif
        lbcLocal.Clear
        'Sort then output new headers and lines
        ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
            ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            ilSortCode = ilLowLimit
            For illoop = ilLowLimit + 1 To UBound(tlSortCode) - 1 Step 1
                If Trim$(tlSortCode(ilSortCode).sKey) <> Trim$(tlSortCode(illoop).sKey) Then
                    ilSortCode = ilSortCode + 1
                    tlSortCode(ilSortCode) = tlSortCode(illoop)
                End If
            Next illoop
            ilSortCode = ilSortCode + 1
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
        End If

        'For ilLoop = 0 To lbcComboMster.ListCount - 1 Step 1
        '    slNameCode = lbcComboMster.List(ilLoop)
        '    ilRet = gParseItem(slNameCode, 1, "\", slName)
        '    If ilRet <> CP_MSG_NONE Then
        '        gPopVehViaRateCardBox = CP_MSG_PARSE
        '        Exit Function
        '    End If
        '    lbcLocal.AddItem Trim$(slName)  'Add ID to list box
        'Next ilLoop
        btrDestroy hlMnf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopVehViaRateCardBox = CP_MSG_PARSE
            Exit Function
        End If
        ilRet = gParseItem(slName, 3, "|", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopVehViaRateCardBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next illoop
    On Error Resume Next
    pvf_rst.Close
    Exit Function
gPopVehViaRateCardBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gRemoveItemFromSortCode         *
'*                                                     *
'*             Created:6/03/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove item from Sort Code List *
'*                                                     *
'*******************************************************
Sub gRemoveItemFromSortCode(ilRemoveIndex As Integer, tlSortCode() As SORTCODE)
    Dim illoop As Integer

    For illoop = ilRemoveIndex To UBound(tlSortCode) - 1 Step 1
        tlSortCode(illoop).sKey = tlSortCode(illoop + 1).sKey
    Next illoop
    ReDim Preserve tlSortCode(0 To UBound(tlSortCode) - 1) As SORTCODE

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gTestChfAdvtExt                 *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if Contract is OK to be    *
'*                     viewed by the user              *
'*      1-11-07 chg to public rtn                      *
'*******************************************************
Public Function gTestChfAdvtExt(frm As Form, ilInSlfCode As Integer, tlChfAdvtExt As CHFADVTEXT, hlVsf As Integer, ilCurrent As Integer) As Integer
    Dim ilFound As Integer
    Dim slDate As String
    Dim slStartDate As String
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length
    Dim llTodayDate As Long
    Dim ilSlfCode As Integer

    llTodayDate = gDateValue(gNow())
    ilSlfCode = ilInSlfCode
    ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length

    ilFound = gCntrOkForUser(hlVsf, ilInSlfCode, tlChfAdvtExt.lVefCode, tlChfAdvtExt.iSlfCode(), tlChfAdvtExt.sSource)
    If ilFound Then
        If ilCurrent = 0 Then   'Current
            gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slDate
            If gDateValue(slDate) < llTodayDate Then
                ilFound = False
            End If
        ElseIf ilCurrent = 2 Then
            If (tlChfAdvtExt.iStartDate(0) <> 0) Or (tlChfAdvtExt.iStartDate(1) <> 0) Or (tlChfAdvtExt.iEndDate(0) <> 0) Or (tlChfAdvtExt.iEndDate(1) <> 0) Then
                gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slDate
                If gDateValue(slStartDate) <= gDateValue(slDate) Then
                    If gDateValue(slDate) < llTodayDate Then
                        ilFound = False
                    End If
                End If
            End If
        ElseIf ilCurrent = 4 Then
            If (tlChfAdvtExt.iStartDate(0) <> 0) Or (tlChfAdvtExt.iStartDate(1) <> 0) Or (tlChfAdvtExt.iEndDate(0) <> 0) Or (tlChfAdvtExt.iEndDate(1) <> 0) Then
                gUnpackDate tlChfAdvtExt.iStartDate(0), tlChfAdvtExt.iStartDate(1), slStartDate
                If gDateValue(slStartDate) > lmInvEndDate Then
                    ilFound = False
                Else
                    gUnpackDate tlChfAdvtExt.iEndDate(0), tlChfAdvtExt.iEndDate(1), slDate
                    If gDateValue(slDate) < lmInvStartDate Then
                        ilFound = False
                    End If
                End If
            End If
        End If
    End If
    gTestChfAdvtExt = ilFound
    Exit Function
gTestChfAdvtExtErr: 'VBC NR
    gTestChfAdvtExt = False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mTestVehType                    *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test if vehicle requested       *
'*                                                     *
'*******************************************************
Private Function mTestVehType(llInVehType As Long, tlVef As VEF, Optional ilIncludeVefCode As Integer = 0) As Integer
    Dim ilOk As Integer
    Dim ilVpfIndex As Integer
    Dim ilVff As Integer
    'Dim ilVpf As Integer
    Dim illoop As Integer
    Dim llVehType As Long

    ilOk = True
    If ilIncludeVefCode > 0 Then
        If ilIncludeVefCode = tlVef.iCode Then
            mTestVehType = ilOk
            Exit Function
        End If
    End If
    llVehType = llInVehType
    If ((llVehType And VEHCONV_WO_FEED) = VEHCONV_WO_FEED) Or ((llVehType And VEHCONV_W_FEED) = VEHCONV_W_FEED) Then
        'If (llVehType And VEHEXCLUDESPORT) <> VEHEXCLUDESPORT Then
        If ((llVehType And VEHEXCLUDESPORT) <> VEHEXCLUDESPORT) And ((llVehType And VEHSPORTMINUELIVE) <> VEHSPORTMINUELIVE) And ((llVehType And VEHMOVESPORTTONONSPORT) <> VEHMOVESPORTTONONSPORT) Then
            llVehType = llVehType Or VEHSPORT
        End If
    End If
    Select Case Trim$(tlVef.sType)
        Case "C"
            If ((llVehType And VEHCONV_WO_FEED) <> VEHCONV_WO_FEED) Or ((llVehType And VEHCONV_W_FEED) <> VEHCONV_W_FEED) Then
                'ilVpfIndex = gVpfFind(Frm, tlVef.iCode)
                ilVpfIndex = -1
                'For ilLoop = 0 To UBound(tgVpf) Step 1
                '    If tlVef.iCode = tgVpf(ilLoop).iVefKCode Then
                    illoop = gBinarySearchVpf(tlVef.iCode)
                    If illoop <> -1 Then
                        ilVpfIndex = illoop
                '        Exit For
                    End If
                'Next ilLoop
                If ilVpfIndex = -1 Then
                    mTestVehType = False
                    Exit Function
                End If
                If (llVehType And VEHCONV_WO_FEED) <> VEHCONV_WO_FEED Then
                    'If tgVpf(ilVpfIndex).iGMnfNCode(1) = 0 Then
                    If tgVpf(ilVpfIndex).iGMnfNCode(0) = 0 Then
                        ilOk = False
                    End If
                End If
                If (llVehType And VEHCONV_W_FEED) <> VEHCONV_W_FEED Then
                    'If tgVpf(ilVpfIndex).iGMnfNCode(1) > 0 Then
                    If tgVpf(ilVpfIndex).iGMnfNCode(0) > 0 Then
                        ilOk = False
                    End If
                End If
            End If
            'If Log vehicle requested, then exclude vehicles that have log vehicle
            If ilOk And ((llVehType And VEHLOG) = VEHLOG) Then
                If (llVehType And VEHLOGVEHICLE) <> VEHLOGVEHICLE Then
                    If tlVef.iVefCode > 0 Then
                        '10/21/12: Add vehicle option about merge
                        'ilOk = False
                        If ((llVehType And VEHTESTLOGMERGE) = VEHTESTLOGMERGE) Then
                            For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                                If tlVef.iCode = tgVff(ilVff).iVefCode Then
                                    If tgVff(ilVff).sMergeTraffic <> "S" Then
                                        ilOk = False
                                    End If
                                    Exit For
                                End If
                            Next ilVff
                        Else
                            ilOk = False
                        End If
                    End If
                End If
            End If
            
        Case "S"
            If (llVehType And VEHSELLING) <> VEHSELLING Then
                ilOk = False
            End If
        
        Case "A"
            If (llVehType And VEHAIRING) <> VEHAIRING Then
                ilOk = False
            End If
        Case "L"
            If (llVehType And VEHLOG) <> VEHLOG Then
                ilOk = False
            End If
        Case "V"
            If (llVehType And VEHVIRTUAL) <> VEHVIRTUAL Then
                ilOk = False
            End If
        Case "T"
            If (llVehType And VEHSIMUL) <> VEHSIMUL Then
                ilOk = False
            End If
        Case "P"
            ilVpfIndex = gBinarySearchVpf(tlVef.iCode)
            If ilVpfIndex <> -1 Then
                If tgVpf(ilVpfIndex).sGMedium <> "P" Then
                    If ((llVehType And VEHPACKAGE) <> VEHPACKAGE) And ((llVehType And VEHSTDPKG) <> VEHSTDPKG) Then
                        ilOk = False
                    Else
                        If ((llVehType And VEHSTDPKG) = VEHSTDPKG) And ((llVehType And VEHPACKAGE) <> VEHPACKAGE) Then
                            If tlVef.lPvfCode <= 0 Then
                                ilOk = False
                            End If
                        End If
                    End If
                Else 'Podcast
                    If ((llVehType And VEHPACKAGE) <> VEHPACKAGE) And ((llVehType And VEHCPMPKG) <> VEHCPMPKG) Then
                        ilOk = False
                    Else
                        If ((llVehType And VEHCPMPKG) = VEHCPMPKG) And ((llVehType And VEHPACKAGE) <> VEHPACKAGE) Then
                            If tlVef.lPvfCode <= 0 Then
                                ilOk = False
                            End If
                        End If
                    End If
                End If
            Else
                ilOk = False
            End If
        Case "R"
            ilOk = False
            If ((llVehType And VEHREP_WO_CLUSTER) = VEHREP_WO_CLUSTER) And ((llVehType And VEHREP_W_CLUSTER) = VEHREP_W_CLUSTER) Then
                ilOk = True
            Else
                If (llVehType And VEHREP_WO_CLUSTER) = VEHREP_WO_CLUSTER Then
                    For illoop = 0 To UBound(igRepVefCode) - 1 Step 1
                        If tlVef.iCode = igRepVefCode(illoop) Then
                            ilOk = True
                            Exit For
                        End If
                    Next illoop
                End If
                If (llVehType And VEHREP_W_CLUSTER) = VEHREP_W_CLUSTER Then
                    For illoop = 0 To UBound(igMktVefCode) - 1 Step 1
                        If tlVef.iCode = igMktVefCode(illoop) Then
                            ilOk = True
                            Exit For
                        End If
                    Next illoop
                End If
            End If
        Case "N"
            If (llVehType And VEHNTR) <> VEHNTR Then
                ilOk = False
            End If
        Case "G"
            'If (llVehType And VEHSPORT) <> VEHSPORT Then
            '    ilOk = False
            'End If
            If (llVehType And VEHSPORTMINUELIVE) = VEHSPORTMINUELIVE Then
                ilVpfIndex = -1
                illoop = gBinarySearchVpf(tlVef.iCode)
                If illoop <> -1 Then
                    ilVpfIndex = illoop
                    If tgVpf(ilVpfIndex).sGenLog = "L" Then
                        ilOk = False
                    End If
                End If
            ElseIf (llVehType And VEHMOVESPORTTONONSPORT) = VEHMOVESPORTTONONSPORT Then
                For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                    If tlVef.iCode = tgVff(ilVff).iVefCode Then
                        If tgVff(ilVff).sMoveSportToNon <> "Y" Then
                            ilOk = False
                        End If
                        Exit For
                    End If
                Next ilVff
            ElseIf (llVehType And VEHMOVESPORTTOSPORT) = VEHMOVESPORTTOSPORT Then
                For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                    If tlVef.iCode = tgVff(ilVff).iVefCode Then
                        If tgVff(ilVff).sMoveSportToSport <> "Y" Then
                            ilOk = False
                        End If
                        Exit For
                    End If
                Next ilVff
            Else
                If (llVehType And VEHSPORT) <> VEHSPORT Then
                    ilOk = False
                Else
                    If tlVef.iVefCode > 0 Then
                        If (llVehType And VEHTESTLOGMERGE) = VEHTESTLOGMERGE Then
                            For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                                If tlVef.iCode = tgVff(ilVff).iVefCode Then
                                    If tgVff(ilVff).sMergeTraffic <> "S" Then
                                        ilOk = False
                                    End If
                                    Exit For
                                End If
                            Next ilVff
                        End If
                    End If
                End If
            End If
    End Select
    If (llVehType And VEHIMPORTAFFILIATESPOTS) <> VEHIMPORTAFFILIATESPOTS Then
        ilVpfIndex = -1
        illoop = gBinarySearchVpf(tlVef.iCode)
        If illoop <> -1 Then
            ilVpfIndex = illoop
            If (Asc(tgVpf(ilVpfIndex).sUsingFeatures1) And IMPORTAFFILIATESPOTS) = IMPORTAFFILIATESPOTS Then
                ilOk = False
            End If
        End If
    End If
    Select Case tlVef.sState
        Case "A"
            If (llVehType And ACTIVEVEH) <> ACTIVEVEH Then
                ilOk = False
            Else
                If tlVef.iVefCode > 0 Then
                    If (llVehType And VEHTESTLOGMERGE) = VEHTESTLOGMERGE Then
                        For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                            If tlVef.iCode = tgVff(ilVff).iVefCode Then
                                If tgVff(ilVff).sMergeTraffic <> "S" Then
                                    ilOk = False
                                End If
                                Exit For
                            End If
                        Next ilVff
                    End If
                End If
            End If
        Case "D"
            If (llVehType And DORMANTVEH) <> DORMANTVEH Then
                ilOk = False
            End If
    End Select
    'Moved to gPopUserVehicleByMkt
    'If ilOk Then
    '    'Market vehicle selection
    '    If ((llVehType And VEHBYMKT) = VEHBYMKT) And (tgSpf.sMktBase = "Y") Then
    '        ilOk = False
    '        For ilLoop = 0 To UBound(igMktCode) - 1 Step 1
    '            If tlVef.iMnfVehGp3Mkt = igMktCode(ilLoop) Then
    '                ilOk = True
    '                Exit For
    '            End If
    '        Next ilLoop
    '    End If
    'End If
    If ilOk Then
        'Market vehicle selection
        If ((llVehType And VEHBYPASSNOLOG) = VEHBYPASSNOLOG) Then
            'For ilLoop = 0 To UBound(tgVpf) Step 1
            '    If tlVef.iCode = tgVpf(ilLoop).iVefKCode Then
                illoop = gBinarySearchVpf(tlVef.iCode)
                If illoop <> -1 Then
                    If tgVpf(illoop).sGenLog = "N" Then
                        ilOk = False
                    End If
            '        Exit For
                End If
            'Next ilLoop
        End If
    End If
    If ilOk Then
        If (llVehType And VEHBYPASSWEGENER_OLA) = VEHBYPASSWEGENER_OLA Then
            ilVpfIndex = -1
            illoop = gBinarySearchVpf(tlVef.iCode)
            If illoop <> -1 Then
                ilVpfIndex = illoop
                If (tgVpf(ilVpfIndex).sWegenerExport = "Y") Or (tgVpf(ilVpfIndex).sOLAExport = "Y") Then
                    ilOk = False
                End If
            End If
        End If
    End If
    If ilOk Then
        If (llVehType And VEHLIVELOG) = VEHLIVELOG Then
            ilVpfIndex = -1
            illoop = gBinarySearchVpf(tlVef.iCode)
            If illoop <> -1 Then
                ilVpfIndex = illoop
                If (tgVpf(ilVpfIndex).sGenLog <> "L") And (tgVpf(ilVpfIndex).sGenLog <> "A") Then
                    ilOk = False
                End If
            End If
        End If
    End If
    If (ilOk) And ((llVehType And VEHONINSERTION) = VEHONINSERTION) And ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
        If (Trim$(tlVef.sType) = "C") Or (Trim$(tlVef.sType) = "S") Or (Trim$(tlVef.sType) = "R") Or (Trim$(tlVef.sType) = "N") Then
            If Not gIsOnInsertions(tlVef.iCode) Then
                ilOk = False
            End If
        Else
            ilOk = False
        End If
    End If
    If (ilOk) And ((llVehType And VEHEXCLUDEIMPORTINVOICESPOTS) = VEHEXCLUDEIMPORTINVOICESPOTS) And ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
        If (Trim$(tlVef.sType) = "C") Or (Trim$(tlVef.sType) = "S") Or (Trim$(tlVef.sType) = "R") Or (Trim$(tlVef.sType) = "N") Then
            If gIsImportInvoicedSpots(tlVef.iCode) Then
                ilOk = False
            End If
        End If
    End If
    
    '--is is a Podcast vehicle?  exclude if has no Programming Defined (no ltf), when vehType includes VEHEXCLUDEPODNOPRGM and Site Option Feature8 = PODADSERVER
    If ilOk And (Trim$(tlVef.sType) = "P" Or Trim$(tlVef.sType) = "C") Then  'Package or Conventional
        'If VEHEXCLUDEPODNOPRGM, Check if Medium ="P" Podcast Vehicle, Check SiteOptions Pod Ad server, if Checked, Exclude vehicle with no Programming (Ltf records)
        If (llVehType And VEHEXCLUDEPODNOPRGM) = VEHEXCLUDEPODNOPRGM Then
            If tgVpf(ilVpfIndex).sGMedium = "P" Then
                If ((Asc(tgSaf(0).sFeatures8) And PODADSERVER) = PODADSERVER) Then
                    If gExistLtf(tlVef.iCode) = False Then
                        ilOk = False
                    End If
                End If
            End If
        End If
    End If

    mTestVehType = ilOk
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchAdf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Advt index into tgCommAdf*
'*                                                     *
'*******************************************************
Public Function gBinarySearchAdf(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim ilRet As Integer
    Dim ilfirstTime As Integer

    ilfirstTime = True
    Do
        ilMin = LBound(tgCommAdf)
        ilMax = UBound(tgCommAdf) - 1
        Do While ilMin <= ilMax
            ilMiddle = (ilMin + ilMax) \ 2
            If ilCode = tgCommAdf(ilMiddle).iCode Then
                'found the match
                gBinarySearchAdf = ilMiddle
                Exit Function
            ElseIf ilCode < tgCommAdf(ilMiddle).iCode Then
                ilMax = ilMiddle - 1
            Else
                'search the right half
                ilMin = ilMiddle + 1
            End If
        Loop
        gBinarySearchAdf = -1
        If Not ilfirstTime Then
            Exit Function
        End If
        ilfirstTime = False
        hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hmAdf
            Exit Function
        End If
        imAdfRecLen = Len(tmAdf)
        tmAdfSrchKey.iCode = ilCode
        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hmAdf
            Exit Function
        End If
        tgCommAdf(UBound(tgCommAdf)).iCode = tmAdf.iCode
        tgCommAdf(UBound(tgCommAdf)).sName = tmAdf.sName
        tgCommAdf(UBound(tgCommAdf)).sAbbr = tmAdf.sAbbr
        tgCommAdf(UBound(tgCommAdf)).iMnfSort = tmAdf.iMnfSort
        tgCommAdf(UBound(tgCommAdf)).sBillAgyDir = tmAdf.sBillAgyDir
        tgCommAdf(UBound(tgCommAdf)).sState = tmAdf.sState
        tgCommAdf(UBound(tgCommAdf)).sAllowRepMG = tmAdf.sAllowRepMG
        tgCommAdf(UBound(tgCommAdf)).sBonusOnInv = tmAdf.sBonusOnInv
        tgCommAdf(UBound(tgCommAdf)).sRepInvGen = tmAdf.sRepInvGen
        tgCommAdf(UBound(tgCommAdf)).sFirstCntrAddr = tmAdf.sCntrAddr(0)
        tgCommAdf(UBound(tgCommAdf)).sPolitical = tmAdf.sPolitical  '5-26-06
        tgCommAdf(UBound(tgCommAdf)).sAddrID = tmAdf.sAddrID
        tgCommAdf(UBound(tgCommAdf)).iTrfCode = tmAdf.iTrfCode
        tgCommAdf(UBound(tgCommAdf)).iMnfInvTerms = tmAdf.iMnfInvTerms
        'ReDim Preserve tgCommAdf(1 To UBound(tgCommAdf) + 1) As ADFEXT
        ReDim Preserve tgCommAdf(0 To UBound(tgCommAdf) + 1) As ADFEXT
        'Sort by code so that binary search can be used
        'If UBound(tgCommAdf) - 1 > 1 Then
        If UBound(tgCommAdf) - 1 > 0 Then
            'ArraySortTyp fnAV(tgCommAdf(), 1), UBound(tgCommAdf) - 1, 0, LenB(tgCommAdf(1)), 0, -1, 0
            ArraySortTyp fnAV(tgCommAdf(), 0), UBound(tgCommAdf), 0, LenB(tgCommAdf(0)), 0, -1, 0
        End If
    Loop
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchRdf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Daypart index into tgMRdf*
'*                                                     *
'*******************************************************
Public Function gBinarySearchRdf(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tgMRdf)
    ilMax = UBound(tgMRdf) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgMRdf(ilMiddle).iCode Then
            'found the match
            gBinarySearchRdf = ilMiddle
            Exit Function
        ElseIf ilCode < tgMRdf(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchRdf = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchAgf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain vehicle Agy into tgCommAgf*
'*                                                     *
'*******************************************************
Public Function gBinarySearchAgf(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim ilRet As Integer
    Dim ilfirstTime As Integer

    ilfirstTime = True
    Do
        ilMin = LBound(tgCommAgf)
        ilMax = UBound(tgCommAgf) - 1
        Do While ilMin <= ilMax
            ilMiddle = (ilMin + ilMax) \ 2
            If ilCode = tgCommAgf(ilMiddle).iCode Then
                'found the match
                gBinarySearchAgf = ilMiddle
                Exit Function
            ElseIf ilCode < tgCommAgf(ilMiddle).iCode Then
                ilMax = ilMiddle - 1
            Else
                'search the right half
                ilMin = ilMiddle + 1
            End If
        Loop
        gBinarySearchAgf = -1
        If Not ilfirstTime Then
            Exit Function
        End If
        ilfirstTime = False
        hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hmAgf
            Exit Function
        End If
        imAgfRecLen = Len(tmAgf)
        tmAgfSrchKey.iCode = ilCode
        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hmAgf
            Exit Function
        End If
        tgCommAgf(UBound(tgCommAgf)).iCode = tmAgf.iCode
        tgCommAgf(UBound(tgCommAgf)).sName = tmAgf.sName
        tgCommAgf(UBound(tgCommAgf)).sCityID = tmAgf.sCityID
        tgCommAgf(UBound(tgCommAgf)).sCreditRestr = tmAgf.sCreditRestr
        tgCommAgf(UBound(tgCommAgf)).iMnfSort = tmAgf.iMnfSort
        tgCommAgf(UBound(tgCommAgf)).sState = tmAgf.sState
        tgCommAgf(UBound(tgCommAgf)).iTrfCode = tmAgf.iTrfCode
        tgCommAgf(UBound(tgCommAgf)).iMnfInvTerms = tmAgf.iMnfInvTerms
        tgCommAgf(UBound(tgCommAgf)).iCommPct = tmAgf.iComm
        tgCommAgf(UBound(tgCommAgf)).sRefID = tmAgf.sAgfRefID
        'ReDim Preserve tgCommAgf(1 To UBound(tgCommAgf) + 1) As AGFEXT
        ReDim Preserve tgCommAgf(0 To UBound(tgCommAgf) + 1) As AGFEXT
        'Sort by code so that binary search can be used
        'If UBound(tgCommAgf) - 1 > 1 Then
        If UBound(tgCommAgf) - 1 > 0 Then
            'ArraySortTyp fnAV(tgCommAgf(), 1), UBound(tgCommAgf) - 1, 0, LenB(tgCommAgf(1)), 0, -1, 0
            ArraySortTyp fnAV(tgCommAgf(), 0), UBound(tgCommAgf), 0, LenB(tgCommAgf(0)), 0, -1, 0
        End If
    Loop
End Function

'
'
'           gBinarySearchANf - find the matching Named Avail code in array that contains
'           index references to Named AVails
'           <input> ilAnfCode = Named avail code to match
 '           return - index to matching named avail entry
'                    -1 if not found
Public Function gBinarySearchAnf(ilAnfCode As Integer, tlAnf() As ANF) As Integer
    Dim ilMiddle As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    ilMin = LBound(tlAnf)
    ilMax = UBound(tlAnf)
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilAnfCode = tlAnf(ilMiddle).iCode Then
            'found the match
            gBinarySearchAnf = ilMiddle
            Exit Function
        ElseIf ilAnfCode < tlAnf(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchAnf = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchSlf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Advt index into tgCommAdf*
'*                                                     *
'*******************************************************
Public Function gBinarySearchPopDnfInfo(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tgPopDnfInfo)
    ilMax = UBound(tgPopDnfInfo) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgPopDnfInfo(ilMiddle).iDnfCode Then
            'found the match
            gBinarySearchPopDnfInfo = ilMiddle
            Exit Function
        ElseIf ilCode < tgPopDnfInfo(ilMiddle).iDnfCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchPopDnfInfo = -1
End Function

Public Function gBinarySearchDnfInfo(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tgBookNameInfo)
    ilMax = UBound(tgBookNameInfo) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgBookNameInfo(ilMiddle).iDnfCode Then
            'found the match
            gBinarySearchDnfInfo = ilMiddle
            Exit Function
        ElseIf ilCode < tgBookNameInfo(ilMiddle).iDnfCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchDnfInfo = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchSlf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Advt index into tgCommAdf*
'*                                                     *
'*******************************************************
Public Function gBinarySearchSlf(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tgMSlf)
    ilMax = UBound(tgMSlf) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgMSlf(ilMiddle).iCode Then
            'found the match
            gBinarySearchSlf = ilMiddle
            Exit Function
        ElseIf ilCode < tgMSlf(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchSlf = -1
End Function

'public function gBinarySearchAVG(ilCode as Integer) as Integer
'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchVef                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain vehicle index into tgMVef*
'*                                                     *
'*******************************************************
Public Function gBinarySearchVef(ilCode As Integer, Optional blCallFromPopVpf = False) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim ilRet As Integer
    Dim ilfirstTime As Integer
    Dim ilVef As Integer

    ilfirstTime = True
    Do
        ilMin = LBound(tgMVef)
        ilMax = UBound(tgMVef) - 1
        Do While ilMin <= ilMax
            ilMiddle = (ilMin + ilMax) \ 2
            If ilCode = tgMVef(ilMiddle).iCode Then
                'found the match
                gBinarySearchVef = ilMiddle
                Exit Function
            ElseIf ilCode < tgMVef(ilMiddle).iCode Then
                ilMax = ilMiddle - 1
            Else
                'search the right half
                ilMin = ilMiddle + 1
            End If
        Loop
        gBinarySearchVef = -1
        If Not ilfirstTime Then
            Exit Function
        End If
        ilfirstTime = False
        hmVef = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hmVef
            Exit Function
        End If
        imVefRecLen = Len(tmVef)
        tmVefSrchKey.iCode = ilCode
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hmVef
            Exit Function
        End If
        tgMVef(UBound(tgMVef)) = tmVef
        'ReDim Preserve tgMVef(1 To UBound(tgMVef) + 1) As VEF
        ReDim Preserve tgMVef(0 To UBound(tgMVef) + 1) As VEF
        'Sort by code so that binary search can be used
        'If UBound(tgMVef) > 2 Then
        If UBound(tgMVef) > 1 Then
            'ArraySortTyp fnAV(tgMVef(), 1), UBound(tgMVef) - 1, 0, LenB(tgMVef(1)), 0, -1, 0
            ArraySortTyp fnAV(tgMVef(), 0), UBound(tgMVef), 0, LenB(tgMVef(0)), 0, -1, 0
        End If
        ReDim tgVefName(0 To UBound(tgMVef)) As VEFNAME
        For ilVef = 0 To UBound(tgMVef) Step 1
            tgVefName(ilVef).sName = UCase(tgMVef(ilVef).sName)
            tgVefName(ilVef).iCode = tgMVef(ilVef).iCode
        Next ilVef
        ArraySortTyp fnAV(tgVefName(), 0), UBound(tgVefName), 0, LenB(tgVefName(0)), 0, LenB(tgVefName(0).sName), 0 'VBC NR
        If Not blCallFromPopVpf Then
            ilRet = gBinarySearchVpf(tmVef.iCode)
        End If
    Loop
End Function

Public Function gBinarySearchVefName(slInName As String) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim ilresult As Integer
    Dim slName As String
    
    On Error GoTo err:
    ilMin = LBound(tgVefName)
    ilMax = UBound(tgVefName) - 1
    slName = UCase(Trim$(slInName))
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        'Need to use Binary so that StrComp(KEX-AM,KEXA-FM,...) yields correct results (-1 not 1 which is what vbTextCompare yields)
        ilresult = StrComp(Trim(tgVefName(ilMiddle).sName), slName, vbBinaryCompare)
        Select Case ilresult
            Case 0:
                gBinarySearchVefName = ilMiddle  ' Found it !
                Exit Function
            Case 1:
                ilMax = ilMiddle - 1
            Case -1:
                ilMin = ilMiddle + 1
        End Select
        
    Loop
err:
    gBinarySearchVefName = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchVpf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain vehicle index into tgVpf *
'*                                                     *
'*******************************************************
Public Function gBinarySearchVpf(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim ilRet As Integer
    Dim ilfirstTime As Integer
    Dim hlVpf As Integer

    ilfirstTime = True
    Do
        ilMin = LBound(tgVpf)
        ilMax = UBound(tgVpf)
        Do While ilMin <= ilMax
            ilMiddle = (ilMin + ilMax) \ 2
            If ilCode = tgVpf(ilMiddle).iVefKCode Then
                'found the match
                gBinarySearchVpf = ilMiddle
                Exit Function
            ElseIf ilCode < tgVpf(ilMiddle).iVefKCode Then
                ilMax = ilMiddle - 1
            Else
                'search the right half
                ilMin = ilMiddle + 1
            End If
        Loop
        gBinarySearchVpf = -1
        If Not ilfirstTime Then
            Exit Function
        End If
        ilfirstTime = False
        hlVpf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hlVpf
            Exit Function
        End If
        imVpfRecLen = Len(tmVpf)
        tmVpfSrchKey.iCode = ilCode
        ilRet = btrGetEqual(hlVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hlVpf
            Exit Function
        End If
        btrDestroy hlVpf
        ReDim Preserve tgVpf(0 To UBound(tgVpf) + 1) As VPF
        tgVpf(UBound(tgVpf)) = tmVpf
        'Sort by code so that binary search can be used
        If UBound(tgVpf) > 1 Then
            ArraySortTyp fnAV(tgVpf(), 0), UBound(tgVpf) + 1, 0, LenB(tgVpf(0)), 0, -1, 0
        End If
    Loop
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchVpfPlus            *
'*                                                     *
'*             Created:9/21/05       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain vehicle index into tgVpf *
'*                     and update array                *
'*                                                     *
'*******************************************************
Public Function gBinarySearchVpfPlus(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim blOpen As Boolean
    'Dim hlVpf As Integer

    ilIndex = -1
    ilMin = LBound(tgVpf)
    ilMax = UBound(tgVpf)
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgVpf(ilMiddle).iVefKCode Then
            'found the match
            ilIndex = ilMiddle
            Exit Do
        ElseIf ilCode < tgVpf(ilMiddle).iVefKCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    'hlVpf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    'ilRet = btrOpen(hlVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet <> BTRV_ERR_NONE Then
    '    btrDestroy hlVpf
    '    gBinarySearchVpfPlus = -1
    '    Exit Function
    'End If
    imVpfRecLen = Len(tmVpf)
    tmVpfSrchKey.iCode = ilCode
    ilRet = btrGetEqual(hgVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet <> BTRV_ERR_NONE Then
        gOpenVpf
        imVpfRecLen = Len(tmVpf)
        tmVpfSrchKey.iCode = ilCode
        ilRet = btrGetEqual(hgVpf, tmVpf, imVpfRecLen, tmVpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        gCloseVpf
        If ilRet <> BTRV_ERR_NONE Then
            'btrDestroy hlVpf
            gBinarySearchVpfPlus = -1
            Exit Function
        End If
    End If
    'btrDestroy hlVpf
    If ilIndex = -1 Then
        ReDim Preserve tgVpf(0 To UBound(tgVpf) + 1) As VPF
        tgVpf(UBound(tgVpf)) = tmVpf
        'Sort by code so that binary search can be used
        If UBound(tgVpf) > 1 Then
            ArraySortTyp fnAV(tgVpf(), 0), UBound(tgVpf) + 1, 0, LenB(tgVpf(0)), 0, -1, 0
        End If
        ilIndex = gBinarySearchVpf(ilCode)
    Else
        tgVpf(ilIndex) = tmVpf
    End If
    gBinarySearchVpfPlus = ilIndex
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchVpf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain vehicle index into tgVff *
'*                                                     *
'*******************************************************
Public Function gBinarySearchVff(ilVefCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim ilRet As Integer
    
    On Error GoTo gBinarySearchVfffErr
    
    ilMin = LBound(tgVff)
    ilMax = UBound(tgVff)
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilVefCode = tgVff(ilMiddle).iVefCode Then
            'found the match
            gBinarySearchVff = ilMiddle
            Exit Function
        ElseIf ilVefCode < tgVff(ilMiddle).iVefCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchVff = -1
    Exit Function
gBinarySearchVfffErr:
    gBinarySearchVff = -1
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gUserAllowedVehicle             *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if specified vehicle  *
'*                     can be seen by User             *
'*                                                     *
'*******************************************************
Public Function gUserAllowedVehicle(ilVefCode As Integer) As Integer
'
'   ilRet = gUserAllowedVehicle(ilVefCode)
'
'   where:
'       ilVefCode (I)- Vehicle code to be checked
'       ilRet (O)- True if User is allowed to access the vehicle;
'                  False if User is not allowed to access the vehicle
'
    Dim illoop As Integer
    Dim ilVef As Integer

    gUserAllowedVehicle = False
    'Check if master user
    If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Then
        gUserAllowedVehicle = True
        Exit Function
    Else
        If tgUrf(0).iMnfHubCode > 0 Then
            ilVef = gBinarySearchVef(ilVefCode)
            If ilVef >= 0 Then
                If tgMVef(ilVef).iMnfHubCode = tgUrf(0).iMnfHubCode Then
                    gUserAllowedVehicle = True
                    Exit Function
                Else
                    gUserAllowedVehicle = False
                    Exit Function
                End If
            Else
                gUserAllowedVehicle = False
                Exit Function
            End If
        End If
        'Check if All vehicles selected for User
        For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
            If (tgUrf(illoop).iVefCode = 0) Then
                gUserAllowedVehicle = True
                Exit Function
            End If
        Next illoop
    End If
    'Check if specified vehicle was defined for User
    For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
        If (tgUrf(illoop).iVefCode = ilVefCode) Then
            gUserAllowedVehicle = True
            Exit Function
        End If
    Next illoop
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainSAF                      *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgSAF array            *
'*                                                     *
'*******************************************************
Public Function gObtainSAF() As Integer
    Dim hlSaf As Integer
    Dim ilRet As Integer
    Dim tlSaf As SAF
    Dim ilSafRecLen As Integer
    Dim ilVefCode As Integer
    Dim ilCode As Integer
    Dim illoop As Integer

    ilSafRecLen = Len(tlSaf)  'btrRecordLength(hlUrf)  'Get and save record length
    hlSaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    ReDim tgSaf(0 To 0) As SAF
    If ilRet <> BTRV_ERR_NONE Then
        gObtainSAF = False
        Exit Function
    End If
    ilRet = btrGetFirst(hlSaf, tlSaf, ilSafRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        tgSaf(UBound(tgSaf)) = tlSaf
        ReDim Preserve tgSaf(0 To UBound(tgSaf) + 1) As SAF
        ilRet = btrGetNext(hlSaf, tlSaf, ilSafRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilRet = btrClose(hlSaf)
    btrDestroy hlSaf
    'Swap safCode and safVefCode temporary
    For illoop = 0 To UBound(tgSaf) - 1 Step 1
        ilVefCode = tgSaf(illoop).iVefCode
        tgSaf(illoop).iVefCode = tgSaf(illoop).iCode
        tgSaf(illoop).iCode = ilVefCode
    Next illoop
    If UBound(tgSaf) - 1 > 0 Then
        ArraySortTyp fnAV(tgSaf(), 0), UBound(tgSaf), 0, LenB(tgSaf(0)), 0, -1, 0
    End If
    For illoop = 0 To UBound(tgSaf) - 1 Step 1
        ilCode = tgSaf(illoop).iVefCode
        tgSaf(illoop).iVefCode = tgSaf(illoop).iCode
        tgSaf(illoop).iCode = ilCode
    Next illoop
    gObtainSAF = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchRdf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Daypart index into tgMRdf*
'*                                                     *
'*******************************************************
Public Function gBinarySearchSaf(ilVefCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tgSaf)
    ilMax = UBound(tgSaf) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilVefCode = tgSaf(ilMiddle).iVefCode Then
            'found the match
            gBinarySearchSaf = ilMiddle
            Exit Function
        ElseIf ilVefCode < tgSaf(ilMiddle).iVefCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchSaf = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainCntr                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Get Contract (Chf), Lines (Clf),*
'*                     Flight (Cff)                    *
'*                                                     *
'*******************************************************
Function gObtainCntrPlusGame(hlChf As Integer, hlClf As Integer, hlCff As Integer, hlCgf As Integer, llChfCode As Long, ilIncludeHistory As Integer, tlChf As CHF, tlClf() As CLFLIST, tlCff() As CFFLIST, tlCgf() As CGFLIST) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilMsfRecLen                   tlMsfSrchKey2                                           *
'******************************************************************************************

    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim ilVef As Integer
    Dim ilCgfUpper As Integer
    Dim ilCgfRecLen As Integer
    Dim tlCgfSrchKey1 As CGFKEY1    'CntrNo; CntRevNo; PropVer
    Dim ilMsfUpper As Integer

    ReDim tlCgf(0 To 0) As CGFLIST
    tlCgf(0).iStatus = -1 'Not Used
    tlCgf(0).iNextCgf = -1
    ilCgfUpper = UBound(tlCgf)
    ReDim tlMsf(0 To 0) As MSFLIST
    tlMsf(0).iStatus = -1 'Not Used
    ilMsfUpper = UBound(tlMsf)

    ilRet = gObtainCntr(hlChf, hlClf, hlCff, llChfCode, ilIncludeHistory, tlChf, tlClf(), tlCff())
    If Not ilRet Then
        gObtainCntrPlusGame = False
        Exit Function
    End If
    For ilClf = LBound(tlClf) To UBound(tlClf) - 1 Step 1
        ilVef = gBinarySearchVef(tlClf(ilClf).ClfRec.iVefCode)
        If ilVef <> -1 Then
            If tgMVef(ilVef).sType = "G" Then
                ilCgfRecLen = Len(tlCgf(0).CgfRec)
                tlCgfSrchKey1.lClfCode = tlClf(ilClf).ClfRec.lCode
                ilRet = btrGetEqual(hlCgf, tlCgf(ilCgfUpper).CgfRec, ilCgfRecLen, tlCgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE) And (tlCgf(ilCgfUpper).CgfRec.lClfCode = tlClf(ilClf).ClfRec.lCode)
                    If tlClf(ilClf).iFirstCgf = -1 Then
                        tlClf(ilClf).iFirstCgf = ilCgfUpper
                    Else
                        tlCgf(ilCgfUpper - 1).iNextCgf = ilCgfUpper
                    End If
                    gUnpackDateLong tlCgf(ilCgfUpper).CgfRec.iAirDate(0), tlCgf(ilCgfUpper).CgfRec.iAirDate(1), tlCgf(ilCgfUpper).lStartDate
                    tlCgf(ilCgfUpper).lEndDate = tlCgf(ilCgfUpper).lStartDate
                    tlCgf(ilCgfUpper).iNextCgf = -1
                    tlCgf(ilCgfUpper).iStatus = 1 'Old and retain
                    ilCgfUpper = ilCgfUpper + 1
                    ReDim Preserve tlCgf(0 To ilCgfUpper) As CGFLIST
                    tlCgf(ilCgfUpper).iStatus = -1 'Not Used
                    tlCgf(ilCgfUpper).iNextCgf = -1
                    ilRet = btrGetNext(hlCgf, tlCgf(ilCgfUpper).CgfRec, ilCgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            Else
                tlClf(ilClf).iFirstCgf = -1
            End If
        Else
            tlClf(ilClf).iFirstCgf = -1
        End If
    Next ilClf
    'If tlChf.lMsfIDCode > 0 Then
'        ilMsfRecLen = Len(tlMsf(0).MsfRec)
'        tlMsfSrchKey2.lChfCode = tlChf.lCode
'        ilRet = btrGetEqual(hlMsf, tlMsf(ilMsfUpper).MsfRec, ilMsfRecLen, tlMsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'        Do While (ilRet = BTRV_ERR_NONE) And (tlMsf(ilMsfUpper).MsfRec.lChfCode = tlChf.lCode)
'            tlMsf(ilMsfUpper).iStatus = 1 'Old and retain
'            ilMsfUpper = ilMsfUpper + 1
'            ReDim Preserve tlMsf(0 To ilMsfUpper) As MSFLIST
'            tlMsf(ilMsfUpper).iStatus = -1 'Not Used
'            ilRet = btrGetNext(hlMsf, tlMsf(ilMsfUpper).MsfRec, ilMsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
    'End If
    gObtainCntrPlusGame = True
    Exit Function
End Function

Public Function gGetGameDates(hlLcf As Integer, hlGhf As Integer, hlGsf As Integer, ilVefCode As Integer, tlTeam() As MNF, tlGsfInfo() As GSFINFO, Optional llSeasonGhfCode As Long = 0)
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim ilTeam As Integer
    Dim slSortGameNo As String
    Dim slSortDate As String
    Dim llSortTime As Long
    Dim slSortTime As String
    Dim tlGhf As GHF        'GHF record image
    Dim tlGhfSrchKey1 As GHFKEY1    'GHF key record image
    Dim ilGhfRecLen As Integer        'GHF record length
    Dim ilLcfRecLen As Integer
    Dim tlLcf As LCF               'LCF record image-current
    Dim tlLcfSrchKey2 As LCFKEY2     'LCF key record image
    Dim tlGsf As GSF        'GSF record image
    Dim tlGsfSrchKey1 As GSFKEY1    'GSF key record image
    Dim ilGsfRecLen As Integer        'GSF record length
    Dim ilGhf As Integer
    ReDim llGhfCode(0 To 0) As Long

    ReDim tlGsfInfo(0 To 0) As GSFINFO
    ilGhfRecLen = Len(tlGhf)
    ilGsfRecLen = Len(tlGsf)
    ilLcfRecLen = Len(tlLcf)

    tlGhfSrchKey1.iVefCode = ilVefCode
    ilRet = btrGetEqual(hlGhf, tlGhf, ilGhfRecLen, tlGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
    Do While (ilRet = BTRV_ERR_NONE) And (tlGhf.iVefCode = ilVefCode)
        If (tlGhf.lCode = llSeasonGhfCode) Or (llSeasonGhfCode = 0) Then
            llGhfCode(UBound(llGhfCode)) = tlGhf.lCode
            ReDim Preserve llGhfCode(0 To UBound(llGhfCode) + 1) As Long
        End If
        ilRet = btrGetNext(hlGhf, tlGhf, ilGhfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If UBound(llGhfCode) <= LBound(llGhfCode) Then
        gGetGameDates = False
        Exit Function
    End If
    'tlLcfSrchKey1.iVefCode = ilVefCode
    'tlLcfSrchKey1.iType = 0
    'ilRet = btrGetGreaterOrEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
    'Do While (ilRet = BTRV_ERR_NONE) And (tlLcf.iVefCode = ilVefCode)
    For ilGhf = 0 To UBound(llGhfCode) - 1 Step 1
        'ilUpper = UBound(tlGsfInfo)
        'tlGsfInfo(ilUpper).iGameNo = tlLcf.iType
        'gUnpackDateLong tlLcf.iLogDate(0), tlLcf.iLogDate(1), tlGsfInfo(ilUpper).lGameDate
        'gUnpackDateForSort tlLcf.iLogDate(0), tlLcf.iLogDate(1), slSortDate
        'gUnpackTimeLong tlLcf.iTime(0, LBound(tlLcf.lLvfCode)), tlLcf.iTime(1, LBound(tlLcf.lLvfCode)), False, llSortTime
        'slSortTime = Trim$(Str$(llSortTime))
        'Do While Len(slSortTime) < 6
        '    slSortTime = "0" & slSortTime
        'Loop
        'slSortGameNo = Trim$(Str$(tlLcf.iType))
        ''6/30/12:  Allow 5 digit event #'s
        'Do While Len(slSortGameNo) < 5  '3
        '    slSortGameNo = "0" & slSortGameNo
        'Loop
        'tlGsfInfo(ilUpper).sKey = slSortDate & slSortTime & slSortGameNo
        'Get Names
        'tlGsfSrchKey1.lGhfCode = tlGhf.lCode
        'tlGsfSrchKey1.iGameNo = tlGsfInfo(ilUpper).iGameNo
        'ilRet = btrGetEqual(hlGsf, tlGsf, ilGsfRecLen, tlGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        'If ilRet = BTRV_ERR_NONE Then
        tlGsfSrchKey1.lghfcode = llGhfCode(ilGhf)
        tlGsfSrchKey1.iGameNo = 0
        ilRet = btrGetGreaterOrEqual(hlGsf, tlGsf, ilGsfRecLen, tlGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation 'VBC NR
        Do While (ilRet = BTRV_ERR_NONE) And (tlGsf.lghfcode = llGhfCode(ilGhf))
            ilUpper = UBound(tlGsfInfo)
            tlGsfInfo(ilUpper).iGameNo = tlGsf.iGameNo
            gUnpackDateLong tlGsf.iAirDate(0), tlGsf.iAirDate(1), tlGsfInfo(ilUpper).lGameDate
            gUnpackDateForSort tlGsf.iAirDate(0), tlGsf.iAirDate(1), slSortDate
            gUnpackTimeLong tlGsf.iAirTime(0), tlGsf.iAirTime(1), False, llSortTime
            slSortTime = Trim$(str$(llSortTime))
            Do While Len(slSortTime) < 6
                slSortTime = "0" & slSortTime
            Loop
            slSortGameNo = Trim$(str$(tlGsf.iGameNo))
            '6/30/12:  Allow 5 digit event #'s
            Do While Len(slSortGameNo) < 5  '3
                slSortGameNo = "0" & slSortGameNo
            Loop
            tlGsfInfo(ilUpper).sKey = slSortDate & slSortTime & slSortGameNo
            tlGsfInfo(ilUpper).sVisitName = ""
            tlGsfInfo(ilUpper).sHomeName = ""
            'For ilTeam = 1 To UBound(tlTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
            For ilTeam = LBound(tlTeam) To UBound(tlTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                If tlGsf.iVisitMnfCode = tlTeam(ilTeam).iCode Then
                    tlGsfInfo(ilUpper).sVisitName = tlTeam(ilTeam).sName
                    If Trim$(tlTeam(ilTeam).sUnitType) <> "" Then
                        tlGsfInfo(ilUpper).sVisitName = tlTeam(ilTeam).sUnitType
                    End If
                    Exit For
                End If
            Next ilTeam
            For ilTeam = LBound(tlTeam) To UBound(tlTeam) - 1 Step 1 'Traffic!lbcAgency.ListCount - 1 Step 1
                If tlGsf.iHomeMnfCode = tlTeam(ilTeam).iCode Then
                    tlGsfInfo(ilUpper).sHomeName = tlTeam(ilTeam).sName
                    If Trim$(tlTeam(ilTeam).sUnitType) <> "" Then
                        tlGsfInfo(ilUpper).sHomeName = tlTeam(ilTeam).sUnitType
                    End If
                    Exit For
                End If
            Next ilTeam
            tlGsfInfo(ilUpper).sGameStatus = tlGsf.sGameStatus
            tlGsfInfo(ilUpper).sLiveLogMerge = tlGsf.sLiveLogMerge
            tlLcfSrchKey2.iVefCode = ilVefCode
            tlLcfSrchKey2.iLogDate(0) = tlGsf.iAirDate(0)
            tlLcfSrchKey2.iLogDate(1) = tlGsf.iAirDate(1)
            ilRet = btrGetEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
            If ilRet = BTRV_ERR_NONE Then
                tlGsfInfo(ilUpper).sAffPost = tlLcf.sAffPost
            Else
                tlGsfInfo(ilUpper).sAffPost = "N"
            End If
            tlGsfInfo(ilUpper).lGsfCode = tlGsf.lCode
            ReDim Preserve tlGsfInfo(0 To ilUpper + 1) As GSFINFO
            ilRet = btrGetNext(hlGsf, tlGsf, ilGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Next ilGhf
    If UBound(tlGsfInfo) - 1 > 0 Then
        ArraySortTyp fnAV(tlGsfInfo(), 0), UBound(tlGsfInfo), 0, LenB(tlGsfInfo(0)), 0, LenB(tlGsfInfo(0).sKey), 0
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainVefIgnoreHub             *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tlMVef                 *
'*                                                     *
'*******************************************************
Public Sub gObtainVefIgnoreHub(tlMVef() As VEF)
'
'   mObtainVef
'   Where:
'       tlMVef() (I)- VEF record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim hmVef As Integer        'Mnf handle
    Dim imVefRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim illoop As Integer

    On Error GoTo 0

    If (Asc(tgSpf.sUsingFeatures3) And USINGHUB) <> USINGHUB Then
        'ReDim tlMVef(1 To UBound(tgMVef)) As VEF
        ReDim tlMVef(0 To UBound(tgMVef)) As VEF
        For illoop = LBound(tgMVef) To UBound(tgMVef) Step 1
            tlMVef(illoop) = tgMVef(illoop)
        Next illoop
        Exit Sub
    End If
    'ReDim tlMVef(1 To 1) As VEF
    ReDim tlMVef(0 To 0) As VEF
    ilUpperBound = UBound(tlMVef)
    ilExtLen = Len(tlMVef(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hmVef) 'Obtain number of records
    btrExtClear hmVef   'Clear any previous extend operation
    ilRet = btrGetFirst(hmVef, tlMVef(ilUpperBound), imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        Exit Sub
    Else
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
    End If
    Call btrExtSetBounds(hmVef, llNoRec, -1, "UC", "VEF", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hmVef, ilOffSet, ilExtLen)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    'ilRet = btrExtGetNextExt(hmVef)    'Extract record
    ilUpperBound = UBound(tlMVef)
    ilRet = btrExtGetNext(hmVef, tlMVef(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            Exit Sub
        End If
        ilUpperBound = UBound(tlMVef)
        ilExtLen = Len(tlMVef(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hmVef, tgCompMnf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmVef, tlMVef(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Or (tgUrf(0).iMnfHubCode <= 0) Or ((tgUrf(0).iMnfHubCode = tlMVef(ilUpperBound).iMnfHubCode) And ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB)) Or ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) <> USINGHUB) Then
                ilUpperBound = ilUpperBound + 1
                'ReDim Preserve tlMVef(1 To ilUpperBound) As VEF
                ReDim Preserve tlMVef(0 To ilUpperBound) As VEF
            End If
            ilRet = btrExtGetNext(hmVef, tlMVef(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmVef, tlMVef(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    'Sort by vefCode
    'If UBound(tlMVef) > 2 Then
    If UBound(tlMVef) > 1 Then
        'ArraySortTyp fnAV(tlMVef(), 1), UBound(tlMVef) - 1, 0, LenB(tlMVef(1)), 0, -1, 0
        ArraySortTyp fnAV(tlMVef(), 0), UBound(tlMVef), 0, LenB(tlMVef(0)), 0, -1, 0
    End If
    Exit Sub
End Sub


Public Function gCntrOkForUser(hlVsf As Integer, ilInSlfCode As Integer, llInChfVefCode As Long, ilInChfSlfCode() As Integer, Optional slCntrSource As String = "") As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slDate                        slStartDate                   llTodayDate               *
'*                                                                                        *
'******************************************************************************************
    Dim ilFound As Integer
    Dim illoop As Integer
    'Dim tlVsf As VSF
    Dim ilVsfReclen As Integer     'Record length
    Dim tlSrchKey As LONGKEY0
    Dim ilUser As Integer
    Dim ilRet As Integer
    Dim ilSlf As Integer
    Dim llLkVsfCode As Long
    Dim ilSlfCode As Integer
    Dim ilAllVehicles As Integer

    ilSlfCode = ilInSlfCode
    ilVsfReclen = Len(tmVsf) 'btrRecordLength(hlSlf)  'Get and save record length
    If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Then
        gCntrOkForUser = True
        Exit Function
    Else
        '2/17/18: Restrict who can see Programmatic Buys
        If (slCntrSource = "P") And (tgUrf(0).sPrgmmaticAlert <> "I") And (tgUrf(0).sPrgmmaticAlert <> "V") Then
            gCntrOkForUser = False
            Exit Function
        End If
        ilAllVehicles = False
        For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
            If (tgUrf(illoop).iVefCode = 0) And (tgUrf(illoop).sDelete <> "Y") Then
                ilAllVehicles = True
                Exit For
            End If
        Next illoop
        If Not ilAllVehicles Then
            If llInChfVefCode > 0 Then
                For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                    If (tgUrf(illoop).iVefCode = llInChfVefCode) Then
                        ilFound = True
                        Exit For
                    End If
                Next illoop
            ElseIf llInChfVefCode < 0 Then
                llLkVsfCode = -llInChfVefCode
                Do While llLkVsfCode > 0
                    tlSrchKey.lCode = llLkVsfCode
                    ilRet = btrGetEqual(hlVsf, tmVsf, ilVsfReclen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        For illoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                            If tmVsf.iFSCode(illoop) > 0 Then
                                For ilUser = LBound(tgUrf) To UBound(tgUrf) Step 1
                                    If (tgUrf(ilUser).iVefCode = tmVsf.iFSCode(illoop)) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilUser
                                If ilFound Then
                                    Exit For
                                End If
                            End If
                        Next illoop
                    End If
                    If ilFound Then
                        Exit Do
                    End If
                    llLkVsfCode = tmVsf.lLkVsfCode
                Loop
            End If
            If ilFound Then
                gCntrOkForUser = True
                Exit Function
            End If
        End If
    End If
    'Salesperson on Order
    If tgUrf(0).iSlfCode > 0 Then
        For ilSlf = LBound(ilInChfSlfCode) To UBound(ilInChfSlfCode) Step 1
            If ilInChfSlfCode(ilSlf) <> 0 Then
                If tgUrf(0).iSlfCode = ilInChfSlfCode(ilSlf) Then
                    gCntrOkForUser = True
                    Exit Function
                End If
            End If
        Next ilSlf
    End If
    If Not ilAllVehicles Then
        gCntrOkForUser = False
        Exit Function
    End If
    'Within same Group
    If (ilSlfCode > 0) And (tgUrf(0).iGroupNo > 0) Then
        For ilSlf = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
            If tgMSlf(ilSlf).iCode = ilSlfCode Then
                If StrComp(tgMSlf(ilSlf).sJobTitle, "S", 1) <> 0 Then
                    ilSlfCode = 0
                End If
                Exit For
            End If
        Next ilSlf
    End If
    ilFound = False
    If ilSlfCode > 0 Then
        For ilSlf = LBound(ilInChfSlfCode) To UBound(ilInChfSlfCode) Step 1
            If ilInChfSlfCode(ilSlf) <> 0 Then
                If ilSlfCode = ilInChfSlfCode(ilSlf) Then
                    gCntrOkForUser = True
                    Exit Function
                End If
            End If
        Next ilSlf
    Else
        If tgUrf(0).iGroupNo > 0 Then
            For ilSlf = LBound(ilInChfSlfCode) To UBound(ilInChfSlfCode) Step 1
                If ilInChfSlfCode(ilSlf) <> 0 Then
                    For illoop = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
                        If tgPopUrf(illoop).iSlfCode = ilInChfSlfCode(ilSlf) Then
                            If tgPopUrf(illoop).iGroupNo = tgUrf(0).iGroupNo Then
                                gCntrOkForUser = True
                                Exit Function
                            End If
                            Exit For
                        End If
                    Next illoop
                End If
            Next ilSlf
        Else
            gCntrOkForUser = True
            Exit Function
        End If
    End If
    gCntrOkForUser = False
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainItf                      *
'*                                                     *
'*             Created:6/03/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Inventory Type           *
'*                                                     *
'*******************************************************
Function gObtainItf(tlITF() As ITF, slITFTag) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slName                                                                                *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  gObtainItfErr                                                                         *
'******************************************************************************************

'
'   ilRet = gObtainItf (MainForm, tlSortCode(), slSortCodeTag)
'   Where:
'       tlITF() (I)- Itf record structure to be created
'       ilRet (O)- True = populated; False = error
'

    Dim slItfStamp As String    'Itf date/time stamp
    Dim slStamp As String       'Concat. stamp
    Dim hlItf As Integer        'Itf handle
    Dim ilItfRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Itf
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim llLen As Long
    Dim ilLowLimit As Integer

    llLen = 0
    slItfStamp = gFileDateTime(sgDBPath & "Itf.btr")
    slStamp = slItfStamp

    'On Error GoTo gObtainItfErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlITF)
    'If ilRet <> 0 Then
    '    slITFTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlITF).Ptr <> 0 Then
        ilLowLimit = LBound(tlITF)
    Else
        slITFTag = ""
        ilLowLimit = 0
    End If

    If slITFTag <> "" Then
        If StrComp(slStamp, slITFTag, 1) = 0 Then
            'If lbcLocal.ListCount > 0 Then
                gObtainItf = True
                Exit Function
            'End If
            'ilPop = False
        End If
    End If
    gObtainItf = True
    slITFTag = slStamp
    hlItf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlItf, "", sgDBPath & "Itf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    ReDim tlITF(ilLowLimit To ilLowLimit) As ITF   'VB list box clear (list box used to retain code number so record can be found)
    ilItfRecLen = Len(tlITF(ilLowLimit)) 'btrRecordLength(hlItf)  'Get and save record length
    ilExtLen = Len(tlITF(ilLowLimit))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlItf) 'Obtain number of records
    btrExtClear hlItf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlItf, tlITF(UBound(tlITF)), ilItfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlItf)
        btrDestroy hlItf
        Exit Function
    End If
    Call btrExtSetBounds(hlItf, llNoRec, -1, "UC", "Itf", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlItf, ilOffSet, ilExtLen)  'Extract iCode field
    ilRet = btrExtGetNext(hlItf, tlITF(UBound(tlITF)), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        ilExtLen = Len(tlITF(ilLowLimit))  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlItf, tlITF(UBound(tlITF)), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ReDim Preserve tlITF(ilLowLimit To UBound(tlITF) + 1) As ITF
            ilRet = btrExtGetNext(hlItf, tlITF(UBound(tlITF)), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlItf, tlITF(UBound(tlITF)), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlItf)
    btrDestroy hlItf
    If UBound(tlITF) - ilLowLimit > ilLowLimit Then
        ArraySortTyp fnAV(tlITF(), ilLowLimit), UBound(tlITF) - ilLowLimit, 0, LenB(tlITF(ilLowLimit)), 0, -1, 0
    End If
    Exit Function
gObtainItfErr: 'VBC NR
    ilRet = btrClose(hlItf)
    btrDestroy hlItf
    gDbg_HandleError "PopSubs: gObtainItf"
'    gObtainItf = CP_MSG_NOSHOW
'    Exit Function
gObtainItfErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainTrf                      *
'*                                                     *
'*             Created:6/03/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Tax Rate Table           *
'*                                                     *
'*******************************************************
Function gObtainTrf() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slName                                                                                *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  gObtainTrfErr                                                                         *
'******************************************************************************************

'
'   ilRet = gObtainTrf (MainForm, tlSortCode(), slSortCodeTag)
'   Where:
'       tgTrf() (I)- TRF record structure to be created
'       ilRet (O)- True = populated; False = error
'

    Dim slTrfStamp As String    'Trf date/time stamp
    Dim slStamp As String       'Concat. stamp
    Dim hlTrf As Integer        'Trf handle
    Dim ilTrfRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Trf
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim llLen As Long
    Dim ilLowLimit As Integer

    llLen = 0
    slTrfStamp = gFileDateTime(sgDBPath & "Trf.btr")
    slStamp = slTrfStamp

    'On Error GoTo gObtainTrfErr2
    'ilRet = 0
    'ilLowLimit = LBound(tgTrf)
    'If ilRet <> 0 Then
    '    sgTrfTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tgTrf).Ptr <> 0 Then
        ilLowLimit = LBound(tgTrf)
    Else
        sgTrfTag = ""
        ilLowLimit = 0
    End If

    If sgTrfTag <> "" Then
        If StrComp(slStamp, sgTrfTag, 1) = 0 Then
            'If lbcLocal.ListCount > 0 Then
                gObtainTrf = True
                Exit Function
            'End If
            'ilPop = False
        End If
    End If
    gObtainTrf = True
    sgTrfTag = slStamp
    hlTrf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlTrf, "", sgDBPath & "Trf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    ReDim tgTrf(ilLowLimit To ilLowLimit) As TRF   'VB list box clear (list box used to retain code number so record can be found)
    ilTrfRecLen = Len(tgTrf(ilLowLimit)) 'btrRecordLength(hlTrf)  'Get and save record length
    ilExtLen = Len(tgTrf(ilLowLimit))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlTrf) 'Obtain number of records
    btrExtClear hlTrf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlTrf, tgTrf(UBound(tgTrf)), ilTrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlTrf)
        btrDestroy hlTrf
        Exit Function
    End If
    Call btrExtSetBounds(hlTrf, llNoRec, -1, "UC", "Trf", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlTrf, ilOffSet, ilExtLen)  'Extract iCode field
    ilRet = btrExtGetNext(hlTrf, tgTrf(UBound(tgTrf)), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        ilExtLen = Len(tgTrf(ilLowLimit))  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlTrf, tgTrf(UBound(tgTrf)), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If Trim$(tgTrf(UBound(tgTrf)).sGrossNet) = "" Then
                If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
                    tgTrf(UBound(tgTrf)).sGrossNet = "N"
                Else
                    tgTrf(UBound(tgTrf)).sGrossNet = "N"
                End If
            End If
            ReDim Preserve tgTrf(ilLowLimit To UBound(tgTrf) + 1) As TRF
            ilRet = btrExtGetNext(hlTrf, tgTrf(UBound(tgTrf)), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlTrf, tgTrf(UBound(tgTrf)), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlTrf)
    btrDestroy hlTrf
    If UBound(tgTrf) - ilLowLimit > ilLowLimit Then
        ArraySortTyp fnAV(tgTrf(), ilLowLimit), UBound(tgTrf) - ilLowLimit, 0, LenB(tgTrf(ilLowLimit)), 0, -1, 0
    End If
    Exit Function
gObtainTrfErr: 'VBC NR
    ilRet = btrClose(hlTrf)
    btrDestroy hlTrf
    gDbg_HandleError "PopSubs: gObtainTrf"
'    gObtainTrf = CP_MSG_NOSHOW
'    Exit Function
gObtainTrfErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopTaxRateBox                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     advertiser names (if direct,    *
'*                     add \Direct to name)            *
'*                                                     *
'*******************************************************
Function gPopTaxRateBox(ilIncludeRate As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopTaxRateBox (ilIncludeRate, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       ilIncludeRate (I)- Include tax rate
'       lbcLocal (I/O)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'

    Dim slStamp As String    'Adf date/time stamp
    'Dim tlAdf As ADF
    Dim slName As String
    Dim ilRet As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilTrf As Integer
    Dim ilLowLimit As Integer

    ilPop = True
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Trf.Btr") & str$(ilIncludeRate)

    'On Error GoTo gPopTaxRateBoxErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlSortCode)
    'If ilRet <> 0 Then
    '    slSortCodeTag = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopTaxRateBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopTaxRateBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilRet = gObtainTrf()
        If ilRet Then
            For ilTrf = LBound(tgTrf) To UBound(tgTrf) - 1 Step 1
                slName = tgTrf(ilTrf).sTax1Name & "|" & tgTrf(ilTrf).sTax2Name & "\" & Trim$(str$(ilTrf))
                tlSortCode(ilSortCode).sKey = slName
                If ilSortCode >= UBound(tlSortCode) Then
                    ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                End If
                ilSortCode = ilSortCode + 1
            Next ilTrf
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
                ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
            End If
        End If
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If ilRet <> CP_MSG_NONE Then
            gPopTaxRateBox = CP_MSG_PARSE
            Exit Function
        End If
        ilTrf = Val(slCode)
        If ilIncludeRate Then
            slName = ""
            If Trim$(tgTrf(ilTrf).sTax1Name) <> "" Then
                slName = Trim$(tgTrf(ilTrf).sTax1Name)
                If tgTrf(ilTrf).lTax1Rate > 0 Then
                    slName = slName & " " & gLongToStrDec(tgTrf(ilTrf).lTax1Rate, 4)
                End If
            End If
            If Trim$(tgTrf(ilTrf).sTax2Name) <> "" Then
                If slName <> "" Then
                    slName = slName & " " & Trim$(tgTrf(ilTrf).sTax2Name)
                Else
                    slName = Trim$(tgTrf(ilTrf).sTax2Name)
                End If
                If tgTrf(ilTrf).lTax2Rate > 0 Then
                    slName = slName & " " & gLongToStrDec(tgTrf(ilTrf).lTax2Rate, 4)
                End If
            End If
        Else
            slName = Trim$(tgTrf(ilTrf).sTax1Name) & " " & Trim$(tgTrf(ilTrf).sTax2Name)
        End If
        slName = Trim$(slName)
        tlSortCode(illoop).sKey = slName & "\" & Trim$(str$(tgTrf(ilTrf).iCode))
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
        lbcLocal.ItemData(lbcLocal.NewIndex) = tgTrf(ilTrf).iCode
    Next illoop
    Exit Function
gPopTaxRateBoxErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gBinarySearchTrf                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Tax index into tgTrf     *
'*                                                     *
'*******************************************************
Public Function gBinarySearchTrf(ilCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tgTrf)
    ilMax = UBound(tgTrf) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tgTrf(ilMiddle).iCode Then
            'found the match
            gBinarySearchTrf = ilMiddle
            Exit Function
        ElseIf ilCode < tgTrf(ilMiddle).iCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchTrf = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetTrfIndexForAgyAdvt          *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine the TrfCode index     *
'*                                                     *
'*******************************************************
Public Function gGetTrfIndexForAgyAdvt(ilAdfCode As Integer, ilAgfCode As Integer) As Integer
    Dim ilAgf As Integer
    Dim ilAdf As Integer
    Dim ilTrfAgyAdvt As Integer

    gGetTrfIndexForAgyAdvt = -1
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) <> TAXONAIRTIME) And ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) <> TAXONNTR) Then
        Exit Function
    End If
    If ilAgfCode > 0 Then
        ilAgf = gBinarySearchAgf(ilAgfCode)
        If ilAgf = -1 Then
            Exit Function
        End If
        If tgCommAgf(ilAgf).iTrfCode <= 0 Then
            Exit Function
        End If
        ilTrfAgyAdvt = gBinarySearchTrf(tgCommAgf(ilAgf).iTrfCode)
        If ilTrfAgyAdvt = -1 Then
            Exit Function
        End If
    Else
        ilAdf = gBinarySearchAdf(ilAdfCode)
        If ilAdf = -1 Then
            Exit Function
        End If
        If tgCommAdf(ilAdf).iTrfCode <= 0 Then
            Exit Function
        End If
        ilTrfAgyAdvt = gBinarySearchTrf(tgCommAdf(ilAdf).iTrfCode)
        If ilTrfAgyAdvt = -1 Then
            Exit Function
        End If
    End If
    gGetTrfIndexForAgyAdvt = ilTrfAgyAdvt
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetAirTimeTaxValues            *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine the tax values        *
'*                                                     *
'*******************************************************
Public Sub gGetAirTimeTaxValues(ilTrfAgyAdvt As Integer, ilVefCode As Integer, llTax1Rate As Long, llTax2Rate As Long, slGrossNet As String)
    Dim ilVef As Integer
    Dim ilTrfVef As Integer

    llTax1Rate = 0
    llTax2Rate = 0
    slGrossNet = ""
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) <> TAXONAIRTIME) Then
        Exit Sub
    End If
    If ilTrfAgyAdvt < 0 Then
        Exit Sub
    End If
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).iTrfCode > 0 Then
            ilTrfVef = gBinarySearchTrf(tgMVef(ilVef).iTrfCode)
            If ilTrfVef <> -1 Then
                If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
                    'Match State
                    If StrComp(Trim$(tgTrf(ilTrfAgyAdvt).sTax1Name), Trim$(tgTrf(ilTrfVef).sTax1Name), vbTextCompare) = 0 Then
                        llTax1Rate = tgTrf(ilTrfAgyAdvt).lTax1Rate
                        llTax2Rate = tgTrf(ilTrfAgyAdvt).lTax2Rate
                        slGrossNet = tgTrf(ilTrfAgyAdvt).sGrossNet
                    End If
                Else
                    llTax1Rate = tgTrf(ilTrfVef).lTax1Rate
                    llTax2Rate = tgTrf(ilTrfVef).lTax2Rate
                    slGrossNet = tgTrf(ilTrfVef).sGrossNet
                End If
                If Trim$(slGrossNet) = "" Then
                    If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
                        slGrossNet = "N"
                    Else
                        slGrossNet = "N"
                    End If
                End If
            End If
        End If
    End If

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetNTRTaxRates                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine the tax values        *
'*                                                     *
'*******************************************************
Public Sub gGetNTRTaxRates(ilNTRTrfCode As Integer, llTax1Rate As Long, llTax2Rate As Long, slGrossNet As String)
    Dim ilTrf As Integer
    llTax1Rate = 0
    llTax2Rate = 0
    slGrossNet = ""
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) <> TAXONNTR) Then
        Exit Sub
    End If
    If ilNTRTrfCode <= 0 Then
        Exit Sub
    End If
    ilTrf = gBinarySearchTrf(ilNTRTrfCode)
    If ilTrf <> -1 Then
        llTax1Rate = tgTrf(ilTrf).lTax1Rate
        llTax2Rate = tgTrf(ilTrf).lTax2Rate
        slGrossNet = tgTrf(ilTrf).sGrossNet
        If Trim$(slGrossNet) = "" Then
            If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
                slGrossNet = "N"
            Else
                slGrossNet = "N"
            End If
        End If
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gGetAirTimeTaxValues            *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine the tax values        *
'*                                                     *
'*******************************************************
Public Sub gGetAirTimeTaxRates(ilAdfCode As Integer, ilAgfCode As Integer, ilVefCode As Integer, llTax1Rate As Long, llTax2Rate As Long, slGrossNet As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVef                                                                                 *
'******************************************************************************************

    Dim ilTrfVef As Integer
    Dim ilTrfAgyAdvt As Integer

    llTax1Rate = 0
    llTax2Rate = 0
    slGrossNet = ""
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) <> TAXONAIRTIME) Then
        Exit Sub
    End If
    ilTrfAgyAdvt = gGetTrfIndexForAgyAdvt(ilAdfCode, ilAgfCode)
    If ilTrfAgyAdvt < 0 Then
        Exit Sub
    End If
    'ilVef = gBinarySearchVef(ilVefCode)
    'If ilVef <> -1 Then
    '    If tgMVef(ilVef).iTrfCode > 0 Then
    '        ilTrfVef = gBinarySearchTrf(tgMVef(ilVef).iTrfCode)
        ilTrfVef = gGetTrfIndexForVeh(ilVefCode)
            If ilTrfVef <> -1 Then
                If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
                    'Match State
                    If StrComp(Trim$(tgTrf(ilTrfAgyAdvt).sTax1Name), Trim$(tgTrf(ilTrfVef).sTax1Name), vbTextCompare) = 0 Then
                        llTax1Rate = tgTrf(ilTrfAgyAdvt).lTax1Rate
                        llTax2Rate = tgTrf(ilTrfAgyAdvt).lTax2Rate
                        slGrossNet = tgTrf(ilTrfAgyAdvt).sGrossNet
                    End If
                Else
                    llTax1Rate = tgTrf(ilTrfVef).lTax1Rate
                    llTax2Rate = tgTrf(ilTrfVef).lTax2Rate
                    slGrossNet = tgTrf(ilTrfVef).sGrossNet
                End If
                If Trim$(slGrossNet) = "" Then
                    If (Asc(tgSpf.sUsingFeatures4) And TAXBYUSA) = TAXBYUSA Then
                        slGrossNet = "N"
                    Else
                        slGrossNet = "N"
                    End If
                End If
            End If
    '    End If
    'End If

End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:gGetTrfIndexForAgyAdvt          *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine the TrfCode index     *
'*                                                     *
'*******************************************************
Public Function gGetTrfIndexForVeh(ilVefCode As Integer) As Integer
    Dim ilVef As Integer
    Dim ilTrfVef As Integer

    ilTrfVef = -1
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
        ilVef = gBinarySearchVef(ilVefCode)
        If ilVef <> -1 Then
            If tgMVef(ilVef).iTrfCode > 0 Then
                ilTrfVef = gBinarySearchTrf(tgMVef(ilVef).iTrfCode)
            End If
        End If
    End If
    gGetTrfIndexForVeh = ilTrfVef
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainPIFDates_ForVef          *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tlPifDates array       *
'*                                                     *
'*******************************************************
Public Function gObtainPIFDates_ForVef(hlPif As Integer, ilVefCode As Integer, tlPifDates() As PIFDATES) As Integer 'VBC NR
    Dim ilRet As Integer 'VBC NR
    Dim tlPif As PIF 'VBC NR
    Dim tlPifSrchKey1 As PIFKEY1 'VBC NR
    Dim ilPifRecLen As Integer 'VBC NR
    Dim ll12312069 As Long 'VBC NR

    ReDim tlPifDates(0 To 0) As PIFDATES 'VBC NR
    ll12312069 = gDateValue("12/31/2069") 'VBC NR
    ilPifRecLen = Len(tlPif) 'VBC NR
    tlPifSrchKey1.iVefCode = ilVefCode 'VBC NR
    gPackDate "1/1/1970", tlPifSrchKey1.iStartDate(0), tlPifSrchKey1.iStartDate(1) 'VBC NR
    tlPifSrchKey1.iSeqNo = 0 'VBC NR
    ilRet = btrGetGreaterOrEqual(hlPif, tlPif, ilPifRecLen, tlPifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation 'VBC NR
    Do While (ilRet = BTRV_ERR_NONE) And (tlPif.iVefCode = ilVefCode) 'VBC NR
        If tlPif.iSeqNo = 1 Then 'VBC NR
            gUnpackDate tlPif.iStartDate(0), tlPif.iStartDate(1), tlPifDates(UBound(tlPifDates)).sStartDate 'VBC NR
            gUnpackDate tlPif.iEndDate(0), tlPif.iEndDate(1), tlPifDates(UBound(tlPifDates)).sEndDate 'VBC NR
            If gDateValue(tlPifDates(UBound(tlPifDates)).sEndDate) = ll12312069 Then 'VBC NR
                tlPifDates(UBound(tlPifDates)).sEndDate = "TFN" 'VBC NR
            End If 'VBC NR
            ReDim Preserve tlPifDates(0 To UBound(tlPifDates) + 1) As PIFDATES 'VBC NR
        End If 'VBC NR
        ilRet = btrGetNext(hlPif, tlPif, ilPifRecLen, BTRV_LOCK_NONE, SETFORREADONLY) 'VBC NR
    Loop 'VBC NR
    gObtainPIFDates_ForVef = True 'VBC NR
End Function 'VBC NR

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainPIF_ForVefDate           *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tlOutPif array         *
'*                                                     *
'*******************************************************
Public Function gObtainPIF_ForVefDate(hlPif As Integer, ilVefCode As Integer, slDate As String, tlOutPif() As PIF) As Integer
    Dim ilRet As Integer
    Dim tlPif As PIF
    Dim tlPifSrchKey1 As PIFKEY1
    Dim ilPifRecLen As Integer
    Dim slFdDate As String
    Dim llFdDate As Long
    Dim llTestDate As Long

    ReDim tlOutPif(0 To 0) As PIF
    ilPifRecLen = Len(tlPif)
    tlPifSrchKey1.iVefCode = ilVefCode
    gPackDate slDate, tlPifSrchKey1.iStartDate(0), tlPifSrchKey1.iStartDate(1)
    tlPifSrchKey1.iSeqNo = 32000
    ilRet = btrGetLessOrEqual(hlPif, tlPif, ilPifRecLen, tlPifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_NONE) And (tlPif.iVefCode = ilVefCode) Then
        gUnpackDate tlPif.iStartDate(0), tlPif.iStartDate(1), slFdDate
        llFdDate = gDateValue(slFdDate)
        tlPifSrchKey1.iVefCode = ilVefCode
        gPackDate slFdDate, tlPifSrchKey1.iStartDate(0), tlPifSrchKey1.iStartDate(1)
        tlPifSrchKey1.iSeqNo = 0
        ilRet = btrGetGreaterOrEqual(hlPif, tlPif, ilPifRecLen, tlPifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tlPif.iVefCode = ilVefCode)
            gUnpackDateLong tlPif.iStartDate(0), tlPif.iStartDate(1), llTestDate
            If llTestDate <> llFdDate Then
                Exit Do
            End If
            tlOutPif(UBound(tlOutPif)) = tlPif
            ReDim Preserve tlOutPif(0 To UBound(tlOutPif) + 1) As PIF
            ilRet = btrGetNext(hlPif, tlPif, ilPifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If

    gObtainPIF_ForVefDate = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainPIF_ForVef               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tlPif array            *
'*                                                     *
'*******************************************************
Public Function gObtainPIF_ForVef(hlPif As Integer, ilVefCode As Integer, tlPif() As PIF) As Integer
    Dim ilRet As Integer
    Dim tlTPif As PIF
    Dim tlPifSrchKey1 As PIFKEY1
    Dim ilPifRecLen As Integer

    ReDim tlPif(0 To 0) As PIF
    ilPifRecLen = Len(tlPif(0))
    tlPifSrchKey1.iVefCode = ilVefCode
    gPackDate "1/1/1970", tlPifSrchKey1.iStartDate(0), tlPifSrchKey1.iStartDate(1)
    tlPifSrchKey1.iSeqNo = 0
    ilRet = btrGetGreaterOrEqual(hlPif, tlTPif, ilPifRecLen, tlPifSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tlTPif.iVefCode = ilVefCode)
        tlPif(UBound(tlPif)) = tlTPif
        ReDim Preserve tlPif(0 To UBound(tlPif) + 1) As PIF
        ilRet = btrGetNext(hlPif, tlTPif, ilPifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    gObtainPIF_ForVef = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainPIF_ForDate              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tlPif array            *
'*                                                     *
'*******************************************************
Public Function gObtainPIF_ForDate(hlPif As Integer, slDate As String, tlPif() As PIF) As Integer
    Dim ilRet As Integer
    Dim tlTPif As PIF
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilPifRecLen As Integer

    ReDim tlPif(0 To 0) As PIF
    llDate = gDateValue(slDate)
    ilPifRecLen = Len(tlPif(0))
    ilRet = btrGetFirst(hlPif, tlTPif, ilPifRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE)
        gUnpackDateLong tlTPif.iStartDate(0), tlTPif.iStartDate(1), llStartDate
        gUnpackDateLong tlTPif.iEndDate(0), tlTPif.iEndDate(1), llEndDate
        If (llDate >= llStartDate) And (llDate <= llEndDate) Then
            tlPif(UBound(tlPif)) = tlTPif
            ReDim Preserve tlPif(0 To UBound(tlPif) + 1) As PIF
        End If
        ilRet = btrGetNext(hlPif, tlTPif, ilPifRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    gObtainPIF_ForDate = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainPaf                      *
'*                                                     *
'*             Created:6/03/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Program Airing Information *
'*                                                     *
'*******************************************************
Function gObtainPaf(hlPaf As Integer, ilVefCode As Integer, slPafStamp As String, tlPaf() As PAF) As Integer
'
'   ilRet = gObtainPaf (MainForm, tlSortCode(), slSortCodeTag)
'   Where:
'       tlPaf() (I)- PAF record structure to be created
'       ilRet (O)- True = populated; False = error
'

    Dim slStamp As String       'Concat. stamp
    Dim ilPafRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Paf
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim llLen As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim ilLowLimit As Integer

    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Paf.btr") & ilVefCode

    'On Error GoTo gObtainPafErr2
    'ilRet = 0
    'ilLowLimit = LBound(tlPaf)
    'If ilRet <> 0 Then
    '    slPafStamp = ""
    '    ilLowLimit = 0
    'End If
    'On Error GoTo 0
    If PeekArray(tlPaf).Ptr <> 0 Then
        ilLowLimit = LBound(tlPaf)
    Else
        slPafStamp = ""
        ilLowLimit = 0
    End If

    If slPafStamp <> "" Then
        If StrComp(slStamp, slPafStamp, 1) = 0 Then
            'If lbcLocal.ListCount > 0 Then
                gObtainPaf = True
                Exit Function
            'End If
            'ilPop = False
        End If
    End If
    gObtainPaf = True
    slPafStamp = slStamp
    ReDim tlPaf(ilLowLimit To ilLowLimit) As PAF   'VB list box clear (list box used to retain code number so record can be found)
    ilPafRecLen = Len(tlPaf(ilLowLimit)) 'btrRecordLength(hlPaf)  'Get and save record length
    ilExtLen = Len(tlPaf(ilLowLimit))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlPaf) 'Obtain number of records
    btrExtClear hlPaf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlPaf, tlPaf(UBound(tlPaf)), ilPafRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        Exit Function
    End If
    Call btrExtSetBounds(hlPaf, llNoRec, -1, "UC", "Paf", "") 'Set extract limits (all records)
    If ilVefCode > 0 Then
        ilOffSet = gFieldOffset("Paf", "PafVefCode")
        tlIntTypeBuff.iType = ilVefCode
        ilRet = btrExtAddLogicConst(hlPaf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
    End If
    ilOffSet = 0
    ilRet = btrExtAddField(hlPaf, ilOffSet, ilExtLen)  'Extract iCode field
    ilRet = btrExtGetNext(hlPaf, tlPaf(UBound(tlPaf)), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        ilExtLen = Len(tlPaf(ilLowLimit))  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlPaf, tlPaf(UBound(tlPaf)), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ReDim Preserve tlPaf(ilLowLimit To UBound(tlPaf) + 1) As PAF
            ilRet = btrExtGetNext(hlPaf, tlPaf(UBound(tlPaf)), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlPaf, tlPaf(UBound(tlPaf)), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    Exit Function
gObtainPafErr2:
    ilRet = 1
    Resume Next
End Function

'  This list must match affpopsubs.bas in affiliate
Public Sub gInitTaskInfo()
    tgTaskInfo(0).sTaskCode = "CSS"
    tgTaskInfo(0).sTaskName = "Contract Spot Scheduler"
    tgTaskInfo(0).sSortCode = "A"
    tgTaskInfo(0).iMenuIndex = 0
    tgTaskInfo(1).sTaskCode = "SSB"
    tgTaskInfo(1).sTaskName = "Station Spot Builder"
    tgTaskInfo(1).sSortCode = "B"
    tgTaskInfo(1).iMenuIndex = 0
    tgTaskInfo(2).sTaskCode = "AEQ"
    tgTaskInfo(2).sTaskName = "Affiliate Export Queue"
    tgTaskInfo(2).sSortCode = "C"
    tgTaskInfo(2).iMenuIndex = 0
    tgTaskInfo(3).sTaskCode = "ASI"
    tgTaskInfo(3).sTaskName = "Affiliate Spot Import"
    tgTaskInfo(3).sSortCode = "D"
    tgTaskInfo(3).iMenuIndex = 0
    tgTaskInfo(4).sTaskCode = "AMB"
    tgTaskInfo(4).sTaskName = "Affiliate Measurement Builder"
    tgTaskInfo(4).sSortCode = "E"
    tgTaskInfo(4).iMenuIndex = 0
    tgTaskInfo(5).sTaskCode = "ARQ"
    tgTaskInfo(5).sTaskName = "Affiliate Report Queue"
    tgTaskInfo(5).sSortCode = "F"
    tgTaskInfo(5).iMenuIndex = 0
    tgTaskInfo(6).sTaskCode = "ASG"
    tgTaskInfo(6).sTaskName = "Avail Summary Generation"
    tgTaskInfo(6).sSortCode = "G"
    tgTaskInfo(6).iMenuIndex = 0
    tgTaskInfo(7).sTaskCode = "SC"
    tgTaskInfo(7).sTaskName = "Set Credit"
    tgTaskInfo(7).sSortCode = "H"
    tgTaskInfo(7).iMenuIndex = 0
    tgTaskInfo(8).sTaskCode = "CE"
    tgTaskInfo(8).sTaskName = "Corporate Export"
    tgTaskInfo(8).sSortCode = "I"
    tgTaskInfo(8).iMenuIndex = 0
    tgTaskInfo(9).sTaskCode = "SFE"
    tgTaskInfo(9).sTaskName = "Sales Force Export"
    tgTaskInfo(9).sSortCode = "J"
    tgTaskInfo(9).iMenuIndex = 0
    tgTaskInfo(10).sTaskCode = "ME"
    tgTaskInfo(10).sTaskName = "Matrix Export"
    tgTaskInfo(10).sSortCode = "K"
    tgTaskInfo(10).iMenuIndex = 0
    tgTaskInfo(11).sTaskCode = "EPE"
    tgTaskInfo(11).sTaskName = "Efficio Projection Export"
    tgTaskInfo(11).sSortCode = "L"
    tgTaskInfo(11).iMenuIndex = 0
    tgTaskInfo(12).sTaskCode = "ERE"
    tgTaskInfo(12).sTaskName = "Efficio Revenue Export"
    tgTaskInfo(12).sSortCode = "M"
    tgTaskInfo(12).iMenuIndex = 0
    tgTaskInfo(13).sTaskCode = "GPE"
    tgTaskInfo(13).sTaskName = "Get Paid Export"
    tgTaskInfo(13).sSortCode = "N"
    tgTaskInfo(13).iMenuIndex = 0
    tgTaskInfo(14).sTaskCode = "BD"
    tgTaskInfo(14).sTaskName = "Backup Data"
    tgTaskInfo(14).sSortCode = "O"
    tgTaskInfo(14).iMenuIndex = 0
    '7-8-15 Add Tableau export
    tgTaskInfo(15).sTaskCode = "TE"
    tgTaskInfo(15).sTaskName = "Tableau Export"
    tgTaskInfo(15).sSortCode = "P"
    tgTaskInfo(15).iMenuIndex = 0
'    '7967 turn on/off here
    tgTaskInfo(16).sTaskCode = "WVM"
    tgTaskInfo(16).sTaskName = "Web Vendor Manager"
    tgTaskInfo(16).sSortCode = "Q"
    tgTaskInfo(16).iMenuIndex = 0
    'D.S. 03/21/18
    tgTaskInfo(17).sTaskCode = "PB"
    tgTaskInfo(17).sTaskName = "Programmatic Buy"
    tgTaskInfo(17).sSortCode = "R"
    tgTaskInfo(17).iMenuIndex = 0
    
    tgTaskInfo(18).sTaskCode = "CAI"                     '1-31-20 this was missing ; copied from affpopsubs
    tgTaskInfo(18).sTaskName = "Compel Auto Import"
    tgTaskInfo(18).sSortCode = "S"
    tgTaskInfo(18).iMenuIndex = 0

    '1-29-20 Add RAB CRM export
    tgTaskInfo(19).sTaskCode = "RE"
    tgTaskInfo(19).sTaskName = "RAB Export"
    tgTaskInfo(19).sSortCode = "T"
    tgTaskInfo(19).iMenuIndex = 0
    
    'TTP 9992 Add Custom Revenue Export
    tgTaskInfo(20).sTaskCode = "CRE"
    tgTaskInfo(20).sTaskName = "Custom Revenue Export"
    tgTaskInfo(20).sSortCode = "U"
    tgTaskInfo(20).iMenuIndex = 0
End Sub

Public Function gGetStationFromVehicle(ilVefCode As Integer) As Integer
    Dim ilShtt As Integer
    Dim ilRet As Integer
    Dim llLimit As Long
    
    On Error GoTo gIsStationVehicleErr
    gGetStationFromVehicle = -1
    llLimit = UBound(tgVehicleStations)
    
    For ilShtt = LBound(tgVehicleStations) To UBound(tgVehicleStations) - 1 Step 1
        If tgVehicleStations(ilShtt).iShttVefCode = ilVefCode Then
            gGetStationFromVehicle = tgVehicleStations(ilShtt).iCode
            Exit Function
        End If
    Next ilShtt
    Exit Function
gIsStationVehicleErr:
    ilRet = gObtainStations()
    Resume Next
End Function

Public Function gIsImportInvoicedSpots(ilVefCode As Integer) As Boolean
    Dim ilVff As Integer
    Dim ilRet As Integer
    Dim ilLimit As Integer
    
    'On Error GoTo gIsImportInvoicedSpotsErr
    'ilLimit = UBound(tgVff)
    If PeekArray(tgVff).Ptr = 0 Then
        ilRet = gVffRead()
    End If
    
    gIsImportInvoicedSpots = False
    'For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
    '    If ilVefCode = tgVff(ilVff).iVefCode Then
    '        If tgVff(ilVff).sPostLogSource = "S" Then
    '            gIsImportInvoicedSpots = True
    '            Exit For
    '        End If
    '    End If
    'Next ilVff
    ilVff = gBinarySearchVff(ilVefCode)
    If ilVff <> -1 Then
        If tgVff(ilVff).sPostLogSource = "S" Then
            gIsImportInvoicedSpots = True
        End If
    End If
    Exit Function
gIsImportInvoicedSpotsErr:
    ilRet = gVffRead()
    Resume Next
End Function

Public Function gIsOnInsertions(ilVefCode As Integer) As Boolean
    Dim ilVff As Integer
    Dim ilRet As Integer
    Dim ilLimit As Integer
    
    'On Error GoTo gIsOnInsertionsErr
    'ilLimit = UBound(tgVff)
    If PeekArray(tgVff).Ptr = 0 Then
        ilRet = gVffRead()
    End If
    gIsOnInsertions = False
    'For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
    '    If ilVefCode = tgVff(ilVff).iVefCode Then
    '        If tgVff(ilVff).sOnInsertions <> "N" Then
    '            gIsOnInsertions = True
    '            Exit For
    '        End If
    '    End If
    'Next ilVff
    ilVff = gBinarySearchVff(ilVefCode)
    If ilVff <> -1 Then
        If tgVff(ilVff).sOnInsertions <> "N" Then
            gIsOnInsertions = True
        End If
    End If
    Exit Function
gIsOnInsertionsErr:
    ilRet = gVffRead()
    Resume Next
End Function

Public Function gIsInsertionExport(ilVefCode As Integer) As Boolean
    Dim ilVff As Integer
    Dim ilRet As Integer
    Dim ilLimit As Integer
    
    'On Error GoTo gIsInsertionExportErr
    'ilLimit = UBound(tgVff)
    If PeekArray(tgVff).Ptr = 0 Then
        ilRet = gVffRead()
    End If
    gIsInsertionExport = False
    '8032
    igStationXmlChoice = STATIONXMLNONE
    'For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
    '    If ilVefCode = tgVff(ilVff).iVefCode Then
    '        If tgVff(ilVff).sOnInsertions <> "N" Then
    '            gIsInsertionExport = True
    '            Exit For
    '        End If
    '    End If
    'Next ilVff
    ilVff = gBinarySearchVff(ilVefCode)
    '8032 add "M" and igStationXmlChoice
    If ilVff <> -1 Then
'        If tgVff(ilVff).sOnXMLInsertion = "W" Or tgVff(ilVff).sOnXMLInsertion = "M" Then
        If tgVff(ilVff).sOnXMLInsertion <> "N" And Trim$(tgVff(ilVff).sOnXMLInsertion) <> "" Then
            gIsInsertionExport = True
            Select Case tgVff(ilVff).sOnXMLInsertion
                Case "W"
                    igStationXmlChoice = STATIONXMLWIDEORBIT
                Case "M"
                    igStationXmlChoice = STATIONXMLMARKETRON
            End Select
        End If
    End If
    Exit Function
gIsInsertionExportErr:
    ilRet = gVffRead()
    Resume Next
End Function

Private Sub mCreateSort(ilSort As Integer, ilFlag As Integer, ilShow As Integer, tlDnf As DNF, slName As String, ilSortCode As Integer, tlSortCode() As SORTCODE)
    Dim slDate As String
    Dim llDate As Long
    Dim slSortDate As String
    
    If ilSort = 0 Then
        gUnpackDate tlDnf.iBookDate(0), tlDnf.iBookDate(1), slDate
        If ilShow = 0 Then
            slName = Trim$(tlDnf.sBookName)
        Else
            slName = Trim$(tlDnf.sBookName) & ": " & slDate
        End If
        slName = slName & "\" & Trim$(str$(tlDnf.iCode))
    Else
        gUnpackDateLong tlDnf.iBookDate(0), tlDnf.iBookDate(1), llDate
        llDate = 99999 - llDate
        slSortDate = Trim$(str$(llDate))
        Do While Len(slSortDate) < 5
            slSortDate = "0" & slSortDate
        Loop
        gUnpackDate tlDnf.iBookDate(0), tlDnf.iBookDate(1), slDate
        If ilShow = 0 Then
            slName = Trim$(tlDnf.sBookName)
        Else
            slName = Trim$(tlDnf.sBookName) & ": " & slDate
        End If
        slName = slSortDate & "|" & slName & "\" & Trim$(str$(tlDnf.iCode))
    End If
    'If Not gOkAddStrToListBox(slName, llLen, True) Then
    '    Exit Do
    'End If
    'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
    tlSortCode(ilSortCode).sKey = slName
    If (ilFlag = 1) Or (ilFlag = 2) Then
        ReDim Preserve tlSortCode(LBound(tlSortCode) To UBound(tlSortCode) + 1) As SORTCODE
    Else
        If ilSortCode >= UBound(tlSortCode) Then
            ReDim Preserve tlSortCode(LBound(tlSortCode) To UBound(tlSortCode) + 100) As SORTCODE
        End If
    End If
    ilSortCode = ilSortCode + 1
End Sub

Private Sub mAddRecSub1(tlMnf As MNF, slMnfType As String, slName As String, ilSortCode As Integer, tlSortCode() As SORTCODE)
    Dim ilAddRec As Integer
    Dim slStr As String
    Dim slTax As String
    Dim slSlspComm As String
    Dim slAcqCost As String
    
    ilAddRec = True
    slName = Trim$(tlMnf.sName)
    If (slMnfType <> "N") And (slMnfType <> "NNS") And (slMnfType <> "Y") And (slMnfType <> "YW") And (slMnfType <> "F") And (slMnfType <> "FG") And (slMnfType <> "1") And (slMnfType <> "2") Then
        slName = slName & "\" & Trim$(str$(tlMnf.iCode))
    End If
    Select Case slMnfType
        Case "I", "INH"
            gPDNToStr tlMnf.sRPU, 2, slStr
            gPDNToStr tlMnf.sSSComm, 4, slSlspComm
            If tlMnf.iGroupNo = 1 Then
                slTax = "Y"
            Else
                slTax = "N"
            End If
            slAcqCost = ""
            '6/7/15: replaced acquisition from site override with Barter in system options
            If (Asc(tgSpf.sOverrideOptions) And SPNTRACQUISITION) = SPNTRACQUISITION Then
                slAcqCost = gLongToStrDec(tlMnf.lCost, 2)
            End If
            If Trim$(tlMnf.sCodeStn) = "Y" Then
                If slMnfType = "INH" Then
                    ilAddRec = False
                End If
                slName = Trim$(tlMnf.sName) & " (Hard Cost)" & "\" & Trim$(str$(tlMnf.iCode))
            End If
            slName = slName & "\" & slStr & "\" & tlMnf.sUnitsPer & "\" & slSlspComm & "\" & slTax & "\" & slAcqCost
        Case "A"
            gPDNToStr tlMnf.sRPU, 2, slStr
            slName = slName & "\" & slStr & "\" & Trim$(str$(tlMnf.iGroupNo))
        Case "C"
            slName = slName & "\" & tlMnf.sUnitType
        Case "S"
            gPDNToStr tlMnf.sSSComm, 4, slStr
            slName = slName & "\" & slStr & "\" & Trim$(str$(tlMnf.iGroupNo))
        Case "N"    'Feed
            Select Case Trim$(tlMnf.sUnitType)
                Case "D"    'Dish
                    slName = slName & "/Dish"
                Case "A"    'Antenna
                    slName = slName & "/Antenna"
                Case "C"    'CD
                    slName = slName & "/CD"
            End Select
            slName = slName & "\" & Trim$(str$(tlMnf.iCode))
        Case "NNS"    'Feed without any subfeeds
            Select Case Trim$(tlMnf.sUnitType)
                Case "D"    'Dish
                    slName = slName & "/Dish"
                Case "A"    'Antenna
                    slName = slName & "/Antenna"
                Case "C"    'CD
                    slName = slName & "/CD"
                Case "S"
                    ilAddRec = False
            End Select
            slName = slName & "\" & Trim$(str$(tlMnf.iCode))
        Case "NOS"  'Subfeed only
            If Trim$(tlMnf.sUnitType) <> "S" Then  '0=No subfeeds
                ilAddRec = False
            End If
        Case "Y"
            slName = Trim$(tlMnf.sUnitType) & " " & Trim(tlMnf.sName) & "\" & Trim$(str$(tlMnf.iCode))
        Case "YW"
            If Left$(Trim$(tlMnf.sUnitType), 1) = "W" Then
                slName = Trim$(tlMnf.sUnitType) & " " & Trim(tlMnf.sName) & "\" & Trim$(str$(tlMnf.iCode))
            Else
                ilAddRec = False
            End If
        Case "D"
            '5/11/10:  Removed sort by group number
            ''5/7/10:  Group number changed from 2 to 3
            If tlMnf.iGroupNo = 0 Then
            '    'slStr = "00"
            '    slStr = "000"
                slStr = "0"
            Else
            '    slStr = Trim$(Str$(tlMnf.iGroupNo))
            '    'Do While Len(slStr) < 2
            '    Do While Len(slStr) < 3
            '        slStr = "0" & slStr
            '    Loop
                slStr = "1"
            End If
            'The Group code is only used to get items sorted in group number order
            slName = slStr & slName
        Case "DS"
            If tlMnf.iGroupNo <> 0 Then
                ilAddRec = False
            End If
        Case "DC"
            If tlMnf.iGroupNo = 0 Then
                ilAddRec = False
            End If
        Case "DP"
            If tlMnf.iGroupNo <> 0 Then
                ilAddRec = False
            Else
                'Eliminate basic demos (12-17;18-20;21-24;18-24;25-34;35-44;45-49;50-54;55-64;65+)
                If (Left(tlMnf.sName, 1) = "M") Or (Left(tlMnf.sName, 1) = "W") Then
                    If InStr(1, tlMnf.sName, "12-17", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "18-20", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "21-24", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "18-24", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "25-34", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "35-44", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "45-49", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "50-54", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "55-64", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "65+", vbTextCompare) > 0 Then
                        ilAddRec = False
                    End If
                End If
            End If
        Case "F", "FNG"
            slName = Trim$(tlMnf.sUnitType) & " " & Trim(tlMnf.sName) & "\" & Trim$(str$(tlMnf.iCode))
        Case "FG"
            slName = Trim$(tlMnf.sUnitType) & "\" & Trim$(str$(tlMnf.iCode))
        Case "1"
            If tlMnf.iGroupNo = 1 Then
                slName = Trim(tlMnf.sName) & "\" & Trim$(str$(tlMnf.iCode))
            Else
                ilAddRec = False
            End If
        Case "2"
            If tlMnf.iGroupNo = 2 Then
                slName = Trim(tlMnf.sName) & "\" & Trim$(str$(tlMnf.iCode))
            Else
                ilAddRec = False
            End If
    End Select
    If Left$(slMnfType, 1) = "H" Then
        If (slMnfType <> "H") And (slMnfType <> "H0") Then
            slStr = Mid$(slMnfType, 2)
            If slStr <> Trim$(tlMnf.sUnitType) Then
                ilAddRec = False
            End If
        Else
            slName = Trim$(tlMnf.sName)
            Select Case Trim$(tlMnf.sUnitType)
                Case "1"
                    slName = slName & "/Participant"
                Case "2"
                    slName = slName & "/Subtotal"
                Case "3"
                    slName = slName & "/Market"
                Case "4"
                    slName = slName & "/Format"
                Case "5"
                    slName = slName & "/Research"
                Case "6"
                    slName = slName & "/Sub-Company"
            End Select
            slName = slName & "\" & Trim$(str$(tlMnf.iCode))
        End If
    End If
    If ilAddRec Then
        'If Not gOkAddStrToListBox(slName, llLen, True) Then
        '    Exit Do
        'End If
        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
        tlSortCode(ilSortCode).sKey = slName
        If ilSortCode >= UBound(tlSortCode) Then
            ReDim Preserve tlSortCode(LBound(tlSortCode) To UBound(tlSortCode) + 100) As SORTCODE
        End If
        ilSortCode = ilSortCode + 1
    End If
End Sub

Private Sub mAddRecSub2(tlMnf As MNF, slMnfType As String, slName As String, ilSortCode As Integer, tlSortCode() As SORTCODE)
    Dim ilAddRec As Integer
    Dim slStr As String
    Dim slTax As String
    Dim slSlspComm As String
    Dim slAcqCost As String
    
    ilAddRec = True
    slName = Trim$(tlMnf.sName)
    If (slMnfType <> "N") And (slMnfType <> "NNS") And (slMnfType <> "Y") And (slMnfType <> "YW") And (slMnfType <> "F") And (slMnfType <> "FG") And (slMnfType <> "1") And (slMnfType <> "2") Then
        slName = slName & "\" & Trim$(str$(tlMnf.iCode))
    End If
    Select Case slMnfType
        Case "I", "INH"
            gPDNToStr tlMnf.sRPU, 2, slStr
            gPDNToStr tlMnf.sSSComm, 4, slSlspComm
            If tlMnf.iGroupNo = 1 Then
                slTax = "Y"
            Else
                slTax = "N"
            End If
            slAcqCost = ""
            '6/7/15: replaced acquisition from site override with Barter in system options
            If (Asc(tgSpf.sOverrideOptions) And SPNTRACQUISITION) = SPNTRACQUISITION Then
                slAcqCost = gLongToStrDec(tlMnf.lCost, 2)
            End If
            If Trim$(tlMnf.sCodeStn) = "Y" Then
                If slMnfType = "INH" Then
                    ilAddRec = False
                End If
                slName = Trim$(tlMnf.sName) & " (Hard Cost)" & "\" & Trim$(str$(tlMnf.iCode))
            End If
            slName = slName & "\" & slStr & "\" & tlMnf.sUnitsPer & "\" & slSlspComm & "\" & slTax & "\" & slAcqCost
        Case "A"
            gPDNToStr tlMnf.sRPU, 2, slStr
            slName = slName & "\" & slStr & "\" & Trim$(str$(tlMnf.iGroupNo))
        Case "C"
            slName = slName & "\" & tlMnf.sUnitType
        Case "S"
            gPDNToStr tlMnf.sSSComm, 4, slStr
            slName = slName & "\" & slStr & "\" & Trim$(str$(tlMnf.iGroupNo))
        Case "N"    'Feed
            Select Case Trim$(tlMnf.sUnitType)
                Case "D"    'Dish
                    slName = slName & "/Dish"
                Case "A"    'Antenna
                    slName = slName & "/Antenna"
                Case "C"    'CD
                    slName = slName & "/CD"
            End Select
            slName = slName & "\" & Trim$(str$(tlMnf.iCode))
        Case "NNS"    'Feed without any subfeeds
            Select Case Trim$(tlMnf.sUnitType)
                Case "D"    'Dish
                    slName = slName & "/Dish"
                Case "A"    'Antenna
                    slName = slName & "/Antenna"
                Case "C"    'CD
                    slName = slName & "/CD"
                Case "S"
                    ilAddRec = False
            End Select
            slName = slName & "\" & Trim$(str$(tlMnf.iCode))
        Case "NOS"  'Subfeed only
            If Trim$(tlMnf.sUnitType) <> "S" Then  '0=No subfeeds
                ilAddRec = False
            End If
        Case "Y"
            slName = Trim$(tlMnf.sUnitType) & " " & Trim(tlMnf.sName) & "\" & Trim$(str$(tlMnf.iCode))
        Case "YW"
            If Left$(Trim$(tlMnf.sUnitType), 1) = "W" Then
                slName = Trim$(tlMnf.sUnitType) & " " & Trim(tlMnf.sName) & "\" & Trim$(str$(tlMnf.iCode))
            Else
                ilAddRec = False
            End If
        Case "D"
            '5/11/10:  Removed group number from sort
            ''5/7/10:  Group number changed from 2 to 3
            If tlMnf.iGroupNo = 0 Then
            '    'slStr = "00"
            '    slStr = "000"
                slStr = "0"
            Else
            '    slStr = Trim$(Str$(tlMnf.iGroupNo))
            '    'Do While Len(slStr) < 2
            '    Do While Len(slStr) < 3
            '        slStr = "0" & slStr
            '    Loop
                slStr = "1"
            End If
            'The Group code is only used to get items sorted in group number order
            slName = slStr & slName
        Case "DS"
            If tlMnf.iGroupNo <> 0 Then
                ilAddRec = False
            End If
        Case "DC"
            If tlMnf.iGroupNo = 0 Then
                ilAddRec = False
            End If
        Case "DP"
            If tlMnf.iGroupNo <> 0 Then
                ilAddRec = False
            Else
                'Eliminate basic demos (12-17;18-20;21-24;18-24;25-34;35-44;45-49;50-54;55-64;65+)
                If (Left(tlMnf.sName, 1) = "M") Or (Left(tlMnf.sName, 1) = "W") Then
                    If InStr(1, tlMnf.sName, "12-17", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "18-20", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "21-24", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "18-24", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "25-34", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "35-44", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "45-49", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "50-54", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "55-64", vbTextCompare) > 0 Then
                        ilAddRec = False
                    ElseIf InStr(1, tlMnf.sName, "65+", vbTextCompare) > 0 Then
                        ilAddRec = False
                    End If
                End If
            End If
        Case "F", "FNG"
            slName = Trim$(tlMnf.sUnitType) & " " & Trim(tlMnf.sName) & "\" & Trim$(str$(tlMnf.iCode))
        Case "FG"
            slName = Trim$(tlMnf.sUnitType) & "\" & Trim$(str$(tlMnf.iCode))
        Case "1"
            If tlMnf.iGroupNo = 1 Then
                slName = Trim(tlMnf.sName) & "\" & Trim$(str$(tlMnf.iCode))
            Else
                ilAddRec = False
            End If
        Case "2"
            If tlMnf.iGroupNo = 2 Then
                slName = Trim(tlMnf.sName) & "\" & Trim$(str$(tlMnf.iCode))
            Else
                ilAddRec = False
            End If
    End Select
    If Left$(slMnfType, 1) = "H" Then
        If (slMnfType <> "H") And (slMnfType <> "H0") Then
            slStr = Mid$(slMnfType, 2)
            If slStr <> Trim$(tlMnf.sUnitType) Then
                ilAddRec = False
            End If
        Else
            slName = Trim$(tlMnf.sName)
            Select Case Trim$(tlMnf.sUnitType)
                Case "1"
                    slName = slName & "/Participant"
                Case "2"
                    slName = slName & "/Subtotal"
                Case "3"
                    slName = slName & "/Market"
                Case "4"
                    slName = slName & "/Format"
                Case "5"
                    slName = slName & "/Research"
                Case "6"
                    slName = slName & "/Sub-Company"
            End Select
            slName = slName & "\" & Trim$(str$(tlMnf.iCode))
        End If
    End If
    If ilAddRec Then
        'If Not gOkAddStrToListBox(slName, llLen, True) Then
        '    Exit Do
        'End If
        'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
        tlSortCode(ilSortCode).sKey = slName
        If ilSortCode >= UBound(tlSortCode) Then
            ReDim Preserve tlSortCode(LBound(tlSortCode) To UBound(tlSortCode) + 100) As SORTCODE
        End If
        ilSortCode = ilSortCode + 1
    End If
End Sub

Public Function gFileChgd(slFileName As String) As Boolean
    Dim slSQLQuery As String
    Dim ilIndex As Integer
    
    ilIndex = -1
    gFileChgd = True
    slSQLQuery = "Select * From fct_File_Chg_Table Where fctFileName = '" & slFileName & "'"
    Set fct_rst = gSQLSelectCall(slSQLQuery)
    If Not fct_rst.EOF Then
        Select Case UCase(slFileName)
            Case "SHTT.MKD"
                ilIndex = SHTTINDEX
            Case "VEF.BTR"
                ilIndex = VEFINDEX
            Case "VPF.BTR"
                ilIndex = VPFINDEX
            Case "CPTT.MKD"
                ilIndex = CPTTINDEX
            Case "ADF.BTR"
                ilIndex = ADFINDEX
        End Select
        If ilIndex <> -1 Then
            If (tgFctChgdInfo(ilIndex).lLastDateChgd <> fct_rst!fctDateChgd) Or (tgFctChgdInfo(ilIndex).lLastTimeChgd <> fct_rst!fctTimeChgd) Then
                gFileChgdSetInternal slFileName
            Else
                gFileChgd = False
            End If
        End If
    Else
        gFileChgdUpdate slFileName, True
    End If
End Function

Public Sub gFileChgdSetInternal(slFileName As String)
    Dim slSQLQuery As String
    Dim ilIndex As Integer
    
    ilIndex = -1
    slSQLQuery = "Select * From fct_File_Chg_Table Where  fctFileName = '" & slFileName & "'"
    Set fct_rst = gSQLSelectCall(slSQLQuery)
    If Not fct_rst.EOF Then
        Select Case UCase(slFileName)
            Case "SHTT.MKD"
                ilIndex = SHTTINDEX
            Case "VEF.BTR"
                ilIndex = VEFINDEX
            Case "VPF.BTR"
                ilIndex = VPFINDEX
            Case "CPTT.MKD"
                ilIndex = CPTTINDEX
            Case "ADF.BTR"
                ilIndex = ADFINDEX
        End Select
        If ilIndex <> -1 Then
            tgFctChgdInfo(ilIndex).lLastDateChgd = fct_rst!fctDateChgd
            tgFctChgdInfo(ilIndex).lLastTimeChgd = fct_rst!fctTimeChgd
        End If
    End If
End Sub

Public Sub gFileChgdUpdate(slFileName As String, blRepopRequired As Boolean)
    Dim slNow As String
    Dim llRet As Long
    Dim slSQLQuery As String
    
    slNow = Now
    slSQLQuery = "Select * From fct_File_Chg_Table Where  fctFileName = '" & slFileName & "'"
    Set fct_rst = gSQLSelectCall(slSQLQuery)
    If Not fct_rst.EOF Then
        slSQLQuery = "UPDATE fct_File_Chg_Table SET "
        slSQLQuery = slSQLQuery & "fctDate = '" & Format$(slNow, sgSQLDateForm) & "', "
        slSQLQuery = slSQLQuery & "fctDateChgd = " & gDateValue(slNow) & ", "
        slSQLQuery = slSQLQuery & "fctTime = '" & Format$(slNow, sgSQLTimeForm) & "', "
        slSQLQuery = slSQLQuery & "fctTimeChgd = " & gTimeToLong(slNow, False) & ", "
        slSQLQuery = slSQLQuery & "fctSystemType = '" & "T" & "', "
        slSQLQuery = slSQLQuery & "fctUrfCode = " & tgUrf(0).iCode
        slSQLQuery = slSQLQuery & " WHERE (fctFileName = '" & slFileName & "')"
        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    Else
        slSQLQuery = "Insert Into FCT_File_Chg_Table ( "
        slSQLQuery = slSQLQuery & "fctCode, "
        slSQLQuery = slSQLQuery & "fctFileName, "
        slSQLQuery = slSQLQuery & "fctDate, "
        slSQLQuery = slSQLQuery & "fctDateChgd, "
        slSQLQuery = slSQLQuery & "fctTime, "
        slSQLQuery = slSQLQuery & "fctTimeChgd, "
        slSQLQuery = slSQLQuery & "fctSystemType, "
        slSQLQuery = slSQLQuery & "fctUrfCode, "
        slSQLQuery = slSQLQuery & "fctUstCode, "
        slSQLQuery = slSQLQuery & "fctUnused "
        slSQLQuery = slSQLQuery & ") "
        slSQLQuery = slSQLQuery & "Values ( "
        slSQLQuery = slSQLQuery & 0 & ", "
        slSQLQuery = slSQLQuery & "'" & gFixQuote(slFileName) & "', "
        slSQLQuery = slSQLQuery & "'" & Format$(slNow, sgSQLDateForm) & "', "
        slSQLQuery = slSQLQuery & gDateValue(slNow) & ", "
        slSQLQuery = slSQLQuery & "'" & Format$(slNow, sgSQLTimeForm) & "', "
        slSQLQuery = slSQLQuery & gTimeToLong(slNow, False) & ", "
        slSQLQuery = slSQLQuery & "'" & "T" & "', "
        slSQLQuery = slSQLQuery & tgUrf(0).iCode & ", "
        slSQLQuery = slSQLQuery & 0 & ", "
        slSQLQuery = slSQLQuery & "'" & "" & "' "
        slSQLQuery = slSQLQuery & ") "
        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    End If
    If llRet = 0 And Not blRepopRequired Then
        gFileChgdSetInternal slFileName
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gMoveCntr                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Move Contract (Chf), Lines (Clf),*
'*                     Flight (Cff) into tgChfCntr,     *
'*                     tgClfCntr and tgCffCntr         *
'*                                                     *
'*******************************************************
Function gMoveCntr(tlInChf As CHF, tlInClf() As CLF, tlInCff() As CFF, Optional blSort As Boolean = True) As Integer

    Dim ilRet As Integer
    Dim llRif As Long
    Dim tlTChf As CHF

    Dim llLnUpper As Long
    Dim llFlUpper As Long
    Dim llClf As Long
    Dim llCff As Long
    Dim slGpSort As String
    Dim slVehName As String
    Dim slVehSort As String
    Dim slDPSort As String
    Dim slLine As String
    Dim slRevNo As String
    Dim slPropVer As String
    Dim ilVef As Integer
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim ilMnfRecLen As Integer
    Dim tlMnfSrchKey As INTKEY0
    Dim hlRdf As Integer
    Dim tlRdf As RDF
    Dim ilRdfRecLen As Integer
    Dim ilDay As Integer
    Dim tlRdfSrchKey As INTKEY0
    Dim slOALineNo As String    'Line number of Parent + A or Line Number of Parent + C or ~~~~ if standard line)
    Dim llEliminateFirst As Long
    Dim llMaxClfCode As Long
    Dim llNext As Long


    ReDim tgClfCntr(0 To 0) As CLFLIST
    ReDim tgCffCntr(0 To 0) As CFFLIST
    ReDim tgLnAvailInfo(0 To 0) As LNAVAILINFO

    If blSort Then
        hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            gMoveCntr = False
            btrDestroy hlMnf
            Exit Function
        End If
        ilMnfRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length
        hlRdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = ilRet
            gMoveCntr = False
            btrDestroy hlMnf
            btrDestroy hlRdf
            Exit Function
        End If
        ilRdfRecLen = Len(tlRdf) 'btrRecordLength(hlMnf)  'Get and save record length
    End If
    ilRet = gObtainVef()
    If Not ilRet Then
        gMoveCntr = False
        If blSort Then
            btrDestroy hlMnf
            btrDestroy hlRdf
        End If
        Exit Function
    End If
    ilRet = gObtainRcfRifRdf()
    If Not ilRet Then
        gMoveCntr = False
        If blSort Then
            btrDestroy hlMnf
            btrDestroy hlRdf
        End If
        Exit Function
    End If
    tgChfCntr = tlInChf
    llLnUpper = UBound(tgClfCntr)
    llFlUpper = UBound(tgCffCntr)
    For llClf = LBound(tlInClf) To UBound(tlInClf) - 1 Step 1
        tgClfCntr(llLnUpper).ClfRec = tlInClf(llClf)
        'Initialize fields
        tgClfCntr(llLnUpper).iStatus = 0    '0=New; 1=Old and retain
        tgClfCntr(llLnUpper).lRecPos = 0
        tgClfCntr(llLnUpper).iFirstCff = -1
        tgClfCntr(llLnUpper).iFirstCgf = -1
        tgClfCntr(llLnUpper).lFirstAvail = -1
        tgClfCntr(llLnUpper).iCancel = False
        tgClfCntr(llLnUpper).iOverride = False
        tgClfCntr(llLnUpper).iGame = False
        tgClfCntr(llLnUpper).iLibBuy = False
        tgClfCntr(llLnUpper).iInPast = False
        tgClfCntr(llLnUpper).lUnbilledDate = 0
        tgClfCntr(llLnUpper).lEndBilledDate = 0
        tgClfCntr(llLnUpper).iChgBilledPrice = False
        tgClfCntr(llLnUpper).iPriDemoBook = 0
        tgClfCntr(llLnUpper).iResearch = False
        tgClfCntr(llLnUpper).iPriceSet = 1
        tgClfCntr(llLnUpper).iSPChg = False
        tgClfCntr(llLnUpper).iOrigAnySpots = 0  '-1
        tgClfCntr(llLnUpper).sCurrentPrice = ""
        tgClfCntr(llLnUpper).iLineSchd = False
        tgClfCntr(llLnUpper).iMonThruToDaysSpots = 0
        tgClfCntr(llLnUpper).bShowLine = True
        tgClfCntr(llLnUpper).bPkExpanded = True 'False
        tgClfCntr(llLnUpper).sLineChgd = "Y"
        
        tgClfCntr(llLnUpper).ClfRec.iNoGames = 0
        tgClfCntr(llLnUpper).ClfRec.iSpotsOrdered = 0
        tgClfCntr(llLnUpper).ClfRec.sSchStatus = " "
        tgClfCntr(llLnUpper).ClfRec.iPriority = -1
        tgClfCntr(llLnUpper).ClfRec.sDelete = "N"
        tgClfCntr(llLnUpper).ClfRec.sPreempt = "P"
        tgClfCntr(llLnUpper).ClfRec.iSpotsWrite = 0
        tgClfCntr(llLnUpper).ClfRec.iSpotsBooked = 0
        tgClfCntr(llLnUpper).ClfRec.iDnfCode = 0
        tgClfCntr(llLnUpper).ClfRec.iMnfDemo = 0
        tgClfCntr(llLnUpper).ClfRec.lCPP = 0
        tgClfCntr(llLnUpper).ClfRec.lCPM = 0
        tgClfCntr(llLnUpper).ClfRec.lGrImp = 0
        tgClfCntr(llLnUpper).ClfRec.sHideCBS = "N"
        tgClfCntr(llLnUpper).ClfRec.iVefCode = 0
        tgClfCntr(llLnUpper).ClfRec.iBBOpenLen = 0
        tgClfCntr(llLnUpper).ClfRec.iBBCloseLen = 0
        tgClfCntr(llLnUpper).ClfRec.sOV2DefinedBits = Chr(&H0)

        
        slOALineNo = "~~~~"
        'Build sort key (Vehicle Group Sort #; Vehicle Sort #; Vehicle Name; Daypart Sort #; Line #
        slVehSort = "999"
        slVehName = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        ilVef = gBinarySearchVef(tgClfCntr(llLnUpper).ClfRec.iVefCode)
        If ilVef <> -1 Then
            slVehName = tgMVef(ilVef).sName
            slVehSort = Trim$(str$(tgMVef(ilVef).iSort))
            If blSort Then
                tlMnfSrchKey.iCode = tgMVef(ilVef).iOwnerMnfCode
                If tlMnf.iCode <> tlMnfSrchKey.iCode Then
                    ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        tlMnf.iGroupNo = 999
                        tlMnf.iCode = 0
                    End If
                End If
                slGpSort = Trim$(str$(tlMnf.iGroupNo))
            Else
                slGpSort = ""
            End If
        End If
        'Next ilVef
        slDPSort = "999"
        If blSort Then
            If (Asc(tgSpf.sUsingFeatures5) And CNTRINVSORTRC) = CNTRINVSORTRC Then     'use rate card items sort code
                For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                    If (tgMRif(llRif).iRcfCode = tgChfCntr.iRcfCode) And (tgMRif(llRif).iVefCode = tgClfCntr(llLnUpper).ClfRec.iVefCode) And (tgMRif(llRif).iRdfCode = tgClfCntr(llLnUpper).ClfRec.iRdfCode) Then
                        If (tgMRif(llRif).iSort > 0) Then
                            slDPSort = Trim$(str$(tgMRif(llRif).iSort))
                        End If
                        Exit For
                    End If
                Next llRif
            End If
        End If
        If (Asc(tgSpf.sUsingFeatures5) And CNTRINVSORTLN) = CNTRINVSORTLN Then
            slDPSort = Trim$(str$(tgClfCntr(llLnUpper).ClfRec.iLine))
        Else
            If slDPSort = "999" Then
                If blSort Then
                    tlRdfSrchKey.iCode = tgClfCntr(llLnUpper).ClfRec.iRdfCode
                    If tlRdf.iCode <> tlRdfSrchKey.iCode Then
                        ilRet = btrGetEqual(hlRdf, tlRdf, ilRdfRecLen, tlRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            If tlRdf.iSortCode > 0 Then
                                slDPSort = Trim$(str$(tlRdf.iSortCode))
                            Else
                                tlRdf.iSortCode = 999
                                tlRdf.iCode = 0
                            End If
                        End If
                    End If
                End If
                If (slDPSort = "999") Then
                    slDPSort = Trim$(str$(tgClfCntr(llLnUpper).ClfRec.iLine))
                End If
            End If
        End If
        Do While Len(slGpSort) < 3
            slGpSort = "0" & slGpSort
        Loop
        Do While Len(slVehSort) < 3
            slVehSort = "0" & slVehSort
        Loop
        Do While Len(slDPSort) < 4
            slDPSort = "0" & slDPSort
        Loop
        slLine = Trim$(str$(tgClfCntr(llLnUpper).ClfRec.iLine))
        Do While Len(slLine) < 4
            slLine = "0" & slLine
        Loop
        slRevNo = Trim$(str$(999 - tgClfCntr(llLnUpper).ClfRec.iCntRevNo))
        Do While Len(slRevNo) < 3
            slRevNo = "0" & slRevNo
        Loop
        slPropVer = Trim$(str$(tgClfCntr(llLnUpper).ClfRec.iPropVer))
        Do While Len(slPropVer) < 3
            slPropVer = "0" & slPropVer
        Loop
        tgClfCntr(llLnUpper).sKey = slOALineNo & "|" & slGpSort & slVehSort & slVehName & slDPSort & slLine & slRevNo & slPropVer
        llLnUpper = llLnUpper + 1
        ReDim Preserve tgClfCntr(0 To llLnUpper) As CLFLIST
        tgClfCntr(llLnUpper).iStatus = -1 'Not Used
        tgClfCntr(llLnUpper).lFirstCff = -1
        tgClfCntr(llLnUpper).lFirstCgf = -1
        tgClfCntr(llLnUpper).lFirstAvail = -1
        tgClfCntr(llLnUpper).lRecPos = 0
    Next llClf

    'Sort then output new headers and lines
    If UBound(tgClfCntr) - 1 > 0 Then
        'Clear last record as it contains the previous record
        llLnUpper = UBound(tgClfCntr)
        ReDim Preserve tgClfCntr(0 To llLnUpper - 1) As CLFLIST
        ReDim Preserve tgClfCntr(0 To llLnUpper) As CLFLIST
    End If
    'Obtain Flight records for line
    For llClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        For llCff = LBound(tlInCff) To UBound(tlInCff) - 1 Step 1
            If tgClfCntr(llClf).ClfRec.iLine = tlInCff(llCff).iClfLine Then
                tgCffCntr(llFlUpper).CffRec = tlInCff(llCff)
                tgCffCntr(llFlUpper).lNextCff = -1
                tgCffCntr(llFlUpper).iStatus = 0 '0=New; 1=Old and retain
                tgCffCntr(llFlUpper).lRecPos = 0
                tgCffCntr(llFlUpper).lAvgAud = 0
                tgCffCntr(llFlUpper).lPriDemoAvgAud = 0
                tgCffCntr(llFlUpper).lPriDemoPop = 0
                tgCffCntr(llFlUpper).iGameNo = 0
                
                If tgCffCntr(llFlUpper).CffRec.sDyWk = "D" Then
                    For ilDay = 0 To 6 Step 1
                        If tgCffCntr(llFlUpper).CffRec.iDay(ilDay) <> 0 Then
                            tgClfCntr(llClf).iOrigAnySpots = 1
                            Exit For
                        End If
                    Next ilDay
                Else
                    If (tgCffCntr(llFlUpper).CffRec.iSpotsWk <> 0) Or (tgCffCntr(llFlUpper).CffRec.iXSpotsWk <> 0) Then
                        tgClfCntr(llClf).iOrigAnySpots = 1
                    End If
                End If
                If tgClfCntr(llClf).lFirstCff = -1 Then
                    tgClfCntr(llClf).lFirstCff = llFlUpper
                Else
                    tgCffCntr(llFlUpper - 1).lNextCff = llFlUpper
                End If
                gUnpackDateLong tgCffCntr(llFlUpper).CffRec.iStartDate(0), tgCffCntr(llFlUpper).CffRec.iStartDate(1), tgCffCntr(llFlUpper).lStartDate    'Week Start date
                gUnpackDateLong tgCffCntr(llFlUpper).CffRec.iEndDate(0), tgCffCntr(llFlUpper).CffRec.iEndDate(1), tgCffCntr(llFlUpper).lEndDate    'Week End date
                llFlUpper = llFlUpper + 1
                ReDim Preserve tgCffCntr(0 To llFlUpper) As CFFLIST
                tgCffCntr(llFlUpper).iStatus = -1 'Not Used
                tgCffCntr(llFlUpper).lNextCff = -1
                tgCffCntr(llFlUpper).lRecPos = 0
            End If
        Next llCff
    Next llClf
    llEliminateFirst = -1
    If (llFlUpper > 32000) Or (llLnUpper > 32000) Then
        'Too many flights, reduce lines
        'Find last contract to retain
        For llClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
            llNext = tgClfCntr(llClf).lFirstCff
            Do While llNext <> -1
                If llNext > 32000 Then
                    Exit For
                End If
                llNext = tgCffCntr(llNext).lNextCff
            Loop
            If llClf > 32000 Then
                Exit For
            End If
            llEliminateFirst = tgClfCntr(llClf).lFirstCff
        Next llClf
        If llEliminateFirst <> -1 Then
            llMaxClfCode = -1
            For llClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
                If tgClfCntr(llClf).lFirstCff = llEliminateFirst Then
                    llMaxClfCode = llClf
                    Exit For
                End If
            Next llClf
            If llMaxClfCode <> -1 Then
                ReDim Preserve tgClfCntr(0 To llMaxClfCode - 1) As CLFLIST
                ReDim Preserve tgClfCntr(0 To llMaxClfCode) As CLFLIST
                tgClfCntr(llMaxClfCode).iStatus = -1 'Not Used
                tgClfCntr(llMaxClfCode).lFirstCff = -1
                tgClfCntr(llMaxClfCode).lFirstCgf = -1
                tgClfCntr(llMaxClfCode).lFirstAvail = -1
                tgClfCntr(llMaxClfCode).lRecPos = 0
            End If
        End If
    End If
    For llClf = LBound(tgClfCntr) To UBound(tgClfCntr) - 1 Step 1
        tgClfCntr(llClf).iFirstCff = tgClfCntr(llClf).lFirstCff
        llNext = tgClfCntr(llClf).lFirstCff
        Do While llNext <> -1
            tgCffCntr(llNext).iNextCff = tgCffCntr(llNext).lNextCff
            llNext = tgCffCntr(llNext).lNextCff
        Loop
        tgClfCntr(llClf).iFirstCgf = tgClfCntr(llClf).lFirstCgf
    Next llClf
    'Sort then output new headers and lines
    If UBound(tgClfCntr) - 1 >= 0 Then
        'Clear last record as it contains the previous record
        tgClfCntr(UBound(tgClfCntr)).iStatus = -1 'Not Used
        tgClfCntr(UBound(tgClfCntr)).iFirstCff = -1
        tgClfCntr(UBound(tgClfCntr)).iFirstCgf = -1
        tgClfCntr(UBound(tgClfCntr)).lFirstCff = -1
        tgClfCntr(UBound(tgClfCntr)).lFirstCgf = -1
        tgClfCntr(UBound(tgClfCntr)).lFirstAvail = -1
        tgClfCntr(UBound(tgClfCntr)).lRecPos = 0
    End If
    If UBound(tgClfCntr) - 1 > 0 Then
        ArraySortTyp fnAV(tgClfCntr(), 0), UBound(tgClfCntr), 0, LenB(tgClfCntr(0)), 0, LenB(tgClfCntr(0).sKey), 0
    End If
    igBtrError = 0
    If blSort Then
        btrDestroy hlMnf
        btrDestroy hlRdf
    End If
    gMoveCntr = True
    Exit Function
End Function

Public Sub gBuildDormantVef()
    Dim ilVef As Integer
    Dim ilUpper As Integer
    ReDim igDormantVef(0 To UBound(tgMVef)) As Integer
    ilUpper = 0
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If tgMVef(ilVef).sState = "D" Then
            igDormantVef(ilUpper) = tgMVef(ilVef).iCode
            ilUpper = ilUpper + 1
        End If
    Next ilVef
    ReDim Preserve igDormantVef(0 To ilUpper) As Integer
End Sub

Public Function gBinarySearchDormantVef(ilVefCode As Integer)
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    Dim ilRet As Integer
    
    On Error GoTo gBinarySearchDormantVefErr
    
    ilMin = LBound(igDormantVef)
    ilMax = UBound(igDormantVef)
    If (ilMax > ilMin) And (ilVefCode > 0) Then
        Do While ilMin <= ilMax
            ilMiddle = (ilMin + ilMax) \ 2
            If ilVefCode = igDormantVef(ilMiddle) Then
                'found the match
                gBinarySearchDormantVef = ilMiddle
                Exit Function
            ElseIf ilVefCode < igDormantVef(ilMiddle) Then
                ilMax = ilMiddle - 1
            Else
                'search the right half
                ilMin = ilMiddle + 1
            End If
        Loop
    End If
    gBinarySearchDormantVef = -1
    Exit Function
gBinarySearchDormantVefErr:
    gBinarySearchDormantVef = -1
    Exit Function
End Function

'6-4-19 tgLogExportLoc contains an array of vefcodes and vehicle pdf export location paths
Public Function gBinarySearchVefForLogExport(ilVefCode As Integer) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tgLogExportLoc)
    ilMax = UBound(tgLogExportLoc) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilVefCode = tgLogExportLoc(ilMiddle).iVefCode Then
            'found the match
            gBinarySearchVefForLogExport = ilMiddle
            Exit Function
        ElseIf ilVefCode < tgLogExportLoc(ilMiddle).iVefCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchVefForLogExport = -1
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopVehFeedBox             *
'*                                                     *
'*             Created:5/20/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with Vehicle  *
'*            which are permitted for the user along with*
'*            Feed define for delivery links           *
'*******************************************************
Function gPopVehFeedBox(frm As Form, llVehType As Long, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String, Optional ilIncludeVefCode As Integer = 0) As Integer
'
'   ilRet = gPopVehFeedBox (MainForm, llVehType, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error
'       llVehType(I)- Right to Left (0 to 15) Bit On means include; Bit Off exclude
'           Conventional without Feed = VEHCONV_WO_FEED = &H1
'           Conventional with Feed = VEHCONV_W_FEED = &H2
'           Selling = VEHSELLING = &H4
'           Airing = VEHAIRING = &H8
'           Log = VEHLOG = &H10     Vehicle with Log Vehicle will be included if VEHLOGVEHICLE is also set on, or exclude if VEHLOGVEHICLE is not set on
'           Virtual = VEHVIRTUAL = &H20
'           Log Vehicle = VEHLOGVEHICLE = &H40  This is only used if VEHLOG set on
'           Package Vehicle = VEHPACKAGE = &H80   This is any kind of Package
'           Standard Package = VEHSTDPKG = &H100  Standard package vehicles only
'           Simulcast Vehicle = VEHSIMUL = &H200  Simulcast vehicle
'
'           Active Vehicles = ACTIVEVEH = &H2000
'           Dormant Vehicles = DORMANTVEF = &H4000
'
'           Old definition (Changed 12/21/95)
'           ilType(I)- 0=selling and conventional(Both); 1=Airing; 2=Unused; 3=All except Log and Dormant;
'                  4=Conventional(Both); 5= Selling; 6 = Airing and Conventional(with feed);
'                  7 = Airing and Conventional(Both); 8=All vehicles; 9 = Airing, Conventional without Log Vehicle and Log Vehicles;
'                  10 = Airing, Selling, Conventional without Log Vehicle and Log Vehicles;
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- True=list was either populated or repopulated
'                  False=List was OK- it didn't require populating
'

    Dim slStamp As String   'Current time stamp
    Dim ilRet As Integer
    Dim slName As String
    Dim illoop As Integer
    'Dim tlSrchKey As INTKEY0
    Dim slNameCode As String
    Dim ilOk As Integer 'Does Vehicle match type
    Dim ilAllVehicles As Integer
    Dim llLen As Long
    Dim ilIndex As Integer
    Dim slGpSort As String
    Dim slVehSort As String
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim ilMnfRecLen As Integer
    Dim tlMnfSrchKey As INTKEY0
    Dim ilSortVehbyGroup As Integer 'True=Sort vehicle by groups; False=Sort by name only
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLowLimit As Integer
    Dim ilVpfIndex As Integer
    Dim ilZone As Integer
    
    ilSortVehbyGroup = False
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Vef.btr") & Trim$(str$(llVehType))


    If PeekArray(tlSortCode).Ptr <> 0 Then
        ilLowLimit = LBound(tlSortCode)
    Else
        slSortCodeTag = ""
        ilLowLimit = 0
    End If

    ilPop = True
    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopVehFeedBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If

    gPopVehFeedBox = CP_MSG_POPREQ
    lbcLocal.Clear
    lbcLocal.Tag = ""
    slSortCodeTag = slStamp
    If ilPop Then
        ilSortCode = ilLowLimit
        ReDim tlSortCode(ilLowLimit To ilLowLimit) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilRet = gObtainVef()
        'DoEvents    'This eliminated error in NWCalls
        For illoop = 0 To 20 Step 1
            ilOk = ilOk
        Next illoop
        If ilRet = False Then
            Exit Function
        End If
        hlMnf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            btrDestroy hlMnf
            Exit Function
        End If
        ilMnfRecLen = Len(tlMnf) 'btrRecordLength(hlMnf)  'Get and save record length

        ilAllVehicles = False
        If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Then
            ilAllVehicles = True
        Else
            For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                If tgUrf(illoop).iVefCode = 0 Then
                    ilAllVehicles = True
                    Exit For
                End If
            Next illoop
        End If
        If ilAllVehicles Then
            For illoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If mTestVehType(llVehType, tgMVef(illoop), ilIncludeVefCode) Then
                    ilVpfIndex = gVpfFind(frm, tgMVef(illoop).iCode)
                    'include only the vehicles that have a zone defined
'                    For ilZone = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
                        If (Trim$(tgVpf(ilVpfIndex).sGZone(0)) <> "") And (tgVpf(ilVpfIndex).iGMnfNCode(0) > 0) Then

                            If ilSortVehbyGroup Then
                                slVehSort = Trim$(str$(tgMVef(illoop).iSort))
                                tlMnfSrchKey.iCode = tgMVef(illoop).iOwnerMnfCode
                                If tlMnf.iCode <> tlMnfSrchKey.iCode Then
                                    ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        tlMnf.iGroupNo = 999
                                    End If
                                End If
                                slGpSort = Trim$(str$(tlMnf.iGroupNo))
                                Do While Len(slGpSort) < 3
                                    slGpSort = "0" & slGpSort
                                Loop
                                Do While Len(slVehSort) < 3
                                    slVehSort = "0" & slVehSort
                                Loop
                            Else
                                slGpSort = "000"
                                slVehSort = "000"
                            End If
                            
                            slName = tgMVef(illoop).sName 'trim cause xxx and xxx y to sort in wrong order Trim$(tlVef.sName)
                            slName = slGpSort & "|" & slVehSort & "|" & slName & "\" & Trim$(str$(tgMVef(illoop).iCode))
                            tlSortCode(ilSortCode).sKey = slName
                            If ilSortCode >= UBound(tlSortCode) Then
                                ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                            End If
                            ilSortCode = ilSortCode + 1
                        End If
'                    Next ilZone
                End If
            Next illoop
        Else
            For illoop = LBound(tgUrf) To UBound(tgUrf) Step 1
                ilIndex = gBinarySearchVef(tgUrf(illoop).iVefCode)
                If ilIndex <> -1 Then
                    If mTestVehType(llVehType, tgMVef(ilIndex), ilIncludeVefCode) Then
                    'include only the vehicles that have a zone defined
'                        For ilZone = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
                            If (Trim$(tgVpf(ilVpfIndex).sGZone(0)) <> "") And (tgVpf(ilVpfIndex).iGMnfNCode(0) > 0) Then
                                If ilSortVehbyGroup Then
                                    slVehSort = Trim$(str$(tgMVef(ilIndex).iSort))
                                    tlMnfSrchKey.iCode = tgMVef(ilIndex).iOwnerMnfCode
                                    If tlMnf.iCode <> tlMnfSrchKey.iCode Then
                                        ilRet = btrGetEqual(hlMnf, tlMnf, ilMnfRecLen, tlMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        If ilRet <> BTRV_ERR_NONE Then
                                            tlMnf.iGroupNo = 999
                                        End If
                                    End If
                                    slGpSort = Trim$(str$(tlMnf.iGroupNo))
                                    Do While Len(slGpSort) < 3
                                        slGpSort = "0" & slGpSort
                                    Loop
                                    Do While Len(slVehSort) < 3
                                        slVehSort = "0" & slVehSort
                                    Loop
                                Else
                                    slGpSort = "000"
                                    slVehSort = "000"
                                End If
                                slName = tgMVef(ilIndex).sName    'Trim$(tlVef.sName)
                                slName = slGpSort & "|" & slVehSort & "|" & slName & "\" & Trim$(str$(tgMVef(ilIndex).iCode))
                                tlSortCode(ilSortCode).sKey = slName
                                If ilSortCode >= UBound(tlSortCode) Then
                                    ReDim Preserve tlSortCode(ilLowLimit To UBound(tlSortCode) + 100) As SORTCODE
                                End If
                                ilSortCode = ilSortCode + 1
                            End If
'                        Next ilZone
                    End If
                    'Exit For
                End If
            Next illoop
        End If
        'Sort then output new headers and lines
        ReDim Preserve tlSortCode(ilLowLimit To ilSortCode) As SORTCODE
        If UBound(tlSortCode) - ilLowLimit > ilLowLimit Then
            ArraySortTyp fnAV(tlSortCode(), ilLowLimit), UBound(tlSortCode) - ilLowLimit, 0, LenB(tlSortCode(ilLowLimit)), 0, LenB(tlSortCode(ilLowLimit).sKey), 0
        End If
        btrDestroy hlMnf
    End If
    llLen = 0
    For illoop = ilLowLimit To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(illoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopVehFeedBox = CP_MSG_PARSE
            Exit Function
        End If
        ilRet = gParseItem(slName, 3, "|", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopVehFeedBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box

    Next illoop
    Exit Function

    btrDestroy hlMnf
    gDbg_HandleError "PopSubs: gPopVehFeedBox"
gPopVehFeedBoxErr2:
    ilRet = 1
    Resume Next
End Function

Public Sub gPopBooks()
    Dim ilUpper As Integer
    Dim slSQLQuery As String
    Dim dnf_rst As ADODB.Recordset

    ReDim tgBookNameInfo(0 To 0) As BOOKNAMEINFO
    slSQLQuery = "Select * from DNF_Demo_Rsrch_Names Order by dnfCode"
    Set dnf_rst = gSQLSelectCall(slSQLQuery)
    Do While Not dnf_rst.EOF
        ilUpper = UBound(tgBookNameInfo)
        tgBookNameInfo(ilUpper).iDnfCode = dnf_rst!dnfCode
        tgBookNameInfo(ilUpper).sName = dnf_rst!dnfBookName
        tgBookNameInfo(ilUpper).sBookDate = Format(dnf_rst!dnfBookDate, sgShowDateForm)
        ReDim Preserve tgBookNameInfo(0 To ilUpper + 1) As BOOKNAMEINFO
        dnf_rst.MoveNext
    Loop
End Sub

Public Sub gOpenVpf()
    Dim ilRet As Integer
    hgVpf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hgVpf, "", sgDBPath & "Vpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imVpfRecLen = Len(tmVpf)
End Sub

Public Sub gCloseVpf()
    Dim ilRet As Integer
    
    On Error Resume Next
    ilRet = btrClose(hgVpf)
    btrDestroy hgVpf
End Sub

Public Function gObtainPcf(hlPcf As Integer, llChfCode As Long, tlPcf() As PCF) As Boolean
'
'   Where:
'       <input>  hlPcf - PCF handle
'                llChfCode - contract internal code
'       <input/output> tlPcf() - Pcf record structure to be created
'
'       boolean return (O)- True = populated; False = error
'

    Dim ilPcfRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Pcf
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim llLen As Long
    Dim tlLongTypeBuff As POPLCODE   'Type field record
    Dim tlSrchKey1 As PCFKEY1
    Dim llUpper As Long

    ReDim tlPcf(0 To 0) As PCF   'VB list box clear (list box used to retain code number so record can be found)
    llUpper = UBound(tlPcf)
    gObtainPcf = True               'assume all ok
    ilExtLen = Len(tlPcf(0))  'Extract operation record size

    gObtainPcf = True               'assume ok
    tlSrchKey1.lChfCode = llChfCode
    tlSrchKey1.iPodCPMID = 0
    tlSrchKey1.iPropVer = 0
    tlSrchKey1.iCntRevNo = 0
    btrExtClear hlPcf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlPcf, tlPcf(llUpper), Len(tlPcf(0)), INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        Exit Function
    End If
    llNoRec = gExtNoRec(ilExtLen)
    Call btrExtSetBounds(hlPcf, llNoRec, -1, "UC", "Pcf", "") 'Set extract limits (all records)
    ilOffSet = gFieldOffset("Pcf", "PcfChfCode")
    tlLongTypeBuff.lCode = llChfCode
    ilRet = btrExtAddLogicConst(hlPcf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlLongTypeBuff, 4)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainPcf = False
        Exit Function
    End If
    ilRet = btrExtAddField(hlPcf, 0, ilExtLen)  'Extract whole record
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlPcf)
        btrDestroy hlPcf
        Exit Function
    End If
    ilRet = btrExtGetNext(hlPcf, tlPcf(llUpper), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        ilExtLen = Len(tlPcf(0))  'Extract operation record size
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlPcf, tlPcf(llUpper), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If (llChfCode > 0 And tlPcf(llUpper).lChfCode = llChfCode) Or (llChfCode = 0) Then
                ReDim Preserve tlPcf(0 To UBound(tlPcf) + 1) As PCF
                llUpper = UBound(tlPcf)
                ilRet = btrExtGetNext(hlPcf, tlPcf(llUpper), ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlPcf, tlPcf(llUpper), ilExtLen, llRecPos)
                Loop
            Else
                Exit Function
            End If
        Loop
    End If
    Exit Function
End Function

'TTP 10725 - Billed and Booked Cal Spots and RAB Cal Spots: not including digital line contract that should be included
'Warning: This is a slow bastard, using gObtainPcfByCntrNo then looping throught the Pcf array saves about 2 minutes on WWO 12 month B&B
Public Function gObtainPcfCPMID(llPcfCode As Long) As Integer
'   Where:
'       <input>  llPcfCode - PCF internal code
'       Integer return PodCPMID
    Dim SQLQuery As String
    Dim rst As ADODB.Recordset
    gObtainPcfCPMID = -1
    SQLQuery = "SELECT pcfPodCPMID FROM pcf_Pod_CPM_Cntr WHERE pcfCode = " & llPcfCode
    Set rst = gSQLSelectCall(SQLQuery)
    If Not rst.EOF Then
        gObtainPcfCPMID = rst!pcfPodCPMID
    End If
    Exit Function
End Function

'TTP 10955 - Billed and Booked Cal Spots: slowness reported, possibly due to including digital lines
'replacing the use of gObtainPcfCPMID by loading all PCF records in an array, so we dont have to query each PCF code individually
Public Function gObtainPcfByCntrNo(llCntrNo As Long, tlPcf() As PCF) As Boolean
'   Where:
'       <input>  llCntrNo - contract number
'       <input/output> tlPcf() - Pcf record structure to be created
'
'       boolean return (O)- True = populated; False = error

    ReDim tlPcf(0 To 0) As PCF
    Dim SQLQuery As String
    Dim rst As ADODB.Recordset
    
    gObtainPcfByCntrNo = False
    SQLQuery = ""
    SQLQuery = SQLQuery & "SELECT "
    SQLQuery = SQLQuery & "pcfCode , pcfChfCode, pcfPodCPMID, pcfVefCode, pcfPriceType, pcfStartDate, pcfEndDate, pcfImpressionGoal, pcfImpressionCap, pcfPodCPM, pcfTotalCost "
    SQLQuery = SQLQuery & "FROM pcf_Pod_CPM_Cntr "
    SQLQuery = SQLQuery & "WHERE "
    SQLQuery = SQLQuery & "pcfChfCode IN ("
    SQLQuery = SQLQuery & "SELECT chfCode "
    SQLQuery = SQLQuery & "FROM CHF_Contract_Header "
    SQLQuery = SQLQuery & "WHERE chfCntrNo = " & llCntrNo
    SQLQuery = SQLQuery & ")"
    
    Set rst = gSQLSelectCall(SQLQuery)
    Do
        tlPcf(UBound(tlPcf)).lCode = rst!pcfCode
        tlPcf(UBound(tlPcf)).lChfCode = rst!pcfChfCode
        tlPcf(UBound(tlPcf)).iPodCPMID = rst!pcfPodCPMID
        tlPcf(UBound(tlPcf)).iVefCode = rst!pcfVefCode
        tlPcf(UBound(tlPcf)).sPriceType = rst!pcfPriceType
        'tlPcf(UBound(tlPcf)).iStartDate(0) = rst!pcfStartDate 'Dont really need these right now, so let's leave them unpopulated
        'tlPcf(UBound(tlPcf)).iStartDate (1)
        'tlPcf(UBound(tlPcf)).iEndDate(0) = rst!pcfEndDate
        'tlPcf(UBound(tlPcf)).iEndDate (1)
        tlPcf(UBound(tlPcf)).lImpressionGoal = rst!pcfImpressionGoal
        tlPcf(UBound(tlPcf)).lImpressionCap = rst!pcfImpressionCap
        tlPcf(UBound(tlPcf)).lPodCPM = rst!pcfPodCPM
        tlPcf(UBound(tlPcf)).lTotalCost = rst!pcfTotalCost
        ReDim Preserve tlPcf(0 To UBound(tlPcf) + 1) As PCF
        rst.MoveNext
    Loop While Not rst.EOF
    Exit Function
End Function

