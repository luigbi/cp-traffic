Attribute VB_Name = "COLLECTSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Collect.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Collect.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Collection subs and functions
Option Explicit
Option Compare Text

Public igPOCodeXFre As Integer '0=agfCode or adfcode to transfer to
Public igPOReturn As Integer   '0=Cancel; 1=transfer to agency, 2= Transfer to advertiser

Public sgInvNoName As String

'Collection data
Public tgCAgf() As AGFADFCOLL
Public sgCAgfStamp As String 'Date/Time stamp of file
Public tgCAdf() As AGFADFCOLL   'Direct only
Public sgCAdfStamp As String 'Date/Time stamp of file
Public tgCAdfAll() As AGFADFCOLL
Public sgCAdfAllStamp As String 'Date/Time stamp of file
Public tgCSlf() As SLFEXT
Public sgCSlfStamp As String 'Date/Time stamp of file
Public tgCCdf() As CDFLIST
Public sgCCdfStamp As String 'Date/Time stamp of file
Public tgCommUrf() As URFEXT
Public sgCommUrfStamp As String 'Date/Time stamp of file
Public tgTransactionCode() As SORTCODE
Public sgTransactionCodeTag As String

Public tgColAdvertiser() As SORTCODE
Public sgColAdvertiserTag As String

Dim imNonPayeeAdfCode() As Integer

Type VEHINV
    iAgfCode As Integer
    iAdfCode As Integer
    iAirVefCode As Integer
    iBillVefCode As Integer
    iPkLineNo As Integer
    lNetPlusTax As Long     '1-16-02
    lTax1 As Long           '1-16-02
    lTax2 As Long           '1-16-02
    sAmount As String          'Net amount
    iMnfGroup As Integer    'Participant
    iMnfItem As Integer
    sAmountGross As String
    lSbfCode As Long        '2-25-04
    iBacklogTrfCode As Integer
    lGsfCode As Long
    lPcfCode As Long    'TTP 10849 - Collections: invoice adjustment applied to a digital invoice (not by vehicle) gets the rvfPcfCode set to the same values instead of using different PcfCodes when there are multiple IN records the AN is applied to
End Type
Type SSPART
    sKey As String * 50 'Sales Source/Participant
    iMnfSSCode As Integer   'Sales Source
    iMnfGroup As Integer    'Participant
    iVefIndex As Integer
    iSSPartLp As Integer
    iProdPct As Integer
    sUpdateRVF As String * 1
End Type
'8-4-10 Moved to RECDEFAL to implement a version of ZeroPurge into GetPaid export
'Type ZEROPURGE
'    lInvNo As Long
'    iAgfCode As Integer
'    iAdfCode As Integer
'    sAmount As String * 20
'    sType As String * 1
'    lFirstLk As Long
'End Type
'
'Type ZPLINK
'    lNextLk As Long
'    lRvfCode As Long
'End Type

'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAdvtCollectBox              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     advertiser name and             *
'*                     specified field                 *
'*                                                     *
'*******************************************************
Function gPopAdvtCollectBox(frm As Form, slSort As String, lbcLocal As control, lbcMster As control) As Integer
'
'   ilRet = gPopAdvtCollectBox (MainForm, slSort, lbcLocal, lbcCtrl)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       slSort(I)- Sort and field selection
'                   "A" = Alphabetic
'                   "B" = % over 90
'                   "C" = Amount owed
'                   "D" = Days since Paid
'                   "E" = Credit restriction
'                   "F" = Payment Rating
'                   "G" = Salesperson
'       lbcLocal (I)- List box to be populated from the master list box
'       lbcCtrl (I)- Master List box Control containing name and code #
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'Agf date/time stamp
    Dim slCAdfStamp As String    'Agf date/time stamp
    Dim hlAdf As Integer        'Adf handle
    Dim ilAdfRecLen As Integer     'Record length
    Dim tlAdf As ADF
    Dim slName As String   'Name plus city
    Dim slListBox As String
    Dim slNameCode As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilList As Integer
    Dim ilBuildTgCAdf As Integer

    slStamp = gFileDateTime(sgDBPath & "adf.Btr") & slSort
    If lbcMster.Tag <> "" Then
        If StrComp(slStamp, lbcMster.Tag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopAdvtCollectBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
        End If
    End If
    ilBuildTgCAdf = True
    slCAdfStamp = gFileDateTime(sgDBPath & "Adf.Btr")
    If sgCAdfAllStamp <> "" Then
        If StrComp(slCAdfStamp, sgCAdfAllStamp, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                ilBuildTgCAdf = False
            End If
        End If
    End If
    gPopAdvtCollectBox = CP_MSG_POPREQ
    lbcMster.Tag = slStamp
    lbcMster.Clear   'VB list box clear (list box used to retain code number so record can be found)
    lbcLocal.Clear
    ilRet = mObtainSalesperson()
    If ilBuildTgCAdf Then
        sgCAdfAllStamp = slCAdfStamp
        ReDim tgCAdfAll(0 To 0) As AGFADFCOLL
        hlAdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopAdvtCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtCollectBox (btrOpen):" & "Adf.Btr", frm
        On Error GoTo 0
        ilAdfRecLen = Len(tlAdf) 'btrRecordLength(hlAgf)  'Get and save record length
        btrExtClear hlAdf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlAdf, tlAdf, ilAdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            On Error GoTo gPopAdvtCollectBoxErr
            gBtrvErrorMsg ilRet, "gPopAdvtCollectBox (btrGetFirst):" & "Adf.Btr", frm
            On Error GoTo 0
            ilRet = mCreateCollectSortRec(hlAdf, "ADFALL", tgCAdfAll())
        End If
        ilRet = btrClose(hlAdf)
        On Error GoTo gPopAdvtCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtCollectBox (btrReset):" & "Adf.Btr", frm
        On Error GoTo 0
        btrDestroy (hlAdf)
    End If
    For ilLoop = LBound(tgCAdfAll) To UBound(tgCAdfAll) - 1 Step 1
        slListBox = mMakeCollectSort("ADF", slSort, tgCAdfAll(ilLoop))
        slListBox = slListBox & "|" & Trim$(str$(ilLoop))
        slListBox = slListBox & "\" & Trim$(str$(tgCAdfAll(ilLoop).iCode))
        lbcMster.AddItem slListBox    'Add ID (retain matching sorted order) and Code number to list box
    Next ilLoop
    For ilList = 0 To lbcMster.ListCount - 1 Step 1
        slNameCode = lbcMster.List(ilList)
        ilRet = gParseItem(slNameCode, 2, "|", slName)  'Obtain Index and code number
        If ilRet <> CP_MSG_NONE Then
            gPopAdvtCollectBox = CP_MSG_PARSE
            Exit Function
        End If
        ilRet = gParseItem(slName, 1, "\", slStr)       'Obtain index
        If ilRet <> CP_MSG_NONE Then
            gPopAdvtCollectBox = CP_MSG_PARSE
            Exit Function
        End If
        ilLoop = Val(slStr)
        slListBox = mMakeCollectListImage("ADF", slSort, tgCAdfAll(ilLoop))
        lbcLocal.AddItem slListBox  'Add ID to list box
    Next ilList
    Exit Function
gPopAdvtCollectBoxErr:
    ilRet = btrClose(hlAdf)
    btrDestroy hlAdf
    gDbg_HandleError "CollectSubs: gPopAdvtCollectBox"
'    gPopAdvtCollectBox = CP_MSG_NOSHOW
'    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAdvtCommentCollectBox       *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     comments and specified field    *
'*                                                     *
'*******************************************************
Function gPopAdvtCommentCollectBox(frm As Form, ilAdvtCode As Integer, lbcMster As control) As Integer
'
'   ilRet = gPopAdvtCommentCollectBox (MainForm, ilAdvtCode, lbcCtrl)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAdvtCode(I)- Advertiser code to obtain comment for
'       lbcCtrl (I)- Master List box Control containing date and time, agency name and comment rec #
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'Cdf date/time stamp
    Dim hlAgf As Integer        'Cdf handle
    Dim ilAgfRecLen As Integer     'Record length
    Dim tlagf As AGF
    Dim tlAgfSrchKey As INTKEY0 'AGF key record image
    Dim hlCdf As Integer        'Cdf handle
    Dim ilCdfRecLen As Integer     'Record length
    Dim tlCdf As CDF
    Dim llNoRec As Long         'Number of records in Sof
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim tlCdfExt As CDFADVTEXT
    Dim slName As String
    Dim slDate As String
    Dim slTime As String
    Dim llTime As Long
    Dim slNameSort As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record

    slStamp = gFileDateTime(sgDBPath & "Cdf.Btr") & "Advt" & Trim$(str$(ilAdvtCode))
    If lbcMster.Tag <> "" Then
        If StrComp(slStamp, lbcMster.Tag, 1) = 0 Then
            gPopAdvtCommentCollectBox = CP_MSG_NOPOPREQ
            Exit Function
        End If
    End If
    gPopAdvtCommentCollectBox = CP_MSG_POPREQ
    lbcMster.Tag = slStamp
    lbcMster.Clear   'VB list box clear (list box used to retain code number so record can be found)
    hlAgf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopAdvtCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtCommentCollectBox (btrOpen):" & "Agf.Btr", frm
    On Error GoTo 0
    ilAgfRecLen = Len(tlagf) 'btrRecordLength(hlAgf)  'Get and save record length
    hlCdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlCdf, "", sgDBPath & "Cdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopAdvtCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtCommentCollectBox (btrOpen):" & "Cdf.Btr", frm
    On Error GoTo 0
    ilCdfRecLen = Len(tlCdf) 'btrRecordLength(hlAgf)  'Get and save record length
    ilExtLen = Len(tlCdfExt)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlCdf) 'Obtain number of records
    btrExtClear hlCdf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlCdf, tlCdf, ilCdfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlAgf)
        On Error GoTo gPopAdvtCommentCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtCommentCollectBox (btrReset):" & "Agf.Btr", frm
        On Error GoTo 0
        btrDestroy hlAgf
        ilRet = btrClose(hlCdf)
        On Error GoTo gPopAdvtCommentCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtCommentCollectBox (btrReset):" & "Cdf.Btr", frm
        On Error GoTo 0
        btrDestroy hlCdf
        Exit Function
    Else
        On Error GoTo gPopAdvtCommentCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtCommentCollectBox (btrReset):" & "Cdf.Btr", frm
        On Error GoTo 0
    End If
    Call btrExtSetBounds(hlCdf, llNoRec, -1, "UC", "CDFADVTEXTPK", CDFADVTEXTPK) 'Set extract limits (all records)
    tlIntTypeBuff.iType = ilAdvtCode
    ilOffSet = gFieldOffset("Cdf", "CdfAdfCode")
    ilRet = btrExtAddLogicConst(hlCdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
    ilOffSet = gFieldOffset("Cdf", "CdfAgfCode")
    ilRet = btrExtAddField(hlCdf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAdvtCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Cdf", "CdfEntryDate")
    ilRet = btrExtAddField(hlCdf, ilOffSet, 4) 'Extract entry date field
    On Error GoTo gPopAdvtCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Cdf", "CdfEntryTime")
    ilRet = btrExtAddField(hlCdf, ilOffSet, 4) 'Extract entry date field
    On Error GoTo gPopAdvtCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
    On Error GoTo 0
    'ilRet = btrExtGetNextExt(hlCdf)    'Extract record
    ilRet = btrExtGetNext(hlCdf, tlCdfExt, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo gPopAdvtCommentCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtCommentCollectBox (btrExtGetNextExt):" & "Cdf.Btr", frm
        On Error GoTo 0
        tlagf.iCode = 0
        tlagf.sName = ""
        ilExtLen = Len(tlCdfExt)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlCdf, tlCdfExt, ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlCdf, tlCdfExt, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            gUnpackDateForSort tlCdfExt.iDateEntrd(0), tlCdfExt.iDateEntrd(1), slDate
            'gUnpackDateLong tlCdfExt.iDateEntrd(0), tlCdfExt.iDateEntrd(1), llDate
            'llDate = 99999 - llDate
            'slDate = Trim$(Str$(llDate))
            'Do While Len(slDate) < 5
            '    slDate = "0" & slDate
            'Loop
            gUnpackTime tlCdfExt.iTimeEntrd(0), tlCdfExt.iTimeEntrd(1), "A", "1", slTime
            llTime = CLng(gTimeToCurrency(slTime, False))
            slTime = Trim$(str$(llTime))
            Do While Len(slTime) < 7
                slTime = "0" & slTime
            Loop
            If tlCdfExt.iAgfCode > 0 Then
                If tlCdfExt.iAgfCode <> tlagf.iCode Then
                    tlAgfSrchKey.iCode = tlCdfExt.iAgfCode
                    ilRet = btrGetEqual(hlAgf, tlagf, ilAgfRecLen, tlAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    On Error GoTo gPopAdvtCommentCollectBoxErr
                    gBtrvErrorMsg ilRet, "gPopAdvtCommentCollectBox (btrGetEqual):" & "Agf.Btr", frm
                    On Error GoTo 0
                End If
                slName = Trim$(tlagf.sName) & ", " & Trim$(tlagf.sCityID)
            Else
                tlagf.iCode = 0
                tlagf.sName = ""
                slName = " "
            End If
            slNameSort = slDate & "\" & slTime & "\" & slName & "\" & Trim$(str$(llRecPos))
            lbcMster.AddItem slNameSort    'Add ID (retain matching sorted order) and Code number to list box
            ilRet = btrExtGetNext(hlCdf, tlCdfExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlCdf, tlCdfExt, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlAgf)
    On Error GoTo gPopAdvtCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtCommentCollectBox (btrReset):" & "Agf.Btr", frm
    On Error GoTo 0
    btrDestroy hlAgf
    ilRet = btrClose(hlCdf)
    On Error GoTo gPopAdvtCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtCommentCollectBox (btrReset):" & "Cdf.Btr", frm
    On Error GoTo 0
    btrDestroy hlCdf
    Exit Function
gPopAdvtCommentCollectBoxErr:
    ilRet = btrClose(hlAgf)
    ilRet = btrClose(hlCdf)
    btrDestroy hlAgf
    btrDestroy hlCdf
    gPopAdvtCommentCollectBox = CP_MSG_NOSHOW
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAdvtFromRvfBox              *
'*                                                     *
'*             Created:7/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     advertiser names given agency   *
'*                     code from receivable file       *
'*                                                     *
'*******************************************************
Function gPopAdvtFromRvfBox(frm As Form, ilAgfCode As Integer, lbcLocal As control, lbcMster As control) As Integer
'
'   ilRet = gPopAdvtFromRvfBox (MainForm, ilAgfCode, lbcLocal, lbcCtrl)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAgfCode(I)- Agency code to find associated advertiser for
'       lbcLocal (I)- List box to be populated from the master list box
'       lbcCtrl (I)- Master List box control containing name and code #
'       ilRet (O)- Error code (0 if no error)
'
'
    Dim slStamp As String    'RvF date/time stamp
    Dim hlRvf As Integer        'RvF handle
    Dim tlRvf As RVF
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Rvf
    Dim llRecPos As Long
    Dim tlRvfExt As RVFEXT
    Dim ilExtRecLen As Integer     'Record length
    Dim slName As String
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilOffSet As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record

    slStamp = gFileDateTime(sgDBPath & "Rvf.Btr") & Trim$(str$(ilAgfCode))
    If lbcMster.Tag <> "" Then
        If StrComp(slStamp, lbcMster.Tag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopAdvtFromRvfBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
        End If
    End If
    gPopAdvtFromRvfBox = CP_MSG_POPREQ
    hlRvf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopAdvtFromRvfBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtFromRvfBox (btrOpen):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlRvf) 'btrRecordLength(hlRvf)  'Get and save record length
    ilExtRecLen = Len(tlRvfExt)
    lbcMster.Clear   'VB list box clear (list box used to retain code number so record can be found)
    lbcLocal.Clear
    lbcMster.Tag = slStamp
    ilRet = gObtainAdvt()
    If Not ilRet Then
        ilRet = btrClose(hlRvf)
        btrDestroy hlRvf
        Exit Function
    End If
    llNoRec = gExtNoRec(ilExtRecLen) 'btrRecords(hlRvf) 'Obtain number of records
    btrExtClear hlRvf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlRvf, tlRvf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlRvf)
        On Error GoTo gPopAdvtFromRvfBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtFromRvfBox (btrReset):" & "Rvf.Btr", frm
        On Error GoTo 0
        btrDestroy hlRvf
        Exit Function
    Else
        On Error GoTo gPopAdvtFromRvfBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtFromRvfBox (btrGetFirst):" & "Rvf.Btr", frm
        On Error GoTo 0
    End If
    Call btrExtSetBounds(hlRvf, llNoRec, -1, "UC", "RVFEXTPK", RVFEXTPK) 'Set extract limits (all records)
    tlIntTypeBuff.iType = ilAgfCode
    ilOffSet = gFieldOffset("Rvf", "RvfAgfCode")
    ilRet = btrExtAddLogicConst(hlRvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
    ilRet = btrExtAddField(hlRvf, ilOffSet, ilExtRecLen)  'Extract agency and advertiser fields (first two fields)
    On Error GoTo gPopAdvtFromRvfBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtFromRvfBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    'ilRet = btrExtGetNextExt(hlRvf)    'Extract record
    ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtRecLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo gPopAdvtFromRvfBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtFromRvfBox (btrExtGetNextExt):" & "Rvf.Btr", frm
        On Error GoTo 0
        'ilRet = btrExtGetFirst(hlRvf, tlRvfExt, ilExtRecLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtRecLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilFound = False
            'For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
            '    If tgCommAdf(ilLoop).iCode = tlRvfExt.iAdfCode Then
                ilLoop = gBinarySearchAdf(tlRvfExt.iAdfCode)
                If ilLoop <> -1 Then
                    ilFound = True
                    If (Trim$(tgCommAdf(ilLoop).sAddrID) <> "") And (tgCommAdf(ilLoop).sBillAgyDir = "D") Then
                        slName = Trim$(tgCommAdf(ilLoop).sName) & ", " & Trim$(tgCommAdf(ilLoop).sAddrID) & "\" & Trim$(str$(tlRvfExt.iAdfCode))
                    Else
                        slName = Trim$(tgCommAdf(ilLoop).sName) & "\" & Trim$(str$(tlRvfExt.iAdfCode))
                    End If
            '        Exit For
                End If
            'Next ilLoop
            If ilFound Then
                gFindMatch slName, 0, lbcMster
                If gLastFound(lbcMster) < 0 Then
                    lbcMster.AddItem slName
                End If
            End If
            ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtRecLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtRecLen, llRecPos)
            Loop
        Loop
        For ilLoop = 0 To lbcMster.ListCount - 1 Step 1
            slNameCode = lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            If ilRet <> CP_MSG_NONE Then
                gPopAdvtFromRvfBox = CP_MSG_PARSE
                Exit Function
            End If
            lbcLocal.AddItem Trim$(slName)  'Add ID to list box
        Next ilLoop
    End If
    ilRet = btrClose(hlRvf)
    On Error GoTo gPopAdvtFromRvfBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtFromRvfBox (btrReset):" & "Rvf.Btr", frm
    On Error GoTo 0
    btrDestroy hlRvf
    Exit Function
gPopAdvtFromRvfBoxErr:
    ilRet = btrClose(hlRvf)
    btrDestroy hlRvf
    gDbg_HandleError "CollectSubs: gPopAdvtFromRvfBox"
'    gPopAdvtFromRvfBox = CP_MSG_NOSHOW
'    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAdvtTransactionBox          *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     transaction and specified field *
'*                                                     *
'*******************************************************
'Function gPopAdvtTransactionBox (Frm As Form, hlRvf As Integer, ilAdvtCode As Integer, ilCashTrade As Integer, lbcMster As Control) As Integer
Function gPopAdvtTransactionBox(frm As Form, hlRvf As Integer, hlChf As Integer, ilAdvtCode As Integer, ilCashTrade As Integer, tlSortCode() As SORTCODE, slSortCodeTag As String, ilNonPayeeAdvt As Integer, ilHistory As Integer) As Integer
'
'   ilRet = gPopAdvtTransactiontBox (MainForm, hlRvf, ilAdvtCode, ilCashTrade, lbcCtrl)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       hlRvf(I)- Handle to RVF or PHF
'       ilAdvtCode(I)- Advertiser code to obtain comment for
'       ilCashTrade(I)- 0=Cash; 1=Trade; 2=Merchandising; 3=Promotion; 4=Installment Revenue
'       lbcCtrl (I)- Master List box Control containing date, advt name and comment rec #
'       ilNonPayee(I)- Selected Advertiser was a Non-Payee (include only transaction associated with the Non-Payee)
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'Rvf date/time stamp
    Dim hlAgf As Integer        'Agf handle
    Dim ilAgfRecLen As Integer     'Record length
    Dim tlagf As AGF
    Dim tlAgfSrchKey As INTKEY0 'AGF key record image
    Dim ilRvfRecLen As Integer     'Record length
    Dim tlRvf As RVF
    Dim llNoRec As Long         'Number of records in Sof
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim tlRvfExt As RVFTRANEXT
    Dim slName As String
    Dim slDate As String
    Dim slInvNo As String
    Dim ilFound As Integer
    Dim slNameSort As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim ilMnfRecLen As Integer
    Dim tlMnfSrchKey As INTKEY0
    Dim slAirVehicle As String
    Dim slBillVehicle As String
    Dim slGpSort As String
    Dim slPkLineNo As String
    Dim slVehSort As String
    Dim slNTRSort As String     '2-7-07
    Dim slParticipantSort As String     '2-7-07
    Dim ilIndex As Integer
    Dim llLen As Long
    Dim ilVef As Integer
    Dim ilSortVehbyGroup As Integer 'True=Sort vehicle by groups; False=Sort by name only
    Dim llSortCode As Long
    Dim tlChf As CHF            'SBF record image
    Dim tlChfSrchKey1 As CHFKEY1 'SBF key record image
    Dim ilChfRecLen As Integer  'SBF record length
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilAgfIndex As Integer

    ilChfRecLen = Len(tlChf)
    ilSortVehbyGroup = False
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Rvf.Btr") & "Advt" & Trim$(str$(ilAdvtCode)) & Trim$(str$(ilCashTrade))
    'If lbcMster.Tag <> "" Then
    '    If StrComp(slStamp, lbcMster.Tag, 1) = 0 Then
    If slSortCodeTag <> "" Then
        If slStamp = slSortCodeTag Then
            gPopAdvtTransactionBox = CP_MSG_NOPOPREQ
            Exit Function
        End If
    End If
    gPopAdvtTransactionBox = CP_MSG_POPREQ
    'lbcMster.Tag = slStamp
    'lbcMster.Clear   'VB list box clear (list box used to retain code number so record can be found)
    llSortCode = 0
    ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    slSortCodeTag = slStamp
    ilRet = gObtainVef()
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
    hlAgf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrOpen):" & "Agf.Btr", frm
    On Error GoTo 0
    ilAgfRecLen = Len(tlagf) 'btrRecordLength(hlAgf)  'Get and save record length
    ilRvfRecLen = Len(tlRvf) 'btrRecordLength(hlRvf)  'Get and save record length
    ilExtLen = Len(tlRvfExt)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlRvf) 'Obtain number of records
    btrExtClear hlRvf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlRvf, tlRvf, ilRvfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlAgf)
        On Error GoTo gPopAdvtTransactionBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrReset):" & "Agf.Btr", frm
        On Error GoTo 0
        btrDestroy hlAgf
        btrDestroy hlMnf
        Exit Function
    Else
        On Error GoTo gPopAdvtTransactionBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrReset):" & "Rvf.Btr", frm
        On Error GoTo 0
    End If
    Call btrExtSetBounds(hlRvf, llNoRec, -1, "UC", "RVFTRANEXTPK", RVFTRANEXTPK) 'Set extract limits (all records)
    If ilHistory Then
         gPackDate sgDateRangeStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
         ilOffSet = gFieldOffset("rvf", "rvfTranDate")
         ilRet = btrExtAddLogicConst(hlRvf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
         gPackDate sgDateRangeEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
         ilOffSet = gFieldOffset("rvf", "rvfTranDate")
         ilRet = btrExtAddLogicConst(hlRvf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    End If
    tlIntTypeBuff.iType = ilAdvtCode
    ilOffSet = gFieldOffset("Rvf", "RvfAdfCode")
    ilRet = btrExtAddLogicConst(hlRvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
    ilOffSet = gFieldOffset("Rvf", "RvfAgfCode")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfAdfCode")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfInvNo")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 4) 'Extract entry date field
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfAirVefCode")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2) 'Extract entry date field
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 4) 'Extract entry date field
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfTranType")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2) 'Extract transaction type
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfCashTrade")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 1) 'Extract Cash/Trade
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfBillVefCode")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2) 'Extract entry date field
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfPkLineNo")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2) 'Extract entry date field
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0

    ilOffSet = gFieldOffset("Rvf", "RvfMnfItem")        '7-7-06
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0

    ilOffSet = gFieldOffset("Rvf", "RvfMnfGroup")        '7-7-06
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0

    ilOffSet = gFieldOffset("Rvf", "RvfType")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 1) 'Extract Cash/Trade
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0

    ilOffSet = gFieldOffset("Rvf", "RvfCntrNo")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 4) 'Extract entry date field
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0

    'ilRet = btrExtGetNextExt(hlRvf)    'Extract record
    ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo gPopAdvtTransactionBoxErr
        gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrExtGetNextExt):" & "Rvf.Btr", frm
        On Error GoTo 0
        tlagf.iCode = 0
        tlagf.sName = ""
        ilExtLen = Len(tlRvfExt)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlRvf, tlRvfExt, ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilFound = False
            If ilCashTrade = 0 Then
                If (tlRvfExt.sCashTrade = "C") Then
                    ilFound = True
                End If
            ElseIf ilCashTrade = 1 Then
                If (tlRvfExt.sCashTrade = "T") Then
                    ilFound = True
                End If
            ElseIf ilCashTrade = 2 Then
                If (tlRvfExt.sCashTrade = "M") Then
                    ilFound = True
                End If
            ElseIf ilCashTrade = 3 Then
                If (tlRvfExt.sCashTrade = "P") Then
                    ilFound = True
                End If
            ElseIf ilCashTrade = 4 Then
                If (tlRvfExt.sCashTrade = "C") And (tlRvfExt.sType = "A") And ((tlRvfExt.sTranType = "HI") Or (tlRvfExt.sTranType = "AN")) Then
                    tlChfSrchKey1.lCntrNo = tlRvfExt.lCntrNo
                    tlChfSrchKey1.iCntRevNo = 32000
                    tlChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hlChf, tlChf, ilChfRecLen, tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                    'If contract missing, assume that transaction entered via Backlog, 5/8/04
                    Do While (ilRet = BTRV_ERR_NONE) And (tlChf.lCntrNo = tlRvfExt.lCntrNo)
                        If tlChf.sSchStatus = "F" Then
                            Exit Do
                        End If
                        ilRet = btrGetNext(hlChf, tlChf, ilChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (tlChf.lCntrNo = tlRvfExt.lCntrNo) And (tlChf.sSchStatus = "F") Then
                        If tlChf.sInstallDefined = "Y" Then
                            ilFound = True
                        End If
                    End If
                End If
            End If
            'When retrieving Non-Payee advertiser only include non-payee transaction
            If ilNonPayeeAdvt Then
                If tlRvfExt.iAgfCode > 0 Then
                    ilFound = False
                End If
            End If
            If ilFound Then
                ilFound = False
                If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Or (tgUrf(0).iMnfHubCode <= 0) Or ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) <> USINGHUB) Then
                    ilFound = True
                Else
                    ilVef = gBinarySearchVef(tlRvfExt.iBillVefCode)
                    If ilVef <> -1 Then
                        If ((tgUrf(0).iMnfHubCode = tgMVef(ilVef).iMnfHubCode) And ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB)) Then
                            ilFound = True
                        End If
                    End If
                End If
            End If
            If ilFound Then
                gUnpackDate tlRvfExt.iTranDate(0), tlRvfExt.iTranDate(1), slDate
                slDate = Trim$(str$(gDateValue(slDate)))
                Do While Len(slDate) < 6
                    slDate = "0" & slDate
                Loop
                slInvNo = Trim$(str$(tlRvfExt.lInvNo))
                Do While Len(slInvNo) < 12
                    slInvNo = "0" & slInvNo
                Loop
                If tlRvfExt.iAgfCode > 0 Then
                    If tlRvfExt.iAgfCode <> tlagf.iCode Then
                        tlAgfSrchKey.iCode = tlRvfExt.iAgfCode
                        'ilRet = btrGetEqual(hlAgf, tlAgf, ilAgfRecLen, tlAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        ilAgfIndex = gBinarySearchAgf(tlRvfExt.iAgfCode)       '5-11-04
                        If ilAgfIndex = -1 Then
                            On Error GoTo gPopAdvtTransactionBoxErr
                            gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrGetEqual):" & "Agf.Btr", frm
                            On Error GoTo 0
                        End If
                    End If
                    'slName = Trim$(tlAgf.sName) & "," & Trim$(tlAgf.sCityID)
                    slName = Trim$(tgCommAgf(ilAgfIndex).sName) & ", " & Trim$(tgCommAgf(ilAgfIndex).sCityID)  '5-11-04
                Else
                    tlagf.iCode = 0
                    tlagf.sName = ""
                    slName = " "
                End If
                '2-7-07 put the ntr type into the sort so they will sort together
                slNTRSort = Trim$(str$(tlRvfExt.iMnfItem))
                Do While Len(slNTRSort) < 5
                    slNTRSort = "0" & slNTRSort
                Loop
                slParticipantSort = Trim$(str$(tlRvfExt.iMnfGroup))
                Do While Len(slParticipantSort) < 5
                    slParticipantSort = "0" & slParticipantSort
                Loop
                If tlRvfExt.iAirVefCode > 0 Then
                    'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If tgMVef(ilIndex).iCode = tlRvfExt.iAirVefCode Then
                        ilIndex = gBinarySearchVef(tlRvfExt.iAirVefCode)
                        If ilIndex <> -1 Then
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
                            slAirVehicle = slGpSort & slVehSort & tgMVef(ilIndex).sName
                        End If
                    'Next ilIndex
                Else
                    slGpSort = "999"
                    slVehSort = "999"
                    slAirVehicle = slGpSort & slVehSort & "                    "
                End If
                If tlRvfExt.iAirVefCode = tlRvfExt.iBillVefCode Then
                    slBillVehicle = slAirVehicle
                ElseIf tlRvfExt.iBillVefCode > 0 Then
                    'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If tgMVef(ilIndex).iCode = tlRvfExt.iBillVefCode Then
                        ilIndex = gBinarySearchVef(tlRvfExt.iBillVefCode)
                        If ilIndex <> -1 Then
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
                            slBillVehicle = slGpSort & slVehSort & tgMVef(ilIndex).sName
                        End If
                    'Next ilIndex
                Else
                    slGpSort = "999"
                    slVehSort = "999"
                    slBillVehicle = slGpSort & slVehSort & "                    "
                End If
                slPkLineNo = Trim$(str$(tlRvfExt.iPkLineNo))
                Do While Len(slPkLineNo) < 4
                    slPkLineNo = "0" & slPkLineNo
                Loop
                slNameSort = slName & "|" & slInvNo & "|" & slInvNo & slBillVehicle & slPkLineNo & slAirVehicle & slParticipantSort & slNTRSort & slDate & "\" & Trim$(str$(llRecPos))
                'If Not gOkAddStrToListBox(slNameSort, llLen, True) Then
                '    Exit Do
                'End If
                'lbcMster.AddItem slNameSort    'Add ID (retain matching sorted order) and Code number to list box
                tlSortCode(llSortCode).sKey = slNameSort
                llSortCode = llSortCode + 1
                ReDim Preserve tlSortCode(0 To llSortCode) As SORTCODE
            End If
            ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    If UBound(tlSortCode) - 1 > 0 Then
        ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
    End If
    ilRet = btrClose(hlAgf)
    On Error GoTo gPopAdvtTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAdvtTransactionBox (btrReset):" & "Agf.Btr", frm
    On Error GoTo 0
    btrDestroy hlAgf
    btrDestroy hlMnf
    Exit Function
gPopAdvtTransactionBoxErr:
    ilRet = btrClose(hlAgf)
    btrDestroy hlAgf
    btrDestroy hlMnf
    'btrDestroy hlRvf
    gPopAdvtTransactionBox = CP_MSG_NOSHOW
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAgencyFromRvfBox            *
'*                                                     *
'*             Created:7/22/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     agency names given advertiser   *
'*                     code from receivable file       *
'*                                                     *
'*******************************************************
Function gPopAgencyFromRvfBox(frm As Form, ilAdfCode As Integer, lbcLocal As control, lbcMster As control) As Integer
'
'   ilRet = gPopAgencyFromRvfBox (MainForm, ilAdfCode, lbcLocal, lbcCtrl)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAdfCode(I)- Agency code to find associated advertiser for
'       lbcLocal (I)- List box to be populated from the master list box
'       lbcCtrl (I)- Master List box control containing name and code #
'       ilRet (O)- Error code (0 if no error)
'
'
    Dim slStamp As String    'RvF date/time stamp
    Dim hlRvf As Integer        'RvF handle
    Dim tlRvf As RVF
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Rvf
    Dim llRecPos As Long
    Dim tlRvfExt As RVFEXT
    Dim ilExtRecLen As Integer     'Record length
    Dim slName As String
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilOffSet As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record

    slStamp = gFileDateTime(sgDBPath & "Rvf.Btr") & Trim$(str$(ilAdfCode))
    If lbcMster.Tag <> "" Then
        If StrComp(slStamp, lbcMster.Tag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopAgencyFromRvfBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
        End If
    End If
    gPopAgencyFromRvfBox = CP_MSG_POPREQ
    hlRvf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopAgencyFromRvfBoxErr
    gBtrvErrorMsg ilRet, "gPopAgencyFromRvfBox (btrOpen):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilRecLen = Len(tlRvf) 'btrRecordLength(hlRvf)  'Get and save record length
    ilExtRecLen = Len(tlRvfExt)
    lbcMster.Clear   'VB list box clear (list box used to retain code number so record can be found)
    lbcLocal.Clear
    lbcMster.Tag = slStamp
    ilRet = gObtainAgency()
    If Not ilRet Then
        ilRet = btrClose(hlRvf)
        btrDestroy hlRvf
        Exit Function
    End If
    llNoRec = gExtNoRec(ilExtRecLen) 'btrRecords(hlRvf) 'Obtain number of records
    btrExtClear hlRvf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlRvf, tlRvf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlRvf)
        On Error GoTo gPopAgencyFromRvfBoxErr
        gBtrvErrorMsg ilRet, "gPopAgencyFromRvfBox (btrReset):" & "Rvf.Btr", frm
        On Error GoTo 0
        btrDestroy hlRvf
        Exit Function
    Else
        On Error GoTo gPopAgencyFromRvfBoxErr
        gBtrvErrorMsg ilRet, "gPopAgencyFromRvfBox (btrGetFirst):" & "Rvf.Btr", frm
        On Error GoTo 0
    End If
    Call btrExtSetBounds(hlRvf, llNoRec, -1, "UC", "RVFEXTPK", RVFEXTPK) 'Set extract limits (all records)
    tlIntTypeBuff.iType = ilAdfCode
    ilOffSet = gFieldOffset("Rvf", "RvfAdfCode")
    ilRet = btrExtAddLogicConst(hlRvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
    ilOffSet = gFieldOffset("Rvf", "RvfAgfCode")
    ilRet = btrExtAddField(hlRvf, ilOffSet, ilExtRecLen)  'Extract agency and advertiser fields (first two fields)
    On Error GoTo gPopAgencyFromRvfBoxErr
    gBtrvErrorMsg ilRet, "gPopAgencyFromRvfBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    'ilRet = btrExtGetNextExt(hlRvf)    'Extract record
    ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtRecLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo gPopAgencyFromRvfBoxErr
        gBtrvErrorMsg ilRet, "gPopAgencyFromRvfBox (btrExtGetNextExt):" & "Rvf.Btr", frm
        On Error GoTo 0
        'ilRet = btrExtGetFirst(hlRvf, tlRvfExt, ilExtRecLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtRecLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilFound = False
            'For ilLoop = LBound(tgCommAgf) To UBound(tgCommAgf) - 1 Step 1
            '    If tgCommAgf(ilLoop).iCode = tlRvfExt.iAgfCode Then
                ilLoop = gBinarySearchAgf(tlRvfExt.iAgfCode)
                If ilLoop <> -1 Then
                    ilFound = True
                    slName = Trim$(tgCommAgf(ilLoop).sName) & "/" & Trim$(tgCommAgf(ilLoop).sCityID) & "\" & Trim$(str$(tlRvfExt.iAgfCode))
            '        Exit For
                End If
            'Next ilLoop
            If ilFound Then
                gFindMatch slName, 0, lbcMster
                If gLastFound(lbcMster) < 0 Then
                    lbcMster.AddItem slName
                End If
            End If
            ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtRecLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtRecLen, llRecPos)
            Loop
        Loop
        For ilLoop = 0 To lbcMster.ListCount - 1 Step 1
            slNameCode = lbcMster.List(ilLoop)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            If ilRet <> CP_MSG_NONE Then
                gPopAgencyFromRvfBox = CP_MSG_PARSE
                Exit Function
            End If
            lbcLocal.AddItem Trim$(slName)  'Add ID to list box
        Next ilLoop
    End If
    ilRet = btrClose(hlRvf)
    On Error GoTo gPopAgencyFromRvfBoxErr
    gBtrvErrorMsg ilRet, "gPopAgencyFromRvfBox (btrReset):" & "Rvf.Btr", frm
    On Error GoTo 0
    btrDestroy hlRvf
    Exit Function
gPopAgencyFromRvfBoxErr:
    ilRet = btrClose(hlRvf)
    btrDestroy hlRvf
    gDbg_HandleError "CollectSubs: gPopAgencyFromRvfBox"
'    gPopAgencyFromRvfBox = CP_MSG_NOSHOW
'    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAgyCollectBox               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     agency names/City ID and        *
'*                     specified field                 *
'*                                                     *
'*******************************************************
Function gPopAgyCollectBox(frm As Form, slSort As String, lbcLocal As control, lbcMster As control) As Integer
'
'   ilRet = gPopAgyCollectBox (MainForm, slSort, lbcLocal, lbcCtrl)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       slSort(I)- Sort and field selection
'                   "A" = Alphabetic
'                   "B" = % over 90
'                   "C" = Amount owed
'                   "D" = Days since Paid
'                   "E" = Credit restriction
'                   "F" = Payment Rating
'                   "G" = Salesperson
'       lbcLocal (I)- List box to be populated from the master list box
'       lbcCtrl (I)- Master List box Control containing name and code #
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'Agf date/time stamp
    Dim slCAgfStamp As String    'Agf date/time stamp
    Dim slCAdfStamp As String    'Agf date/time stamp
    Dim hlAgf As Integer        'Agf handle
    Dim hlAdf As Integer        'Adf handle
    Dim ilAgfRecLen As Integer     'Record length
    Dim ilAdfRecLen As Integer     'Record length
    Dim tlagf As AGF
    Dim tlAdf As ADF
    Dim slName As String   'Name plus city
    Dim slListBox As String
    Dim slNameCode As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilList As Integer
    Dim ilBuildTgCAgf As Integer
    Dim ilBuildTgCAdf As Integer

    slStamp = gFileDateTime(sgDBPath & "Agf.Btr") & gFileDateTime(sgDBPath & "adf.Btr") & slSort
    If lbcMster.Tag <> "" Then
        If StrComp(slStamp, lbcMster.Tag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopAgyCollectBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
        End If
    End If
    ilBuildTgCAgf = True
    slCAgfStamp = gFileDateTime(sgDBPath & "Agf.Btr")
    If sgCAgfStamp <> "" Then
        If StrComp(slCAgfStamp, sgCAgfStamp, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                ilBuildTgCAgf = False
            End If
        End If
    End If
    ilBuildTgCAdf = True
    slCAdfStamp = gFileDateTime(sgDBPath & "Adf.Btr")
    If sgCAdfStamp <> "" Then
        If StrComp(slCAdfStamp, sgCAdfStamp, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                ilBuildTgCAdf = False
            End If
        End If
    End If
    gPopAgyCollectBox = CP_MSG_POPREQ
    lbcMster.Tag = slStamp
    lbcMster.Clear   'VB list box clear (list box used to retain code number so record can be found)
    lbcLocal.Clear
    ilRet = mObtainSalesperson()
    If ilBuildTgCAgf Then
        sgCAgfStamp = slCAgfStamp
        ReDim tgCAgf(0 To 0) As AGFADFCOLL
        hlAgf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopAgyCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyCollectBox (btrOpen):" & "Agf.Btr", frm
        On Error GoTo 0
        ilAgfRecLen = Len(tlagf) 'btrRecordLength(hlAgf)  'Get and save record length
        btrExtClear hlAgf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlAgf, tlagf, ilAgfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        'Agency and advertiser will use the same ext record structure
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            On Error GoTo gPopAgyCollectBoxErr
            gBtrvErrorMsg ilRet, "gPopAgyCollectBox (btrGetFirst):" & "Agf.Btr", frm
            On Error GoTo 0
            ilRet = mCreateCollectSortRec(hlAgf, "AGF", tgCAgf())
        End If
        ilRet = btrClose(hlAgf)
        On Error GoTo gPopAgyCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyCollectBox (btrReset):" & "Agf.Btr", frm
        On Error GoTo 0
        btrDestroy (hlAgf)
    End If
    If ilBuildTgCAdf Then
        mNonPayeeAdfCode
        sgCAdfStamp = slCAdfStamp
        ReDim tgCAdf(0 To 0) As AGFADFCOLL
        hlAdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopAgyCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyCollectBox (btrOpen):" & "Adf.Btr", frm
        On Error GoTo 0
        ilAdfRecLen = Len(tlAdf) 'btrRecordLength(hlAgf)  'Get and save record length
        btrExtClear hlAdf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlAdf, tlAdf, ilAdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            On Error GoTo gPopAgyCollectBoxErr
            gBtrvErrorMsg ilRet, "gPopAgyCollectBox (btrGetFirst):" & "Adf.Btr", frm
            On Error GoTo 0
            ilRet = mCreateCollectSortRec(hlAdf, "ADFDIR", tgCAdf())
        End If
        ilRet = btrClose(hlAdf)
        On Error GoTo gPopAgyCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyCollectBox (btrReset):" & "Adf.Btr", frm
        On Error GoTo 0
        btrDestroy (hlAdf)
    End If
    For ilLoop = LBound(tgCAgf) To UBound(tgCAgf) - 1 Step 1
        slListBox = mMakeCollectSort("AGF", slSort, tgCAgf(ilLoop))
        slListBox = slListBox & "|" & "G" & Trim$(str$(ilLoop))
        slListBox = slListBox & "\" & Trim$(str$(tgCAgf(ilLoop).iCode))
        lbcMster.AddItem slListBox    'Add ID (retain matching sorted order) and Code number to list box
    Next ilLoop
    For ilLoop = LBound(tgCAdf) To UBound(tgCAdf) - 1 Step 1
        slListBox = mMakeCollectSort("ADF", slSort, tgCAdf(ilLoop))
        slListBox = slListBox & "|" & "D" & Trim$(str$(ilLoop))
        slListBox = slListBox & "\" & Trim$(str$(tgCAdf(ilLoop).iCode))
        lbcMster.AddItem slListBox    'Add ID (retain matching sorted order) and Code number to list box
    Next ilLoop
    For ilList = 0 To lbcMster.ListCount - 1 Step 1
        slNameCode = lbcMster.List(ilList)
        ilRet = gParseItem(slNameCode, 2, "|", slName)  'Obtain Index and code number
        If ilRet <> CP_MSG_NONE Then
            gPopAgyCollectBox = CP_MSG_PARSE
            Exit Function
        End If
        ilRet = gParseItem(slName, 1, "\", slStr)       'Obtain index
        If ilRet <> CP_MSG_NONE Then
            gPopAgyCollectBox = CP_MSG_PARSE
            Exit Function
        End If
        ilLoop = Val(right$(slStr, Len(slStr) - 1))
        If Asc(slStr) = Asc("G") Then
            slListBox = mMakeCollectListImage("AGF", slSort, tgCAgf(ilLoop))
        Else
            slListBox = mMakeCollectListImage("ADF", slSort, tgCAdf(ilLoop))
        End If
        lbcLocal.AddItem slListBox  'Add ID to list box
    Next ilList
    Exit Function
gPopAgyCollectBoxErr:
    ilRet = btrClose(hlAgf)
    ilRet = btrClose(hlAdf)
    btrDestroy hlAgf
    btrDestroy hlAdf
    gDbg_HandleError "CollectSubs: gPopAgyCollectBox"
'    gPopAgyCollectBox = CP_MSG_NOSHOW
'    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAgyCommentCollectBox        *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     comments and specified field    *
'*                                                     *
'*******************************************************
Function gPopAgyCommentCollectBox(frm As Form, ilAgyCode As Integer, lbcMster As control) As Integer
'
'   ilRet = gPopAgyCommentCollectBox (MainForm, ilAgyCode, lbcCtrl)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAgyCode(I)- Agency code
'       lbcCtrl (I)- Master List box Control containing date and time, advt name and comment rec #
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'Cdf date/time stamp
    Dim hlAdf As Integer        'Cdf handle
    Dim ilAdfRecLen As Integer     'Record length
    Dim tlAdf As ADF
    Dim tlAdfSrchKey As INTKEY0 'ADF key record image
    Dim hlCdf As Integer        'Cdf handle
    Dim ilCdfRecLen As Integer     'Record length
    Dim tlCdf As CDF
    Dim llNoRec As Long         'Number of records in Sof
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim tlCdfExt As CDFAGYEXT
    Dim slName As String
    Dim slDate As String
    Dim slTime As String
    Dim llTime As Long
    Dim slNameSort As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    slStamp = gFileDateTime(sgDBPath & "Cdf.Btr") & "Agy" & Trim$(str$(ilAgyCode))
    If lbcMster.Tag <> "" Then
        If StrComp(slStamp, lbcMster.Tag, 1) = 0 Then
            gPopAgyCommentCollectBox = CP_MSG_NOPOPREQ
            Exit Function
        End If
    End If
    gPopAgyCommentCollectBox = CP_MSG_POPREQ
    lbcMster.Tag = slStamp
    lbcMster.Clear   'VB list box clear (list box used to retain code number so record can be found)
    hlAdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopAgyCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyCommentCollectBox (btrOpen):" & "Adf.Btr", frm
    On Error GoTo 0
    ilAdfRecLen = Len(tlAdf) 'btrRecordLength(hlAdf)  'Get and save record length
    hlCdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlCdf, "", sgDBPath & "Cdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopAgyCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyCommentCollectBox (btrOpen):" & "Cdf.Btr", frm
    On Error GoTo 0
    ilCdfRecLen = Len(tlCdf) 'btrRecordLength(hlAgf)  'Get and save record length
    ilExtLen = Len(tlCdfExt)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlCdf) 'Obtain number of records
    btrExtClear hlCdf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlCdf, tlCdf, ilCdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlAdf)
        On Error GoTo gPopAgyCommentCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyCommentCollectBox (btrReset):" & "Adf.Btr", frm
        On Error GoTo 0
        btrDestroy hlAdf
        ilRet = btrClose(hlCdf)
        On Error GoTo gPopAgyCommentCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyCommentCollectBox (btrReset):" & "Cdf.Btr", frm
        On Error GoTo 0
        btrDestroy hlCdf
        Exit Function
    Else
        On Error GoTo gPopAgyCommentCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyCommentCollectBox (btrReset):" & "Cdf.Btr", frm
        On Error GoTo 0
    End If
    Call btrExtSetBounds(hlCdf, llNoRec, -1, "UC", "CDFAGYEXTPK", CDFAGYEXTPK) 'Set extract limits (all records)
    tlIntTypeBuff.iType = ilAgyCode
    ilOffSet = gFieldOffset("Cdf", "CdfAgfCode")
    ilRet = btrExtAddLogicConst(hlCdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
    ilOffSet = gFieldOffset("Cdf", "CdfAdfCode")
    ilRet = btrExtAddField(hlCdf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAgyCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Cdf", "CdfEntryDate")
    ilRet = btrExtAddField(hlCdf, ilOffSet, 4) 'Extract entry date field
    On Error GoTo gPopAgyCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Cdf", "CdfEntryTime")
    ilRet = btrExtAddField(hlCdf, ilOffSet, 4) 'Extract entry date field
    On Error GoTo gPopAgyCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
    On Error GoTo 0
    'ilRet = btrExtGetNextExt(hlCdf)    'Extract record
    ilRet = btrExtGetNext(hlCdf, tlCdfExt, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo gPopAgyCommentCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyCommentCollectBox (btrExtGetNextExt):" & "Cdf.Btr", frm
        On Error GoTo 0
        tlAdf.iCode = 0
        tlAdf.sName = ""
        ilExtLen = Len(tlCdfExt)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlCdf, tlCdfExt, ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlCdf, tlCdfExt, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            gUnpackDateForSort tlCdfExt.iDateEntrd(0), tlCdfExt.iDateEntrd(1), slDate
            'gUnpackDateLong tlCdfExt.iDateEntrd(0), tlCdfExt.iDateEntrd(1), llDate
            'llDate = 99999 - llDate
            'slDate = Trim$(Str$(llDate))
            'Do While Len(slDate) < 5
            '    slDate = "0" & slDate
            'Loop
            gUnpackTime tlCdfExt.iTimeEntrd(0), tlCdfExt.iTimeEntrd(1), "A", "1", slTime
            llTime = CLng(gTimeToCurrency(slTime, False))
            slTime = Trim$(str$(llTime))
            Do While Len(slTime) < 7
                slTime = "0" & slTime
            Loop
            If tlCdfExt.iAdfCode > 0 Then
                If tlCdfExt.iAdfCode <> tlAdf.iCode Then
                    tlAdfSrchKey.iCode = tlCdfExt.iAdfCode
                    ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    On Error GoTo gPopAgyCommentCollectBoxErr
                    gBtrvErrorMsg ilRet, "gPopAgyCommentCollectBox (btrGetEqual):" & "Adf.Btr", frm
                    On Error GoTo 0
                End If
                If (tlAdf.sBillAgyDir = "D") And (Trim$(tlAdf.sName) <> "") Then
                    slName = Trim$(tlAdf.sName) & ", " & Trim$(tlAdf.sAddrID)
                Else
                    slName = Trim$(tlAdf.sName)
                End If
            Else
                tlAdf.iCode = 0
                tlAdf.sName = ""
                slName = " "
            End If
            slNameSort = slDate & "\" & slTime & "\" & slName & "\" & Trim$(str$(llRecPos))
            lbcMster.AddItem slNameSort    'Add ID (retain matching sorted order) and Code number to list box
            ilRet = btrExtGetNext(hlCdf, tlCdfExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlCdf, tlCdfExt, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlAdf)
    On Error GoTo gPopAgyCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyCommentCollectBox (btrReset):" & "Adf.Btr", frm
    On Error GoTo 0
    btrDestroy hlAdf
    ilRet = btrClose(hlCdf)
    On Error GoTo gPopAgyCommentCollectBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyCommentCollectBox (btrReset):" & "Cdf.Btr", frm
    On Error GoTo 0
    btrDestroy hlCdf
    Exit Function
gPopAgyCommentCollectBoxErr:
    ilRet = btrClose(hlAdf)
    ilRet = btrClose(hlCdf)
    btrDestroy hlAdf
    btrDestroy hlCdf
    gPopAgyCommentCollectBox = CP_MSG_NOSHOW
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gPopAgyTransactionBox           *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     transaction and specified field *
'*                                                     *
'*******************************************************
'Function gPopAgyTransactionBox (Frm As Form, hlRvf As Integer, ilAgyCode As Integer, ilCashTrade As Integer, lbcMster As Control) As Integer
Function gPopAgyTransactionBox(frm As Form, hlRvf As Integer, hlChf As Integer, ilAgyCode As Integer, ilCashTrade As Integer, tlSortCode() As SORTCODE, slSortCodeTag As String, ilHistory As Integer) As Integer
'
'   ilRet = gPopAgyTransactionBox (MainForm, hlRvf, ilAgyCode, ilCashTrade, lbcCtrl)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       hlRvf(I)- RVF or PHF handle
'       ilAgyCode(I)- Agency code
'       ilCashTrade(I)- 0=Cash; 1=Trade; 2=Merchandising; 3=Promotion; 4=Installment Revenue
'       lbcCtrl (I)- Master List box Control containing date, advt name and comment rec #
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'Rvf date/time stamp
    Dim hlAdf As Integer        'Adf handle
    Dim ilAdfRecLen As Integer     'Record length
    Dim tlAdf As ADF
    Dim tlAdfSrchKey As INTKEY0 'ADF key record image
    Dim ilRvfRecLen As Integer     'Record length
    Dim tlRvf As RVF
    Dim llNoRec As Long         'Number of records in Sof
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim tlRvfExt As RVFTRANEXT
    Dim slName As String
    Dim slDate As String
    Dim slInvNo As String
    Dim ilFound As Integer
    Dim slNameSort As String
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim hlMnf As Integer
    Dim tlMnf As MNF
    Dim ilMnfRecLen As Integer
    Dim tlMnfSrchKey As INTKEY0
    Dim slAirVehicle As String
    Dim slBillVehicle As String
    Dim slGpSort As String
    Dim slPkLineNo As String
    Dim slVehSort As String
    Dim slNTRSort As String     '7-7-06
    Dim slParticipantSort As String     '2-7-07
    Dim ilIndex As Integer
    Dim llLen As Long
    Dim ilSortVehbyGroup As Integer 'True=Sort vehicle by groups; False=Sort by name only
    Dim llSortCode As Long
    Dim ilVef As Integer
    Dim tlChf As CHF            'SBF record image
    Dim tlChfSrchKey1 As CHFKEY1 'SBF key record image
    Dim ilChfRecLen As Integer  'SBF record length
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilAdfIndex As Integer

    ilChfRecLen = Len(tlChf)
    ilSortVehbyGroup = False
    llLen = 0
    slStamp = gFileDateTime(sgDBPath & "Rvf.Btr") & "Agy" & Trim$(str$(ilAgyCode)) & Trim$(str$(ilCashTrade))
    'If lbcMster.Tag <> "" Then
    '    If StrComp(slStamp, lbcMster.Tag, 1) = 0 Then
    If slSortCodeTag <> "" Then
        If slStamp = slSortCodeTag Then
            gPopAgyTransactionBox = CP_MSG_NOPOPREQ
            Exit Function
        End If
    End If
    gPopAgyTransactionBox = CP_MSG_POPREQ
    'lbcMster.Tag = slStamp
    'lbcMster.Clear   'VB list box clear (list box used to retain code number so record can be found)
    llSortCode = 0
    ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
    slSortCodeTag = slStamp
    ilRet = gObtainVef()
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
    hlAdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrOpen):" & "Adf.Btr", frm
    On Error GoTo 0
    ilAdfRecLen = Len(tlAdf) 'btrRecordLength(hlAdf)  'Get and save record length
    ilRvfRecLen = Len(tlRvf) 'btrRecordLength(hlAgf)  'Get and save record length
    ilExtLen = Len(tlRvfExt)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlRvf) 'Obtain number of records
    btrExtClear hlRvf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlRvf, tlRvf, ilRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlAdf)
        On Error GoTo gPopAgyTransactionBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrReset):" & "Adf.Btr", frm
        On Error GoTo 0
        btrDestroy hlAdf
        btrDestroy hlMnf
        Exit Function
    Else
        On Error GoTo gPopAgyTransactionBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrReset):" & "Rvf.Btr", frm
        On Error GoTo 0
    End If
    Call btrExtSetBounds(hlRvf, llNoRec, -1, "UC", "RVFTRANEXTPK", RVFTRANEXTPK) 'Set extract limits (all records)
    If ilHistory Then
         gPackDate sgDateRangeStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
         ilOffSet = gFieldOffset("rvf", "rvfTranDate")
         ilRet = btrExtAddLogicConst(hlRvf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
         gPackDate sgDateRangeEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
         ilOffSet = gFieldOffset("rvf", "rvfTranDate")
         ilRet = btrExtAddLogicConst(hlRvf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
    End If
    tlIntTypeBuff.iType = ilAgyCode
    ilOffSet = gFieldOffset("Rvf", "RvfAgfCode")
    ilRet = btrExtAddLogicConst(hlRvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
    ilOffSet = gFieldOffset("Rvf", "RvfAgfCode")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfAdfCode")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfInvNo")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 4) 'Extract entry date field
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfAirVefCode")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfTranDate")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 4) 'Extract entry date field
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfTranType")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2) 'Extract transaction type
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfCashTrade")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 1) 'Extract Cash/Trade
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfBillVefCode")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0
    ilOffSet = gFieldOffset("Rvf", "RvfPkLineNo")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0

    ilOffSet = gFieldOffset("Rvf", "RvfMnfItem")        '7-7-06
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0

    ilOffSet = gFieldOffset("Rvf", "RvfMnfGroup")        '7-7-06
    ilRet = btrExtAddField(hlRvf, ilOffSet, 2)  'Extract First Name field
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0

    ilOffSet = gFieldOffset("Rvf", "RvfType")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 1) 'Extract Cash/Trade
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0

    ilOffSet = gFieldOffset("Rvf", "RvfCntrNo")
    ilRet = btrExtAddField(hlRvf, ilOffSet, 4) 'Extract entry date field
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtAddField):" & "Rvf.Btr", frm
    On Error GoTo 0

    'ilRet = btrExtGetNextExt(hlRvf)    'Extract record
    ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        On Error GoTo gPopAgyTransactionBoxErr
        gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrExtGetNextExt):" & "Rvf.Btr", frm
        On Error GoTo 0
        tlAdf.iCode = 0
        tlAdf.sName = ""
        ilExtLen = Len(tlRvfExt)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlRvf, tlRvfExt, ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilFound = False
            If ilCashTrade = 0 Then
                If (tlRvfExt.sCashTrade = "C") Then
                    ilFound = True
                End If
            ElseIf ilCashTrade = 1 Then
                If (tlRvfExt.sCashTrade = "T") Then
                    ilFound = True
                End If
            ElseIf ilCashTrade = 2 Then
                If (tlRvfExt.sCashTrade = "M") Then
                    ilFound = True
                End If
            ElseIf ilCashTrade = 3 Then
                If (tlRvfExt.sCashTrade = "P") Then
                    ilFound = True
                End If
            ElseIf ilCashTrade = 4 Then
                If (tlRvfExt.sCashTrade = "C") And (tlRvfExt.sType = "A") And ((tlRvfExt.sTranType = "HI") Or (tlRvfExt.sTranType = "AN")) Then
                    tlChfSrchKey1.lCntrNo = tlRvfExt.lCntrNo
                    tlChfSrchKey1.iCntRevNo = 32000
                    tlChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hlChf, tlChf, ilChfRecLen, tlChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                    'If contract missing, assume that transaction entered via Backlog, 5/8/04
                    Do While (ilRet = BTRV_ERR_NONE) And (tlChf.lCntrNo = tlRvfExt.lCntrNo)
                        If tlChf.sSchStatus = "F" Then
                            Exit Do
                        End If
                        ilRet = btrGetNext(hlChf, tlChf, ilChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (tlChf.lCntrNo = tlRvfExt.lCntrNo) And (tlChf.sSchStatus = "F") Then
                        If tlChf.sInstallDefined = "Y" Then
                            ilFound = True
                        End If
                    End If
                End If
            End If
            If ilFound Then
                ilFound = False
                If (tgUrf(0).iCode = 1) Or (tgUrf(0).iCode = 2) Or (tgUrf(0).iMnfHubCode <= 0) Or ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) <> USINGHUB) Then
                    ilFound = True
                Else
                    ilVef = gBinarySearchVef(tlRvfExt.iBillVefCode)
                    If ilVef <> -1 Then
                        If ((tgUrf(0).iMnfHubCode = tgMVef(ilVef).iMnfHubCode) And ((Asc(tgSpf.sUsingFeatures3) And USINGHUB) = USINGHUB)) Then
                            ilFound = True
                        End If
                    End If
                End If
            End If
            If ilFound Then
                gUnpackDate tlRvfExt.iTranDate(0), tlRvfExt.iTranDate(1), slDate
                slDate = Trim$(str$(gDateValue(slDate)))
                Do While Len(slDate) < 6
                    slDate = "0" & slDate
                Loop
                slInvNo = Trim$(str$(tlRvfExt.lInvNo))
                Do While Len(slInvNo) < 12
                    slInvNo = "0" & slInvNo
                Loop
                If tlRvfExt.iAdfCode > 0 Then
                    If tlRvfExt.iAdfCode <> tlAdf.iCode Then
                        tlAdfSrchKey.iCode = tlRvfExt.iAdfCode
                        'ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        ilAdfIndex = gBinarySearchAdf(tlRvfExt.iAdfCode)   '5-11-04
                        If ilAdfIndex = -1 Then
                            On Error GoTo gPopAgyTransactionBoxErr
                            gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrGetEqual):" & "Adf.Btr", frm
                            On Error GoTo 0
                        End If
                    End If
                    'slName = Trim$(tlAdf.sName)
                    If (Trim$(tgCommAdf(ilAdfIndex).sAddrID) <> "") And (tgCommAdf(ilAdfIndex).sBillAgyDir = "D") Then
                        slName = Trim$(tgCommAdf(ilAdfIndex).sName) & ", " & Trim$(tgCommAdf(ilAdfIndex).sAddrID)
                    Else
                        slName = Trim$(tgCommAdf(ilAdfIndex).sName)        '5-11-04
                    End If
                Else
                    tlAdf.iCode = 0
                    tlAdf.sName = ""
                    slName = " "
                End If
                '7-7-06 put the ntr type into the sort so they will sort together
                slNTRSort = Trim$(str$(tlRvfExt.iMnfItem))
                Do While Len(slNTRSort) < 5
                    slNTRSort = "0" & slNTRSort
                Loop
                slParticipantSort = Trim$(str$(tlRvfExt.iMnfGroup))
                Do While Len(slParticipantSort) < 5
                    slParticipantSort = "0" & slParticipantSort
                Loop
                If tlRvfExt.iAirVefCode > 0 Then
                    'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If tgMVef(ilIndex).iCode = tlRvfExt.iAirVefCode Then
                        ilIndex = gBinarySearchVef(tlRvfExt.iAirVefCode)
                        If ilIndex <> -1 Then
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
                            slAirVehicle = slGpSort & slVehSort & tgMVef(ilIndex).sName
                        End If
                    'Next ilIndex
                Else
                    slGpSort = "999"
                    slVehSort = "999"
                    slAirVehicle = slGpSort & slVehSort & "                    "
                End If
                If tlRvfExt.iAirVefCode = tlRvfExt.iBillVefCode Then
                    slBillVehicle = slAirVehicle
                ElseIf tlRvfExt.iBillVefCode > 0 Then
                    'For ilIndex = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If tgMVef(ilIndex).iCode = tlRvfExt.iBillVefCode Then
                        ilIndex = gBinarySearchVef(tlRvfExt.iBillVefCode)
                        If ilIndex <> -1 Then
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
                            slBillVehicle = slGpSort & slVehSort & tgMVef(ilIndex).sName
                        End If
                    'Next ilIndex
                Else
                    slGpSort = "999"
                    slVehSort = "999"
                    slBillVehicle = slGpSort & slVehSort & "                    "
                End If
                slPkLineNo = Trim$(str$(tlRvfExt.iPkLineNo))
                Do While Len(slPkLineNo) < 4
                    slPkLineNo = "0" & slPkLineNo
                Loop

                slNameSort = slName & "|" & slInvNo & "|" & slInvNo & slBillVehicle & slPkLineNo & slAirVehicle & slParticipantSort & slNTRSort & slDate & "\" & Trim$(str$(llRecPos))
                'If Not gOkAddStrToListBox(slNameSort, llLen, True) Then
                '    Exit Do
                'End If
                'lbcMster.AddItem slNameSort    'Add ID (retain matching sorted order) and Code number to list box
                tlSortCode(llSortCode).sKey = slNameSort
                llSortCode = llSortCode + 1
                ReDim Preserve tlSortCode(0 To llSortCode) As SORTCODE
            End If
            ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlRvf, tlRvfExt, ilExtLen, llRecPos)
            Loop
        Loop
    End If
    If UBound(tlSortCode) - 1 > 0 Then
        ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
    End If
    ilRet = btrClose(hlAdf)
    On Error GoTo gPopAgyTransactionBoxErr
    gBtrvErrorMsg ilRet, "gPopAgyTransactionBox (btrReset):" & "Adf.Btr", frm
    On Error GoTo 0
    btrDestroy hlAdf
    btrDestroy hlMnf
    Exit Function
gPopAgyTransactionBoxErr:
    ilRet = btrClose(hlAdf)
    btrDestroy hlAdf
    btrDestroy hlMnf
    'btrDestroy hlRvf
    gPopAgyTransactionBox = CP_MSG_NOSHOW
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gPopCommentCollectBox           *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     comments and specified field    *
'*                                                     *
'*******************************************************
Function gPopCommentCollectBox(frm As Form, slType As String, slSort As String, slOldestDate As String, lbcLocal As control, lbcMster As control) As Integer
'
'   ilRet = gPopCommentCollectBox (MainForm, slSort, slOldestDate, lbcLocal, lbcCtrl)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       slType (I)-"AGF" = Agency comments; "ADF" = Advertiser comments
'       slSort(I)- Sort and field selection
'                   "A" = Action date for current user
'                   "B" = Action date for any user
'                   "C" = Entry date for current user
'                   "D" = Entry date for any user
'       slOldestDate(I)- oldest date for selecting comments
'       lbcLocal (I)- List box to be populated from the master list box
'       lbcCtrl (I)- Master List box Control containing name and code #
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'Cdf date/time stamp
    Dim slCCdfStamp As String    'Cdf date/time stamp
    Dim hlCdf As Integer        'Cdf handle
    Dim ilCdfRecLen As Integer     'Record length
    Dim tlCdf As CDF
    Dim slName As String   'Name plus city
    Dim slListBox As String
    Dim slNameCode As String
    Dim slStr As String
    Dim llNoRec As Long         'Number of records in Sof
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilList As Integer
    Dim ilBuildTgCCdf As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer
    Dim slDate As String
    Dim llOldestDate As Long
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim ilMatchFound As Integer
    Dim slDate1 As String
    Dim slDate2 As String
    Dim slTime1 As String
    Dim slTime2 As String
    Dim slTime As String
    Dim llTime As Long

    slStamp = gFileDateTime(sgDBPath & "Cdf.Btr") & slSort & slOldestDate
    If lbcMster.Tag <> "" Then
        If StrComp(slStamp, lbcMster.Tag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopCommentCollectBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
        End If
    End If
    ilBuildTgCCdf = True
    slCCdfStamp = gFileDateTime(sgDBPath & "Cdf.Btr") & slOldestDate
    If sgCCdfStamp <> "" Then
        If StrComp(slCCdfStamp, sgCCdfStamp, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                ilBuildTgCCdf = False
            End If
        End If
    End If
    gPopCommentCollectBox = CP_MSG_POPREQ
    If slOldestDate = "" Then
        llOldestDate = 0
    Else
        llOldestDate = gDateValue(slOldestDate)
    End If
    lbcMster.Tag = slStamp
    lbcMster.Clear   'VB list box clear (list box used to retain code number so record can be found)
    lbcLocal.Clear
    ilRet = gObtainAgency()
    ilRet = gObtainAdvt()
    ilRet = mObtainUser()
    If ilBuildTgCCdf Then
        sgCCdfStamp = slCCdfStamp
        ReDim tgCCdf(0 To 0) As CDFLIST
        hlCdf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlCdf, "", sgDBPath & "Cdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopCommentCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopCommentCollectBox (btrOpen):" & "Cdf.Btr", frm
        On Error GoTo 0
        ilCdfRecLen = Len(tlCdf) 'btrRecordLength(hlAgf)  'Get and save record length
        ilUpperBound = UBound(tgCCdf)
        ilExtLen = Len(tgCCdf(ilUpperBound).tCdfExt)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlCdf) 'Obtain number of records
        btrExtClear hlCdf   'Clear any previous extend operation
        ilRet = btrGetFirst(hlCdf, tlCdf, ilCdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_END_OF_FILE Then
            On Error GoTo gPopCommentCollectBoxErr
            gBtrvErrorMsg ilRet, "gPopCommentCollectBox (btrGetFirst):" & "Cdf.Btr", frm
            On Error GoTo 0
            Call btrExtSetBounds(hlCdf, llNoRec, -1, "UC", "CDFEXTPK", CDFEXTPK) 'Set extract limits (all records)
            If slOldestDate <> "" Then
                If (slSort = "A") Or (slSort = "B") Then
                    gPackDate slOldestDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Cdf", "CdfActionDate")
                    ilRet = btrExtAddLogicConst(hlCdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                Else
                    gPackDate slOldestDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Cdf", "CdfEntryDate")
                    ilRet = btrExtAddLogicConst(hlCdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                End If
            End If
            ilOffSet = gFieldOffset("Cdf", "CdfAgfCode")
            ilRet = btrExtAddField(hlCdf, ilOffSet, 2)  'Extract iCode field
            On Error GoTo gPopCommentCollectBoxErr
            gBtrvErrorMsg ilRet, "gPopCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
            On Error GoTo 0
            ilOffSet = gFieldOffset("Cdf", "CdfAdfCode")
            ilRet = btrExtAddField(hlCdf, ilOffSet, 2)  'Extract First Name field
            On Error GoTo gPopCommentCollectBoxErr
            gBtrvErrorMsg ilRet, "gPopCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
            On Error GoTo 0
            ilOffSet = gFieldOffset("Cdf", "CdfActionDate")
            ilRet = btrExtAddField(hlCdf, ilOffSet, 4) 'Extract Last Name field
            On Error GoTo gPopCommentCollectBoxErr
            gBtrvErrorMsg ilRet, "gPopCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
            On Error GoTo 0
            ilOffSet = gFieldOffset("Cdf", "CdfEntryDate")
            ilRet = btrExtAddField(hlCdf, ilOffSet, 4) 'Extract Entry date field
            On Error GoTo gPopCommentCollectBoxErr
            gBtrvErrorMsg ilRet, "gPopCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
            On Error GoTo 0
            ilOffSet = gFieldOffset("Cdf", "CdfEntryTime")
            ilRet = btrExtAddField(hlCdf, ilOffSet, 4) 'Extract Entry time field
            On Error GoTo gPopCommentCollectBoxErr
            gBtrvErrorMsg ilRet, "gPopCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
            On Error GoTo 0
            ilOffSet = gFieldOffset("Cdf", "CdfUrfCode")
            ilRet = btrExtAddField(hlCdf, ilOffSet, 2) 'Extract Last Name field
            On Error GoTo gPopCommentCollectBoxErr
            gBtrvErrorMsg ilRet, "gPopCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
            On Error GoTo 0
            'ilRet = btrExtGetNextExt(hlCdf)    'Extract record
            ilUpperBound = UBound(tgCCdf)
            ilRet = btrExtGetNext(hlCdf, tgCCdf(ilUpperBound).tCdfExt, ilExtLen, tgCCdf(ilUpperBound).lRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                On Error GoTo gPopCommentCollectBoxErr
                gBtrvErrorMsg ilRet, "gPopCommentCollectBox (btrExtAddField):" & "Cdf.Btr", frm
                On Error GoTo 0
                ilUpperBound = UBound(tgCCdf)
                ilExtLen = Len(tgCCdf(ilUpperBound).tCdfExt)  'Extract operation record size
                'ilRet = btrExtGetFirst(hlCdf, tgCCdf(ilUpperBound).tCdfExt, ilExtLen, tgCCdf(ilUpperBound).lRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlCdf, tgCCdf(ilUpperBound).tCdfExt, ilExtLen, tgCCdf(ilUpperBound).lRecPos)
                Loop
                Do While ilRet = BTRV_ERR_NONE
                    'Only retain the most recent comment
                    ilMatchFound = False
                    For ilLoop = LBound(tgCCdf) To UBound(tgCCdf) - 1 Step 1
                        If (tgCCdf(ilLoop).tCdfExt.iAgfCode = tgCCdf(ilUpperBound).tCdfExt.iAgfCode) And (tgCCdf(ilLoop).tCdfExt.iAdfCode = tgCCdf(ilUpperBound).tCdfExt.iAdfCode) Then
                            ilMatchFound = True
                            If (slSort = "A") Or (slSort = "B") Then
                                gUnpackDate tgCCdf(ilLoop).tCdfExt.iActionDate(0), tgCCdf(ilLoop).tCdfExt.iActionDate(1), slDate1
                                gUnpackDate tgCCdf(ilUpperBound).tCdfExt.iActionDate(0), tgCCdf(ilUpperBound).tCdfExt.iActionDate(1), slDate2
                                slTime1 = "12:00AM"
                                slTime2 = "12:00AM"
                            Else    'Entry date
                                gUnpackDate tgCCdf(ilLoop).tCdfExt.iDateEntrd(0), tgCCdf(ilLoop).tCdfExt.iDateEntrd(1), slDate1
                                gUnpackDate tgCCdf(ilUpperBound).tCdfExt.iDateEntrd(0), tgCCdf(ilUpperBound).tCdfExt.iDateEntrd(1), slDate2
                                gUnpackTime tgCCdf(ilLoop).tCdfExt.iTimeEntrd(0), tgCCdf(ilLoop).tCdfExt.iTimeEntrd(1), "A", "1", slTime1
                                gUnpackTime tgCCdf(ilUpperBound).tCdfExt.iTimeEntrd(0), tgCCdf(ilUpperBound).tCdfExt.iTimeEntrd(1), "A", "1", slTime2
                            End If
                            If gDateValue(slDate2) > gDateValue(slDate1) Then
                                tgCCdf(ilLoop).tCdfExt = tgCCdf(ilUpperBound).tCdfExt
                            ElseIf (gDateValue(slDate2) = gDateValue(slDate1)) And (gTimeToCurrency(slTime2, False) > gTimeToCurrency(slTime1, False)) Then
                                tgCCdf(ilLoop).tCdfExt = tgCCdf(ilUpperBound).tCdfExt
                            End If
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilMatchFound Then
                        ilUpperBound = ilUpperBound + 1
                        ReDim Preserve tgCCdf(0 To ilUpperBound) As CDFLIST
                    End If
                    ilRet = btrExtGetNext(hlCdf, tgCCdf(ilUpperBound).tCdfExt, ilExtLen, tgCCdf(ilUpperBound).lRecPos)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlCdf, tgCCdf(ilUpperBound).tCdfExt, ilExtLen, tgCCdf(ilUpperBound).lRecPos)
                    Loop
                Loop
            End If
        End If
        ilRet = btrClose(hlCdf)
        On Error GoTo gPopCommentCollectBoxErr
        gBtrvErrorMsg ilRet, "gPopCommentCollectBox (btrReset):" & "Cdf.Btr", frm
        On Error GoTo 0
        btrDestroy (hlCdf)
    End If
    For ilLoop = LBound(tgCCdf) To UBound(tgCCdf) - 1 Step 1
        ilFound = False
        slListBox = ""
        If (slSort = "A") Or (slSort = "C") Then
            For ilIndex = LBound(tgUrf) To UBound(tgUrf) Step 1
                If tgUrf(ilIndex).iCode = tgCCdf(ilLoop).tCdfExt.iUrfCode Then
                    ilFound = True
'                    slListBox = tgUrf(ilIndex).sName    'Don't trim as dates must line up
                    Exit For
                End If
            Next ilIndex
        Else
            For ilIndex = LBound(tgCommUrf) To UBound(tgCommUrf) - 1 Step 1
                If tgCommUrf(ilIndex).iCode = tgCCdf(ilLoop).tCdfExt.iUrfCode Then
                    ilFound = True
'                    slListBox = tgCommUrf(ilIndex).sName    'Don't trim as dates must line up
                    Exit For
                End If
            Next ilIndex
        End If
        If ilFound Then
            If (slSort = "A") Or (slSort = "B") Then
                gUnpackDate tgCCdf(ilLoop).tCdfExt.iActionDate(0), tgCCdf(ilLoop).tCdfExt.iActionDate(1), slDate
                If gDateValue(slDate) < llOldestDate Then
                    ilFound = False
                Else
                    gUnpackDateForSort tgCCdf(ilLoop).tCdfExt.iActionDate(0), tgCCdf(ilLoop).tCdfExt.iActionDate(1), slDate
                    slListBox = slListBox & slDate
                End If
            Else
                gUnpackDate tgCCdf(ilLoop).tCdfExt.iDateEntrd(0), tgCCdf(ilLoop).tCdfExt.iDateEntrd(1), slDate
                If gDateValue(slDate) < llOldestDate Then
                    ilFound = False
                Else
                    gUnpackDateForSort tgCCdf(ilLoop).tCdfExt.iDateEntrd(0), tgCCdf(ilLoop).tCdfExt.iDateEntrd(1), slDate
                    slListBox = slListBox & slDate
                    gUnpackTime tgCCdf(ilLoop).tCdfExt.iTimeEntrd(0), tgCCdf(ilLoop).tCdfExt.iTimeEntrd(1), "A", "1", slTime
                    llTime = CLng(gTimeToCurrency(slTime, False))
                    slTime = Trim$(str$(llTime))
                    Do While Len(slTime) < 7
                        slTime = "0" & slTime
                    Loop
                End If
            End If
        End If
        If ilFound Then
            ilFound = False
            If slType = "AGF" Then
                If tgCCdf(ilLoop).tCdfExt.iAgfCode <> 0 Then
                    'For ilIndex = LBound(tgCommAgf) To UBound(tgCommAgf) - 1 Step 1
                    '    If tgCommAgf(ilIndex).iCode = tgCCdf(ilLoop).tCdfExt.iAgfCode Then
                        ilIndex = gBinarySearchAgf(tgCCdf(ilLoop).tCdfExt.iAgfCode)
                        If ilIndex <> -1 Then
                            ilFound = True
                            slListBox = slListBox & " " & tgCommAgf(ilIndex).sName & ", " & tgCommAgf(ilIndex).sCityID
                    '        Exit For
                        End If
                    'Next ilIndex
                Else
                    'For ilIndex = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                    '    If tgCommAdf(ilIndex).iCode = tgCCdf(ilLoop).tCdfExt.iAdfCode Then
                        ilIndex = gBinarySearchAdf(tgCCdf(ilLoop).tCdfExt.iAdfCode)
                        If ilIndex <> -1 Then
                            If tgCommAdf(ilIndex).sBillAgyDir = "D" Then
                                ilFound = True
                                If (Trim$(tgCommAdf(ilIndex).sAddrID) <> "") Then
                                    slListBox = slListBox & " " & Trim$(tgCommAdf(ilIndex).sName) & ", " & Trim$(tgCommAdf(ilIndex).sAddrID) & "/Direct"
                                Else
                                    slListBox = slListBox & " " & Trim$(tgCommAdf(ilIndex).sName) & "/Direct"
                                End If
'                            Else    'Don't include advertiser which use to be direct
'                                ilFound = True
'                                slListBox = slListBox & " " & tgCommAdf(ilIndex).sName
                            End If
                    '        Exit For
                        End If
                    'Next ilIndex
                End If
            Else
                'For ilIndex = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                '    If tgCommAdf(ilIndex).iCode = tgCCdf(ilLoop).tCdfExt.iAdfCode Then
                    ilIndex = gBinarySearchAdf(tgCCdf(ilLoop).tCdfExt.iAdfCode)
                    If ilIndex <> -1 Then
                        ilFound = True
                        If tgCommAdf(ilIndex).sBillAgyDir = "D" Then
                            If (Trim$(tgCommAdf(ilIndex).sAddrID) <> "") Then
                                slListBox = slListBox & " " & Trim$(tgCommAdf(ilIndex).sName) & ", " & Trim$(tgCommAdf(ilIndex).sAddrID) & "/Direct"
                            Else
                                slListBox = slListBox & " " & Trim$(tgCommAdf(ilIndex).sName) & "/Direct"
                            End If
                        Else
                            'slListBox = slListBox & " " & tgCommAdf(ilIndex).sName & "/Direct"
                            slListBox = slListBox & " " & Trim$(tgCommAdf(ilIndex).sName)
                        End If
                '        Exit For
                    End If
                'Next ilIndex
            End If
        End If
        If ilFound Then
            slListBox = slListBox & "|" & Trim$(str$(ilLoop))
'            slListBox = slListBox & "\" & Trim$(Str$(tgCCdf(ilLoop).lRecPos))
            If (slType = "AGF") And (tgCCdf(ilLoop).tCdfExt.iAgfCode > 0) Then
                slListBox = slListBox & "\" & Trim$(str$(tgCCdf(ilLoop).tCdfExt.iAgfCode))
            Else
                slListBox = slListBox & "\" & Trim$(str$(tgCCdf(ilLoop).tCdfExt.iAdfCode))
            End If
            lbcMster.AddItem slListBox    'Add ID (retain matching sorted order) and Code number to list box
        End If
    Next ilLoop
    For ilList = 0 To lbcMster.ListCount - 1 Step 1
        slNameCode = lbcMster.List(ilList)
        ilRet = gParseItem(slNameCode, 2, "|", slName)  'Obtain Index and code number
        If ilRet <> CP_MSG_NONE Then
            gPopCommentCollectBox = CP_MSG_PARSE
            Exit Function
        End If
        ilRet = gParseItem(slName, 1, "\", slStr)       'Obtain index
        If ilRet <> CP_MSG_NONE Then
            gPopCommentCollectBox = CP_MSG_PARSE
            Exit Function
        End If
        ilLoop = Val(slStr)
        slListBox = ""
        If (slSort = "A") Or (slSort = "C") Then
            For ilIndex = LBound(tgUrf) To UBound(tgUrf) Step 1
                If tgUrf(ilIndex).iCode = tgCCdf(ilLoop).tCdfExt.iUrfCode Then
'                    slListBox = Trim$(tgUrf(ilIndex).sName)    'Don't trim as dates must line up
                    Exit For
                End If
            Next ilIndex
        Else
            For ilIndex = LBound(tgCommUrf) To UBound(tgCommUrf) - 1 Step 1
                If tgCommUrf(ilIndex).iCode = tgCCdf(ilLoop).tCdfExt.iUrfCode Then
                    slListBox = Trim$(gDecryptField(tgCommUrf(ilIndex).sName))
                    Exit For
                End If
            Next ilIndex
        End If
        If (slSort = "A") Or (slSort = "B") Then
            gUnpackDate tgCCdf(ilLoop).tCdfExt.iActionDate(0), tgCCdf(ilLoop).tCdfExt.iActionDate(1), slDate
            slListBox = slListBox & " " & slDate
        Else
            gUnpackDate tgCCdf(ilLoop).tCdfExt.iDateEntrd(0), tgCCdf(ilLoop).tCdfExt.iDateEntrd(1), slDate
            slListBox = slListBox & " " & slDate
        End If
        If slType = "AGF" Then
            If tgCCdf(ilLoop).tCdfExt.iAgfCode > 0 Then
                'For ilIndex = LBound(tgCommAgf) To UBound(tgCommAgf) - 1 Step 1
                '    If tgCommAgf(ilIndex).iCode = tgCCdf(ilLoop).tCdfExt.iAgfCode Then
                    ilIndex = gBinarySearchAgf(tgCCdf(ilLoop).tCdfExt.iAgfCode)
                    If ilIndex <> -1 Then
                        slListBox = slListBox & " " & Trim$(tgCommAgf(ilIndex).sName) & ", " & Trim$(tgCommAgf(ilIndex).sCityID)
                '        Exit For
                    End If
                'Next ilIndex
            Else
                'For ilIndex = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                '    If tgCommAdf(ilIndex).iCode = tgCCdf(ilLoop).tCdfExt.iAdfCode Then
                    ilIndex = gBinarySearchAdf(tgCCdf(ilLoop).tCdfExt.iAdfCode)
                    If ilIndex <> -1 Then
                        If tgCommAdf(ilIndex).sBillAgyDir = "D" Then
                            If Trim$(tgCommAdf(ilIndex).sAddrID) <> "" Then
                                slListBox = slListBox & " " & Trim$(tgCommAdf(ilIndex).sName) & ", " & Trim$(tgCommAdf(ilIndex).sAddrID) & "/Direct"
                            Else
                                slListBox = slListBox & " " & Trim$(tgCommAdf(ilIndex).sName) & "/Direct"
                            End If
                        Else
                            slListBox = slListBox & " " & Trim$(tgCommAdf(ilIndex).sName)
                        End If
                '        Exit For
                    End If
                'Next ilIndex
            End If
        Else
            'For ilIndex = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
            '    If tgCommAdf(ilIndex).iCode = tgCCdf(ilLoop).tCdfExt.iAdfCode Then
                ilIndex = gBinarySearchAdf(tgCCdf(ilLoop).tCdfExt.iAdfCode)
                If ilIndex <> -1 Then
                    If tgCommAdf(ilIndex).sBillAgyDir = "D" Then
                        If Trim$(tgCommAdf(ilIndex).sAddrID) <> "" Then
                            slListBox = slListBox & " " & Trim$(tgCommAdf(ilIndex).sName) & ", " & Trim$(tgCommAdf(ilIndex).sAddrID) & "/Direct"
                        Else
                            slListBox = slListBox & " " & Trim$(tgCommAdf(ilIndex).sName) & "/Direct"
                        End If
                    Else
                        slListBox = slListBox & " " & Trim$(tgCommAdf(ilIndex).sName)
                    End If
            '        Exit For
                End If
            'Next ilIndex
        End If
        lbcLocal.AddItem Trim$(slListBox)    'Add ID (retain matching sorted order) and Code number to list box
    Next ilList
    Exit Function
gPopCommentCollectBoxErr:
    ilRet = btrClose(hlCdf)
    btrDestroy hlCdf
    gPopCommentCollectBox = CP_MSG_NOSHOW
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateCollectSortRec           *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Create AGFADFCOLL structure     *
'*                                                     *
'*******************************************************
Private Function mCreateCollectSortRec(hlFile As Integer, slType As String, tlColl() As AGFADFCOLL) As Integer
'
'   ilRet = mCreateCollectSortRec(hlFile, slType, tlColl())
'   Where
'       hlFile (I)- file handle (agf or adf)
'       slType (I)- "AGF"= Agency; "ADFDIR"= Direct Advertiser;
'                   "ADFXDIR" = All advertiser which are not direct; "ADFALL" = All Advertiser
'       tlColl() (I)- AGFADFCOLL record structure to be created
'       ilRet (I)- True=No errors; False=Error
'
    Dim ilRet As Integer
    Dim ilUpperBound As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim ilExtLen As Integer
    Dim llRecPos As Long
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim slFileName As String
    Dim ilLoop As Integer
    Dim ilAdd As Integer

    ilUpperBound = UBound(tlColl)
    ilExtLen = Len(tlColl(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlFile) 'Obtain number of records
    Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "AGFADFCOLLPK", AGFADFCOLLPK) 'Set extract limits (all records)
    If slType = "ADFDIR" Then
        'Get all record, then test if Direct or Non-Payee
'        tlCharTypeBuff.sType = "D"
'        ilOffset = gFieldOffset("Adf", "AdfBilAgyDir")
'        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        slFileName = "ADF"
    ElseIf slType = "ADFXDIR" Then
        tlCharTypeBuff.sType = "D"
        ilOffSet = gFieldOffset("Adf", "AdfBilAgyDir")
        ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        slFileName = "ADF"
    ElseIf slType = "ADFALL" Then
        slFileName = "ADF"
    Else
        slFileName = slType
    End If
    ilOffSet = gFieldOffset(slFileName, slFileName & "Code")
    ilRet = btrExtAddField(hlFile, ilOffSet, 2)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        mCreateCollectSortRec = False
        Exit Function
    End If
    ilOffSet = gFieldOffset(slFileName, slFileName & "Name")
    ilRet = btrExtAddField(hlFile, ilOffSet, 30)  'Extract Name field
    If ilRet <> BTRV_ERR_NONE Then
        mCreateCollectSortRec = False
        Exit Function
    End If
    If slFileName = "AGF" Then
        ilOffSet = gFieldOffset(slFileName, slFileName & "City")
        'Obtain 1 extra bytes to match advertiser
        ilRet = btrExtAddField(hlFile, ilOffSet, 1) 'Extract city ID
        ilOffSet = gFieldOffset(slFileName, slFileName & "City")
        'Obtain 4 extra bytes to match advertiser
        ilRet = btrExtAddField(hlFile, ilOffSet, 9) 'Extract city ID
    Else
        ilOffSet = gFieldOffset(slFileName, slFileName & "BilAgyDir")
        ilRet = btrExtAddField(hlFile, ilOffSet, 1) 'Extract city ID
        ilOffSet = gFieldOffset(slFileName, slFileName & "AddrID")
        ilRet = btrExtAddField(hlFile, ilOffSet, 9) 'Extract city ID
    End If
    If ilRet <> BTRV_ERR_NONE Then
        mCreateCollectSortRec = False
        Exit Function
    End If
    ilOffSet = gFieldOffset(slFileName, slFileName & "SlfCode")
    ilRet = btrExtAddField(hlFile, ilOffSet, 2) 'Salesperson  code
    If ilRet <> BTRV_ERR_NONE Then
        mCreateCollectSortRec = False
        Exit Function
    End If
    ilOffSet = gFieldOffset(slFileName, slFileName & "CreditRestr")
    ilRet = btrExtAddField(hlFile, ilOffSet, 1) 'Extract Credit restriction
    If ilRet <> BTRV_ERR_NONE Then
        mCreateCollectSortRec = False
        Exit Function
    End If
    ilOffSet = gFieldOffset(slFileName, slFileName & "CreditLimit")
    ilRet = btrExtAddField(hlFile, ilOffSet, 4) 'Extract Credit limit
    If ilRet <> BTRV_ERR_NONE Then
        mCreateCollectSortRec = False
        Exit Function
    End If
    ilOffSet = gFieldOffset(slFileName, slFileName & "PaymRating")
    ilRet = btrExtAddField(hlFile, ilOffSet, 1) 'Extract Payment rating
    If ilRet <> BTRV_ERR_NONE Then
        mCreateCollectSortRec = False
        Exit Function
    End If
    ilOffSet = gFieldOffset(slFileName, slFileName & "Pct90")
    ilRet = btrExtAddField(hlFile, ilOffSet, 2) 'Extract % over 90
    If ilRet <> BTRV_ERR_NONE Then
        mCreateCollectSortRec = False
        Exit Function
    End If
    ilOffSet = gFieldOffset(slFileName, slFileName & "CurrAR")
    ilRet = btrExtAddField(hlFile, ilOffSet, 6) 'Extract Current A/R
    If ilRet <> BTRV_ERR_NONE Then
        mCreateCollectSortRec = False
        Exit Function
    End If
    ilOffSet = gFieldOffset(slFileName, slFileName & "DateLstPaym")
    ilRet = btrExtAddField(hlFile, ilOffSet, 4) 'Extract Date of last payment
    If ilRet <> BTRV_ERR_NONE Then
        mCreateCollectSortRec = False
        Exit Function
    End If
    ilOffSet = gFieldOffset(slFileName, slFileName & "AvgToPay")
    ilRet = btrExtAddField(hlFile, ilOffSet, 2) 'Extract average number of days to pay
    If ilRet <> BTRV_ERR_NONE Then
        mCreateCollectSortRec = False
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlFile)    'Extract record
    ilUpperBound = UBound(tlColl)
    ilRet = btrExtGetNext(hlFile, tlColl(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            mCreateCollectSortRec = False
            Exit Function
        End If
        ilExtLen = Len(tlColl(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlFile, tlColl(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlFile, tlColl(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If slType = "ADFDIR" Then
                If Left$(tlColl(ilUpperBound).sBillAgyDir, 1) = "D" Then
                    ilAdd = True
                Else
                    ilAdd = False
                    For ilLoop = LBound(imNonPayeeAdfCode) To UBound(imNonPayeeAdfCode) - 1 Step 1
                        If tlColl(ilUpperBound).iCode = imNonPayeeAdfCode(ilLoop) Then
                            tlColl(ilUpperBound).sBillAgyDir = "N"
                            ilAdd = True
                        End If
                    Next ilLoop
                End If
            Else
                ilAdd = True
            End If
            If ilAdd Then
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tlColl(LBound(tlColl) To ilUpperBound) As AGFADFCOLL
            End If
            ilRet = btrExtGetNext(hlFile, tlColl(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlFile, tlColl(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    mCreateCollectSortRec = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeCollectListImage           *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Make list image for collect     *
'*                     list box                        *
'*                                                     *
'*******************************************************
Private Function mMakeCollectListImage(slType As String, slSort As String, tlColl As AGFADFCOLL) As String
'
'   slBox = mMakeCollectListImage(slType, slSort, tlColl)
'   Where
'       slType (I)- "AGF"= Agency; "ADF"= Advertiser
'       slSort (I)- see gPopAgyCollectBox
'       tlColl (I)- AGFADFCOLL record structure to be converted into sort string
'       slBox (I)- string to be added to list box
'
    Dim slName As String
    Dim slCurrAR As String
    Dim slListBox As String
    Dim slStr As String
    Dim slDate As String
    Dim ilLoop As Integer
    If slType = "AGF" Then
        slName = Trim$(tlColl.sName) & ", " & Trim$(Left$(tlColl.sCityID, 5))
    Else
        If Asc(tlColl.sBillAgyDir) = Asc("D") Then
            slStr = "/Direct"
        ElseIf Trim$(tlColl.sBillAgyDir) = "N" Then
            slStr = "/Non-Payee"
        Else
            slStr = ""
        End If
        If Trim$(tlColl.sCityID) = "" Then
            slName = Trim$(tlColl.sName) & slStr
        Else
            slName = Trim$(tlColl.sName) & ", " & Trim$(tlColl.sCityID) & slStr
        End If
    End If
    gPDNToStr tlColl.sCurrAR, 2, slCurrAR
    gFormatStr slCurrAR, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 2, slCurrAR
    Select Case slSort
        Case "A"    'Alphabetic
            slListBox = slName & " " & slCurrAR
        Case "B"    'by % over 90
            'gPDNToStr tlColl.sPct90, 0, slStr
            slStr = str$(tlColl.iPct90)
            slListBox = slStr & "%" & " " & slName & " " & slCurrAR
        Case "C"    'Amount owed
            slListBox = slCurrAR & " " & slName
        Case "D"    'Days since paid
            gUnpackDate tlColl.iDateLstPaym(0), tlColl.iDateLstPaym(1), slStr
            If Trim$(slStr) <> "" Then
                slDate = Format$(gNow(), "m/d/yy")
                slListBox = Trim$(str$(gDateValue(slDate) - gDateValue(slStr))) & " " & slName & " " & slCurrAR
            Else
                slListBox = "0" & " " & slName & " " & slCurrAR
            End If
        Case "E"    'Average days to pay
            slListBox = Trim$(str$(tlColl.iAvgToPay)) & " " & slName & " " & slCurrAR
        Case "F"    'Credit Restriction
            'gPDNToStr tlColl.sCreditLimit, 2, slStr
            slStr = gLongToStrDec(tlColl.lCreditLimit, 2)
            gFormatStr slStr, FMTLEAVEBLANK + FMTCOMMA + FMTDOLLARSIGN, 2, slStr
            Select Case tlColl.sCreditRestr
                Case "N"
                    slStr = "Unrestricted: "
                Case "L"
                    slStr = "Limit " & slStr & ": "
                Case "W"
                    slStr = "Cash in Adv- Week: "
                Case "M"
                    slStr = "Cash in Adv- Month: "
                Case "T"
                    slStr = "Cash in Adv- Quarter: "
                Case "P"
                    slStr = "No Orders: "
                Case Else
                    slStr = ""
            End Select
            slListBox = slStr & slName & " " & slCurrAR
        Case "G"    'Payment rating
            Select Case tlColl.sPaymRating
                Case "0"
                    slStr = "Quick Pay: "
                Case "1"
                    slStr = "Normal Pay: "
                Case "2"
                    slStr = "Slow Pay: "
                Case "3"
                    slStr = "Difficult: "
                Case "4"
                    slStr = "In Collection: "
                Case Else
                    slStr = ""
            End Select
            slListBox = slStr & slName & " " & slCurrAR
        Case "H"    'Salesperson
            slListBox = ""
            For ilLoop = LBound(tgCSlf) To UBound(tgCSlf) - 1 Step 1
                If tlColl.iSlfCode = tgCSlf(ilLoop).iCode Then
                    slListBox = Trim$(tgCSlf(ilLoop).sLastName) & "," & Trim$(tgCSlf(ilLoop).sFirstName) & ": "
                    Exit For
                End If
            Next ilLoop
            slListBox = slListBox & slName & " " & slCurrAR
    End Select
    mMakeCollectListImage = slListBox
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeCollectSort                *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Make sort record for collect    *
'*                     from AGFADFCOLL structure       *
'*                                                     *
'*******************************************************
Private Function mMakeCollectSort(slType As String, slSort As String, tlColl As AGFADFCOLL) As String
'
'   slBox = mMakeCollectSort(slType, slSort, tlColl)
'   Where
'       slType (I)- "AGF"= Agency; "ADF"= Advertiser
'       slSort (I)- see gPopAgyCollectBox
'       tlColl (I)- AGFADFCOLL record structure to be converted into sort string
'       slBox (I)- string to be sorted
'
    Dim slName As String * 42
    Dim slCurrAR As String
    Dim slListBox As String
    Dim slStr As String
    Dim ilLoop As Integer
    If slType = "AGF" Then
        slName = Trim$(tlColl.sName) & ", " & Trim$(Left$(tlColl.sCityID, 5))
    Else
        If Asc(tlColl.sBillAgyDir) = Asc("D") Then
            slName = Trim$(tlColl.sName) & "/Direct"
        ElseIf Trim$(tlColl.sBillAgyDir) = "N" Then
            slName = Trim$(tlColl.sName) & "/Non-Payee"
        Else
            slName = tlColl.sName
        End If
    End If
    gPDNToStr tlColl.sCurrAR, 2, slCurrAR
    Do While Len(slCurrAR) < 12
        slCurrAR = "0" & slCurrAR
    Loop
    Select Case slSort
        Case "A"    'Alphabetic
            slListBox = ""
        Case "B"    'by % over 90
            'gPDNToStr tlColl.sPct90, 0, slStr
            slStr = str$(tlColl.iPct90)
            slStr = gSubStr("1000", slStr)
            Do While Len(slStr) < 4
                slStr = "0" & slStr
            Loop
            slListBox = slStr
        Case "C"    'Amount owed
            gPDNToStr tlColl.sCurrAR, 2, slStr
            If InStr(slStr, "-") Then
                slStr = right$(slStr, Len(slStr) - 1)
                Do While Len(slStr) < 13
                    slStr = "0" & slStr
                Loop
                slStr = "-" & slStr
            Else
                slStr = gSubStr("1000000000.00", slStr)
                Do While Len(slStr) < 13
                    slStr = "0" & slStr
                Loop
                slStr = "+" & slStr
            End If
            slListBox = slStr
        Case "D"    'Days since paid
            gUnpackDateForSort tlColl.iDateLstPaym(0), tlColl.iDateLstPaym(1), slListBox
            If slListBox = "" Then  'Sort to end of list
                slListBox = "999999"
            End If
        Case "E"    'Average days to pay
            slStr = Trim$(str$(tlColl.iAvgToPay))
            slStr = gSubStr("10000", slStr)
            Do While Len(slStr) < 5
                slStr = "0" & slStr
            Loop
            slListBox = slStr
        Case "F"    'Credit Restriction
            'gPDNToStr tlColl.sCreditLimit, 2, slStr
            slStr = gLongToStrDec(tlColl.lCreditLimit, 2)
            slStr = gSubStr("1000000000.00", slStr)
            Do While Len(slStr) < 13
                slStr = "0" & slStr
            Loop
            Select Case tlColl.sCreditRestr
                Case "N"
                    slStr = "6" & slStr
                Case "L"
                    slStr = "5" & slStr
                Case "W"
                    slStr = "4" & slStr
                Case "M"
                    slStr = "3" & slStr
                Case "T"
                    slStr = "2" & slStr
                Case "P"
                    slStr = "1" & slStr
                Case Else
                    slStr = "0" & slStr
            End Select
            slListBox = slStr
        Case "G"    'Payment rating
            Select Case tlColl.sPaymRating
                Case "0"
                    slStr = "5"
                Case "1"
                    slStr = "4"
                Case "2"
                    slStr = "3"
                Case "3"
                    slStr = "2"
                Case "4"
                    slStr = "1"
                Case Else
                    slStr = "0"
            End Select
            slListBox = slStr
        Case "H"    'Salesperson
            'Find match
            slListBox = "ZZZZZZZZZZZZZZZZZZZZ"  'set so non-defined salesperson sort to bottom
            For ilLoop = LBound(tgCSlf) To UBound(tgCSlf) - 1 Step 1
                If tlColl.iSlfCode = tgCSlf(ilLoop).iCode Then
                    slListBox = Trim$(tgCSlf(ilLoop).sLastName) & "," & Trim$(tgCSlf(ilLoop).sFirstName)
                    Exit For
                End If
            Next ilLoop
    End Select
    mMakeCollectSort = slListBox & slName & slCurrAR
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainSalesperson              *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgCSlf for collection  *
'*                                                     *
'*******************************************************
Private Function mObtainSalesperson() As Integer
'
'   ilRet = mObtainSalesperson ()
'   Where:
'       tgCSlf() (I)- SLFEXT record structure to be created
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
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer

    slStamp = gFileDateTime(sgDBPath & "Slf.Btr")
    If sgCSlfStamp <> "" Then
        If StrComp(slStamp, sgCSlfStamp, 1) = 0 Then
            If UBound(tgCSlf) > 0 Then
                mObtainSalesperson = True
                Exit Function
            End If
        End If
    End If
    ReDim tgCSlf(0 To 0) As SLFEXT
    hlSlf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mObtainSalesperson = False
        ilRet = btrClose(hlSlf)
        btrDestroy hlSlf
        Exit Function
    End If
    ilRecLen = Len(tlSlf) 'btrRecordLength(hlSlf)  'Get and save record length
    sgCSlfStamp = slStamp
    ilUpperBound = UBound(tgCSlf)
    ilExtLen = Len(tgCSlf(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSlf) 'Obtain number of records
    btrExtClear hlSlf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlSlf, tlSlf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlSlf)
        btrDestroy hlSlf
        mObtainSalesperson = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mObtainSalesperson = False
            ilRet = btrClose(hlSlf)
            btrDestroy hlSlf
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlSlf, llNoRec, -1, "UC", "SLFEXTPK", SLFEXTPK) 'Set extract limits (all records)
    ilOffSet = gFieldOffset("Slf", "SlfCode")
    ilRet = btrExtAddField(hlSlf, ilOffSet, 2)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainSalesperson = False
        ilRet = btrClose(hlSlf)
        btrDestroy hlSlf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Slf", "SlfFirstName")
    ilRet = btrExtAddField(hlSlf, ilOffSet, 20)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainSalesperson = False
        ilRet = btrClose(hlSlf)
        btrDestroy hlSlf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Slf", "SlfLastName")
    ilRet = btrExtAddField(hlSlf, ilOffSet, 20) 'Extract Last Name field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainSalesperson = False
        ilRet = btrClose(hlSlf)
        btrDestroy hlSlf
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlSlf)    'Extract record
    ilUpperBound = UBound(tgCSlf)
    ilRet = btrExtGetNext(hlSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            mObtainSalesperson = False
            ilRet = btrClose(hlSlf)
            btrDestroy hlSlf
            Exit Function
        End If
        ilUpperBound = UBound(tgCSlf)
        ilExtLen = Len(tgCSlf(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilUpperBound = ilUpperBound + 1
            ReDim Preserve tgCSlf(0 To ilUpperBound) As SLFEXT
            ilRet = btrExtGetNext(hlSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSlf, tgCSlf(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlSlf)
    btrDestroy hlSlf
    mObtainSalesperson = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainUser                     *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgCommUrf for          *
'*                     collection                      *
'*                                                     *
'*******************************************************
Private Function mObtainUser() As Integer
'
'   ilRet = mObtainUser ()
'   Where:
'       tgCommUrf() (I)- URFEXT record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Slf date/time stamp
    Dim hlUrf As Integer        'Slf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Sof
    Dim tlUrf As URF
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer

    slStamp = gFileDateTime(sgDBPath & "Urf.Btr")
    If sgCommUrfStamp <> "" Then
        If StrComp(slStamp, sgCommUrfStamp, 1) = 0 Then
            If UBound(tgCommUrf) > 0 Then
                mObtainUser = True
                Exit Function
            End If
        End If
    End If
    ReDim tgCommUrf(0 To 0) As URFEXT
    hlUrf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlUrf, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mObtainUser = False
        ilRet = btrClose(hlUrf)
        btrDestroy hlUrf
        Exit Function
    End If
    ilRecLen = Len(tlUrf) 'btrRecordLength(hlUrf)  'Get and save record length
    sgCommUrfStamp = slStamp
    ilUpperBound = UBound(tgCommUrf)
    ilExtLen = Len(tgCommUrf(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlUrf) 'Obtain number of records
    btrExtClear hlUrf   'Clear any previous extend operation
    ilRet = btrGetFirst(hlUrf, tlUrf, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlUrf)
        btrDestroy hlUrf
        mObtainUser = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            mObtainUser = False
            ilRet = btrClose(hlUrf)
            btrDestroy hlUrf
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlUrf, llNoRec, -1, "UC", "URFEXTPK", URFEXTPK) 'Set extract limits (all records)
    ilOffSet = gFieldOffset("Urf", "UrfCode")
    ilRet = btrExtAddField(hlUrf, ilOffSet, 2)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainUser = False
        ilRet = btrClose(hlUrf)
        btrDestroy hlUrf
        Exit Function
    End If
    ilOffSet = gFieldOffset("Urf", "UrfRept")
    ilRet = btrExtAddField(hlUrf, ilOffSet, 20)  'Extract First Name field
    If ilRet <> BTRV_ERR_NONE Then
        mObtainUser = False
        ilRet = btrClose(hlUrf)
        btrDestroy hlUrf
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlUrf)    'Extract record
    ilUpperBound = UBound(tgCommUrf)
    ilRet = btrExtGetNext(hlUrf, tgCommUrf(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            mObtainUser = False
            ilRet = btrClose(hlUrf)
            btrDestroy hlUrf
            Exit Function
        End If
        ilUpperBound = UBound(tgCommUrf)
        ilExtLen = Len(tgCommUrf(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlUrf, tgCommUrf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlUrf, tgCommUrf(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            ilUpperBound = ilUpperBound + 1
            ReDim Preserve tgCommUrf(0 To ilUpperBound) As URFEXT
            ilRet = btrExtGetNext(hlUrf, tgCommUrf(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlUrf, tgCommUrf(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlUrf)
    btrDestroy hlUrf
    mObtainUser = True
    Exit Function
End Function

Public Sub mNonPayeeAdfCode()
'    Dim hlFile As Integer
    Dim ilIndex As Integer

    ReDim imNonPayeeAdfCode(0 To 0) As Integer
    'Remove test to see if transaction exist for possible Non-Payee to speed up populating list box
    '5/28/04
'    imRvfPhfRecLen = Len(tmRvfPhf)
'    hlRvf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
'    ilRet = btrOpen(hlRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    hlPhf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
'    ilRet = btrOpen(hlPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    For ilIndex = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
        If (tgCommAdf(ilIndex).sBillAgyDir <> "D") And (Trim$(Left$(tgCommAdf(ilIndex).sFirstCntrAddr, 1)) <> "") Then
'            ilFound = False
'            tmRvfPhfSrchKey1.iAdfCode = tgCommAdf(ilIndex).iCode
'            ilRet = btrGetGreaterOrEqual(hlRvf, tmRvfPhf, imRvfPhfRecLen, tmRvfPhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
'            Do While (ilRet = BTRV_ERR_NONE) And (tmRvfPhf.iAdfCode = tgCommAdf(ilIndex).iCode)
'                If tmRvfPhf.iAgfCode = 0 Then
'                    ilFound = True
                    imNonPayeeAdfCode(UBound(imNonPayeeAdfCode)) = tgCommAdf(ilIndex).iCode
                    ReDim Preserve imNonPayeeAdfCode(0 To UBound(imNonPayeeAdfCode) + 1) As Integer
'                    Exit Do
'                End If
'                ilRet = btrGetNext(hlRvf, tmRvfPhf, imRvfPhfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'            Loop
'            If Not ilFound Then
'                tmRvfPhfSrchKey1.iAdfCode = tgCommAdf(ilIndex).iCode
'                ilRet = btrGetGreaterOrEqual(hlPhf, tmRvfPhf, imRvfPhfRecLen, tmRvfPhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
'                Do While (ilRet = BTRV_ERR_NONE) And (tmRvfPhf.iAdfCode = tgCommAdf(ilIndex).iCode)
'                    If tmRvfPhf.iAgfCode = 0 Then
'                        imNonPayeeAdfCode(UBound(imNonPayeeAdfCode)) = tgCommAdf(ilIndex).iCode
'                        ReDim Preserve imNonPayeeAdfCode(0 To UBound(imNonPayeeAdfCode) + 1) As Integer
'                        Exit Do
'                    End If
'                    ilRet = btrGetNext(hlPhf, tmRvfPhf, imRvfPhfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                Loop
'            End If
        End If
    Next ilIndex
'    ilRet = btrClose(hlRvf)
'    btrDestroy (hlRvf)
'    ilRet = btrClose(hlPhf)
'    btrDestroy (hlPhf)
    Exit Sub
'    imRvfPhfRecLen = Len(tmRvfPhf)
'    For ilPass = 0 To 1 Step 1
'        If ilPass = 0 Then
'            hlFile = CBtrvTable(ONEHANDLE) 'CBtrvTable()
'            ilRet = btrOpen(hlFile, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'        Else
'            hlFile = CBtrvTable(ONEHANDLE) 'CBtrvTable()
'            ilRet = btrOpen(hlFile, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'        End If
'        ilAdfCode = -1
'        ilExtRecLen = Len(tlRvfExt)
'        llNoRec = gExtNoRec(ilExtRecLen) 'btrRecords(hlRvf) 'Obtain number of records
'        btrExtClear hlFile   'Clear any previous extend operation
'        tmRvfPhfSrchKey0.iAgfCode = 0
'        ilRet = btrGetGreaterOrEqual(hlFile, tmRvfPhf, imRvfPhfRecLen, tmRvfPhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)
''        Do While (ilRet <> BTRV_ERR_END_OF_FILE) And (tmRvfPhf.iAgfCode = 0)
'        If (ilRet = BTRV_ERR_NONE) And (tmRvfPhf.iAgfCode = 0) Then
'            Call btrExtSetBounds(hlFile, llNoRec, -1, "UC", "RVFEXTPK", RVFEXTPK) 'Set extract limits (all records)
'            tlIntTypeBuff.iType = 0
'            ilOffset = gFieldOffset("Rvf", "RvfAgfCode")
'            ilRet = btrExtAddLogicConst(hlFile, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
'            ilRet = btrExtAddField(hlFile, ilOffset, ilExtRecLen)  'Extract agency and advertiser fields (first two fields)
'            If ilRet = BTRV_ERR_NONE Then
'                ilRet = btrExtGetNext(hlFile, tlRvfExt, ilExtRecLen, llRecPos)
'                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
'                    If ilRet = BTRV_ERR_NONE Then
'                        Do While ilRet = BTRV_ERR_REJECT_COUNT
'                            ilRet = btrExtGetNext(hlFile, tlRvfExt, ilExtRecLen, llRecPos)
'                        Loop
'                        Do While ilRet = BTRV_ERR_NONE
'                            If ilAdfCode <> tlRvfExt.iAdfCode Then
'                                ilIndex = gBinarySearchAdf(tlRvfExt.iAdfCode)
'                                If ilIndex <> -1 Then
'                                    If tgCommAdf(ilIndex).sBillAgyDir <> "D" Then
'                                        ilFound = False
'                                        For ilLoop = LBound(imNonPayeeAdfCode) To UBound(imNonPayeeAdfCode) - 1 Step 1
'                                            If imNonPayeeAdfCode(ilLoop) = tgCommAdf(ilIndex).iCode Then
'                                                ilFound = True
'                                                Exit For
'                                            End If
'                                        Next ilLoop
'                                        If Not ilFound Then
'                                            imNonPayeeAdfCode(UBound(imNonPayeeAdfCode)) = tgCommAdf(ilIndex).iCode
'                                            ReDim Preserve imNonPayeeAdfCode(0 To UBound(imNonPayeeAdfCode) + 1) As Integer
'                                        End If
'                                    End If
'                                End If
'                            End If
'                            ilAdfCode = tlRvfExt.iAdfCode
'                            ilRet = btrExtGetNext(hlFile, tlRvfExt, ilExtRecLen, llRecPos)
'                            Do While ilRet = BTRV_ERR_REJECT_COUNT
'                                ilRet = btrExtGetNext(hlFile, tlRvfExt, ilExtRecLen, llRecPos)
'                            Loop
'                        Loop
'                    End If
'                End If
'            End If
''            ilRet = btrGetNext(hlFile, tmRvfPhf, imRvfPhfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
''        Loop
'        End If
'        ilRet = btrClose(hlFile)
'        btrDestroy (hlFile)
'    Next ilPass
'    ilRet = btrClose(hlFile)
'    btrDestroy (hlFile)
End Sub
