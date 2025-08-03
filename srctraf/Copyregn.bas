Attribute VB_Name = "CopyRegnSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Copyregn.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  lmLowLimit                                                                            *
'*                                                                                        *
'* Public Variables (Removed)                                                             *
'*  tgOwners                      sgOwnersStamp                                           *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  RAFKEY3                       ARTTKEY1                      ARTTKEY2                  *
'*  MKTKEY1                       MKTKEY2                                                 *
'*                                                                                        *
'* Public Procedures (Marked)                                                             *
'*  gObtainOwners                                                                         *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Copy.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Copy subs and functions
Option Explicit
Option Compare Text

Dim imLowLimit As Integer


'Public igAdfCode As Integer
'********************************************************
'
'Region Area file definition
'
'*********************************************************
'Raf record layout
Type RAF
    lCode                 As Long            ' Autoincrement number
    iAdfCode              As Integer         ' Advertiser code number
    sName                 As String * 80     ' Region Name
    sState                As String * 1      ' A=Active; D=Dormant
    iDateEntrd(0 To 1)    As Integer         ' Date entered
    iDateDormant(0 To 1)  As Integer         ' Date set to Dormant
    iUrfCode              As Integer         ' Last user who modified product
                                             ' names
    lRegionCode           As Long            ' Cross Reference Region Code (map
                                             ' stations to regions), only for
                                             ' rafType = R
    sType                 As String * 1      ' Region Type(R=Regional Copy;
                                             ' C=Split Copy; N=Split Network P=Pod-Target)
    sCategory             As String * 1      ' Split Network and Split Copy
                                             ' Category(M=Market; N=State Name;
                                             ' Z=Zip Code; O=Owner and
                                             ' S=Station).  For rafType = C or
                                             ' N)
    sInclExcl             As String * 1      ' Include/Exclude (I=Include;
                                             ' E=Exclude).  For rafType = C or N
    sShowNoProposal       As String * 1      ' Show On Proposal (Y/N) for
                                             ' category S only.  For rafType = C
                                             ' or N
    sShowOnOrder       As String * 1      ' Show On Order (Y/N) for category
                                             ' S only.  For rafType = C or N
    sShowOnInvoice        As String * 1      ' Show On Invoice (Y/N) for
                                             ' Category S only.  For rafType = C
                                             ' or N
    iAudPct               As Integer         ' Audience Percent (xxx.xx).  Default 100.00.  Treat 0 as 100.
    sAbbr                 As String * 5      'Abbreviation.  Split Networks only (rafType = N)
    'sUnused               As String * 3
    sAssigned             As String * 1     'Assigned to Rotation (Y/N).  Test for N.  Used to know if region definition can be changed
    sUnused               As String * 1        'Unused
    sCustom               As String * 1      'Custom Region (Y/N).  Test for Y. For Type = C
End Type
'Raf key record layout- use LONGKEY0
'Type RAFKEY0
'    lCode As Long
'End Type
'Raf key record layout
Type RAFKEY1
    iAdfCode              As Integer
    sType                 As String * 1
End Type
'Raf key record layout- use LONGKEY0
'Type RAFKEY2
'    lRegionCode As Long
'End Type
Type RAFKEY3 'VBC NR
    iAdfCode              As Integer 'VBC NR
    sCustom               As String * 1      'Custom Region (Y/N).  Test for Y 'VBC NR
End Type 'VBC NR
Dim tmRaf As RAF
Type RAFKEY4
    sType                 As String * 1
End Type

Type RAFDUPL
    lCode As Long
    sName As String * 80
    iCount As Integer
End Type

'******************************************************************************
' SEF_Split_Entity Record Definition
'
'******************************************************************************
Type SEF
    lCode                 As Long            ' Autoincrement number
    lRafCode              As Long            ' RAF Reference
    sName                 As String * 40     ' State or Zip Code
    iIntCode              As Integer         ' Integer Refence Code.  mktCode
                                             ' (Market); shttCode (Station);
                                             ' arttCode (Owner)
    lLongCode             As Long            ' Long Reference code
                                             ' thfCode(Podcast-Target)
    sCategory             As String * 1      ' Split Copy Category(M=Market;
                                             ' N=State Name; Z=Zip Code; O=Owner
                                             ' and S=Station).  For rafType = C
                                             ' or N); P=Pod-Target
    sInclExcl             As String * 1      ' Include/Exclude (I=Include;
                                             ' E=Exclude).
    iSeqNo                As Integer         ' Sequence number (used to retain
                                             ' order of definitions)
    sUnused               As String * 12     ' Unused
End Type





'Type SEFKEY0- use LONGKEY0
'    lCode                 As Long
'End Type

Type SEFKEY1
    lRafCode              As Long
    iSeqNo                As Integer         ' Sequence number
End Type

Type SEFKEY2
    sCategory             As String * 1
    lLongCode             As Long
End Type
'******************************************************************************
' artt Record Definition
'
'******************************************************************************
Type ARTT
    lCode                 As Long            ' Auto Increment
    sFirstName            As String * 20     ' First Name.  Blank for arttType =
                                             ' P.
    sLastName             As String * 60     ' Last Name. For arttType = P, this
                                             ' will contain the whole name.  The
                                             ' First Name will be blank.
    sPhone                As String * 20     ' Phone Number
    sFax                  As String * 20     ' Fax #
    sEMail                As String * 70     ' E-Mail Address
    iState                As Integer         ' 0=Active; 1=Dormant
    iUsfCode              As Integer         ' User Code reference that added or
                                             ' changed this record
    sAddress(0 To 2)      As String * 40     ' Address field 1
    sAddressState         As String * 40     ' State Name (i.e Oregon; New
                                             ' York,..)
    sZip                  As String * 20     ' Zip code
    sCountry              As String * 40     ' Country name
    sType                 As String * 1      ' R=Affiliate Rep record;
                                             ' A=Administrator Records (from
                                             ' Site Option); P=Personnel record
                                             ' (Like Program Director; Music
                                             ' Director and Traffic Director.
                                             ' The arttTntCode indicates the
                                             ' persons title).
    iTntCode              As Integer         ' Reference to Title Name table
    iShttCode             As Integer         ' Station Reference Code.  Used for
                                             ' Personnel records (arttType =
                                             ' "P").
    sAffContact           As String * 1      ' Affiliate Contact (Y or N).
    aISCI2Contact         As String * 1      ' ISCI # 2 Contact.  Export with ISCI as record type F.  Y or N. Test for Y
    sWebEMail             As String * 1      ' Contact to receive Web E-Mail
    sEMailToWeb           As String * 1      ' E-Mail sent to Web. I=Insert into
                                             ' Web Table; U=Update Web Table; Y
                                             ' or Blank=Sent To Web
    iWebEMailRefID        As Integer         ' This ID is used to keep the Web and Affiliate system in sync with the E-Mail address
    sEMailRights          As String * 1      ' E-Mail Rights.M=Master Accept/Reject; A=Alternate Accept/Reject; N or Blank=None
    sUnused               As String * 19
End Type


'Type ARTTKEY0- use LONGKEY0
'    lCode                 As Long
'End Type

Type ARTTKEY1 'VBC NR
    sType                 As String * 1 'VBC NR
End Type 'VBC NR

Type ARTTKEY2 'VBC NR
    iShttCode             As Integer 'VBC NR
    sType                 As String * 1 'VBC NR
End Type 'VBC NR


'******************************************************************************
' mkt Record Definition
'
'******************************************************************************
Type MKT
    iCode                 As Integer         ' Auto Increment DMA
    sName                 As String * 60     ' Market Name
    iRank                 As Integer         ' Market Rank
    sBIA                  As String * 10     ' BIA ID Code for the market
    sArb                  As String * 10     ' Arbitron ID Code for the Market
    iUsfCode              As Integer         ' Reference to User
    sGroupName            As String * 10     ' Group Name
    sUnused               As String * 10     ' Unused
End Type


'Type MKTKEY0- use INTKEY0
'    iCode                 As Integer
'End Type

Type MKTKEY1 'VBC NR
    sBIA                  As String * 10 'VBC NR
End Type 'VBC NR

Type MKTKEY2 'VBC NR
    sArb                  As String * 10 'VBC NR
End Type 'VBC NR


'******************************************************************************
' met Record Definition
'
'******************************************************************************
Type MET
    iCode                 As Integer         'MSA
    sName                 As String * 60     ' Metro Name
    iRank                 As Integer         ' Metro Rank
    sGroupName            As String * 10     ' Group Name
    iUstCode              As Integer
    sUnused               As String * 10
End Type


'Type METKEY0- use INTKEY0
'    iCode                 As Integer
'End Type

'******************************************************************************
' FMT_Station_Format Record Definition
'
'******************************************************************************
Type FMT
    iCode                 As Integer
    sName                 As String * 60
    iUstCode              As Integer
    sGroupName            As String * 10     ' Group Name
    sUnused               As String * 10
End Type


'Type FMTKEY0- Use INTKEY0
'    iCode                 As Integer
'End Type

'******************************************************************************
' SNT Record Definition
'
'******************************************************************************
Type SNT
    iCode                 As Integer         ' Auto Increment.  State Names
    sName                 As String * 40     ' State or Provinces name
    sPostalName           As String * 2      ' Postal Name
    sGroupName            As String * 10     ' Group Name
    iUstCode              As Integer         ' Affiliate user reference code
    sUnused               As String * 20     ' Unused
End Type


'Type SNTKEY0- Use INTKEY0
'    iCode                 As Integer
'End Type


'******************************************************************************
' TZT Record Definition
'
'******************************************************************************
Type TZT
    iCode                 As Integer         ' Auto Increment code.  Time zone
                                             ' definition table.  Used to obtain
                                             ' Group name by time zone
    sName                 As String * 40     ' Time zone name
    sGroupName            As String * 10     ' Group Name
    sCSIName              As String * 3      ' CSI Name for time zone (EST, CST,
                                             ' MST or PST)
    iDisplaySeqNo         As Integer         'Display sequence number
    iUstCode              As Integer         ' Affiliate user reference code
    sUnused               As String * 20     ' Unused
End Type


'Type TZTKEY0- Use INTKEY0
'    iCode                 As Integer
'End Type


Public tgMarkets() As MKT
Public sgMarketsStamp As String
Public tgMSAMarkets() As MET
Public sgMSAMarketsStamp As String
Public tgFormats() As FMT
Public sgFormatsStamp As String
Public tgStates() As SNT
Public tgTimeZones() As TZT
Public tgVehicleStations() As SHTTINFO








'*******************************************************
'*                                                     *
'*      Procedure Name:gPopRegion                      *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     advertiser Region Area names    *
'*                                                     *
'*******************************************************
Function gObtainRegion(ilAdvtCode As Integer, slType As String, ilDormant As Integer, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gObtainRegion (ilAdvt, ilDormant, tlSortCode(), slSortCodeTag)
'   Where:
'       ilAdvt (I)- Advertise code value
'       ilDormant(I)- True=Include dormant Regions; False=Exclude dormant dormant
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'Adf date/time stamp
    Dim hlRaf As Integer        'Adf handle
    Dim ilRecLen As Integer     'Record length
    Dim slName As String
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlSrchKey As RAFKEY1  'Raf key record image
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilLoop As Integer
    Dim ilNumberDuplFd As Integer
    Dim ilMaxDuplNumber As Integer
    ReDim llClearAdfCode(0 To 0) As Long
    ReDim tlRafDupl(0 To 0) As RAFDUPL
    Dim tlRafSrchKey As LONGKEY0

    ilPop = True
    llLen = 0
    ilMaxDuplNumber = 0
    slStamp = gFileDateTime(sgDBPath & "Raf.Btr") & Trim$(str$(ilAdvtCode))
    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            'If lbcLocal.ListCount > 0 Then
                gObtainRegion = CP_MSG_NOPOPREQ
                Exit Function
            'End If
            ilPop = False
        End If
    End If
    gObtainRegion = CP_MSG_POPREQ
    slSortCodeTag = slStamp
    If ilPop Then
        hlRaf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'On Error GoTo gObtainRegionErr
        'gBtrvErrorMsg ilRet, "gObtainRegion (btrOpen): Raf.Btr", Frm
        'On Error GoTo 0
        ilRecLen = Len(tmRaf) 'btrRecordLength(hlRaf)  'Get and save record length
        ilSortCode = 0
        ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tmRaf)
        llNoRec = gExtNoRec(ilExtLen)
        btrExtClear hlRaf
        Call btrExtSetBounds(hlRaf, llNoRec, -1, "UC", "RAF", "") 'Set extract limits (all records)
        tlSrchKey.iAdfCode = ilAdvtCode
        ilRet = btrGetGreaterOrEqual(hlRaf, tmRaf, ilRecLen, tlSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
        If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
            ilRet = btrClose(hlRaf)
            'On Error GoTo gObtainRegionErr
            'gBtrvErrorMsg ilRet, "gObtainRegion (btrReset):" & "Raf.Btr", Frm
            'On Error GoTo 0
            btrDestroy hlRaf
            Exit Function
        End If
'        tlCharTypeBuff.sType = slType
'        ilOffset = gFieldOffset("Raf", "RafType")
'        ilRet = btrExtAddLogicConst(hlRaf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
'        tlIntTypeBuff.iType = ilAdvtCode
'        'ilOffset = GetOffSetForInt(tmRaf, tmRaf.iAdfCode) 'gFieldOffset("Raf", "RafAdfCode")
'        ilOffset = gFieldOffset("RAF", "RafAdfCode")
'        ilRet = btrExtAddLogicConst(hlRaf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
'        Call btrExtSetBounds(hlRaf, llNoRec, -1, "UC", "RAF", "") 'Set extract limits (all records)
        If ilAdvtCode <> -1 Then
            tlCharTypeBuff.sType = slType
            ilOffSet = gFieldOffset("Raf", "RafType")
            ilRet = btrExtAddLogicConst(hlRaf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
            tlIntTypeBuff.iType = ilAdvtCode
            'ilOffset = GetOffSetForInt(tmRaf, tmRaf.iAdfCode) 'gFieldOffset("Raf", "RafAdfCode")
            ilOffSet = gFieldOffset("RAF", "RafAdfCode")
            ilRet = btrExtAddLogicConst(hlRaf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        Else
            tlCharTypeBuff.sType = slType
            ilOffSet = gFieldOffset("Raf", "RafType")
            ilRet = btrExtAddLogicConst(hlRaf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        End If
        Call btrExtSetBounds(hlRaf, llNoRec, -1, "UC", "RAF", "") 'Set extract limits (all records)
        ilOffSet = 0
        ilRet = btrExtAddField(hlRaf, ilOffSet, ilRecLen)  'Extract iCode field
        'On Error GoTo gObtainRegionErr
        'gBtrvErrorMsg ilRet, "gObtainRegion (btrExtAddField):" & "Raf.Btr", Frm
        'On Error GoTo 0
        ilRet = btrExtGetNext(hlRaf, tmRaf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            'On Error GoTo gObtainRegionErr
            'gBtrvErrorMsg ilRet, "gObtainRegion (btrExtGetNextExt):" & "Raf.Btr", Frm
            'On Error GoTo 0
            ilExtLen = Len(tmRaf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlRaf, tmRaf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If (slType = "N") Then
                    If (tmRaf.iAdfCode > 0) Then
                        llClearAdfCode(UBound(llClearAdfCode)) = tmRaf.lCode
                        ReDim Preserve llClearAdfCode(0 To UBound(llClearAdfCode) + 1) As Long
                    End If
                    ilNumberDuplFd = 0
                    For ilLoop = 0 To UBound(tlRafDupl) - 1 Step 1
                        If StrComp(Trim$(tlRafDupl(ilLoop).sName), Trim$(tmRaf.sName), vbTextCompare) = 0 Then
                            ilNumberDuplFd = ilNumberDuplFd + 1
                        End If
                    Next ilLoop
                    tlRafDupl(UBound(tlRafDupl)).lCode = tmRaf.lCode
                    tlRafDupl(UBound(tlRafDupl)).sName = tmRaf.sName
                    tlRafDupl(UBound(tlRafDupl)).iCount = ilNumberDuplFd
                    ReDim Preserve tlRafDupl(0 To UBound(tlRafDupl) + 1) As RAFDUPL
                    If ilNumberDuplFd > ilMaxDuplNumber Then
                        ilMaxDuplNumber = ilNumberDuplFd
                    End If
                    If ilNumberDuplFd > 0 Then
                        tmRaf.sName = Trim$(tmRaf.sName) & " " & Trim$(str$(ilNumberDuplFd))
                    End If
                End If
                slName = tmRaf.sName
                If Trim$(tmRaf.sState) = "D" Then
                    slName = Trim$(slName) & "/Dormant"
                    Do While Len(slName) < Len(tmRaf.sName)
                        slName = slName & " "
                    Loop
                End If
                slName = slName & "\" & Trim$(str$(tmRaf.lCode)) & "\" & Trim$(tmRaf.sState)
                'If Not gOkAddStrToListBox(slName, llLen, True) Then
                '    Exit Do
                'End If
                'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                If (ilDormant) Or (tmRaf.sState = "A") Then
                    tlSortCode(ilSortCode).sKey = slName
                    If ilSortCode >= UBound(tlSortCode) Then
                        ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 100) As SORTCODE
                    End If
                    ilSortCode = ilSortCode + 1
                End If
                ilRet = btrExtGetNext(hlRaf, tmRaf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlRaf, tmRaf, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(0 To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - 1 > 0 Then
                ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
            End If
        End If
        For ilLoop = 0 To UBound(llClearAdfCode) - 1 Step 1
            tlRafSrchKey.lCode = llClearAdfCode(ilLoop)
            ilRet = btrGetEqual(hlRaf, tmRaf, ilExtLen, tlRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tmRaf.iAdfCode = 0
                ilRet = btrUpdate(hlRaf, tmRaf, ilExtLen)
            End If
        Next ilLoop
        If ilMaxDuplNumber > 0 Then
            For ilLoop = 0 To UBound(tlRafDupl) - 1 Step 1
                If tlRafDupl(ilLoop).iCount > 0 Then
                    tlRafSrchKey.lCode = tlRafDupl(ilLoop).lCode
                    ilRet = btrGetEqual(hlRaf, tmRaf, ilExtLen, tlRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        tmRaf.sName = Trim$(tmRaf.sName) & " " & Trim$(str$(tlRafDupl(ilLoop).iCount))
                        ilRet = btrUpdate(hlRaf, tmRaf, ilExtLen)
                    End If
                End If
            Next ilLoop
        End If
        ilRet = btrClose(hlRaf)
        'On Error GoTo gObtainRegionErr
        'gBtrvErrorMsg ilRet, "gObtainRegion (btrReset):" & "Raf.Btr", Frm
        'On Error GoTo 0
        btrDestroy hlRaf
    End If

    Exit Function

    ilRet = btrClose(hlRaf)
    btrDestroy hlRaf
    gObtainRegion = CP_MSG_NOSHOW
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gPopRegionBox                    *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate list Box with          *
'*                     advertiser Region Area names    *
'*                                                     *
'*******************************************************
Function gPopRegionBox(frm As Form, ilAdvtCode As Integer, slType As String, ilDormant As Integer, lbcLocal As Control, tlSortCode() As SORTCODE, slSortCodeTag As String) As Integer
'
'   ilRet = gPopRegionBox (MainForm, ilAdvt, ilDormant, lbcLocal, tlSortCode(), slSortCodeTag)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       ilAdvt (I)- Advertise code value or -1
'       ilDormant(I)- True=Include dormant Regions; False=Exclude dormant dormant
'       lbcLocal (I)- List box to be populated from the master list box
'       tlSortCode (I/O)- Sorted List containing name and code #
'       slSortCodeTag(I/O)- Date/Time stamp for tlSortCode
'       ilRet (O)- Error code (0 if no error)
'
    Dim slStamp As String    'Adf date/time stamp
    Dim hlRaf As Integer        'Adf handle
    Dim ilRecLen As Integer     'Record length
    Dim slName As String
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slNameCode As String
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim tlSrchKey As RAFKEY1  'Raf key record image
    Dim llLen As Long
    Dim ilSortCode As Integer
    Dim ilPop As Integer
    Dim ilNumberDuplFd As Integer
    Dim ilMaxDuplNumber As Integer
    ReDim llClearAdfCode(0 To 0) As Long
    ReDim tlRafDupl(0 To 0) As RAFDUPL
    Dim tlRafSrchKey As LONGKEY0

    ilPop = True
    llLen = 0
    ilMaxDuplNumber = 0
    slStamp = gFileDateTime(sgDBPath & "Raf.Btr") & Trim$(str$(ilAdvtCode))
    If slSortCodeTag <> "" Then
        If StrComp(slStamp, slSortCodeTag, 1) = 0 Then
            If lbcLocal.ListCount > 0 Then
                gPopRegionBox = CP_MSG_NOPOPREQ
                Exit Function
            End If
            ilPop = False
        End If
    End If
    gPopRegionBox = CP_MSG_POPREQ
    lbcLocal.Clear
    slSortCodeTag = slStamp
    If ilPop Then
        hlRaf = CBtrvTable(ONEHANDLE) 'CBtrvTable()
        ilRet = btrOpen(hlRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gPopRegionBoxErr
        gBtrvErrorMsg ilRet, "gPopRegionBox (btrOpen): Raf.Btr", frm
        On Error GoTo 0
        ilRecLen = Len(tmRaf) 'btrRecordLength(hlRaf)  'Get and save record length
        ilSortCode = 0
        ReDim tlSortCode(0 To 0) As SORTCODE   'VB list box clear (list box used to retain code number so record can be found)
        ilExtLen = Len(tmRaf)
        llNoRec = gExtNoRec(ilExtLen)
        btrExtClear hlRaf
        Call btrExtSetBounds(hlRaf, llNoRec, -1, "UC", "RAF", "") 'Set extract limits (all records)
        tlSrchKey.iAdfCode = ilAdvtCode
        ilRet = btrGetGreaterOrEqual(hlRaf, tmRaf, ilRecLen, tlSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
        If (ilRet = BTRV_ERR_END_OF_FILE) Or (ilRet = BTRV_ERR_KEY_NOT_FOUND) Then
            ilRet = btrClose(hlRaf)
            On Error GoTo gPopRegionBoxErr
            gBtrvErrorMsg ilRet, "gPopRegionBox (btrReset):" & "Raf.Btr", frm
            On Error GoTo 0
            btrDestroy hlRaf
            Exit Function
        End If
'        tlCharTypeBuff.sType = slType
'        ilOffset = gFieldOffset("Raf", "RafType")
'        ilRet = btrExtAddLogicConst(hlRaf, BTRV_KT_STRING, ilOffset, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
'        tlIntTypeBuff.iType = ilAdvtCode
'        'ilOffset = GetOffSetForInt(tmRaf, tmRaf.iAdfCode) 'gFieldOffset("Raf", "RafAdfCode")
'        ilOffset = gFieldOffset("RAF", "RAFADFCODE")
'        ilRet = btrExtAddLogicConst(hlRaf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
'        Call btrExtSetBounds(hlRaf, llNoRec, -1, "UC", "RAF", "") 'Set extract limits (all records)
        If ilAdvtCode <> -1 Then
            tlCharTypeBuff.sType = slType
            ilOffSet = gFieldOffset("Raf", "RafType")
            ilRet = btrExtAddLogicConst(hlRaf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlCharTypeBuff, 1)
            tlIntTypeBuff.iType = ilAdvtCode
            'ilOffset = GetOffSetForInt(tmRaf, tmRaf.iAdfCode) 'gFieldOffset("Raf", "RafAdfCode")
            ilOffSet = gFieldOffset("RAF", "RAFADFCODE")
            ilRet = btrExtAddLogicConst(hlRaf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        Else
            tlCharTypeBuff.sType = slType
            ilOffSet = gFieldOffset("Raf", "RafType")
            ilRet = btrExtAddLogicConst(hlRaf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)
        End If
        Call btrExtSetBounds(hlRaf, llNoRec, -1, "UC", "RAF", "") 'Set extract limits (all records)
        ilOffSet = 0
        ilRet = btrExtAddField(hlRaf, ilOffSet, ilRecLen)  'Extract iCode field
        On Error GoTo gPopRegionBoxErr
        gBtrvErrorMsg ilRet, "gPopRegionBox (btrExtAddField):" & "Raf.Btr", frm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlRaf, tmRaf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo gPopRegionBoxErr
            gBtrvErrorMsg ilRet, "gPopRegionBox (btrExtGetNextExt):" & "Raf.Btr", frm
            On Error GoTo 0
            ilExtLen = Len(tmRaf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlRaf, tmRaf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If (slType = "N") Then
                    If (tmRaf.iAdfCode > 0) Then
                        llClearAdfCode(UBound(llClearAdfCode)) = tmRaf.lCode
                        ReDim Preserve llClearAdfCode(0 To UBound(llClearAdfCode) + 1) As Long
                    End If
                    ilNumberDuplFd = 0
                    For ilLoop = 0 To UBound(tlRafDupl) - 1 Step 1
                        If StrComp(Trim$(tlRafDupl(ilLoop).sName), Trim$(tmRaf.sName), vbTextCompare) = 0 Then
                            ilNumberDuplFd = ilNumberDuplFd + 1
                        End If
                    Next ilLoop
                    tlRafDupl(UBound(tlRafDupl)).lCode = tmRaf.lCode
                    tlRafDupl(UBound(tlRafDupl)).sName = tmRaf.sName
                    tlRafDupl(UBound(tlRafDupl)).iCount = ilNumberDuplFd
                    ReDim Preserve tlRafDupl(0 To UBound(tlRafDupl) + 1) As RAFDUPL
                    If ilNumberDuplFd > ilMaxDuplNumber Then
                        ilMaxDuplNumber = ilNumberDuplFd
                    End If
                    If ilNumberDuplFd > 0 Then
                        tmRaf.sName = Trim$(tmRaf.sName) & " " & Trim$(str$(ilNumberDuplFd))
                    End If
                End If
                slName = tmRaf.sName
                slName = slName & "\" & Trim$(str$(tmRaf.lCode))
                'If Not gOkAddStrToListBox(slName, llLen, True) Then
                '    Exit Do
                'End If
                'lbcMster.AddItem slName    'Add ID (retain matching sorted order) and Code number to list box
                If (ilDormant) Or (tmRaf.sState = "A") Then
                    tlSortCode(ilSortCode).sKey = slName
                    If ilSortCode >= UBound(tlSortCode) Then
                        ReDim Preserve tlSortCode(0 To UBound(tlSortCode) + 100) As SORTCODE
                    End If
                    ilSortCode = ilSortCode + 1
                End If
                ilRet = btrExtGetNext(hlRaf, tmRaf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlRaf, tmRaf, ilExtLen, llRecPos)
                Loop
            Loop
            'Sort then output new headers and lines
            ReDim Preserve tlSortCode(0 To ilSortCode) As SORTCODE
            If UBound(tlSortCode) - 1 > 0 Then
                ArraySortTyp fnAV(tlSortCode(), 0), UBound(tlSortCode), 0, LenB(tlSortCode(0)), 0, LenB(tlSortCode(0).sKey), 0
            End If
        End If
        For ilLoop = 0 To UBound(llClearAdfCode) - 1 Step 1
            tlRafSrchKey.lCode = llClearAdfCode(ilLoop)
            ilRet = btrGetEqual(hlRaf, tmRaf, ilExtLen, tlRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tmRaf.iAdfCode = 0
                ilRet = btrUpdate(hlRaf, tmRaf, ilExtLen)
            End If
        Next ilLoop
        If ilMaxDuplNumber > 0 Then
            For ilLoop = 0 To UBound(tlRafDupl) - 1 Step 1
                If tlRafDupl(ilLoop).iCount > 0 Then
                    tlRafSrchKey.lCode = tlRafDupl(ilLoop).lCode
                    ilRet = btrGetEqual(hlRaf, tmRaf, ilExtLen, tlRafSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet = BTRV_ERR_NONE Then
                        tmRaf.sName = Trim$(tmRaf.sName) & " " & Trim$(str$(tlRafDupl(ilLoop).iCount))
                        ilRet = btrUpdate(hlRaf, tmRaf, ilExtLen)
                    End If
                End If
            Next ilLoop
        End If
        ilRet = btrClose(hlRaf)
        On Error GoTo gPopRegionBoxErr
        gBtrvErrorMsg ilRet, "gPopRegionBox (btrReset):" & "Raf.Btr", frm
        On Error GoTo 0
        btrDestroy hlRaf
    End If
    llLen = 0
    For ilLoop = 0 To UBound(tlSortCode) - 1 Step 1
        slNameCode = tlSortCode(ilLoop).sKey    'lbcMster.List(ilLoop)
        ilRet = gParseItem(slNameCode, 1, "\", slName)
        If ilRet <> CP_MSG_NONE Then
            gPopRegionBox = CP_MSG_PARSE
            Exit Function
        End If
        slName = Trim$(slName)
        If Not gOkAddStrToListBox(slName, llLen, True) Then
            Exit For
        End If
        lbcLocal.AddItem slName  'Add ID to list box
    Next ilLoop

    Exit Function
gPopRegionBoxErr:
    ilRet = btrClose(hlRaf)
    btrDestroy hlRaf
    gDbg_HandleError "CopyRegnSubs: gPopRegionBox"
'    gPopRegionBox = CP_MSG_NOSHOW
'    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainStations                 *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgStations             *
'*                                                     *
'*******************************************************
Function gObtainStations() As Integer
'
'   ilRet = gObtainStations ()
'   Where:
'       tgStations() (O)- SHTTINFO record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlShtt As Integer        'Mnf handle
    Dim ilShttRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlShtt As SHTT
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperShtt As Integer
    Dim ilUpperVefShtt As Integer
    Dim ilShtt As Integer
    'Dim hlAtt As Integer
    'Dim ilAttRecLen As Integer     'Record length
    'Dim tlAtt As ATT
    'Dim tlAttSrchKey As INTKEY0
    Dim ilAdd As Integer
    Dim ilVef As Integer
    Dim rst_vef As ADODB.Recordset
    
    '8132 Dan moved
'    bgStationAreVehicles = False
    
    'If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) <> SPLITNETWORKS) And ((Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) <> SPLITCOPY) Then
    '    ReDim tgStations(1 To 1) As SHTTINFO
    '    gObtainStations = True
    '    Exit Function
    'End If

    slStamp = gFileDateTime(sgDBPath & "Shtt.mkd")

    '11/26/17: Check Changed date/time
    If Not gFileChgd("shtt.mkd") Then
        gObtainStations = True
        Exit Function
    End If

    'On Error GoTo gObtainStationsErr2
    'ilRet = 0
    'imLowLimit = LBound(tgStations)
    'If ilRet <> 0 Then
    '    sgStationsStamp = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tgStations).Ptr <> 0 Then
        imLowLimit = LBound(tgStations)
    Else
        sgStationsStamp = ""
        imLowLimit = 0
    End If

    '6/9/18: replaced with gFileChgd
    'If sgStationsStamp <> "" Then
    '    If StrComp(slStamp, sgStationsStamp, 1) = 0 Then
    '        'If UBound(tgStations) > 1 Then
    '            gObtainStations = True
    '            Exit Function
    '        'End If
    '    End If
    'End If
    bgStationAreVehicles = False
    'ReDim tgStations(1 To 1) As SHTT
    'ReDim tgStations(1 To 20000) As SHTTINFO
    ReDim tgStations(0 To 20000) As SHTTINFO
    'ReDim tgVehicleStations(1 To 20000) As SHTTINFO
    ReDim tgVehicleStations(0 To 20000) As SHTTINFO
    hlShtt = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlShtt, "", sgDBPath & "Shtt.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainStations = False
        ilRet = btrClose(hlShtt)
        btrDestroy hlShtt
        Exit Function
    End If

    ilShttRecLen = Len(tlShtt) 'btrRecordLength(hlShtt)  'Get and save record length
    'hlAtt = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    'ilRet = btrOpen(hlAtt, "", sgDBPath & "Att.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet <> BTRV_ERR_NONE Then
    '    gObtainStations = False
    '    ilRet = btrClose(hlShtt)
    '    btrDestroy hlShtt
    '    ilRet = btrClose(hlAtt)
    '    btrDestroy hlAtt
    '    Exit Function
    'End If
    'ilAttRecLen = Len(tlAtt) 'btrRecordLength(hlShtt)  'Get and save record length
    sgStationsStamp = slStamp
    'ilUpperShtt = UBound(tgStations)
    ilUpperShtt = LBound(tgStations)
    ilUpperVefShtt = LBound(tgVehicleStations)
    ilExtLen = Len(tlShtt)  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlShtt) 'Obtain number of records
    btrExtClear hlShtt   'Clear any previous extend operation
    ilRet = btrGetFirst(hlShtt, tlShtt, ilShttRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlShtt)
        btrDestroy hlShtt
        'ilRet = btrClose(hlAtt)
        'btrDestroy hlAtt
        gObtainStations = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainStations = False
            ilRet = btrClose(hlShtt)
            btrDestroy hlShtt
            'ilRet = btrClose(hlAtt)
            'btrDestroy hlAtt
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlShtt, llNoRec, -1, "UC", "SHTT", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlShtt, ilOffSet, ilShttRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainStations = False
        ilRet = btrClose(hlShtt)
        btrDestroy hlShtt
        'ilRet = btrClose(hlAtt)
        'btrDestroy hlAtt
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlShtt)    'Extract record
    ilRet = btrExtGetNext(hlShtt, tlShtt, ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainStations = False
            ilRet = btrClose(hlShtt)
            btrDestroy hlShtt
            'ilRet = btrClose(hlAtt)
            'btrDestroy hlAtt
            Exit Function
        End If
        ilExtLen = Len(tlShtt)  'Extract operation record size
        'ilRet = btrExtGetFirst(hlShtt, tgStations(ilUpperShtt), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlShtt, tlShtt, ilExtLen, llRecPos)
        Loop

        ''Dick, I commented out one line of code and added two lines.  You can see all 3 lines in the left margin.
        'ReDim Preserve tgStations(1 To 20000) As SHTT

        Do While ilRet = BTRV_ERR_NONE
            If tlShtt.iType = 0 Then
                'ilAdd = False
                ''If ((Asc(tgSpf.sUsingFeatures7) And WEGENEREXPORT) <> WEGENEREXPORT) And ((Asc(tgSpf.sUsingFeatures7) And OLAEXPORT) <> OLAEXPORT) Then
                'If (Asc(tgSpf.sUsingFeatures7) And OLAEXPORT) <> OLAEXPORT Then
                '    'tlAttSrchKey.iCode = tgStations(ilUpperShtt).iCode
                '    'If tlAttSrchKey.iCode > 0 Then
                '    '    ilRet = btrGetEqual(hlAtt, tlAtt, ilAttRecLen, tlAttSrchKey, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                '    '    If ilRet = BTRV_ERR_NONE Then
                '    '        ilAdd = True
                '    '        '12/26/08: Don't add station twice
                '    '        'ilUpperShtt = ilUpperShtt + 1
                '    '        ''Dick
                '    '        ''ReDim Preserve tgStations(1 To ilUpperShtt) As SHTT
                '    '        'If ilUpperShtt > UBound(tgStations) Then
                '    '        '    ReDim Preserve tgStations(1 To ilUpperShtt + 1000) As SHTT
                '    '        'End If
                '    '    End If
                '    'End If
                '    If tlShtt.sAgreementExist = "Y" Then
                '        ilAdd = True
                '    End If
                'Else
                '    ilAdd = True
                'End If
                'If ilAdd Then
                    tgStations(ilUpperShtt).iCode = tlShtt.iCode
                    tgStations(ilUpperShtt).sCallLetters = tlShtt.sCallLetters
                    '12/28/15: Replace state with user specified state
                    'tgStations(ilUpperShtt).sState = tlShtt.sState
                    If (Asc(tgSaf(0).sFeatures3) And SPLITCOPYLICENSE) = SPLITCOPYLICENSE Then 'Require Station Posting Prior to Invoicing
                        tgStations(ilUpperShtt).sState = tlShtt.sStateLic
                    ElseIf (Asc(tgSaf(0).sFeatures3) And SPLITCOPYPHYSICAL) = SPLITCOPYPHYSICAL Then
                        tgStations(ilUpperShtt).sState = tlShtt.sONState
                    Else
                        tgStations(ilUpperShtt).sState = tlShtt.sState
                    End If
                    tgStations(ilUpperShtt).sTimeZone = tlShtt.sTimeZone
                    tgStations(ilUpperShtt).sAgreementExist = tlShtt.sAgreementExist
                    tgStations(ilUpperShtt).lPermStationID = tlShtt.lPermStationID
                    tgStations(ilUpperShtt).iMktCode = tlShtt.iMktCode
                    tgStations(ilUpperShtt).iFmtCode = tlShtt.iFmtCode
                    tgStations(ilUpperShtt).iTztCode = tlShtt.iTztCode
                    tgStations(ilUpperShtt).iMetCode = tlShtt.iMetCode
                    tgStations(ilUpperShtt).iShttVefCode = 0
                    tgStations(ilUpperShtt).lAudP12Plus = tlShtt.lAudP12Plus        '12-17-18
                    tgStations(ilUpperShtt).lMultiCastGroupID = tlShtt.lMultiCastGroupID
'                    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
'                        If UCase$(Trim$(tlShtt.sCallLetters)) = UCase$(Trim$(tgMVef(ilVef).sName)) Then
'                            bgStationAreVehicles = True
'                            tgStations(ilUpperShtt).iShttVefCode = tgMVef(ilVef).iCode
'                            tgVehicleStations(ilUpperVefShtt) = tgStations(ilUpperShtt)
'                            ilUpperVefShtt = ilUpperVefShtt + 1
'                            If ilUpperVefShtt > UBound(tgVehicleStations) Then
'                                'ReDim Preserve tgVehicleStations(1 To ilUpperVefShtt + 1000) As SHTTINFO
'                                ReDim Preserve tgVehicleStations(0 To ilUpperVefShtt + 1000) As SHTTINFO
'                            End If
'                            Exit For
'                        End If
'                    Next ilVef
                    ilVef = gBinarySearchVefName(UCase$(Trim$(tlShtt.sCallLetters)))
                    If ilVef <> -1 Then
                        bgStationAreVehicles = True
                        tgStations(ilUpperShtt).iShttVefCode = tgVefName(ilVef).iCode
                        tgVehicleStations(ilUpperVefShtt) = tgStations(ilUpperShtt)
                        ilUpperVefShtt = ilUpperVefShtt + 1
                        If ilUpperVefShtt > UBound(tgVehicleStations) Then
                            'ReDim Preserve tgVehicleStations(1 To ilUpperVefShtt + 1000) As SHTTINFO
                            ReDim Preserve tgVehicleStations(0 To ilUpperVefShtt + 1000) As SHTTINFO
                        End If
                    End If
                    If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Or ((Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) = SPLITCOPY) Then
                        If (tlShtt.sAgreementExist = "Y") Or ((Asc(tgSpf.sUsingFeatures7) And OLAEXPORT) = OLAEXPORT) Or ((Asc(tgSpf.sUsingFeatures5) And REMOTEEXPORT) = REMOTEEXPORT) Then
                            ilUpperShtt = ilUpperShtt + 1
                            If ilUpperShtt > UBound(tgStations) Then
                                'ReDim Preserve tgStations(1 To ilUpperShtt + 1000) As SHTTINFO
                                ReDim Preserve tgStations(0 To ilUpperShtt + 1000) As SHTTINFO
                            End If
                        End If
                    End If
                'End If
            End If
            ilRet = btrExtGetNext(hlShtt, tlShtt, ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlShtt, tlShtt, ilExtLen, llRecPos)
            Loop
        Loop
        'Dick
        'ReDim Preserve tgStations(1 To ilUpperShtt) As SHTTINFO
        ReDim Preserve tgStations(0 To ilUpperShtt) As SHTTINFO
        'ReDim Preserve tgVehicleStations(1 To ilUpperVefShtt) As SHTTINFO
        ReDim Preserve tgVehicleStations(0 To ilUpperVefShtt) As SHTTINFO
    End If
    ilRet = btrClose(hlShtt)
    btrDestroy hlShtt
    'ilRet = btrClose(hlAtt)
    'btrDestroy hlAtt
    ReDim tgMktSort(LBound(tgStations) To UBound(tgStations)) As REGIONINTSORT
    ReDim tgMSAMktSort(LBound(tgStations) To UBound(tgStations)) As REGIONINTSORT
    'ReDim tgOwnerSort(LBound(tgStations) To UBound(tgStations)) As REGIONINTSORT
    'ReDim tgZipSort(LBound(tgStations) To UBound(tgStations)) As REGIONSTRSORT
    ReDim tgStateSort(LBound(tgStations) To UBound(tgStations)) As REGIONSTRSORT
    ReDim tgFmtSort(LBound(tgStations) To UBound(tgStations)) As REGIONINTSORT
    ReDim tgTztSort(LBound(tgStations) To UBound(tgStations)) As REGIONINTSORT
    For ilShtt = LBound(tgStations) To UBound(tgStations) - 1 Step 1
        tgMktSort(ilShtt).iIntCode = tgStations(ilShtt).iMktCode
        tgMktSort(ilShtt).iShttCode = tgStations(ilShtt).iCode
        If tgStations(ilShtt).iMetCode > 0 Then
            tgMSAMktSort(ilShtt).iIntCode = tgStations(ilShtt).iMetCode
            tgMSAMktSort(ilShtt).iShttCode = tgStations(ilShtt).iCode
        End If
        'tgOwnerSort(ilShtt).iIntCode = tgStations(ilShtt).iOwnerArttCode
        'tgOwnerSort(ilShtt).iShttCode = tgStations(ilShtt).iCode
        'tgZipSort(ilShtt).sStr = tgStations(ilShtt).sZip
        'tgZipSort(ilShtt).iShttCode = tgStations(ilShtt).iCode
        tgStateSort(ilShtt).sStr = tgStations(ilShtt).sState
        tgStateSort(ilShtt).iShttCode = tgStations(ilShtt).iCode
        tgFmtSort(ilShtt).iIntCode = tgStations(ilShtt).iFmtCode
        tgFmtSort(ilShtt).iShttCode = tgStations(ilShtt).iCode
        tgTztSort(ilShtt).iIntCode = tgStations(ilShtt).iTztCode
        tgTztSort(ilShtt).iShttCode = tgStations(ilShtt).iCode
    Next ilShtt
    If UBound(tgMktSort) - 1 > 1 Then
        ArraySortTyp fnAV(tgMktSort(), 1), UBound(tgMktSort) - 1, 0, LenB(tgMktSort(1)), 0, -1, 0
    End If
    If UBound(tgMSAMktSort) - 1 > 1 Then
        ArraySortTyp fnAV(tgMSAMktSort(), 1), UBound(tgMSAMktSort) - 1, 0, LenB(tgMSAMktSort(1)), 0, -1, 0
    End If
    'If UBound(tgOwnerSort) - 1 > 1 Then
    '    ArraySortTyp fnAV(tgOwnerSort(), 1), UBound(tgOwnerSort) - 1, 0, LenB(tgOwnerSort(1)), 0, -1, 0
    'End If
    'If UBound(tgZipSort) - 1 > 1 Then
    '    ArraySortTyp fnAV(tgZipSort(), 1), UBound(tgZipSort) - 1, 0, LenB(tgZipSort(1)), 0, LenB(tgZipSort(1).sStr), 0
    'End If
    If UBound(tgStateSort) - 1 > 1 Then
        ArraySortTyp fnAV(tgStateSort(), 1), UBound(tgStateSort) - 1, 0, LenB(tgStateSort(1)), 0, LenB(tgStateSort(1).sStr), 0
    End If
    If UBound(tgFmtSort) - 1 > 1 Then
        ArraySortTyp fnAV(tgFmtSort(), 1), UBound(tgFmtSort) - 1, 0, LenB(tgFmtSort(1)), 0, -1, 0
    End If
    If UBound(tgTztSort) - 1 > 1 Then
        ArraySortTyp fnAV(tgTztSort(), 1), UBound(tgTztSort) - 1, 0, LenB(tgTztSort(1)), 0, -1, 0
    End If
    gObtainStations = True
    Exit Function
gObtainStationsErr2:
    ilRet = 1
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainFormats                  *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgFormats              *
'*                                                     *
'*******************************************************
Function gObtainFormats() As Integer
'
'   ilRet = gObtainFormats ()
'   Where:
'       tgFormats() (I)- MNFCOMPEXT record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlFmt As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlFmt As FMT
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer

    slStamp = gFileDateTime(sgDBPath & "Fmt.mkd")

    'On Error GoTo gObtainFormatsErr2
    'ilRet = 0
    'imLowLimit = LBound(tgFormats)
    'If ilRet <> 0 Then
    '    sgFormatsStamp = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tgFormats).Ptr <> 0 Then
        imLowLimit = LBound(tgFormats)
    Else
        sgFormatsStamp = ""
        imLowLimit = 0
    End If

    If sgFormatsStamp <> "" Then
        If StrComp(slStamp, sgFormatsStamp, 1) = 0 Then
            'If UBound(tgFormats) > 1 Then
                gObtainFormats = True
                Exit Function
            'End If
        End If
    End If
    'ReDim tgFormats(1 To 1) As FMT
    ReDim tgFormats(0 To 0) As FMT
    hlFmt = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlFmt, "", sgDBPath & "Fmt.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainFormats = False
        ilRet = btrClose(hlFmt)
        btrDestroy hlFmt
        Exit Function
    End If
    ilRecLen = Len(tlFmt) 'btrRecordLength(hlFmt)  'Get and save record length
    sgFormatsStamp = slStamp
    ilUpperBound = UBound(tgFormats)
    ilExtLen = Len(tgFormats(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlFmt) 'Obtain number of records
    btrExtClear hlFmt   'Clear any previous extend operation
    ilRet = btrGetFirst(hlFmt, tlFmt, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlFmt)
        btrDestroy hlFmt
        gObtainFormats = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainFormats = False
            ilRet = btrClose(hlFmt)
            btrDestroy hlFmt
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlFmt, llNoRec, -1, "UC", "FMT", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlFmt, ilOffSet, ilRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainFormats = False
        ilRet = btrClose(hlFmt)
        btrDestroy hlFmt
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlFmt)    'Extract record
    ilUpperBound = UBound(tgFormats)
    ilRet = btrExtGetNext(hlFmt, tgFormats(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainFormats = False
            ilRet = btrClose(hlFmt)
            btrDestroy hlFmt
            Exit Function
        End If
        ilUpperBound = UBound(tgFormats)
        ilExtLen = Len(tgFormats(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlFmt, tgMarket(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlFmt, tgFormats(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If Trim(tgFormats(ilUpperBound).sName) <> "" Then
                ilUpperBound = ilUpperBound + 1
                'ReDim Preserve tgFormats(1 To ilUpperBound) As FMT
                ReDim Preserve tgFormats(0 To ilUpperBound) As FMT
            End If
            ilRet = btrExtGetNext(hlFmt, tgFormats(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlFmt, tgFormats(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlFmt)
    btrDestroy hlFmt
    gObtainFormats = True
    Exit Function
gObtainFormatsErr2:
    ilRet = 1
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainMarkets                  *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgMarkets              *
'*                                                     *
'*******************************************************
Function gObtainMarkets() As Integer
'
'   ilRet = gObtainMarkets ()
'   Where:
'       tgMarkets() (I)- MNFCOMPEXT record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlMkt As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlMkt As MKT
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer

    slStamp = gFileDateTime(sgDBPath & "Mkt.mkd")

    'On Error GoTo gObtainMarketsErr2
    'ilRet = 0
    'imLowLimit = LBound(tgMarkets)
    'If ilRet <> 0 Then
    '    sgMarketsStamp = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tgMarkets).Ptr <> 0 Then
        imLowLimit = LBound(tgMarkets)
    Else
        sgMarketsStamp = ""
        imLowLimit = 0
    End If

    If sgMarketsStamp <> "" Then
        If StrComp(slStamp, sgMarketsStamp, 1) = 0 Then
            'If UBound(tgMarkets) > 1 Then
                gObtainMarkets = True
                Exit Function
            'End If
        End If
    End If
    'ReDim tgMarkets(1 To 1) As MKT
    ReDim tgMarkets(0 To 0) As MKT
    hlMkt = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlMkt, "", sgDBPath & "Mkt.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainMarkets = False
        ilRet = btrClose(hlMkt)
        btrDestroy hlMkt
        Exit Function
    End If
    ilRecLen = Len(tlMkt) 'btrRecordLength(hlMkt)  'Get and save record length
    sgMarketsStamp = slStamp
    ilUpperBound = UBound(tgMarkets)
    ilExtLen = Len(tgMarkets(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlMkt) 'Obtain number of records
    btrExtClear hlMkt   'Clear any previous extend operation
    ilRet = btrGetFirst(hlMkt, tlMkt, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlMkt)
        btrDestroy hlMkt
        gObtainMarkets = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainMarkets = False
            ilRet = btrClose(hlMkt)
            btrDestroy hlMkt
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlMkt, llNoRec, -1, "UC", "MKT", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlMkt, ilOffSet, ilRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainMarkets = False
        ilRet = btrClose(hlMkt)
        btrDestroy hlMkt
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlMkt)    'Extract record
    ilUpperBound = UBound(tgMarkets)
    ilRet = btrExtGetNext(hlMkt, tgMarkets(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainMarkets = False
            ilRet = btrClose(hlMkt)
            btrDestroy hlMkt
            Exit Function
        End If
        ilUpperBound = UBound(tgMarkets)
        ilExtLen = Len(tgMarkets(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlMkt, tgMarket(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlMkt, tgMarkets(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If Trim(tgMarkets(ilUpperBound).sName) <> "" Then
                ilUpperBound = ilUpperBound + 1
                'ReDim Preserve tgMarkets(1 To ilUpperBound) As MKT
                ReDim Preserve tgMarkets(0 To ilUpperBound) As MKT
            End If
            ilRet = btrExtGetNext(hlMkt, tgMarkets(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlMkt, tgMarkets(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlMkt)
    btrDestroy hlMkt
    gObtainMarkets = True
    Exit Function
gObtainMarketsErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainOwners                   *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgOwners               *
'*                                                     *
'*******************************************************
Function gObtainOwners() As Integer 'VBC NR
''
''   ilRet = gObtainOwners ()
''   Where:
''       tgOwners() (I)- MNFCOMPEXT record structure to be created
''       ilRet (O)- True = populated; False = error
''
'    Dim slStamp As String    'Mnf date/time stamp
'    Dim hlArtt As Integer        'Mnf handle
'    Dim ilRecLen As Integer     'Record length
'    Dim llNoRec As Long         'Number of records in Mnf
'    Dim tlArtt As ARTT
'    Dim ilExtLen As Integer
'    Dim llRecPos As Long        'Record location
'    Dim ilRet As Integer
'    Dim ilOffset As Integer
'    Dim ilUpperBound As Integer
'
'    slStamp = gFileDateTime(sgDBPath & "Artt.mkd")
'
'    On Error GoTo gObtainOwnersErr2
'    ilRet = 0
'    imLowLimit = LBound(tgOwners)
'    If ilRet <> 0 Then
'        sgOwnersStamp = ""
'    End If
'    On Error GoTo 0
'
'    If sgOwnersStamp <> "" Then
'        If StrComp(slStamp, sgOwnersStamp, 1) = 0 Then
'            'If UBound(tgOwners) > 1 Then
'                gObtainOwners = True
'                Exit Function
'            'End If
'        End If
'    End If
'    ReDim tgOwners(1 To 1) As ARTT
'    hlArtt = CBtrvTable(ONEHANDLE) 'CBtrvTable()
'    ilRet = btrOpen(hlArtt, "", sgDBPath & "Artt.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet <> BTRV_ERR_NONE Then
'        gObtainOwners = False
'        ilRet = btrClose(hlArtt)
'        btrDestroy hlArtt
'        Exit Function
'    End If
'    ilRecLen = Len(tlArtt) 'btrRecordLength(hlArtt)  'Get and save record length
'    sgOwnersStamp = slStamp
'    ilUpperBound = UBound(tgOwners)
'    ilExtLen = Len(tgOwners(ilUpperBound))  'Extract operation record size
'    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlArtt) 'Obtain number of records
'    btrExtClear hlArtt   'Clear any previous extend operation
'    ilRet = btrGetFirst(hlArtt, tlArtt, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'    If ilRet = BTRV_ERR_END_OF_FILE Then
'        ilRet = btrClose(hlArtt)
'        btrDestroy hlArtt
'        gObtainOwners = True
'        Exit Function
'    Else
'        If ilRet <> BTRV_ERR_NONE Then
'            gObtainOwners = False
'            ilRet = btrClose(hlArtt)
'            btrDestroy hlArtt
'            Exit Function
'        End If
'    End If
'    Call btrExtSetBounds(hlArtt, llNoRec, -1, "UC", "ARTT", "") 'Set extract limits (all records)
'    ilOffset = 0
'    ilRet = btrExtAddField(hlArtt, ilOffset, ilRecLen)  'Extract iCode field
'    If ilRet <> BTRV_ERR_NONE Then
'        gObtainOwners = False
'        ilRet = btrClose(hlArtt)
'        btrDestroy hlArtt
'        Exit Function
'    End If
'    'ilRet = btrExtGetNextExt(hlArtt)    'Extract record
'    ilUpperBound = UBound(tgOwners)
'    ilRet = btrExtGetNext(hlArtt, tgOwners(ilUpperBound), ilExtLen, llRecPos)
'    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
'        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
'            gObtainOwners = False
'            ilRet = btrClose(hlArtt)
'            btrDestroy hlArtt
'            Exit Function
'        End If
'        ilUpperBound = UBound(tgOwners)
'        ilExtLen = Len(tgOwners(ilUpperBound))  'Extract operation record size
'        'ilRet = btrExtGetFirst(hlArtt, tgOwners(ilUpperBound), ilExtLen, llRecPos)
'        Do While ilRet = BTRV_ERR_REJECT_COUNT
'            ilRet = btrExtGetNext(hlArtt, tgOwners(ilUpperBound), ilExtLen, llRecPos)
'        Loop
'        Do While ilRet = BTRV_ERR_NONE
'            If (tgOwners(ilUpperBound).sType = "O") And (Trim$(tgOwners(ilUpperBound).sLastName) <> "") Then
'                ilUpperBound = ilUpperBound + 1
'                ReDim Preserve tgOwners(1 To ilUpperBound) As ARTT
'            End If
'            ilRet = btrExtGetNext(hlArtt, tgOwners(ilUpperBound), ilExtLen, llRecPos)
'            Do While ilRet = BTRV_ERR_REJECT_COUNT
'                ilRet = btrExtGetNext(hlArtt, tgOwners(ilUpperBound), ilExtLen, llRecPos)
'            Loop
'        Loop
'    End If
'    ilRet = btrClose(hlArtt)
'    btrDestroy hlArtt
'    gObtainOwners = True
'    Exit Function
'gObtainOwnersErr2:
'    ilRet = 1
'    Resume Next
End Function 'VBC NR

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainStates                  *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgStates              *
'*                                                     *
'*******************************************************
Function gObtainStates() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStamp                                                                               *
'******************************************************************************************

'
'   ilRet = gObtainStates ()
'   Where:
'       tgStates() (I)- MNFCOMPEXT record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim hlSnt As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlSnt As SNT
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer

    'On Error GoTo gObtainStatesErr2
    'ilRet = 0
    'imLowLimit = LBound(tgStates)
    'If ilRet = 0 Then
    '    gObtainStates = True
    '    Exit Function
    'End If
    'On Error GoTo 0
    If PeekArray(tgStates).Ptr <> 0 Then
        imLowLimit = LBound(tgStates)
        gObtainStates = True
        Exit Function
    Else
        imLowLimit = 0
    End If

    ReDim tgStates(0 To 0) As SNT
    hlSnt = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlSnt, "", sgDBPath & "Snt.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainStates = False
        ilRet = btrClose(hlSnt)
        btrDestroy hlSnt
        Exit Function
    End If
    ilRecLen = Len(tlSnt) 'btrRecordLength(hlSnt)  'Get and save record length
    ilUpperBound = UBound(tgStates)
    ilExtLen = Len(tgStates(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSnt) 'Obtain number of records
    btrExtClear hlSnt   'Clear any previous extend operation
    ilRet = btrGetFirst(hlSnt, tlSnt, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlSnt)
        btrDestroy hlSnt
        gObtainStates = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainStates = False
            ilRet = btrClose(hlSnt)
            btrDestroy hlSnt
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlSnt, llNoRec, -1, "UC", "Snt", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlSnt, ilOffSet, ilRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainStates = False
        ilRet = btrClose(hlSnt)
        btrDestroy hlSnt
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlSnt)    'Extract record
    ilUpperBound = UBound(tgStates)
    ilRet = btrExtGetNext(hlSnt, tgStates(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainStates = False
            ilRet = btrClose(hlSnt)
            btrDestroy hlSnt
            Exit Function
        End If
        ilUpperBound = UBound(tgStates)
        ilExtLen = Len(tgStates(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlSnt, tgMarket(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlSnt, tgStates(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If Trim(tgStates(ilUpperBound).sName) <> "" Then
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tgStates(0 To ilUpperBound) As SNT
            End If
            ilRet = btrExtGetNext(hlSnt, tgStates(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSnt, tgStates(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlSnt)
    btrDestroy hlSnt
    gObtainStates = True
    Exit Function
gObtainStatesErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainTimeZones                  *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgTimeZones              *
'*                                                     *
'*******************************************************
Function gObtainTimeZones() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStamp                                                                               *
'******************************************************************************************

'
'   ilRet = gObtainTimeZones ()
'   Where:
'       tgTimeZones() (I)- MNFCOMPEXT record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim hlTzt As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlTzt As TZT
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer

    'On Error GoTo gObtainTimeZonesErr2
    'ilRet = 0
    'imLowLimit = LBound(tgTimeZones)
    'If ilRet = 0 Then
    '    gObtainTimeZones = True
    '    Exit Function
    'End If
    'On Error GoTo 0
    If PeekArray(tgTimeZones).Ptr <> 0 Then
        imLowLimit = LBound(tgTimeZones)
        gObtainTimeZones = True
        Exit Function
    Else
        imLowLimit = 0
    End If

    ReDim tgTimeZones(0 To 0) As TZT
    hlTzt = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlTzt, "", sgDBPath & "Tzt.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainTimeZones = False
        ilRet = btrClose(hlTzt)
        btrDestroy hlTzt
        Exit Function
    End If
    ilRecLen = Len(tlTzt) 'btrRecordLength(hlTzt)  'Get and save record length
    ilUpperBound = UBound(tgTimeZones)
    ilExtLen = Len(tgTimeZones(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlTzt) 'Obtain number of records
    btrExtClear hlTzt   'Clear any previous extend operation
    ilRet = btrGetFirst(hlTzt, tlTzt, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlTzt)
        btrDestroy hlTzt
        gObtainTimeZones = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainTimeZones = False
            ilRet = btrClose(hlTzt)
            btrDestroy hlTzt
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlTzt, llNoRec, -1, "UC", "Tzt", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlTzt, ilOffSet, ilRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainTimeZones = False
        ilRet = btrClose(hlTzt)
        btrDestroy hlTzt
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlTzt)    'Extract record
    ilUpperBound = UBound(tgTimeZones)
    ilRet = btrExtGetNext(hlTzt, tgTimeZones(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainTimeZones = False
            ilRet = btrClose(hlTzt)
            btrDestroy hlTzt
            Exit Function
        End If
        ilUpperBound = UBound(tgTimeZones)
        ilExtLen = Len(tgTimeZones(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlTzt, tgMarket(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlTzt, tgTimeZones(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If Trim(tgTimeZones(ilUpperBound).sName) <> "" Then
                ilUpperBound = ilUpperBound + 1
                ReDim Preserve tgTimeZones(0 To ilUpperBound) As TZT
            End If
            ilRet = btrExtGetNext(hlTzt, tgTimeZones(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlTzt, tgTimeZones(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlTzt)
    btrDestroy hlTzt
    gObtainTimeZones = True
    Exit Function
gObtainTimeZonesErr2:
    ilRet = 1
    Resume Next
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainMSAMarkets               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Populate tgMSAMarkets           *
'*                                                     *
'*******************************************************
Function gObtainMSAMarkets() As Integer
'
'   ilRet = gObtainMSAMarkets ()
'   Where:
'       tgMSAMarkets() (I)- MNFCOMPEXT record structure to be created
'       ilRet (O)- True = populated; False = error
'
    Dim slStamp As String    'Mnf date/time stamp
    Dim hlMet As Integer        'Mnf handle
    Dim ilRecLen As Integer     'Record length
    Dim llNoRec As Long         'Number of records in Mnf
    Dim tlMet As MET
    Dim ilExtLen As Integer
    Dim llRecPos As Long        'Record location
    Dim ilRet As Integer
    Dim ilOffSet As Integer
    Dim ilUpperBound As Integer

    slStamp = gFileDateTime(sgDBPath & "Met.mkd")

    'On Error GoTo gObtainMSAMarketsErr2
    'ilRet = 0
    'imLowLimit = LBound(tgMSAMarkets)
    'If ilRet <> 0 Then
    '    sgMSAMarketsStamp = ""
    'End If
    'On Error GoTo 0
    If PeekArray(tgMSAMarkets).Ptr <> 0 Then
        imLowLimit = LBound(tgMSAMarkets)
    Else
        sgMSAMarketsStamp = ""
        imLowLimit = 0
    End If

    If sgMSAMarketsStamp <> "" Then
        If StrComp(slStamp, sgMSAMarketsStamp, 1) = 0 Then
            'If UBound(tgMSAMarkets) > 1 Then
                gObtainMSAMarkets = True
                Exit Function
            'End If
        End If
    End If
    'ReDim tgMSAMarkets(1 To 1) As MET
    ReDim tgMSAMarkets(0 To 0) As MET
    hlMet = CBtrvTable(ONEHANDLE) 'CBtrvTable()
    ilRet = btrOpen(hlMet, "", sgDBPath & "Met.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gObtainMSAMarkets = False
        ilRet = btrClose(hlMet)
        btrDestroy hlMet
        Exit Function
    End If
    ilRecLen = Len(tlMet) 'btrRecordLength(hlMet)  'Get and save record length
    sgMSAMarketsStamp = slStamp
    ilUpperBound = UBound(tgMSAMarkets)
    ilExtLen = Len(tgMSAMarkets(ilUpperBound))  'Extract operation record size
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlMet) 'Obtain number of records
    btrExtClear hlMet   'Clear any previous extend operation
    ilRet = btrGetFirst(hlMet, tlMet, ilRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_END_OF_FILE Then
        ilRet = btrClose(hlMet)
        btrDestroy hlMet
        gObtainMSAMarkets = True
        Exit Function
    Else
        If ilRet <> BTRV_ERR_NONE Then
            gObtainMSAMarkets = False
            ilRet = btrClose(hlMet)
            btrDestroy hlMet
            Exit Function
        End If
    End If
    Call btrExtSetBounds(hlMet, llNoRec, -1, "UC", "Met", "") 'Set extract limits (all records)
    ilOffSet = 0
    ilRet = btrExtAddField(hlMet, ilOffSet, ilRecLen)  'Extract iCode field
    If ilRet <> BTRV_ERR_NONE Then
        gObtainMSAMarkets = False
        ilRet = btrClose(hlMet)
        btrDestroy hlMet
        Exit Function
    End If
    'ilRet = btrExtGetNextExt(hlMet)    'Extract record
    ilUpperBound = UBound(tgMSAMarkets)
    ilRet = btrExtGetNext(hlMet, tgMSAMarkets(ilUpperBound), ilExtLen, llRecPos)
    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
            gObtainMSAMarkets = False
            ilRet = btrClose(hlMet)
            btrDestroy hlMet
            Exit Function
        End If
        ilUpperBound = UBound(tgMSAMarkets)
        ilExtLen = Len(tgMSAMarkets(ilUpperBound))  'Extract operation record size
        'ilRet = btrExtGetFirst(hlMet, tgMSAMarket(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hlMet, tgMSAMarkets(ilUpperBound), ilExtLen, llRecPos)
        Loop
        Do While ilRet = BTRV_ERR_NONE
            If Trim(tgMSAMarkets(ilUpperBound).sName) <> "" Then
                ilUpperBound = ilUpperBound + 1
                'ReDim Preserve tgMSAMarkets(1 To ilUpperBound) As MET
                ReDim Preserve tgMSAMarkets(0 To ilUpperBound) As MET
            End If
            ilRet = btrExtGetNext(hlMet, tgMSAMarkets(ilUpperBound), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlMet, tgMSAMarkets(ilUpperBound), ilExtLen, llRecPos)
            Loop
        Loop
    End If
    ilRet = btrClose(hlMet)
    btrDestroy hlMet
    gObtainMSAMarkets = True
    Exit Function
gObtainMSAMarketsErr2:
    ilRet = 1
    Resume Next
End Function
