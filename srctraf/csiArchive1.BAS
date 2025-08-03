Attribute VB_Name = "csiArchive1"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: csiArchive.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Global Variables for Set Advertiser/Agency Credit
Option Explicit
Option Compare Text

Declare Function csiArchiveInit Lib "csi_utils.dll" (ByVal slDBPath$) As Integer
Declare Function csiArchiveAndRemoveAst Lib "csi_utils.dll" (ByVal llAttCode&, ByVal slStartDate$, ByVal slEndDate$, ByVal ilExportType%, ByRef llLstCount&, ByRef llSdfCount&) As Integer
Declare Function csiOpenArchiveFileName Lib "csi_utils.dll" (ByVal slArchivePathFileName$) As Integer
Declare Function csiCloseArchiveFile Lib "csi_utils.dll" () As Integer
Declare Function csiArchiveDeInit Lib "csi_utils.dll" () As Integer


'Affiliate Spot Archive
    'Affiliate Spots
    Public hgDBAst As Integer
    Public tgDBAst As AST            'AST record image
    Public tgDBAstSrchKey1 As ASTKEY1 'AST key record image
    Public igDBAstRecLen As Integer     'AST record length
    
    'Marketorn export
    Public hgDBAet As Integer
    Public tgDBAet As AET            'AET record image
    Public tgDBAetSrchKey2 As LONGKEY0 'AET key record image
    Public igDBAetRecLen As Integer     'AET record length
    
    'Affiliate Log
    Public hgDBLst As Integer
    Public tgDBLst As LST            'LST record image
    Public tgDBLstSrchKey0 As LONGKEY0 'LST key record image
    Public tgDBLstSrchKey2 As LSTKEY2 'LST key record image
    Public igDBLstRecLen As Integer     'LST record length
    
    'CP Posting
    Public hgDBCptt As Integer
    Public tgDBCptt As CPTT            'CPTT record image
    Public tgDBCpttSrchKey1 As CPTTKEY1 'CPTT key record image
    Public tgDBCpttSrchKey2 As CPTTKEY2 'CPTT key record image
    Public igDBCpttRecLen As Integer     'CPTT record length
    
    'CCT CP Comments
    Public hgDBCct As Integer
    Public tgDBCct As CCT               'CCT record image
    Public tgDBCctSrchKey0 As LONGKEY0   'CCT key record image
    Public igDBCctRecLen As Integer     'CCT record length
    
    'Agreements
    Public hgDBAtt As Integer
    Public tgDBAtt As ATT            'ATT record image
    Public tgDBAttSrchKey0 As LONGKEY0 'ATT key record image
    Public tgDBAttSrchKey1 As INTKEY0
    Public igDBAttRecLen As Integer     'ATT record length

    'Pedge Info
    Public hgDBDat As Integer
    Public tgDBDat As DAT            'ATT record image
    Public tgDBDatSrchKey0 As LONGKEY0 'ATT key record image
    Public tgDBDatSrchKey1 As LONGKEY0 'ATT key record image
    Public igDBDatRecLen As Integer     'ATT record length

    'Station
    Public hgDBShtt As Integer
    Public tgDBShtt As SHTT            'ATT record image
    Public tgDBShttSrchKey0 As INTKEY0 'ATT key record image
    Public igDBShttRecLen As Integer     'ATT record length

    'Export Summary
    Public hgDBEsf As Integer
    Public tgDBEsf As ESF            'ESF record image
    Public tgDBEsfSrchKey0 As LONGKEY0 'ESF key record image
    Public igDBEsfRecLen As Integer     'ESF record length
    
    'Export Detail
    Public hgDBEdf As Integer
    Public tgDBEdf As EDF            'EDF record image
    Public tgDBEdfSrchKey1 As LONGKEY0 'EDF key record image
    Public igDBEdfRecLen As Integer     'EDF record length
    
    'Web Processing Info
    Public hgDBWebl As Integer
    Public tgDBWebl As WEBL            'WEBL record image
    Public tgDBWeblSrchKey0 As LONGKEY0 'WEBL key record image
    Public igDBWeblRecLen As Integer     'WEBL record length

'Proposal Archive
    'Contract header
    Public tgDBChfSrchKey0 As LONGKEY0    'chf key record image
    Public tgDBChfSrchKey1 As CHFKEY1    'chf key record image
    Public tgDBChfSrchKey2 As INTKEY0    'chf key record image
    Public tgDBChfSrchKey3 As CHFKEY3    'chf key record image
    Public hgDBChf As Integer    'Contract header file handle
    Public igDBChfRecLen As Integer        'Chf record length
    Public tgDBChf As CHF
    
    'Contract Lines
    Public tgDBClf As CLF        'Clf record image
    Public tgDBClfSrchKey0 As CLFKEY0
    Public tgDBClfSrchKey1 As CLFKEY1
    Public tgDBClfSrchKey2 As LONGKEY0    'Clf key record image
    Public hgDBClf As Integer    'Contract Line file handle
    Public igDBClfRecLen As Integer        'CLF record length
    
    'Contract Lines flights
    Public tgDBCff As CFF        'Cff record image
    Public tgDBCffSrchKey As CFFKEY0
    Public hgDBCff As Integer    'Contract Line file handle
    Public igDBCffRecLen As Integer        'CfF record length
    
    'Contract Lines vehicles
    Public tgDBVsf As VSF        'Clf record image
    Public tgDBVsfSrchKey As LONGKEY0
    Public hgDBVsf As Integer    'Contract Line file handle
    Public igDBVsfRecLen As Integer        'CLF record length
    
    'Contract Lines flights
    Public tgDBSbf As SBF        'Cff record image
    Public tgDBSbfSrchKey0 As SBFKEY0
    Public hgDBSbf As Integer    'Contract Line file handle
    Public igDBSbfRecLen As Integer        'CfF record length
    
    'Comment record-Header/Line
    Public hgDBCxf As Integer
    Public tgDBCxf As CXF            'CXF record image
    Public tgDBCxfSrchKey As LONGKEY0  'CXF key record image
    Public igDBCxfRecLen As Integer      'CXF record length
    
    'Extra Fields
    Public hgDBEff As Integer
    Public tgDBEff As EFF            'CGF record image
    Public tgDBEffSrchKey As LONGKEY0  'CGF key record image
    Public igDBEffRecLen As Integer      'CGF record length
    
    'Game
    Public hgDBCgf As Integer
    Public tgDBCgf As CGF            'CGF record image
    Public tgDBCgfSrchKey As LONGKEY0  'CGF key record image
    Public tgDBCgfSrchKey1 As CGFKEY1
    Public igDBCgfRecLen As Integer      'CGF record length
    
    'Multi-Media
    Public hgDBMsf As Integer
    Public tgDBMsf As MSF            'MSF record image
    Public tgDBMsfSrchKey As LONGKEY0  'MSF key record image
    Public tgDBMsfSrchKey2 As MSFKEY2
    Public igDBMsfRecLen As Integer      'MSF record length
    
    'Multi-Media
    Public hgDBMgf As Integer
    Public tgDBMgf As MGF            'MGF record image
    Public tgDBMgfSrchKey As LONGKEY0  'MGF key record image
    Public tgDBMgfSrchKey1 As MGFKEY1
    Public igDBMgfRecLen As Integer      'MGF record length
    
'Traffic Spot Archive
    Public hgDBSdf As Integer
    Public tgDBSdf As SDF
    Public tgDBSdfSrchKey3 As LONGKEY0
    Public tgDBSdfSrchKey4 As SDFKEY4
    Public tgDBSdfSrchKey5 As LONGKEY0
    Public igDBSdfRecLen As Integer

    Public hgDBSmf As Integer
    Public tgDBSmf As SMF
    Public tgDBSmfSrchKey2 As LONGKEY0
    Public igDBSmfRecLen As Integer
    
    Public hgDBStf As Integer
    Public tgDBStf As STF
    Public tgDBStfSrchKey0 As STFKEY0
    Public igDBStfRecLen As Integer
    
    Public hgDBMtf As Integer
    Public tgDBMtf As MTF
    Public tgDBMtfSrchKey As LONGKEY0
    Public igDBMtfRecLen As Integer
    
    Public hgDBPsf As Integer
    Public tgDBPsf As SDF
    Public tgDBPsfSrchKey4 As SDFKEY4
    Public igDBPsfRecLen As Integer
    
    Public hgDBRsf As Integer
    Public tgDBRsf As RSF
    Public tgDBRsfSrchKey0 As LONGKEY0
    Public tgDBRsfSrchKey1 As LONGKEY0
    Public igDBRsfRecLen As Integer
    
    Public hgDBTzf As Integer
    Public tgDBTzf As TZF
    Public tgDBTzfSrchKey As LONGKEY0
    Public igDBTzfRecLen As Integer
    
'Programming Archive
    'Log Calender
    Public tgDBLcfSrchKey2 As LCFKEY2
    Public hgDBLcf As Integer
    Public igDBLcfRecLen As Integer
    Public tgDBLcf As LCF
    
    'Spot Summary
    Public tgDBSsf As SSF
    Public tgDBSsfSrchKey2 As SSFKEY2
    Public hgDBSsf As Integer
    Public igDBSsfRecLen As Integer
    
    'Selling/Airing Links
    Public tgDBVlf As VLF
    Public tgDBVlfSrchKey0 As VLFKEY0
    Public tgDBVlfSrchKey1 As VLFKEY1
    Public hgDBVlf As Integer
    Public igDBVlfRecLen As Integer
    
    'Conflict Links
    Public tgDBVcf As VCF
    Public tgDBVcfSrchKey0 As VCFKEY0
    Public hgDBVcf As Integer
    Public igDBVcfRecLen As Integer
    
    'Delivery Links
    Public tgDBDlf As DLF
    Public tgDBDlfSrchKey0 As DLFKEY0
    Public hgDBDlf As Integer
    Public igDBDlfRecLen As Integer
    
    'Engineering Link
    Public tgDBEgf As DLF
    Public tgDBEgfSrchKey0 As DLFKEY0
    Public hgDBEgf As Integer
    Public igDBEgfRecLen As Integer
    
    'Library Titles
    Public hgDBLtf As Integer
    Public tgDBLtf As LTF
    Public tgDBLtfSrchKey As INTKEY0
    Public igDBLtfRecLen As Integer
    
    'Library versions
    Public hgDBLvf As Integer
    Public tgDBLvf As LVF
    Public tgDBLvfSrchKey0 As LONGKEY0
    Public tgDBLvfSrchKey1 As LVFKEY1
    Public igDBLvfRecLen As Integer
    
    'Library Events
    Public hgDBLef As Integer
    Public tgDBLef As LEF
    Public tgDBLefSrchKey1 As LONGKEY0
    Public igDBLefRecLen As Integer
    
    'Library Events Comments
    Public hgDBCef As Integer
    Public tgDBCef As CEF            'CEF record image
    Public tgDBCefSrchKey0 As LONGKEY0  'CEF key record image
    Public igDBCefRecLen As Integer      'CEF record length
    
'Copy Archive
    Public hgDBCrf As Integer
    Public tgDBCrf As CRF
    Public tgDBCrfSrchKey As LONGKEY0
    Public tgDBCrfSrchKey1 As CRFKEY1
    Public igDBCrfRecLen As Integer

    Public hgDBCnf As Integer
    Public tgDBCnf As CNF
    Public tgDBCnfSrchKey As CNFKEY0
    Public igDBCnfRecLen As Integer
    
    Public hgDBCif As Integer
    Public tgDBCif As CIF
    Public tgDBCifSrchKey As LONGKEY0
    Public igDBCifRecLen As Integer
    
    Public hgDBCsf As Integer
    Public tgDBCsf As CSF
    Public tgDBCsfSrchKey As LONGKEY0
    Public igDBCsfRecLen As Integer
    
    Public hgDBCpf As Integer
    Public tgDBCpf As CPF
    Public tgDBCpfSrchKey As LONGKEY0
    Public igDBCpfRecLen As Integer
    
    Public hgDBCyf As Integer
    Public tgDBCyf As CYF
    Public tgDBCyfSrchKey As CYFKEY0
    Public igDBCyfRecLen As Integer
    
    Public hgDBRaf As Integer
    Public tgDBRaf As RAF
    Public tgDBRafSrchKey As LONGKEY0
    Public igDBRafRecLen As Integer
        
    Public hgDBSef As Integer
    Public tgDBSef As SEF
    Public tgDBSefSrchKey1 As LONGKEY0
    Public igDBSefRecLen As Integer
    
    Public hgDBSif As Integer
    Public tgDBSif As SIF
    Public tgDBSifSrchKey As LONGKEY0
    Public tgDBSifSrchKey1 As INTKEY0
    Public igDBSifRecLen As Integer
        
    Public hgDBCaf As Integer
    Public tgDBCaf As CAF
    Public tgDBCafSrchKey As LONGKEY0
    Public tgDBCafSrchKey1 As CAFKEY1
    Public igDBCafRecLen As Integer
        
'Advertiser and Agency Archive
    'Advertiser
    Public hgDBAdf As Integer
    Public tgDBAdf As ADF            'ADF record image
    Public tgDBAdfSrchKey As INTKEY0 'ADF key record image
    Public igDBAdfRecLen As Integer     'ADF record length
    
    'Agency
    Public hgDBAgf As Integer
    Public tgDBAgf As AGF            'AGF record image
    Public tgDBAgfSrchKey As INTKEY0 'AGF key record image
    Public igDBAgfRecLen As Integer     'AGF record length

    Public hgDBPrf As Integer
    Public tgDBPrf As PRF            'PRF record image
    Public tgDBPrfSrchKey1 As INTKEY0 'PRF key record image
    Public igDBPrfRecLen As Integer     'PRF record length
    
    Public hgDBPnf As Integer
    Public tgDBPnf As PNF
    Public tgDBPnfSrchKey1 As INTKEY0
    Public tgDBPnfSrchKey2 As INTKEY0
    Public igDBPnfRecLen As Integer

    Public hgDBRvf As Integer
    Public tgDBRvf As RVF
    Public tgDBRvfSrchKey0 As INTKEY0
    Public tgDBRvfSrchKey1 As INTKEY0
    Public igDBRvfRecLen As Integer

    Public hgDBPhf As Integer
    Public tgDBPhfSrchKey3 As RVFKEY3

    Public hgDBCdf As Integer
    Public tgDBCdf As CDF
    Public tgDBCdfSrchKey0 As INTKEY0
    Public tgDBCdfSrchKey1 As INTKEY0
    Public igDBCdfRecLen As Integer

'Blackout
    'Blackout
    Public hgDBBof As Integer
    Public tgDBBof As BOF
    Public tgDBBofSrchKey0 As LONGKEY0
    Public tgDBBofSrchKey1 As BOFKEY1
    Public igDBBofRecLen As Integer

'Projections
    'Projection
    Public hgDBPjf As Integer
    Public tgDBPjf As PJF
    Public tgDBPjfSrchKey0 As PJFKEY0
    Public tgDBPjfSrchKey1 As LONGKEY0
    Public igDBPjfRecLen As Integer
    
''Report Queue
'    'Report Queue
'    Public hgDBRqf As Integer
'    Public tgDBRqf As RQF
'    Public tgDBRqfSrchKey0 As LONGKEY0
'    Public igDBRqfRecLen As Integer
'
'    'Report Parameters
'    Public hgDBRff As Integer
'    Public tgDBRff As RFF
'    Public tgDBRffSrchKey0 As LONGKEY0
'    Public tgDBRffSrchKey1 As RFFKEY1
'    Public igDBRffRecLen As Integer
'
'Type CRFARCHIVEINFO
'    sKey As String * 10
'    lCrfCode As Long
'    lDate As Long
'End Type
'
'Type DBVEHICLE
'    sType As String * 1
'    iVefCode As Integer
'End Type
'
'Public tgStations() As SHTT
'Public sgStationsStamp As String
'
'
'Type STATUSTYPES
'    sName As String * 30
'    iPledged As Integer '0=Live; 1=Delayed; 2=Not Carried; 3=No Pledge
'    iStatus As Integer
'End Type
'
'Public tgStatusTypes(0 To 12) As STATUSTYPES
'
'Type ADFCREDIT
'    iCode As Integer
'    sUnbilled As String * 12
'    sProjected As String * 12
'    s1Wk As String * 12
'    s4Wks As String * 12
'    s13Wks As String * 12
'End Type
'Type AGFCREDIT
'    iCode As Integer
'    sUnbilled As String * 12
'    sProjected As String * 12
'    s1Wk As String * 12
'    s4Wks As String * 12
'    s13Wks As String * 12
'End Type
'Type SDFCREDITEXT
'    lCode As Long           'Spot Reference Number
'    iVefCode As Integer     'Vehicle Code (combos not allowed)
'    lChfCode As Long        'Contract code (used to get Agency; Advertiser)
'    iLineNo As Integer      'Line number
'    iDate(0 To 1) As Integer    'Schedule or missed Date of spot
'    sSchStatus As String * 1    'S=Scheduled, M=Missed, R=Ready to schd MG, U=Unscheduled MG,
'    sPriceType As String * 1 'L=Line; P=Post Log N/C  old=>T=True; N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
'    sBill As String * 1     'Billed status (Y= Yes; N=No)
'End Type
'
'Type SCR
'    lCode As Long                   'auto code
'    iGenDate(0 To 1) As Integer     'Report generation date
'    lGenTime As Long                'report generation time
'    iStrLen As Integer  'String length (required by LVar)
'    sScript As String * 5002   'Last two bytes after the comment must be 0
'
'End Type
'
'Type SCRKEY1
'    iGenDate(0 To 1) As Integer   'Generation Date
'    lGenTime As Long
'End Type



Public Sub gCreateAlert(slMessage As String, ilLogMsg)
    Dim ilRet As Integer
    
    'tgDBCef.iStrLen = Len(Trim$(slMessage))
    tgDBCef.sComment = Trim$(slMessage) & Chr$(0) '& Chr$(0) 'sgTB
    igDBCefRecLen = Len(tgDBCef)    '5 + Len(Trim$(tgDBCef.sComment)) + 2   '5 = fixed record length; 2 is the length of the record which is part of the variable record
    tgDBCef.lCode = 0 'Autoincrement
    ilRet = btrInsert(hgDBCef, tgDBCef, igDBCefRecLen, INDEXKEY0)
    If ilRet = BTRV_ERR_NONE Then
        ilRet = gAlertAdd("M", "", 0, 0, Trim$(Str$(tgDBCef.lCode)))
    End If
    If ilLogMsg Then
        gLogMsg slMessage, "RepNetLink.txt", False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gFindMatch                      *
'*                                                     *
'*             Created:10/20/93      By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Find matching match within list*
'*                      box                            *
'*                                                     *
'*******************************************************
Sub gFindMatch(slInMatch As String, ilFirstAllowedIndex As Integer, lbcList As Control)
'
'   gFindMatch slInMatch, ilFirstAllowedIndex, lbcList
'   Where:
'       slInMatch(I)- String to be found
'       ilFirstAllowedIndex(I)- First legal index
'       lbcList(I)- list box containing strings
'
'       lbcList.LastFound(O)- Found index
'
    Dim ilFound As Integer
    Dim slMatch As String
    Dim ilPass As Integer   'Pass: 0= Trim string; 1=Leave leading blanks
    Dim ilRow As Integer
    Dim slNext As String

    For ilPass = 0 To 1 Step 1
        ilFound = False
        If ilPass = 0 Then
            slMatch = Trim$(slInMatch)
        Else
            slMatch = RTrim$(slInMatch)
        End If
        gFndFirst lbcList, slMatch
        Do
            If gLastFound(lbcList) >= ilFirstAllowedIndex Then
                If StrComp(lbcList.List(gLastFound(lbcList)), slMatch, 1) = 0 Then
                    Exit Sub
                Else
                    ilFound = False
                    gFndNext lbcList, slMatch
                End If
            Else
                If ilPass = 1 Then
                    Exit Sub
                Else
                    Exit Do
                End If
            End If
        Loop While Not ilFound
    Next ilPass
End Sub


Public Sub gFndFirst(lbcList As Control, slInMatch As String)
    cgList = lbcList
    If TypeOf lbcList Is ComboBox Then
        igFoundRow = SendMessageByString(lbcList.hwnd, CB_FINDSTRING, -1, slInMatch)
    Else
        igFoundRow = SendMessageByString(lbcList.hwnd, LB_FINDSTRING, -1, slInMatch)
    End If
End Sub

Public Sub gFndNext(lbcList As Control, slInMatch As String)
    Dim slNext As String
    Dim ilTestRow As Integer
    
    If cgList <> lbcList Then
        igFoundRow = -1
        Exit Sub
    End If
    ilTestRow = igFoundRow + 1
    Do While ilTestRow < lbcList.ListCount
        slNext = lbcList.List(ilTestRow)
        If InStr(1, slNext, slInMatch, vbTextCompare) = 1 Then
            igFoundRow = ilTestRow
            Exit Sub
        End If
        ilTestRow = ilTestRow + 1
    Loop
    igFoundRow = -1

End Sub


Public Function gLastFound(lbcList As Control) As Integer
    If cgList <> lbcList Then
        gLastFound = -1
    Else
        gLastFound = igFoundRow
    End If
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:gSetPhoneNo                     *
'*                                                     *
'*             Created:4/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Insert phone number into mask   *
'*                     field.  Works with and without  *
'*                     extension                       *
'*                                                     *
'*******************************************************
Sub gSetPhoneNo(slPhone As String, mkcCtrl As Control)
'
'   gSetPhoneNo mkcCtrl, sPhone
'   Where:
'       sPhone (I)- Phone number to be inserted into the mask control
'       mkcCtrl (O)- Mask edit control to receive the phone number
'

    'Clear field as blanks will not place old values
    If Len(slPhone) <= 10 Then
        Traffic!mkcPhone.Mask = "(AAA) AAA-AAAA"
        Traffic!mkcPhone.MaxLength = 14
    Else
        Traffic!mkcPhone.Mask = "(AAA) AAA-AAAA Ext(AAAA)"
        Traffic!mkcPhone.MaxLength = 24
    End If
    Traffic!mkcPhone.SelStart = 0
    Traffic!mkcPhone.SelLength = Traffic!mkcPhone.MaxLength
    Traffic!mkcPhone.SelText = ""
    Traffic!mkcPhone.SelStart = 1
    Traffic!mkcPhone.SelLength = 3
    Traffic!mkcPhone.SelText = Mid$(slPhone, 1, 3)
    Traffic!mkcPhone.SelStart = 6
    Traffic!mkcPhone.SelLength = 3
    Traffic!mkcPhone.SelText = Mid$(slPhone, 4, 3)
    Traffic!mkcPhone.SelStart = 10
    Traffic!mkcPhone.SelLength = 4
    Traffic!mkcPhone.SelText = Mid$(slPhone, 7, 4)
    If Len(slPhone) <= 10 Then
        mkcCtrl.Text = Traffic!mkcPhone.Text
        Exit Sub
    End If
    Traffic!mkcPhone.SelStart = 19
    Traffic!mkcPhone.SelLength = 4
    Traffic!mkcPhone.SelText = Mid$(slPhone, 11, 4)
    mkcCtrl.Text = Traffic!mkcPhone.Text
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gFormatPhoneNo                  *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove _ characters from number *
'*                                                     *
'*******************************************************
Function gFormatPhoneNo(slPhone As String) As String
    Dim slStr As String
    Dim ilPos As Integer
    Dim slTmp As String
    Dim slPhoneImage As String

    If Len(slPhone) <= 10 Then
        slPhoneImage = "(___) ___-____"
        gSetPhoneNo slPhone, Traffic!mkcTempFaxNo
        slStr = Traffic!mkcTempFaxNo.Text
    Else
        slPhoneImage = "(___) ___-____ Ext(____)"
        gSetPhoneNo slPhone, Traffic!mkcTempPhoneNo
        slStr = Traffic!mkcTempPhoneNo.Text
    End If
    If slStr = slPhoneImage Then
        slStr = ""
    Else
        If InStr(slStr, "(____)") <> 0 Then 'Test for missing extension
            ilPos = InStr(slStr, "Ext(")
            slStr = Left$(slStr, ilPos - 1)
        End If
        If InStr(slStr, "(___)") <> 0 Then  'Test for missing area code
            slStr = right$(slStr, Len(slStr) - 5)
        End If
        ilPos = InStr(slStr, "_")
        Do While ilPos > 0
            slTmp = Left$(slStr, ilPos - 1) & Mid$(slStr, ilPos + 1)
            slStr = slTmp
            ilPos = InStr(slStr, "_")
        Loop
    End If
    gFormatPhoneNo = slStr
End Function

Function gDeleteCxf(llCxfCode As Long) As Integer
    Dim ilRet As Integer
    
    If llCxfCode > 0 Then
        igDBCxfRecLen = Len(tgDBCxf) '5011
        tgDBCxfSrchKey.lCode = llCxfCode
        ilRet = btrGetEqual(hgDBCxf, tgDBCxf, igDBCxfRecLen, tgDBCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hgDBCxf)
        End If
        gDeleteCxf = ilRet
    Else
        ilRet = BTRV_ERR_NONE
    End If
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name: gCtrlGotFocus                  *
'*                                                     *
'*             Created:4/12/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Text in the control is          *
'*                     highlighted                     *
'*                                                     *
'*******************************************************
Sub gCtrlGotFocus(Ctrl As Control)
'
'   gCtrlGotFocus Ctrl
'   Where:
'       Ctrl (I)- control for which text will be highlighted
'

'    Traffic.plcHelp.Caption = " " & Ctrl.Tag
    On Error GoTo gCtrlGotFocusErr:
    If TypeOf Ctrl Is TextBox Then
        Ctrl.SelStart = 0
        Ctrl.SelLength = Len(Ctrl.Text)
    ElseIf TypeOf Ctrl Is MaskEdBox Then
        Ctrl.SelStart = 0
        Ctrl.SelLength = Len(Ctrl.Text)
    ElseIf TypeOf Ctrl Is ComboBox Then
        Ctrl.SelStart = 0
        Ctrl.SelLength = Len(Ctrl.Text)
    End If
    Exit Sub
gCtrlGotFocusErr:
    Exit Sub
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:gCheckKeyasc                    *
'*                                                     *
'*             Created:4/16/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Check for illegal key           *
'*                                                     *
'*******************************************************
Function gCheckEnteredChar(ilKeyAscii As Integer) As Integer
    '92=Backslash (\);47=Forwardslash (/); 94=Caret (^); 124=Verical Bar (|); 91=Sq Bracket ([); 93=Sq Bracket (]); 59=Semi-colon(;)
    'If (ilKeyAscii < 32) Or (ilKeyAscii > 126) Or (ilKeyAscii = 92) Or (ilKeyAscii = 94) Or (ilKeyAscii = 124) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Or (ilKeyAscii = 59) Then
    'Allow LF, CR and Backspace.
    '                 LF                     CR                     Backspace
    If (ilKeyAscii <> 10) And (ilKeyAscii <> 13) And (ilKeyAscii <> 8) Then
    
        '                                      "                    \                    /                    ^                    |                     [                    ]                    ;
        If (ilKeyAscii < 32) Or (ilKeyAscii > 126) Or (ilKeyAscii = 34) Or (ilKeyAscii = 92) Or (ilKeyAscii = 47) Or (ilKeyAscii = 94) Or (ilKeyAscii = 124) Or (ilKeyAscii = 91) Or (ilKeyAscii = 93) Or (ilKeyAscii = 59) Then
            Beep
            gCheckEnteredChar = False
            Exit Function
        End If
    End If
    gCheckEnteredChar = True
End Function
