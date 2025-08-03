VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form ExportGetPaid 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate Export GetPaid"
   ClientHeight    =   2100
   ClientLeft      =   465
   ClientTop       =   2340
   ClientWidth     =   7860
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "ExportGetPaid.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2100
   ScaleWidth      =   7860
   Begin VB.Timer tmcSetTime 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6615
      Top             =   1260
   End
   Begin VB.ListBox lbcACodes 
      Height          =   1620
      ItemData        =   "ExportGetPaid.frx":08CA
      Left            =   180
      List            =   "ExportGetPaid.frx":08D1
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   270
      Visible         =   0   'False
      Width           =   1695
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   285
      Left            =   2085
      TabIndex        =   4
      Top             =   810
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmcTimer 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   6345
      Top             =   165
   End
   Begin VB.CommandButton cmcGen 
      Appearance      =   0  'Flat
      Caption         =   "&Generate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2385
      TabIndex        =   1
      Top             =   1140
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   1140
      Width           =   1050
   End
   Begin VB.Label lacTo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   150
      TabIndex        =   0
      Top             =   1620
      Width           =   5820
   End
   Begin VB.Label lacInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1860
      TabIndex        =   3
      Top             =   420
      Visible         =   0   'False
      Width           =   3390
   End
End
Attribute VB_Name = "ExportGetPaid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software®, Do not copy
'
' File Name: ExportGetPaid.Frm
'
' Release: 1.0
' Modified By: Rick Gibson  07/27/2007
' Description:
'   This file contains the GetPaid Export
Option Explicit
Option Compare Text

Dim lmTotalNoRecs As Long
Dim lmProcessedNoRecs As Long
Dim hmTo As Integer   'From file hanle
Dim hmToM As Integer   'From file hanle
Dim tmRvf As RVF            'RVF or PHF record image
Dim tmRvfSrchKey As RVFKEY0 'RVF key record image
Dim imRvfRecLen As Integer  'RVF record length
Dim hmRvf As Integer        'RVF or PHF file handle
Dim tmRvfSrchKey1 As RVFKEY1 'RVF key record image
Dim tmRvfSrchKey2 As LONGKEY0
Dim tmRvfSrchKey5 As RVFKEY5 'RVF key record image
'Product
Dim hmPrf As Integer        'Prf Handle
Dim tmPrf As PRF
Dim imPrfRecLen As Integer      'Prf record length
Dim tmPrfSrchKey As LONGKEY0  'Prf key record image
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmAgf As Integer        'AGF Handle
Dim tmAgf As AGF
Dim imAgfRecLen As Integer      'AGF record length
Dim hmAdf As Integer        'ADF Handle
Dim tmAdf As ADF
Dim imAdfRecLen As Integer      'ADF record length
Dim tmSrchKey As LONGKEY0
Dim tmPnfSrchKey As LONGKEY0
Dim tmSlfSrchKey As LONGKEY0
Dim hmPnf As Integer        'Pnf Handle
Dim tmPnf As PNF
Dim imPnfRecLen As Integer      'Pnf record length
Dim hmSlf As Integer        'Slf Handle
Dim tmSlf As SLF
Dim imSlfRecLen As Integer      'Slf record length

Dim hmSaf As Integer
Dim tmSaf As SAF            'Schedule Attributes record image
Dim imSafRecLen As Integer

Dim imTestSystem As Integer
Dim imTerminate As Integer
Dim imUserInput  As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imExportGetPaid As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmNowDate As Long
Dim lmInvNo As String
Dim imLoop As Integer

Dim imShowHelpMsg As Integer    'True=Show help message; False=Ignore help message system

Dim tmDetailRec(0 To 16) As FixedRec        '6-25-13 add 1 field for vehicle name, 9-10-13 add terms description, 9-18-13 Add Tax field
Dim tmMasterRec(0 To 15) As FixedRec

'Zero balance variables and buffers
Dim tmSbf As SBF            'SBF record image
Dim tmSbfSrchKey0 As SBFKEY0
Dim tmSbfSrchKey1 As LONGKEY0 'SBF key record image
Dim imSbfRecLen As Integer  'SBF record length
Dim hmSbf As Integer        'Special Billing file handle

Dim imDefaultTermsCode As Integer
Dim imTermsNumber() As Integer      'dimensioned to size of number of terms defined (from mnf, type J).  It will contain a number if the terms contain one (i.e. 30 for Net 30 days, 60 for Net 60 days.)
                                    'if no number in terms field, then a 0 is stored

Dim tmZP() As ZEROPURGE
Dim tmZPLink() As ZPLINK
Dim tmTransToIgnore() As SEQSORTTYPE 'array of transaction record # to ignore in export
                            'this contains the rvf internal code
'search the array that stores the RVF codes so that the export will not create transactions
'that are zero-balanced
Private Function mBinarySearchRVF(llRvfCode As Long, tlTransToIgnore() As SEQSORTTYPE) As Long
Dim llMiddle As Long
Dim llMin As Long
Dim llMax As Long
    llMin = LBound(tlTransToIgnore)
    llMax = UBound(tlTransToIgnore)
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llRvfCode = tlTransToIgnore(llMiddle).lSdfCode Then
            'found the match
            mBinarySearchRVF = llMiddle
            Exit Function
        ElseIf llRvfCode < tlTransToIgnore(llMiddle).lSdfCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchRVF = -1
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mZeroBalance for GetPaid Export *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Transfer all zero balanced     *
'*                      invoices and trades to revenue *
'*                      history                        *
'*            5/7/97 dh: Change to zero balance all    *
'*                      transactions as of previous    *
'*                      closing period                 *
'*      copy of mZeroPurge from Collections.  Change to*
'*      keep track of all invoices that will be purged *
'*      w/o testing of dates.  Test to determine all   *
'*      transactions that can be purged and build an   *
'       array of the IN transactions.  These trans.    *
'*      will not be written to export                  *
'*******************************************************
Private Sub mZeroBalance()
    Dim tlRvf As RVF            'RVF or PHF record image
    Dim ilRvfRecLen As Integer  'RVF record length
    Dim hlRvf As Integer        'RVF or PHF file handle
    Dim hlPhf As Integer        'RVF or PHF file handle
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim llRvfRecPos As Long
    Dim slDate As String
    Dim llEndPrevPeriod As Long
    Dim slEndPrevPeriod As String
    Dim slNet As String
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim llCashCount As Long  '2-8-02 chg from int. to long
    Dim ilTradeCount  As Integer
    Dim ilMerchCount As Integer
    Dim ilPromoCount  As Integer
    Dim slNowDate As String
    Dim slMsg As String
    Dim ilPass As Integer
    Dim slStr As String
    Dim ilBypassRvf As Integer
    Dim llNextLk As Long
    Dim llIgnoreCount As Long
    Dim slCode As String

    slNowDate = Format$(gNow(), "m/d/yy")
    hlRvf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlRvf, "", sgDBPath & "RVF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    hlPhf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlPhf, "", sgDBPath & "PHF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    'Passes:
    '        1 = Match on Invoice Numbers, Agency, Advertiser and Cash/Trade
    '        2 = Match on Contract Number, Agency, Advertiser and Cash/Trade
    '        3 = Zero in InvNo, CntrNo, slfCode, Air/Bill Vehicle and Match on Agency, Advertiser and Cash/Trade
    '        4 = Zero in InvNo, CntrNo, slfCode, Air/Bill Vehicle and Match on Aging, Agency, Advertiser and Cash/Trade
    llCashCount = 0         '2-8-02 chg from int. to long
    ilTradeCount = 0
    ilMerchCount = 0
    ilPromoCount = 0
    ReDim tmTransToIgnore(0 To 0) As SEQSORTTYPE
    llIgnoreCount = 0

    For ilPass = 1 To 4 Step 1
        Screen.MousePointer = vbHourglass
        ReDim tmZP(0 To 0) As ZEROPURGE
        ReDim tmZPLink(0 To 0) As ZPLINK
        tmZP(0).lInvNo = 0
        tmZP(0).iAgfCode = -1
        tmZP(0).iAdfCode = -1
        tmZP(0).sAmount = ".00"
        tmZP(0).sType = ""
        tmZP(0).lFirstLk = -1
        ilUpper = 0
        ilRvfRecLen = Len(tlRvf)
        gUnpackDate tgSpf.iRPRP(0), tgSpf.iRPRP(1), slEndPrevPeriod
        'gUnpackDate tgSpf.iRCRP(0), tgSpf.iRCRP(1), slEndPrevPeriod
        llEndPrevPeriod = gDateValue(slEndPrevPeriod)
        ilRet = btrGetFirst(hlRvf, tlRvf, ilRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE)
            If (tlRvf.sCashTrade <> "T") Or (tgSpf.sRUseTrade = "Y") Then
                'Ignore trade
                gUnpackDate tlRvf.iTranDate(0), tlRvf.iTranDate(1), slDate
                'All transactions will be tested for zero balance, regardless of how current
                'If gDateValue(slDate) <= llEndPrevPeriod Then
                    ilFound = False
                    ilBypassRvf = False
                    If ilPass = 1 Then
                        'Remove salesperson = 0 test as salesperson is now set sometimes
                        'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                        If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                            ilBypassRvf = True
                        Else
                            For ilLoop = LBound(tmZP) To UBound(tmZP) - 1 Step 1
                                If (tmZP(ilLoop).lInvNo = tlRvf.lInvNo) And (tmZP(ilLoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(ilLoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(ilLoop).sType = tlRvf.sCashTrade) Then
                                    ilIndex = ilLoop
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                    ElseIf ilPass = 2 Then
                        'Remove salesperson = 0 test as salesperson is now set sometimes
                        'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                        If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                            ilBypassRvf = True
                        Else
                            For ilLoop = LBound(tmZP) To UBound(tmZP) - 1 Step 1
                                If (tmZP(ilLoop).lInvNo = tlRvf.lCntrNo) And (tmZP(ilLoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(ilLoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(ilLoop).sType = tlRvf.sCashTrade) Then
                                    ilIndex = ilLoop
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                    ElseIf ilPass = 3 Then
                        'Remove salesperson = 0 test as salesperson is now set sometimes
                        'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                        If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                            For ilLoop = LBound(tmZP) To UBound(tmZP) - 1 Step 1
                                'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) And (tmZP(ilLoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(ilLoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(ilLoop).sType = tlRvf.sCashTrade) Then

                                If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) And (tmZP(ilLoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(ilLoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(ilLoop).sType = tlRvf.sCashTrade) Then

                                    ilIndex = ilLoop
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                        Else
                            ilBypassRvf = True
                        End If
                    ElseIf ilPass = 4 Then
                        'Remove salesperson = 0 test as salesperson is now set sometimes
                        'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                        If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) Then
                            For ilLoop = LBound(tmZP) To UBound(tmZP) - 1 Step 1
                                'If (tlRvf.lInvNo = 0) And (tlRvf.lCntrNo = 0) And (tlRvf.iSlfCode = 0) And (tlRvf.iAirVefCode = 0) And (tlRvf.iBillVefCode = 0) And (tmZP(ilLoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(ilLoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(ilLoop).sType = tlRvf.sCashTrade) And (tmZP(ilLoop).lInvNo = (CLng(tlRvf.iAgePeriod) * 10000 + tlRvf.iAgingYear)) Then
                                If (tmZP(ilLoop).iAgfCode = tlRvf.iAgfCode) And (tmZP(ilLoop).iAdfCode = tlRvf.iAdfCode) And (tmZP(ilLoop).sType = tlRvf.sCashTrade) And (tmZP(ilLoop).lInvNo = (CLng(tlRvf.iAgePeriod) * 10000 + tlRvf.iAgingYear)) Then
                                    ilIndex = ilLoop
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilLoop
                        Else
                            ilBypassRvf = True
                        End If
                    End If
                    If Not ilBypassRvf Then
                        If Not mRepBilled(tlRvf) Then
                            ilBypassRvf = True
                        End If
                    End If
                    If Not ilBypassRvf Then
                        gPDNToStr tlRvf.sNet, 2, slNet
                        If Not ilFound Then
                            ilIndex = UBound(tmZP)
                            ReDim Preserve tmZP(0 To ilIndex + 1) As ZEROPURGE
                            If ilPass = 1 Then
                                tmZP(ilIndex).lInvNo = tlRvf.lInvNo
                            ElseIf ilPass = 2 Then
                                tmZP(ilIndex).lInvNo = tlRvf.lCntrNo
                            ElseIf ilPass = 3 Then
                                tmZP(ilIndex).lInvNo = -1
                            ElseIf ilPass = 4 Then
                                tmZP(ilIndex).lInvNo = CLng(tlRvf.iAgePeriod) * 10000 + tlRvf.iAgingYear
                            End If
                            tmZP(ilIndex).iAgfCode = tlRvf.iAgfCode
                            tmZP(ilIndex).iAdfCode = tlRvf.iAdfCode
                            tmZP(ilIndex).sAmount = slNet
                            tmZP(ilIndex).sAmount = gAddStr(Trim$(tmZP(ilIndex).sAmount), gLongToStrDec(tlRvf.lTax1, 2))
                            tmZP(ilIndex).sAmount = gAddStr(Trim$(tmZP(ilIndex).sAmount), gLongToStrDec(tlRvf.lTax2, 2))
                            tmZP(ilIndex).sType = tlRvf.sCashTrade
                            tmZP(ilIndex).lFirstLk = UBound(tmZPLink)
                            tmZPLink(UBound(tmZPLink)).lNextLk = -1
                            tmZPLink(UBound(tmZPLink)).lRvfCode = tlRvf.lCode
                            ReDim Preserve tmZPLink(0 To UBound(tmZPLink) + 1) As ZPLINK
                            tmZP(ilIndex + 1).lInvNo = 0
                            tmZP(ilIndex + 1).iAdfCode = -1
                            tmZP(ilIndex + 1).iAdfCode = -1
                            tmZP(ilIndex + 1).sAmount = ".00"
                            tmZP(ilIndex + 1).sType = ""
                            tmZP(ilIndex + 1).lFirstLk = -1
                        Else
                            tmZP(ilIndex).sAmount = gAddStr(Trim$(tmZP(ilIndex).sAmount), slNet)
                            tmZP(ilIndex).sAmount = gAddStr(Trim$(tmZP(ilIndex).sAmount), gLongToStrDec(tlRvf.lTax1, 2))
                            tmZP(ilIndex).sAmount = gAddStr(Trim$(tmZP(ilIndex).sAmount), gLongToStrDec(tlRvf.lTax2, 2))
                            llNextLk = tmZP(ilIndex).lFirstLk
                            Do
                                If tmZPLink(llNextLk).lNextLk = -1 Then
                                    tmZPLink(llNextLk).lNextLk = UBound(tmZPLink)
                                    tmZPLink(UBound(tmZPLink)).lNextLk = -1
                                    tmZPLink(UBound(tmZPLink)).lRvfCode = tlRvf.lCode
                                    ReDim Preserve tmZPLink(0 To UBound(tmZPLink) + 1) As ZPLINK
                                    Exit Do
                                End If
                                llNextLk = tmZPLink(llNextLk).lNextLk
                            Loop 'While llNextLk <> -1
                        End If
                    End If
                'End If         rvfTranDate <= llEndPrevPeriod
            End If
            ilRet = btrGetNext(hlRvf, tlRvf, ilRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        For ilLoop = LBound(tmZP) To UBound(tmZP) - 1 Step 1
            'Transfer zero balance
            If (gCompNumberStr(Trim$(tmZP(ilLoop).sAmount), ".00") = 0) Then
                ReDim lmZBRvfCode(0 To 0) As Long
                
                llNextLk = tmZP(ilLoop).lFirstLk
                Do
'                    lmZBRvfCode(UBound(lmZBRvfCode)) = tmZPLink(llNextLk).lRvfCode
'                    ReDim Preserve lmZBRvfCode(0 To UBound(lmZBRvfCode) + 1) As Long
                    tmTransToIgnore(llIgnoreCount).lSdfCode = tmZPLink(llNextLk).lRvfCode
                    slCode = Trim$(str$(tmZPLink(llNextLk).lRvfCode))
                    Do While Len(slCode) < 10
                        slCode = "0" & slCode
                    Loop
                    tmTransToIgnore(llIgnoreCount).sKey = slCode
                    ReDim Preserve tmTransToIgnore(0 To llIgnoreCount + 1) As SEQSORTTYPE
                    llIgnoreCount = llIgnoreCount + 1
                   
                    If tmZPLink(llNextLk).lNextLk = -1 Then
                        Exit Do
                    End If
                    llNextLk = tmZPLink(llNextLk).lNextLk
                Loop
            End If
        Next ilLoop
    Next ilPass
    'Transfer Trade or Net of Zero
    ilRet = btrGetFirst(hlRvf, tlRvf, ilRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE)
        gUnpackDate tlRvf.iTranDate(0), tlRvf.iTranDate(1), slDate
        If (tlRvf.sCashTrade <> "T") Or (tgSpf.sRUseTrade = "Y") Then
            'All transactions are tested for zero balance, regardless of how current
            'If gDateValue(slDate) <= llEndPrevPeriod Then
                
                    'Transfer transaction with a net of zero
                    gPDNToStr tlRvf.sNet, 2, slNet
                    'Combine taxes with net
                    slNet = gAddStr(Trim$(slNet), gLongToStrDec(tlRvf.lTax1, 2))
                    slNet = gAddStr(Trim$(slNet), gLongToStrDec(tlRvf.lTax2, 2))

                    If (gCompNumberStr(Trim$(slNet), ".00") = 0) Then
                        tmTransToIgnore(llIgnoreCount).lSdfCode = tlRvf.lCode
                        slCode = Trim$(str$(tlRvf.lCode))
                        Do While Len(slCode) < 10
                            slCode = "0" & slCode
                        Loop
                        tmTransToIgnore(llIgnoreCount).sKey = slCode
                        ReDim Preserve tmTransToIgnore(0 To llIgnoreCount + 1) As SEQSORTTYPE
                        llIgnoreCount = llIgnoreCount + 1
                    End If
            'End If         rvftrandate <= llEndPrevPeriod
        Else
            'Trade- transfer amount
            'All transactions are tested for zero balance, regardless of how current
            'If gDateValue(slDate) <= llEndPrevPeriod Then
                tmTransToIgnore(llIgnoreCount).lSdfCode = tlRvf.lCode
                slCode = Trim$(str$(tlRvf.lCode))
                Do While Len(slCode) < 10
                    slCode = "0" & slCode
                Loop
                tmTransToIgnore(llIgnoreCount).sKey = slCode
                ReDim Preserve tmTransToIgnore(0 To llIgnoreCount + 1) As SEQSORTTYPE
                llIgnoreCount = llIgnoreCount + 1
            'End If         rvftrandate <= llPrevEndPeriod
        End If
        ilRet = btrGetNext(hlRvf, tlRvf, ilRvfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    
    'All Zero Balance Invoices created in array, sort by its internal code for faster access
    llIgnoreCount = UBound(tmTransToIgnore)
    If llIgnoreCount > 0 Then
        ArraySortTyp fnAV(tmTransToIgnore(), 0), llIgnoreCount, 0, LenB(tmTransToIgnore(0)), 0, LenB(tmTransToIgnore(0).sKey), 0
    End If

    ilRet = btrClose(hlRvf)
    btrDestroy hlRvf
    ilRet = btrClose(hlPhf)
    btrDestroy hlPhf
    Erase tmZP, tmZPLink
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:m RepBilled                     *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if Rep Billed        *
'*******************************************************
Private Function mRepBilled(tlRvf As RVF) As Integer
    Dim ilBillVefCode As Integer
    Dim ilVef As Integer
    Dim ilRet As Integer
    Dim ilCount As Integer

    mRepBilled = True
    ilCount = 0
    'Only check NTR transactions
    If (sgRepDef = "Y") And (tlRvf.sTranType = "IN") Then
        'Bypass Rep vehicles that are associated with NTR transaction
        If (tgSpf.sUsingNTR <> "Y") Or (tlRvf.iMnfItem <= 0) Then
            ilBillVefCode = tlRvf.iBillVefCode
            ilVef = gBinarySearchVef(ilBillVefCode)
            If ilVef <> -1 Then
                If tgMVef(ilVef).sType = "R" Then
                    tmChfSrchKey1.lCntrNo = tlRvf.lCntrNo
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                    'If contract missing, assume that transaction entered via Backlog, 5/8/04
                    If (ilRet <> BTRV_ERR_NONE) Or (tmChf.lCntrNo <> tlRvf.lCntrNo) Then
                        Exit Function
                    End If
                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tlRvf.lCntrNo)
                        If tmChf.sSchStatus = "F" Then
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (tmChf.lCntrNo = tlRvf.lCntrNo) And (tmChf.sSchStatus = "F") Then
                        'Billing
                        tmSbfSrchKey0.lChfCode = tmChf.lCode
                        tmSbfSrchKey0.iDate(0) = tlRvf.iTranDate(0)
                        tmSbfSrchKey0.iDate(1) = tlRvf.iTranDate(1)
                        tmSbfSrchKey0.sTranType = "T"
                        ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCode = tmSbf.lChfCode)
                            If (tmSbf.iDate(0) <> tlRvf.iTranDate(0)) Or (tmSbf.iDate(1) <> tlRvf.iTranDate(1)) Then
                                Exit Do
                            End If
                            If tmSbf.sTranType = "T" Then
                                If tmSbf.sBilled <> "Y" Then
                                    mRepBilled = False
                                    Exit Function
                                End If
                                ilCount = ilCount + 1
                            End If
                            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                    If ilCount <= 0 Then
                        mRepBilled = False
                    End If
                End If
            End If
        End If
    End If
    Exit Function
End Function

Private Sub cmcCancel_Click()
    If imExportGetPaid Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
'
'
'           1-29-01 Remove everything in file before getting export
Private Sub cmcGen_Click()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slToFile As String
    Dim slDateTime As String

    lacInfo.Visible = False
    If imExportGetPaid Then
        Exit Sub
    End If

    cmcCancel.Caption = "&Cancel"
    imExportGetPaid = True
    lacInfo.Visible = True
    lacInfo.Caption = ""
    lmTotalNoRecs = btrRecords(hmRvf)

    lmProcessedNoRecs = 0
    plcGauge.Visible = True
    plcGauge.Value = 0
    
        ilRet = 0
        lbcACodes.Clear 'clear listbox

        slToFile = sgExportPath & "ARMAST.TXT"
        'If DoesFileExist(slToFile) Then
        If gFileExist(slToFile) = 0 Then
            Kill slToFile
        End If
        
    'On Error GoTo cmcExportErr:
    'slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        'hmTo = FreeFile
        'Open slToFile For Append As hmTo
        ilRet = gFileOpen(slToFile, "Append", hmTo)
        If ilRet <> 0 Then
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            'edcTo.SetFocus
            Exit Sub
        End If
    Else
        ilRet = 0
        'hmTo = FreeFile
        'Open slToFile For Output As hmTo
        ilRet = gFileOpen(slToFile, "Output", hmTo)
        If ilRet <> 0 Then
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            'edcTo.SetFocus
            Exit Sub
        End If
    End If
       
        Screen.MousePointer = vbHourglass
        mZeroBalance            '8-4-10 build array of transactions that are zero balance to exclude from export
        mReadRecDetail

    lmTotalNoRecs = lbcACodes.ListCount
    lmProcessedNoRecs = 0
    plcGauge.Visible = True
    plcGauge.Value = 0
    
    slToFile = sgExportPath & "ARCUST.TXT"
    'If DoesFileExist(slToFile) Then
    If gFileExist(slToFile) = 0 Then
        Kill slToFile
    End If
    
    
    
    If ilRet = 0 Then
        'hmTo = FreeFile
        'Open slToFile For Append As hmTo
        ilRet = gFileOpen(slToFile, "Append", hmTo)
        If ilRet <> 0 Then
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            'edcTo.SetFocus
            Exit Sub
        End If
    Else
        ilRet = 0
        'hmTo = FreeFile
        'Open slToFile For Output As hmTo
        ilRet = gFileOpen(slToFile, "Output", hmTo)
        If ilRet <> 0 Then
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOKOnly + vbCritical + vbApplicationModal, "Open Error"
            'edcTo.SetFocus
            Exit Sub
        End If
    End If
    
        mReadRecMaster
        

    imExportGetPaid = False
    cmcCancel.Caption = "&Done"
    cmcCancel.SetFocus
    cmcGen.Enabled = False

 
    Screen.MousePointer = vbDefault

    Exit Sub
'cmcExportErr:
'
'    ilRet = Err.Number
'    gLogMsg "Error Generating Export - cmcGen", "ExportGetPaid.txt", False
'    Resume Next

End Sub
'Function WriteToFile(FileType As Integer)
'If FileType = 1 Then
'    If ilRet = 0 Then
'        hmTo = FreeFile
'        Open smToFile For Append As hmTo
'        If ilRet <> 0 Then
'            MsgBox "Open " & smToFileMaster & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'            Exit Function
'        End If
'    Else
'        ilRet = 0
'        hmTo = FreeFile
'        Open slToFile For Output As hmTo
'        If ilRet <> 0 Then
'            MsgBox "Open " & smToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'            Exit Function
'        End If
'    End If
'Else
'    If ilRet = 0 Then
'        hmTo = FreeFile
'        Open smToFile For Append As hmTo
'        If ilRet <> 0 Then
'            MsgBox "Open " & smToFileMaster & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'            Exit Function
'        End If
'    Else
'        ilRet = 0
'        hmTo = FreeFile
'        Open slToFile For Output As hmTo
'        If ilRet <> 0 Then
'            MsgBox "Open " & smToFileMaster & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
'            Exit Function
'        End If
'    End If
'End If
'End Function



Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
   ' cmcGen_Click
    tmcTimer.Enabled = True
    If Not imUserInput Then
        gOpenTmf
        tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
        tmcSetTime.Enabled = True
    End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    If Not imUserInput Then
        tmcSetTime.Enabled = False
        gCloseTmf
    End If
    
    ilRet = btrClose(hmRvf)
    btrDestroy hmRvf
    ilRet = btrClose(hmPrf)
    btrDestroy hmPrf
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    ilRet = btrClose(hmPnf)
    btrDestroy hmPnf
    ilRet = btrClose(hmSlf)
    btrDestroy hmSlf
    
    Set ExportGetPaid = Nothing   'Remove data segment
    
    btrStopAppl
    End
End Sub


'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'
'      2-02-01 Move Opens to gcmcGen; right before the
'           generation of summary records.  The timer
'           wasnt activated until after opening; therefore
'           files could remain open and the backup
'           aborted
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim slDate As String
    Dim ilTemp As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slOneChar As String
    Dim ilPos As Integer
    imTerminate = False
    mParseCmmdLine


    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    imExportGetPaid = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoRecs = 0
    lmProcessedNoRecs = 0
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)

    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen)", ExportGetPaid
    On Error GoTo 0
    imRvfRecLen = Len(tmRvf)
    
    hmPrf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen)", ExportGetPaid
    On Error GoTo 0
    imPrfRecLen = Len(tmPrf)
    
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen)", ExportGetPaid
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
      
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen)", ExportGetPaid
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)
    
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen)", ExportGetPaid
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    
    hmPnf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmPnf, "", sgDBPath & "Pnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen)", ExportGetPaid
    On Error GoTo 0
    imPnfRecLen = Len(tmPnf)
    
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitExportErr
    gBtrvErrorMsg ilRet, "cmcGen (btrOpen)", ExportGetPaid
    On Error GoTo 0
    imSlfRecLen = Len(tmSlf)
    
    lacTo.Caption = sgDBPath
    gCenterStdAlone ExportGetPaid
    ilRet = gObtainAdvt()   'Build into tgCommAdf
    If ilRet = False Then
        imTerminate = True
    End If
    ilRet = gObtainAgency() 'Build into tgCommAgf
    If ilRet = False Then
        imTerminate = True
    End If

    ilRet = gObtainSalesperson() 'Build into tgMSlf
    If ilRet = False Then
        imTerminate = True
    End If

    sgExclMnfStamp = ""
    ilRet = gObtainMnfForType("Y", sgExclMnfStamp, tgExclMnf())
    
    sgShareBudgetStamp = ""
    ilRet = gObtainMnfForType("J", sgShareBudgetStamp, tgShareBudget())         'Terms
    ReDim imTermsNumber(LBound(tgShareBudget) To UBound(tgShareBudget))
    '9-6-13 loop thru the MNF items for the terms and find its number; if no number, then store 0.
    'i.e. 30 for Net 30 days, 60 for net 60 days, 0 for Cash in Advance
    On Error GoTo mInitNumberErr
    For ilLoop = LBound(tgShareBudget) To UBound(tgShareBudget) - 1
        imTermsNumber(ilLoop) = 0
        slStr = ""
        ilPos = 1
        Do While ilPos < 20
            If IsReallyNumeric(Mid$(tgShareBudget(ilLoop).sName, ilPos, 1)) = True Then
                slStr = slStr & Trim$(Mid$(tgShareBudget(ilLoop).sName, ilPos, 1))
            End If
            ilPos = ilPos + 1
        Loop
        If Trim$(Len(slStr)) > 0 Then
            imTermsNumber(ilLoop) = Val(slStr)
        End If
    Next ilLoop
    
    'set the default mnfcode to Terms.  If none exists, hard coded to 15 days net
    imDefaultTermsCode = tgSpf.iMnfInvTerms         '9-8-13
    If imDefaultTermsCode = 0 Then                  'if default terms not set in site, set it to 15 days
        For ilLoop = LBound(imTermsNumber) To UBound(imTermsNumber) - 1
            If imTermsNumber(ilLoop) = 15 Then              'see if the 15 days net has been built into array
                imDefaultTermsCode = tgShareBudget(ilLoop).iCode
                Exit For
            End If
        Next ilLoop
        If imDefaultTermsCode = 0 Then               'default terms code still not found, force to be 15 days net
            imDefaultTermsCode = -15                'use the default, its not in mnf and is hard-coded as default (15 nets days)
        End If
    End If
    
    gAnyRepDef

    If ((Asc(tgSpf.sUsingFeatures6) And GETPAIDEXPORT) <> GETPAIDEXPORT) Then
        If igBkgdProg = 4 Then
            imTerminate = True
        Else
            cmcGen.Enabled = False
        End If

    End If
    
    Screen.MousePointer = vbDefault
    gLogMsg "mInit succeeded ", "ExportGetPaid.txt", False

    Exit Sub
mInitExportErr:
    imTerminate = True
    gLogMsg "mInit failed ", "ExportGetPaid.txt", False
    ilRet = Err
    Resume Next
    
mInitNumberErr:
    ilRet = Err
    Resume Next
mInitErr:
    On Error GoTo 0
    imTerminate = True
    gLogMsg "mInit failed ", "ExportGetPaid.txt", False
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
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
    Dim ilPos As Integer
    Dim ilSpace As Integer
    Dim slStartIn As String
    Dim ilRet As Integer
    
    'test the spf getpaidexport
    slCommand = Command$
    igStdAloneMode = True 'False  'Switch from/to stand alone mode
    sgCallAppName = ""
    imTestSystem = False
    igDirectCall = 0
    sgIniPath = ""
    
    ilPos = InStr(1, slCommand, "/IniLoc:", vbTextCompare)

    If ilPos > 0 Then
 
        igDirectCall = -1
        ilSpace = InStr(ilPos, slCommand, " ")

        If ilSpace = 0 Then
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8))
        Else
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8, ilSpace - ilPos - 3))
        End If
        If right$(sgIniPath, 1) <> "\" Then
            sgIniPath = sgIniPath & "\"
        End If
    End If
    ilPos = InStr(1, slCommand, "/UserInput", 1)
    If ilPos > 0 Then
       ' imTestSystem = True
        imUserInput = True
        igBkgdProg = 0
        'MsgBox "Current Mode: Shortcut.", vbExclamation, ""
      '  gLogMsg "Current Mode: UserInput.", "ExportGetPaid.txt", False
    Else
        imUserInput = False
        igBkgdProg = 4
        Me.Visible = False
        'MsgBox "Current Mode: Background.", vbExclamation, ""
       ' gLogMsg "Current Mode: Background.", "ExportGetPaid.txt", False
    End If

    '7-25-13 allow to run standalone in test
    slStartIn = CurDir$
    If InStr(1, slStartIn, "Test", vbTextCompare) = 0 Then
        imTestSystem = False
    Else
        imTestSystem = True
    End If
    
    slStr = "Guide"
    slCommand = "Traffic\Guide"
    gInitStdAlone ExportGetPaid, slStr, imTestSystem
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "GetPaid.Frm"
    
    If igBkgdProg = 0 Then
        gLogMsg "Current Mode: UserInput.", "ExportGetPaid.txt", False
    Else
        gLogMsg "Current Mode: Background.", "ExportGetPaid.txt", False
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*            Created:5/17/93       By:D. LeVine       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: terminate form                 *
'*                                                     *
'*******************************************************
Private Sub mTerminate()
'
'   mTerminate
'   Where:
'

    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload ExportGetPaid
    igManUnload = NO
End Sub

Private Sub tmcSetTime_Timer()
    gUpdateTaskMonitor 0, "GPE"
End Sub

Private Sub tmcTimer_Timer()
    tmcTimer.Enabled = False
    
    If imExportGetPaid Then
        Exit Sub
    End If
    If imTerminate Then
        mTerminate
    Else
        mCheckGG
        If igGGFlag = 0 Then
            imTerminate = True
            mTerminate
            Exit Sub
        End If
        If Not imUserInput Then
            gUpdateTaskMonitor 1, "GPE"
            cmcGen_Click
            gUpdateTaskMonitor 2, "GPE"
            mTerminate
        End If
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRecDetail                  *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRecDetail() As Integer
'
'   iRet = mReadRec()
'   Where:
'       ilSelectIndex (I) - list box index
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOk As Integer
    Dim llPercent As Long
    Dim ilOffSet As Integer
    Dim tlCashTradeBuff As POPCHARTYPE   'Type field record
    
    llNoRec = gExtNoRec(imRvfRecLen) 'btrRecords(hlAdf) 'Obtain number of records
    btrExtClear hmRvf   'Clear any previous extend operation
    ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY5, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Call btrExtSetBounds(hmRvf, llNoRec, -1, "UC", "RVF", "")  'Set extract limits (all records)
    
    tlCashTradeBuff.sType = "C"
    ilOffSet = gFieldOffset("Rvf", "rvfCashTrade")
    ilRet = btrExtAddLogicConst(hmRvf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCashTradeBuff, 1)
    If ilRet = BTRV_ERR_NONE Then
     
        ilRet = btrExtAddField(hmRvf, 0, imRvfRecLen)  'Extract record
        If ilRet = BTRV_ERR_NONE Then
        ' If i comment out the next line I only end up get one record
            ilRet = btrExtGetNext(hmRvf, tmRvf, imRvfRecLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                    imRvfRecLen = Len(tmRvf)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmRvf, tmRvf, imRvfRecLen, llRecPos)
                    Loop
                    Do While ilRet = BTRV_ERR_NONE
                        ilRet = mBinarySearchRVF(tmRvf.lCode, tmTransToIgnore())      'determine if this transaction is zero-balanced
                                                                    'if so, do not export it
                        If ilRet < 0 Then                           'not found, ok to export
                            mMakeExportRecDetail tmRvf
                        Else
                            ilRet = ilRet
                        End If
                        
                        lmProcessedNoRecs = lmProcessedNoRecs + 1
                        llPercent = (lmProcessedNoRecs * 100) \ lmTotalNoRecs
                        If llPercent >= 100 Then
                            llPercent = 100
                        End If
                        plcGauge.Value = llPercent

                        ilRet = btrExtGetNext(hmRvf, tmRvf, imRvfRecLen, llRecPos)
                        
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hmRvf, tmRvf, imRvfRecLen, llRecPos)
                        Loop
                    Loop
                End If
            End If
        End If
    End If



    mReadRecDetail = True
    gLogMsg "mReadRecDetail succeeded ", "ExportGetPaid.txt", False
    Exit Function
            
    On Error GoTo 0
    mReadRecDetail = False
    gLogMsg "mReadRecDetail failed ", "ExportGetPaid.txt", False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRecMaster                  *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRecMaster() As Integer
'
'       Read items from listbox


    Dim llPercent As Long
    Dim ilLoop As Integer
    Dim slACode As Integer
    Dim slAType As String
    
    For ilLoop = 0 To lbcACodes.ListCount - 1 Step 1 'Get first record as starting point of extend operation

        slACode = Val(Mid$(lbcACodes.List(ilLoop), 5))
        slAType = Trim$(Left$(lbcACodes.List(ilLoop), 4))
        If slACode = "14" Then
        slACode = slACode
        End If
        If slAType = "AGNY" Then
            mMakeMasterAgy slACode
        Else
            mMakeMasterAdv slACode
        End If

        lmProcessedNoRecs = lmProcessedNoRecs + 1
        llPercent = (lmProcessedNoRecs * 100) \ lmTotalNoRecs
        If llPercent >= 100 Then
            llPercent = 100
        End If
        plcGauge.Value = llPercent
    Next ilLoop
    
    
    



    mReadRecMaster = True
    gLogMsg "mReadRecMaster succeeded ", "ExportGetPaid.txt", False
    Exit Function
            
    On Error GoTo 0
    mReadRecMaster = False
    gLogMsg "mReadRecMaster failed ", "ExportGetPaid.txt", False
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeExportRec                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize stand alone mode    *
'*                                                     *
'*******************************************************
Private Sub mMakeExportRecDetail(tlRvf As RVF)
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slAdfName As String
    Dim slCode As String
    Dim slBuyer As String
    Dim slCntrNo As String
    Dim slInvDate As String
    Dim slCheckNo As String
    Dim slTranType As String
    Dim slNet As String
    Dim slSlsp As String
    Dim slRec As String
    Dim slAType As String
    Dim ilIsPositive As Integer
    Dim slNewCode As String
    Dim slRef As String
    Dim slPrf As String
    Dim slPInvNo As String
    Dim slVehicleName As String * 50
    Dim ilTermsCode As Integer
    Dim ilInx As Integer
    Dim ilTerms As Integer
    Dim slTax As String
    Dim llTax As Long

    'slName = ""
    'For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1

        If lmInvNo = Trim$(tlRvf.lInvNo) Then
            imLoop = imLoop + 1
            slPInvNo = Trim$(tlRvf.lInvNo) & "-" & imLoop
        Else
            slPInvNo = Trim$(tlRvf.lInvNo) & "-" & 1
            imLoop = 1
        End If

'''''' lmInvNo = Trim$(tlRvf.lInvNo)
        tmDetailRec(0).sFieldValue = slPInvNo 'Trim$(tlRvf.lInvNo) '
        tmDetailRec(0).iStartPos = 1
        tmDetailRec(0).iLength = 20
        tmDetailRec(0).bLeadingZeroes = False
        tmDetailRec(0).sJustification = 0
        tmDetailRec(0).sDataType = 0
                
If tlRvf.iAgfCode = 296 Then
ilRet = ilRet
End If
        If tlRvf.iAgfCode > 0 Then
            slCode = tlRvf.iAgfCode
            slAType = "AGNY"
            ilInx = gBinarySearchAgf(tlRvf.iAgfCode)
            ilTermsCode = tgCommAgf(ilInx).iMnfInvTerms
        Else
            slCode = tlRvf.iAdfCode
            slAType = "ADVT"
            'direct advertiser
            ilInx = gBinarySearchAdf(tlRvf.iAdfCode)
            ilTermsCode = tgCommAdf(ilInx).iMnfInvTerms
        End If
         If ilTermsCode = 0 Then            'if no terms code, take the default
            ilTermsCode = imDefaultTermsCode
        End If
        
        slNewCode = slAType & slCode

        gFindMatch slNewCode, 0, lbcACodes
        If gLastFound(lbcACodes) >= 0 Then
            'lbcACodes.ListIndex = gLastFound(lbcACodes)
           ' mSetCommands
        Else
            lbcACodes.AddItem (slNewCode)
            'lbcACodes.Text = slNewCode    'Name
        End If
        
        'get Advt name
        If tlRvf.iAdfCode = 0 Then
           slAdfName = ""
        Else
            ilLoop = gBinarySearchAdf(tlRvf.iAdfCode)
            If tgCommAdf(ilLoop).iCode = tlRvf.iAdfCode Then
                If ilLoop <> -1 Then
                    slAdfName = Trim$(tgCommAdf(ilLoop).sName)
                End If
            End If
        End If

        tmDetailRec(1).sFieldValue = slNewCode
        tmDetailRec(1).iStartPos = 21
        tmDetailRec(1).iLength = 20
        tmDetailRec(1).bLeadingZeroes = False
        tmDetailRec(1).sJustification = 0
        tmDetailRec(1).sDataType = 0
        
        slInvDate = ""
        gUnpackDateByFormat tlRvf.iTranDate(0), tlRvf.iTranDate(1), "mmddyyyy", slInvDate
        tmDetailRec(2).sFieldValue = slInvDate
        tmDetailRec(2).iStartPos = 41
        tmDetailRec(2).iLength = 8
        tmDetailRec(2).bLeadingZeroes = False
        tmDetailRec(2).sJustification = 0
        tmDetailRec(2).sDataType = 0

        
        tmDetailRec(15).sFieldValue = sgDefaultTerms                'terms description default
        tmDetailRec(15).iStartPos = 613
        tmDetailRec(15).iLength = 35
        tmDetailRec(15).bLeadingZeroes = False
        tmDetailRec(15).sJustification = 0
        tmDetailRec(15).sDataType = 0
        
        ilTerms = 0
        For ilLoop = LBound(tgShareBudget) To UBound(tgShareBudget) - 1     'match the internal mnf terms code with all the term entries.  Will need for full 20 char description,
                                                                            'but also need to put the number of days the term is defined for.  If Cash in Advance, or other terms
                                                                            'without #, then store 0
            If tgShareBudget(ilLoop).iCode = ilTermsCode Then
                ilTerms = imTermsNumber(ilLoop)
                tmDetailRec(15).sFieldValue = tgShareBudget(ilLoop).sName
                Exit For
            End If
        Next ilLoop
        
        If ilTermsCode < 0 Then     'negative, use default of 15 days net
            tmDetailRec(3).sFieldValue = "15"
        Else
            tmDetailRec(3).sFieldValue = str$(ilTerms)
        End If
        'tmDetailRec(3).sFieldValue = "15"          used to be hard-coded
        tmDetailRec(3).iStartPos = 69
        tmDetailRec(3).iLength = 4
        tmDetailRec(3).bLeadingZeroes = False
        tmDetailRec(3).sJustification = 1
        tmDetailRec(3).sDataType = 0
        
        slNet = ""
        gPDNToStr tlRvf.sNet, 2, slNet
        If Left$(slNet, 1) = "-" Then
            ilIsPositive = 0
        Else
            ilIsPositive = 1
        End If

        tmDetailRec(10).sFieldValue = slNet
        tmDetailRec(10).iStartPos = 49
        tmDetailRec(10).iLength = 20
        tmDetailRec(10).bLeadingZeroes = False
        tmDetailRec(10).sJustification = 1
        tmDetailRec(10).sDataType = 0
        
        tmDetailRec(4).sFieldValue = slNet
        tmDetailRec(4).iStartPos = 73
        tmDetailRec(4).iLength = 20
        tmDetailRec(4).bLeadingZeroes = False
        tmDetailRec(4).sJustification = 1
        tmDetailRec(4).sDataType = 0

        slRef = ""

        For ilLoop = LBound(tgExclMnf) To UBound(tgExclMnf) - 1 Step 1
            If Trim$(tgExclMnf(ilLoop).sUnitType) = tlRvf.sTranType Then
                slRef = Trim$(tgExclMnf(ilLoop).sName)
                Exit For
            End If
        Next ilLoop
            
            
        slTranType = Trim$(tlRvf.sTranType)
        Select Case slTranType
            Case Is = "IN", "PI"
            slTranType = "I"
            Case Is = "WI"
            slTranType = "I"
            Case Is = "PO"
            slTranType = "U"
            Case Is = "AN"
                If ilIsPositive Then
                    slTranType = "C"
                Else
                    slTranType = "D"
                End If
            Case Else
                If ilIsPositive Then
                    slTranType = "C"
                Else
                    slTranType = "D"
                End If
            
        End Select
        
  
            tmDetailRec(5).sFieldValue = slRef
            tmDetailRec(5).iStartPos = 93
            tmDetailRec(5).iLength = 30
            tmDetailRec(5).bLeadingZeroes = False
            tmDetailRec(5).sJustification = 0
            tmDetailRec(5).sDataType = 0


                
        slPrf = ""
        If tlRvf.lPrfCode > 0 Then
            tmPrfSrchKey.lCode = tlRvf.lPrfCode
            ilRet = btrGetEqual(hmPrf, tmPrf, imPrfRecLen, tmPrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If (ilRet = BTRV_ERR_NONE) Then
                slPrf = Trim$(tmPrf.sName)

            End If
        End If
        
            tmDetailRec(6).sFieldValue = slPrf
            tmDetailRec(6).iStartPos = 123
            tmDetailRec(6).iLength = 54
            tmDetailRec(6).bLeadingZeroes = False
            tmDetailRec(6).sJustification = 0
            tmDetailRec(6).sDataType = 0
        
            tmDetailRec(7).sFieldValue = slTranType
            tmDetailRec(7).iStartPos = 247
            tmDetailRec(7).iLength = 1
            tmDetailRec(7).bLeadingZeroes = False
            tmDetailRec(7).sJustification = 0
            tmDetailRec(7).sDataType = 0
            '6/10/15: Check number changed to string
            'If tlRvf.lCheckNo > 0 Then
            '    slCheckNo = Trim$(str$(tlRvf.lCheckNo))
            'Else
            '    slCheckNo = ""
            'End If
            slCheckNo = Trim$(tlRvf.sCheckNo)
            tmDetailRec(8).sFieldValue = slCheckNo
            tmDetailRec(8).iStartPos = 248
            tmDetailRec(8).iLength = 20
            tmDetailRec(8).bLeadingZeroes = False
            tmDetailRec(8).sJustification = 0
            tmDetailRec(8).sDataType = 0
        
        slSlsp = ""
        Dim ilLoop2 As Integer
        If tlRvf.iSlfCode > 0 Then
        
            For ilLoop2 = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
                If tgMSlf(ilLoop2).iCode = tlRvf.iSlfCode Then
                 
                    slSlsp = Trim$(Left$(tgMSlf(ilLoop2).sLastName, 8)) & "," & Trim$(Left$(tgMSlf(ilLoop2).sFirstName, 1))
                    tmDetailRec(9).sFieldValue = slSlsp
                    tmDetailRec(9).iStartPos = 288
                    tmDetailRec(9).iLength = 10
                    tmDetailRec(9).bLeadingZeroes = False
                    tmDetailRec(9).sJustification = 0
                    tmDetailRec(9).sDataType = 0
                    Exit For
                End If
            Next ilLoop2
        Else            '7-18-13 handle 0 starting position when slsp doesnt exist
            tmDetailRec(9).sFieldValue = " "
            tmDetailRec(9).iStartPos = 288
            tmDetailRec(9).iLength = 1
            tmDetailRec(9).bLeadingZeroes = False
            tmDetailRec(9).sJustification = 0
            tmDetailRec(9).sDataType = 0
        End If
        


            tmDetailRec(11).sFieldValue = slAdfName
            tmDetailRec(11).iStartPos = 438
            tmDetailRec(11).iLength = 35
            tmDetailRec(11).bLeadingZeroes = False
            tmDetailRec(11).sJustification = 0
            tmDetailRec(11).sDataType = 0

            If tlRvf.lCntrNo > 0 Then
                slCntrNo = Trim$(str$(tlRvf.lCntrNo))
            Else
                slCntrNo = ""
            End If
            tmDetailRec(12).sFieldValue = slCntrNo
            tmDetailRec(12).iStartPos = 473
            tmDetailRec(12).iLength = 35
            tmDetailRec(12).bLeadingZeroes = False
            tmDetailRec(12).sJustification = 0
            tmDetailRec(12).sDataType = 0
            
            tmChfSrchKey1.lCntrNo = tlRvf.lCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tlRvf.lCntrNo) And (tmChf.sSchStatus = "A")
                ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Loop
            slBuyer = ""
            If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tlRvf.lCntrNo) Then
                slBuyer = Trim$(tmChf.sBuyer)
  
            End If
            tmDetailRec(13).sFieldValue = slBuyer 'tmChf.sBuyer
            tmDetailRec(13).iStartPos = 508
            tmDetailRec(13).iLength = 35
            tmDetailRec(13).bLeadingZeroes = False
            tmDetailRec(13).sJustification = 0
            tmDetailRec(13).sDataType = 0

            slVehicleName = ""
            ilRet = gBinarySearchVef(tlRvf.iBillVefCode)
            If ilRet >= 0 Then
                slVehicleName = Trim$(tgMVef(ilRet).sName)
            End If
                
            tmDetailRec(14).sFieldValue = slVehicleName 'vehicle name
            tmDetailRec(14).iStartPos = 928
            tmDetailRec(14).iLength = 50
            tmDetailRec(14).bLeadingZeroes = False
            tmDetailRec(14).sJustification = 0
            tmDetailRec(14).sDataType = 0
            
            'tmDetailRec(15) = terms info
            
            llTax = tlRvf.lTax1 + tlRvf.lTax2
            slTax = gLongToStrDec(llTax, 2)
            tmDetailRec(16).sFieldValue = slTax                '9-18-13
            tmDetailRec(16).iStartPos = 187
            tmDetailRec(16).iLength = 20
            tmDetailRec(16).bLeadingZeroes = False
            tmDetailRec(16).sJustification = 1
            tmDetailRec(16).sDataType = 0


        lmInvNo = Trim$(tlRvf.lInvNo)
    
        slRec = gCreateFixedRec(1233, tmDetailRec())
        Print #hmTo, slRec


End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeMasterAdv                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize stand alone mode    *
'*                                                     *
'*******************************************************
Private Sub mMakeMasterAdv(slACode As Integer)

    Dim ilRet As Integer
    Dim slAName As String
    Dim slRec As String
    Dim slAddr1 As String
    Dim slAddr2 As String
    Dim slAddr3 As String
    Dim slCity As String
    Dim slState As String
    Dim slZip As String
    Dim slCredit As String
    Dim slPnfName As String
    Dim slPnfPhoneNumber As String
    Dim slSalesman As String
    Dim slLastPay As String
    Dim slLastInv As String
    Dim slCurBal As String
    Dim slHiCredit As String
    Dim ilPos As Integer
    
    slRec = slACode

        tmMasterRec(0).sFieldValue = "ADVT" & slRec
        tmMasterRec(0).iStartPos = 1
        tmMasterRec(0).iLength = 20
        tmMasterRec(0).bLeadingZeroes = False
        tmMasterRec(0).sJustification = 0
        tmMasterRec(0).sDataType = 0

        slHiCredit = ""
        slCurBal = ""
        slLastInv = ""
        slLastPay = ""
        slZip = ""
        slCity = ""
        slState = ""
        
        tmSrchKey.lCode = slACode
        
            ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slAName = tmAdf.sName '-41 -35
                gUnpackDateByFormat tmAdf.iDateLstPaym(0), tmAdf.iDateLstPaym(1), "mmddyyyy", slLastPay '-611 -8
                gUnpackDateByFormat tmAdf.iDateLstPaym(0), tmAdf.iDateLstPaym(1), "mmddyyyy", slLastInv '-619 -8
                gPDNToStr tmAdf.sCurrAR, 2, slCurBal '-637 -20
                gPDNToStr tmAdf.sHiCredit, 2, slHiCredit '-813 -20
                                          
                If Len(Trim$(tmAdf.sBillAddr(0))) > 0 Then
                    slAddr1 = Trim$(tmAdf.sBillAddr(0)) '-76 -50
                    slAddr2 = Trim$(tmAdf.sBillAddr(1)) '-126 -50
                    slAddr3 = Trim$(tmAdf.sBillAddr(2)) '-176 -50
                Else
                    slAddr1 = Trim$(tmAdf.sCntrAddr(0)) '-76 -50
                    slAddr2 = Trim$(tmAdf.sCntrAddr(1)) '-126 -50
                    slAddr3 = Trim$(tmAdf.sCntrAddr(2)) '-176 -50
                End If
                
                Dim FoundAdd As Boolean ' Checking for CSZ City State Zip
                FoundAdd = False

                If Not FoundAdd Then
                    If IsReallyNumeric(right$(slAddr2, 5)) Then 'If Zip has 5 Chars
                        If NoComma(slAddr2) Then ' If no comma in address field then put all address field in the city and Nothing in State
                            If Left$(right$(slAddr2, 5), 1) = "-" Then
                                slZip = Trim$(right$(slAddr2, 10))
                                slCity = Trim$(slAddr2) '-226 -50
                                slState = GetState(slAddr2) '""  '-276 -50
                                FoundAdd = True
                            Else
                                slZip = Trim$(right$(slAddr2, 5))
                                slCity = Trim$(slAddr2) '-226 -50
                                slState = GetState(slAddr2) '""  '-276 -50
                                FoundAdd = True
                            End If
                        Else
                            slZip = Trim$(right$(slAddr2, 5))
                            slCity = Trim$(Left(slAddr2, InStr(1, slAddr2, ",") - 1)) '-226 -50
                            'slState = Mid([slAddr2], InStr([slAddr2], ",") + 2, 2)
                            slState = Mid([slAddr2], InStr([slAddr2], ",") + 1)
                            slState = Mid(Trim$(slState), 1, 2)
                            FoundAdd = True
                        End If
                    End If
                If IsReallyNumeric(right$(slAddr2, 5)) Then
                    If Left$(right$(slAddr2, 5), 1) = "-" Then   'If Zip has 10 Chars
                        slZip = Trim$(right$(slAddr2, 10))
                        'slCity = Trim$(Left(slAddr2, InStr(1, slAddr2, ",") - 1)) '-226 -50
                        slCity = Trim$(Mid(slAddr2, 1, Len(slAddr2) - 10))        '6-26-13 handle no comma in address
                        If InStr(1, slAddr2, ",", vbBinaryCompare) > 0 Then     'comma exists for state
                            slState = Mid([slAddr2], InStr([slAddr2], ",") + 2, 2)  '-276 -50
                        Else
                            slState = ""
                        End If
                        FoundAdd = True
                    End If
                End If
                End If
                If Not FoundAdd Then
                    If NoComma(slAddr3) Then ' If no comma in address field then put all address field in the city and Nothing in State
                        If Left$(right$(slAddr3, 5), 1) = "-" Then
                            slZip = Trim$(right$(slAddr3, 10))
                            slCity = Trim$(slAddr3) '-226 -50
                            slState = GetState(slAddr3) '""  '-276 -50
                            FoundAdd = True
                        Else
                            slZip = Trim$(right$(slAddr3, 5))
                            slCity = Trim$(slAddr3) '-226 -50
                            slState = GetState(slAddr3) '""  '-276 -50
                            FoundAdd = True
                        End If
                    Else
                        If IsReallyNumeric(right$(slAddr3, 5)) Then
                            slZip = Trim$(right$(slAddr3, 5))
                            slCity = Trim$(Left(slAddr3, InStr(1, slAddr3, ",") - 1)) '-226 -50
                            'slState = Mid([slAddr3], InStr([slAddr3], ",") + 2, 2)  '-276 -50
                            slState = Mid(slAddr3, InStr(slAddr3, ",") + 1)
                            slState = Mid(Trim$(slState), 1, 2)
                            FoundAdd = True
                        End If
        
                        If Left$(right$(slAddr3, 5), 1) = "-" Then
                            slZip = Trim$(right$(slAddr3, 10))
                            slCity = Trim$(Left(slAddr3, InStr(1, slAddr3, ",") - 1)) '-226 -50
                            slState = Mid([slAddr3], InStr([slAddr3], ",") + 2, 2)   '-276 -50
                            FoundAdd = True
                        End If
                    End If
                End If

               If tmAdf.sCreditRestr = "L" Then
                    slCredit = gLongToStrDec(tmAdf.lCreditLimit, 2) '-541 -20
               End If
              

            End If

    tmPnfSrchKey.lCode = tmAdf.iPnfBuyer
    ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
         slPnfName = tmPnf.sName ' - 571 - 20
         slPnfPhoneNumber = tmPnf.sPhone ' gFormatPhoneNo(tmPnf.sPhone) ' - 396 - 35

    Else
        slPnfName = ""
        slPnfPhoneNumber = ""
    End If
    
    tmSlfSrchKey.lCode = tmAdf.iSlfCode
    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        slSalesman = Left$(Trim$(tmSlf.sLastName), 8) & "," & LeftB$(Trim$(tmSlf.sFirstName), 1) '- 561 - 10
    Else
        slSalesman = ""
    End If

            
        tmMasterRec(1).sFieldValue = slAName
        tmMasterRec(1).iStartPos = 41
        tmMasterRec(1).iLength = 35
        tmMasterRec(1).bLeadingZeroes = False
        tmMasterRec(1).sJustification = 0
        tmMasterRec(1).sDataType = 0
        
        tmMasterRec(2).sFieldValue = slAddr1
        tmMasterRec(2).iStartPos = 76
        tmMasterRec(2).iLength = 50
        tmMasterRec(2).bLeadingZeroes = False
        tmMasterRec(2).sJustification = 0
        tmMasterRec(2).sDataType = 0
        
        tmMasterRec(3).sFieldValue = slAddr2
        tmMasterRec(3).iStartPos = 126
        tmMasterRec(3).iLength = 50
        tmMasterRec(3).bLeadingZeroes = False
        tmMasterRec(3).sJustification = 0
        tmMasterRec(3).sDataType = 0
        
        tmMasterRec(4).sFieldValue = slAddr3
        tmMasterRec(4).iStartPos = 176
        tmMasterRec(4).iLength = 50
        tmMasterRec(4).bLeadingZeroes = False
        tmMasterRec(4).sJustification = 0
        tmMasterRec(4).sDataType = 0

        tmMasterRec(5).sFieldValue = slCity
        tmMasterRec(5).iStartPos = 226
        tmMasterRec(5).iLength = 50
        tmMasterRec(5).bLeadingZeroes = False
        tmMasterRec(5).sJustification = 0
        tmMasterRec(5).sDataType = 0

        tmMasterRec(6).sFieldValue = slState
        tmMasterRec(6).iStartPos = 276
        tmMasterRec(6).iLength = 50
        tmMasterRec(6).bLeadingZeroes = False
        tmMasterRec(6).sJustification = 0
        tmMasterRec(6).sDataType = 0
        
        tmMasterRec(7).sFieldValue = slZip
        tmMasterRec(7).iStartPos = 326
        tmMasterRec(7).iLength = 20
        tmMasterRec(7).bLeadingZeroes = False
        tmMasterRec(7).sJustification = 0
        tmMasterRec(7).sDataType = 0
        
        tmMasterRec(8).sFieldValue = slPnfPhoneNumber
        tmMasterRec(8).iStartPos = 396
        tmMasterRec(8).iLength = 35
        tmMasterRec(8).bLeadingZeroes = False
        tmMasterRec(8).sJustification = 0
        tmMasterRec(8).sDataType = 0
        
        tmMasterRec(9).sFieldValue = slCredit
        tmMasterRec(9).iStartPos = 541
        tmMasterRec(9).iLength = 20
        tmMasterRec(9).bLeadingZeroes = False
        tmMasterRec(9).sJustification = 0
        tmMasterRec(9).sDataType = 0
                
        tmMasterRec(10).sFieldValue = slSalesman
        tmMasterRec(10).iStartPos = 561
        tmMasterRec(10).iLength = 10
        tmMasterRec(10).bLeadingZeroes = False
        tmMasterRec(10).sJustification = 0
        tmMasterRec(10).sDataType = 0
        
        tmMasterRec(11).sFieldValue = slPnfName
        tmMasterRec(11).iStartPos = 571
        tmMasterRec(11).iLength = 20
        tmMasterRec(11).bLeadingZeroes = False
        tmMasterRec(11).sJustification = 0
        tmMasterRec(11).sDataType = 0
        
        tmMasterRec(12).sFieldValue = slLastPay
        tmMasterRec(12).iStartPos = 611
        tmMasterRec(12).iLength = 8
        tmMasterRec(12).bLeadingZeroes = False
        tmMasterRec(12).sJustification = 0
        tmMasterRec(12).sDataType = 0
        
        tmMasterRec(13).sFieldValue = slLastInv
        tmMasterRec(13).iStartPos = 619
        tmMasterRec(13).iLength = 8
        tmMasterRec(13).bLeadingZeroes = False
        tmMasterRec(13).sJustification = 0
        tmMasterRec(13).sDataType = 0
        
        tmMasterRec(14).sFieldValue = slCurBal
        tmMasterRec(14).iStartPos = 637
        tmMasterRec(14).iLength = 20
        tmMasterRec(14).bLeadingZeroes = False
        tmMasterRec(14).sJustification = 1
        tmMasterRec(14).sDataType = 0
        
        tmMasterRec(15).sFieldValue = slHiCredit
        tmMasterRec(15).iStartPos = 813
        tmMasterRec(15).iLength = 20
        tmMasterRec(15).bLeadingZeroes = False
        tmMasterRec(15).sJustification = 1
        tmMasterRec(15).sDataType = 0


        slRec = gCreateFixedRec(1233, tmMasterRec())
        Print #hmTo, slRec
        
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMakeMasterAgy                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize stand alone mode    *
'*                                                     *
'*******************************************************
Private Sub mMakeMasterAgy(slACode As Integer)

    Dim ilRet As Integer
    Dim slRec As String
    Dim slAName As String
    Dim slAddr1 As String
    Dim slAddr2 As String
    Dim slAddr3 As String
    Dim slCity As String
    Dim slState As String
    Dim slZip As String
    Dim slCredit As String
    Dim slPnfName As String
    Dim slPnfPhoneNumber As String
    Dim slSalesman As String
    Dim slLastPay As String
    Dim slLastInv As String
    Dim slCurBal As String
    Dim slHiCredit As String
    Dim ilPos As Integer
    
    slRec = slACode

        tmMasterRec(0).sFieldValue = "AGNY" & slRec
        tmMasterRec(0).iStartPos = 1
        tmMasterRec(0).iLength = 20
        tmMasterRec(0).bLeadingZeroes = False
        tmMasterRec(0).sJustification = 0
        tmMasterRec(0).sDataType = 0

        slHiCredit = ""
        slCurBal = ""
        slLastInv = ""
        slLastPay = ""
        slZip = ""
        slCity = ""
        slState = ""
        
        tmSrchKey.lCode = slACode
            If slACode = 125 Then
            Dim dd
            dd = dd
            
            End If
            ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                slAName = tmAgf.sName '-41 -35
                slLastPay = tmAgf.iDateLstPaym(0)
                slLastPay = tmAgf.iDateLstPaym(1)
                gUnpackDateByFormat tmAgf.iDateLstPaym(0), tmAgf.iDateLstPaym(1), "mmddyyyy", slLastPay '-611 -8
                gUnpackDateByFormat tmAgf.iDateLstPaym(0), tmAgf.iDateLstPaym(1), "mmddyyyy", slLastInv '-619 -8
                gPDNToStr tmAgf.sCurrAR, 2, slCurBal '-637 -20
                gPDNToStr tmAgf.sHiCredit, 2, slHiCredit '-813 -20
                                          
                If Len(Trim$(tmAgf.sBillAddr(0))) > 0 Then
                    slAddr1 = Trim$(tmAgf.sBillAddr(0)) '-76 -50
                    slAddr2 = Trim$(tmAgf.sBillAddr(1)) '-126 -50
                    slAddr3 = Trim$(tmAgf.sBillAddr(2)) '-176 -50
                Else
                    slAddr1 = Trim$(tmAgf.sCntrAddr(0)) '-76 -50
                    slAddr2 = Trim$(tmAgf.sCntrAddr(1)) '-126 -50
                    slAddr3 = Trim$(tmAgf.sCntrAddr(2)) '-176 -50
                End If
                                
                Dim FoundAdd As Boolean ' Checking for CSZ City State Zip
                FoundAdd = False

                If Not FoundAdd Then
                    If IsReallyNumeric(right$(slAddr2, 5)) Then 'If Zip has 5 Chars
                        If NoComma(slAddr2) Then ' If no comma in address field then put all address field in the city and Nothing in State
                            If Left$(right$(slAddr2, 5), 1) = "-" Then
                                slZip = Trim$(right$(slAddr2, 10))
                                slCity = Trim$(slAddr2) '-226 -50
                                slState = GetState(slAddr2) '""  '-276 -50
                                FoundAdd = True
                            Else
                                slZip = Trim$(right$(slAddr2, 5))
                                slCity = Trim$(slAddr2) '-226 -50
                                slState = GetState(slAddr2) '""  '-276 -50
                                FoundAdd = True
                            End If
                        Else
                            slZip = Trim$(right$(slAddr2, 5))
                            slCity = Trim$(Left(slAddr2, InStr(1, slAddr2, ",") - 1)) '-226 -50
                            slState = Mid([slAddr2], InStr([slAddr2], ",") + 2, 2)
                            FoundAdd = True
                        End If
                    End If
                
                    If Left$(right$(slAddr2, 5), 1) = "-" Then   'If Zip has 10 Chars
                        slZip = Trim$(right$(slAddr2, 10))        'strip off the zip code
                        'slCity = Trim$(Left(slAddr2, InStr(1, slAddr2, ",") - 1)) '-226 -50
                        slCity = Trim$(Mid(slAddr2, 1, Len(slAddr2) - 10))        '6-26-13 handle no comma in address
                        If InStr(1, slAddr2, ",", vbBinaryCompare) > 0 Then     'comma exists for state
                            slState = Mid([slAddr2], InStr([slAddr2], ",") + 2, 2)  '-276 -50
                        Else
                            slState = ""
                        End If
                        FoundAdd = True
                    End If
                End If
                If Not FoundAdd Then
                    If NoComma(slAddr3) Then ' If no comma in address field then put all address field in the city and Nothing in State
                        If Left$(right$(slAddr3, 5), 1) = "-" Then
                            slZip = Trim$(right$(slAddr3, 10))
                            slCity = Trim$(slAddr3) '-226 -50
                            slState = GetState(slAddr3) '""  '-276 -50
                            FoundAdd = True
                        Else
                            slZip = Trim$(right$(slAddr3, 5))
                            slCity = Trim$(slAddr3) '-226 -50
                            slState = GetState(slAddr3) '""  '-276 -50
                            FoundAdd = True
                        End If
                    Else
                        If IsReallyNumeric(right$(slAddr3, 5)) Then
                            slZip = Trim$(right$(slAddr3, 5))
                            slCity = Trim$(Left(slAddr3, InStr(1, slAddr3, ",") - 1)) '-226 -50
                            slState = Mid([slAddr3], InStr([slAddr3], ",") + 2, 2)  '-276 -50
                            FoundAdd = True
                        End If
        
                        If Left$(right$(slAddr3, 5), 1) = "-" Then
                            slZip = Trim$(right$(slAddr3, 10))
                            slCity = Trim$(Left(slAddr3, InStr(1, slAddr3, ",") - 1)) '-226 -50
                            slState = Mid([slAddr3], InStr([slAddr3], ",") + 2, 2)   '-276 -50
                            FoundAdd = True
                        End If
                    End If
                End If

            If tmAgf.sCreditRestr = "L" Then
                 slCredit = gLongToStrDec(tmAgf.lCreditLimit, 2) '-541 -20
            End If
              
    
            End If

    tmPnfSrchKey.lCode = tmAgf.iPnfBuyer
    ilRet = btrGetEqual(hmPnf, tmPnf, imPnfRecLen, tmPnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
         slPnfName = tmPnf.sName ' - 571 - 20
        If Left$(tmPnf.sPhone, 1) = "_" Then
         slPnfPhoneNumber = ""
        Else
         slPnfPhoneNumber = Left$(tmPnf.sPhone, 3) & "-" & Mid$(tmPnf.sPhone, 4, 7)  ' gFormatPhoneNo(tmPnf.sPhone) ' - 396 - 35
        End If
    Else
        slPnfName = ""
        slPnfPhoneNumber = ""
    End If
    
    tmSlfSrchKey.lCode = tmAgf.iSlfCode
    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    If ilRet = BTRV_ERR_NONE Then
        slSalesman = Left$(Trim$(tmSlf.sLastName), 8) & "," & LeftB$(Trim$(tmSlf.sFirstName), 1) '- 561 - 10
    Else
        slSalesman = ""
    End If

            
        tmMasterRec(1).sFieldValue = slAName
        tmMasterRec(1).iStartPos = 41
        tmMasterRec(1).iLength = 35
        tmMasterRec(1).bLeadingZeroes = False
        tmMasterRec(1).sJustification = 0
        tmMasterRec(1).sDataType = 0
        
        tmMasterRec(2).sFieldValue = slAddr1
        tmMasterRec(2).iStartPos = 76
        tmMasterRec(2).iLength = 50
        tmMasterRec(2).bLeadingZeroes = False
        tmMasterRec(2).sJustification = 0
        tmMasterRec(2).sDataType = 0
        
        tmMasterRec(3).sFieldValue = slAddr2
        tmMasterRec(3).iStartPos = 126
        tmMasterRec(3).iLength = 50
        tmMasterRec(3).bLeadingZeroes = False
        tmMasterRec(3).sJustification = 0
        tmMasterRec(3).sDataType = 0
        
        tmMasterRec(4).sFieldValue = slAddr3
        tmMasterRec(4).iStartPos = 176
        tmMasterRec(4).iLength = 50
        tmMasterRec(4).bLeadingZeroes = False
        tmMasterRec(4).sJustification = 0
        tmMasterRec(4).sDataType = 0

        tmMasterRec(5).sFieldValue = slCity
        tmMasterRec(5).iStartPos = 226
        tmMasterRec(5).iLength = 50
        tmMasterRec(5).bLeadingZeroes = False
        tmMasterRec(5).sJustification = 0
        tmMasterRec(5).sDataType = 0

        tmMasterRec(6).sFieldValue = slState
        tmMasterRec(6).iStartPos = 276
        tmMasterRec(6).iLength = 50
        tmMasterRec(6).bLeadingZeroes = False
        tmMasterRec(6).sJustification = 0
        tmMasterRec(6).sDataType = 0
        
        tmMasterRec(7).sFieldValue = slZip
        tmMasterRec(7).iStartPos = 326
        tmMasterRec(7).iLength = 20
        tmMasterRec(7).bLeadingZeroes = False
        tmMasterRec(7).sJustification = 0
        tmMasterRec(7).sDataType = 0
        
        tmMasterRec(8).sFieldValue = slPnfPhoneNumber
        tmMasterRec(8).iStartPos = 396
        tmMasterRec(8).iLength = 35
        tmMasterRec(8).bLeadingZeroes = False
        tmMasterRec(8).sJustification = 0
        tmMasterRec(8).sDataType = 0
        
        tmMasterRec(9).sFieldValue = slCredit
        tmMasterRec(9).iStartPos = 541
        tmMasterRec(9).iLength = 20
        tmMasterRec(9).bLeadingZeroes = False
        tmMasterRec(9).sJustification = 0
        tmMasterRec(9).sDataType = 0
                
        tmMasterRec(10).sFieldValue = slSalesman
        tmMasterRec(10).iStartPos = 561
        tmMasterRec(10).iLength = 10
        tmMasterRec(10).bLeadingZeroes = False
        tmMasterRec(10).sJustification = 0
        tmMasterRec(10).sDataType = 0
        
        tmMasterRec(11).sFieldValue = slPnfName
        tmMasterRec(11).iStartPos = 571
        tmMasterRec(11).iLength = 20
        tmMasterRec(11).bLeadingZeroes = False
        tmMasterRec(11).sJustification = 0
        tmMasterRec(11).sDataType = 0
        
        tmMasterRec(12).sFieldValue = slLastPay
        tmMasterRec(12).iStartPos = 611
        tmMasterRec(12).iLength = 8
        tmMasterRec(12).bLeadingZeroes = False
        tmMasterRec(12).sJustification = 0
        tmMasterRec(12).sDataType = 0
        
        tmMasterRec(13).sFieldValue = slLastInv
        tmMasterRec(13).iStartPos = 619
        tmMasterRec(13).iLength = 8
        tmMasterRec(13).bLeadingZeroes = False
        tmMasterRec(13).sJustification = 0
        tmMasterRec(13).sDataType = 0
        
        tmMasterRec(14).sFieldValue = slCurBal
        tmMasterRec(14).iStartPos = 637
        tmMasterRec(14).iLength = 20
        tmMasterRec(14).bLeadingZeroes = False
        tmMasterRec(14).sJustification = 1
        tmMasterRec(14).sDataType = 0
        
        tmMasterRec(15).sFieldValue = slHiCredit
        tmMasterRec(15).iStartPos = 813
        tmMasterRec(15).iLength = 20
        tmMasterRec(15).bLeadingZeroes = False
        tmMasterRec(15).sJustification = 1
        tmMasterRec(15).sDataType = 0


        slRec = gCreateFixedRec(1233, tmMasterRec())
        Print #hmTo, slRec
        
End Sub

Private Function mCheckRecLength() As Integer
    If Not gRecLengthOk("Rvf.Btr", Len(tmRvf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Prf.Btr", Len(tmPrf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Adf.Btr", Len(tmAdf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Agf.Btr", Len(tmAgf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Chf.Btr", Len(tmChf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Pnf.Btr", Len(tmPnf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Slf.Btr", Len(tmSlf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    mCheckRecLength = True
End Function

Private Sub mCheckGG()
    Dim ilRet As Integer
    Dim ilField1 As Integer
    Dim ilField2 As Integer
    Dim llNow As Long
    Dim llDate As Long
    Dim slStr As String
    Dim ilLoop As Integer
    
    On Error Resume Next
    
    'If imLastHourGGChecked = Hour(Now) Then
    '    Exit Sub
    'End If
    'imLastHourGGChecked = Hour(Now)
    igGGFlag = 1
    hmSaf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSaf, "", sgDBPath & "Saf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    imSafRecLen = Len(tmSaf)
    ilRet = btrGetFirst(hmSaf, tmSaf, imSafRecLen, 0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSaf
        Exit Sub
    End If
    
    ilField1 = Asc(tmSaf.sName)
    slStr = Mid$(tmSaf.sName, 2, 5)
    llDate = Val(slStr)
    ilField2 = Asc(Mid$(tmSaf.sName, 11, 1))
    llNow = gDateValue(Format$(Now, "m/d/yy"))
    If (ilField1 = 0) And (ilField2 = 1) Then
        If llDate <= llNow Then
            ilField2 = 0
        End If
    End If
    If (ilField1 = 0) And (ilField2 = 0) Then
        igGGFlag = 0
    End If
    'gSetRptGGFlag tmSaf
    btrDestroy hmSaf
End Sub

