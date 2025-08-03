VERSION 5.00
Begin VB.Form SetCrdit 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Set Credit"
   ClientHeight    =   285
   ClientLeft      =   4260
   ClientTop       =   3645
   ClientWidth     =   1800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "SETCRDIT.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   285
   ScaleWidth      =   1800
   WindowState     =   1  'Minimized
   Begin VB.Timer tmcSetTime 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1320
      Top             =   60
   End
   Begin VB.Timer tmcTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3270
      Top             =   435
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2505
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   195
      Visible         =   0   'False
      Width           =   525
   End
End
Attribute VB_Name = "SetCrdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SetCrdit.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Set Advertiser/Agency Credit code
Option Explicit
Option Compare Text
'Program library dates Field Areas
Dim tmAgf As AGF        'AGf record image
Dim tmAgfSrchKey As INTKEY0    'AGf key record image
Dim hmAgf As Integer    'Agency file handle
Dim imAgfRecLen As Integer        'Cif record length
Dim tmAgfCredit() As AGFCREDIT
Dim tmAdf As ADF        'Adf record image
Dim tmAdfSrchKey As INTKEY0    'Adf key record image
Dim hmAdf As Integer    'Advertiser file handle
Dim imAdfRecLen As Integer        'ADF record length
Dim tmAdfCredit() As ADFCREDIT
Dim tmSpf As SPF        'Spf record image
Dim tmSpfSrchKey As INTKEY0    'Spf key record image
Dim hmSpf As Integer    'Site preference file handle
Dim imSpfRecLen As Integer        'SPF record length
Dim tmChf As CHF        'Chf record image
Dim tmChfSrchKey As LONGKEY0    'Chf key record image
Dim hmCHF As Integer    'Contract file handle
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChfAdvtExt() As CHFADVTEXT
Dim tmClf As CLF        'Clf record image
Dim tmClfSrchKey As CLFKEY0    'Clf key record image
Dim hmClf As Integer    'Contract Line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClfSC() As CLFLIST
Dim tmCff As CFF        'Cff record image
Dim hmCff As Integer    'Line Flight file handle
Dim imCffRecLen As Integer        'CFF record length
Dim tmCffSC() As CFFLIST      'CFF record image
Dim tmSbf As SBF        'Sbf record image
Dim tmSbfSrchKey As SBFKEY0    'Sbf key record image
Dim hmSbf As Integer    'Special Bill file handle
Dim imSbfRecLen As Integer        'SBF record length
'Used to indicate that updating credit is occurring on another terminate
Dim tmUrf As URF        'URf record image
Dim hmUrf As Integer    'User file handle
Dim imUrfRecLen As Integer        'URF record length
Dim imTerminate As Integer  'True = terminating task, False= OK
Dim imFirstFocus As Integer
Dim imPopReqd As Integer         'Flag indicating if cbcSelect was populated
Dim imBypassSetting As Integer      'In cbcSelect--- bypass mSetCommands (when user entering new name, don't want cbcSelect disabled)
Dim smNowDate As String
Dim lmNowDate As Long
Dim smCommandDate As String
Dim lmAgfNoRecs As Long
Dim lmAdfNoRecs As Long
Dim imOkToUnload As Integer
Dim bmForceRun As Boolean

Private Sub edcLinkDestHelpMsg_Change()
    igParentRestarted = True
End Sub
Private Sub Form_Activate()
'    gShowBranner
End Sub


Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    If App.PrevInstance Then
        End
    End If
    gSetBkgdMode
    igBkgdProg = 2
    mInit
'    If imTerminate Then
'        mTerminate
'    Else
        tmcTimer.Enabled = True
        gOpenTmf
        tmcSetTime.Interval = 1000 * MONITORTIMEINTERVAL
        tmcSetTime.Enabled = True
'    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ilRes As Integer
    Dim ilRet As Integer
    Cancel = 0
    If (Not imOkToUnload) And (UnloadMode > 1) Then
        ilRes = MsgBox("Select OK to Terminate Setting of Credit Values, Select Cancel to Continue Setting Credit Values and Stop the Exit from Windows", vbOKCancel + vbQuestion + vbDefaultButton2, "Set Credit")
        If ilRes = vbCancel Then
            Cancel = 1  'Any non-zero value stop the unload
        Else
            'Do
            '    ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            '    tmSpf.sLkCredit = ""
            '    ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
            'Loop While ilRet = BTRV_ERR_CONFLICT
        End If
        'ilRet = btrUnlock(hmUrf, BTRV_UNLOCK_MULTIPLE)
    Else
        'Do
        '    ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        '    tmSpf.sLkCredit = ""
        '    ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
        'Loop While ilRet = BTRV_ERR_CONFLICT
        ''ilRet = btrUnlock(hmUrf, BTRV_UNLOCK_MULTIPLE)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    
    tmcSetTime.Enabled = False
    If igLogActivityStatus = 32123 Then
        igLogActivityStatus = -32123
        gUserActivityLog "", ""
    End If
    
    gCloseTmf
    
    Erase tmAgfCredit
    Erase tmAdfCredit
    Erase tmClfSC
    Erase tmCffSC
    btrExtClear hmSbf   'Clear any previous extend operation
    ilRet = btrClose(hmSbf)
    btrDestroy hmSbf
    btrExtClear hmCff   'Clear any previous extend operation
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    btrExtClear hmClf   'Clear any previous extend operation
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    btrExtClear hmCHF   'Clear any previous extend operation
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAgf)
    btrDestroy hmAgf
    btrExtClear hmAdf   'Clear any previous extend operation
    ilRet = btrClose(hmAdf)
    btrDestroy hmAdf
    btrExtClear hmSpf   'Clear any previous extend operation
    ilRet = btrClose(hmSpf)
    btrDestroy hmSpf
    btrExtClear hmUrf   'Clear any previous extend operation
    ilRet = btrClose(hmUrf) 'This will remove the Lock
    btrDestroy hmUrf
    
    btrStopAppl
    Set SetCrdit = Nothing   'Remove data segment
    End
End Sub
Private Sub mComputeCredit()
    Dim ilRet As Integer
    Dim slDate As String
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim llDate As Long
    'Dim llNoRec As Long         'Number of records in Sof
    'Dim ilExtLen As Integer
    'Dim llRecPos As Long        'Record location
    Dim slSpotPrice As String
    'Dim llSpotDate As Long
    'Dim ilOffset As Integer
    Dim ilLoop As Integer
    'Dim ilUsingExt As Integer
    Dim ilFound As Integer
    Dim ilCode As Integer
    Dim slAmount As String
    Dim ilConflict As Integer
    Dim ilCounter As Integer
    Dim slCntrStatus As String
    Dim ilHOType As Integer
    Dim slCntrType As String
    Dim llClfStartDate As Long
    Dim llClfEndDate As Long
    Dim ilClf As Integer
    Dim llCffStartDate As Long
    Dim llCffEndDate As Long
    Dim ilCff As Integer
    Dim ilDay As Integer
    Dim slPrice As String
    Dim ilChf As Integer
    Dim slAgyRate As String
    
    If (tmSpf.sRUnbilled <> "N") Or (tmSpf.iRNoWks <> 0) Then
        If tmSpf.sRUnbilled = "Y" Then
            gUnpackDate tmSpf.iBLastStdMnth(0), tmSpf.iBLastStdMnth(1), slDate
            If slDate <> "" Then
                llStartDate = gDateValue(slDate) + 1
            Else
                llStartDate = lmNowDate + 1
            End If
        Else
            llStartDate = lmNowDate + 1
        End If
        If tmSpf.iRNoWks >= 13 Then
            llEndDate = lmNowDate + 7 * tmSpf.iRNoWks
        Else
            llEndDate = lmNowDate + 13 * 7
        End If
        slStartDate = Format$(llStartDate, "m/d/yy")
        slEndDate = Format$(llEndDate, "m/d/yy")
    
'        ilRet = btrGetFirst(hmVef, tmVef, imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'        Do While ilRet = BTRV_ERR_NONE
'            DoEvents
'            If (tmVef.sType = "S") Or (tmVef.sType = "C") Then
'                DoEvents
'                ilExtLen = Len(tmSdf)   'Len(tmSdfCreditExt)  'Extract operation record size
'                llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
'                btrExtClear hmSdf   'Clear any previous extend operation
'                tmSdfSrchKey.iVefCode = tmVef.iCode
'                gPackDate slStartDate, tmSdfSrchKey.iDate(0), tmSdfSrchKey.iDate(1)
'                tmSdfSrchKey.iTime(0) = 0
'                tmSdfSrchKey.iTime(1) = 0
'                tmSdfSrchKey.sSchStatus = " "
'                DoEvents
'                ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'                If ilUsingExt Then
'                    If (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tmVef.iCode) Then
'                        'Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC") 'Set extract limits (all records)
'                        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
'                        DoEvents
'                        tlIntTypeBuff.iType = tmVef.iCode
'                        ilOffset = gFieldOffset("Sdf", "SdfVefCode")
'                        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
'                        If ilRet <> BTRV_ERR_NONE Then
'                            Exit Sub
'                        End If
'                        DoEvents
'                        gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
'                        ilOffset = gFieldOffset("Sdf", "SdfDate")
'                        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
'                        If ilRet <> BTRV_ERR_NONE Then
'                            Exit Sub
'                        End If
'                        DoEvents
'                        gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
'                        ilOffset = gFieldOffset("Sdf", "SdfDate")
'                        ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
'                        If ilRet <> BTRV_ERR_NONE Then
'                            Exit Sub
'                        End If
'                        ilOffset = 0
'                        ilRet = btrExtAddField(hmSdf, ilOffset, Len(tmSdf))  'Extract iCode field
'                        If ilRet <> BTRV_ERR_NONE Then
'                            Exit Sub
'                        End If
'                        'DoEvents
'                        'ilOffset = gFieldOffset("Sdf", "SdfCode")
'                        'ilRet = btrExtAddField(hmSdf, ilOffset, 4)  'Extract iCode field
'                        'If ilRet <> BTRV_ERR_NONE Then
'                        '    Exit Sub
'                        'End If
'                        'DoEvents
'                        'ilOffset = gFieldOffset("Sdf", "SdfVefCode")
'                        'ilRet = btrExtAddField(hmSdf, ilOffset, 2)  'Extract iCode field
'                        'If ilRet <> BTRV_ERR_NONE Then
'                        '    Exit Sub
'                        'End If
'                        'DoEvents
'                        'ilOffset = gFieldOffset("Sdf", "SdfChfCode")
'                        'ilRet = btrExtAddField(hmSdf, ilOffset, 4)  'Extract iCode field
'                        'If ilRet <> BTRV_ERR_NONE Then
'                        '    Exit Sub
'                        'End If
'                        'DoEvents
'                        'ilOffset = gFieldOffset("Sdf", "SdfLineNo")
'                        'ilRet = btrExtAddField(hmSdf, ilOffset, 2)  'Extract iCode field
'                        'If ilRet <> BTRV_ERR_NONE Then
'
'                        '    Exit Sub
'                        'End If
'                        'DoEvents
'                        'ilOffset = gFieldOffset("Sdf", "SdfDate")
'                        'ilRet = btrExtAddField(hmSdf, ilOffset, 4)  'Extract iCode field
'                        'If ilRet <> BTRV_ERR_NONE Then
'                        '    Exit Sub
'                        'End If
'                        'DoEvents
'                        'ilOffset = gFieldOffset("Sdf", "SdfSchStatus")
'                        'ilRet = btrExtAddField(hmSdf, ilOffset, 1)  'Extract iCode field
'                        'If ilRet <> BTRV_ERR_NONE Then
'                        '    Exit Sub
'                        'End If
'                        'DoEvents
'                        'ilOffset = gFieldOffset("Sdf", "SdfPriceType")
'                        'ilRet = btrExtAddField(hmSdf, ilOffset, 1)  'Extract iCode field
'                        'If ilRet <> BTRV_ERR_NONE Then
'                        '    Exit Sub
'                        'End If
'                        'ilOffset = gFieldOffset("Sdf", "SdfBill")
'                        'ilRet = btrExtAddField(hmSdf, ilOffset, 1)  'Extract iCode field
'                        'If ilRet <> BTRV_ERR_NONE Then
'                        '    Exit Sub
'                        'End If
'                        'ilRet = btrExtGetNextExt(hmSdf)  'Extract records
'                        DoEvents
'                        'ilRet = btrExtGetNext(hmSdf, tmSdfCreditExt, ilExtLen, llRecPos)
'                        ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
'                        DoEvents
'                        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
'                            If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
'                                ilExtLen = Len(tmSdf)   'Len(tmSdfCreditExt)  'Extract operation record size
'                                'ilRet = btrExtGetFirst(hlAdf, tgCommAdf(ilUpperBound), ilExtLen, llRecPos)
'                                Do While ilRet = BTRV_ERR_REJECT_COUNT
'                                    'ilRet = btrExtGetNext(hmSdf, tmSdfCreditExt, ilExtLen, llRecPos)
'                                    ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
'                                Loop
'                                Do While ilRet = BTRV_ERR_NONE
'                                    DoEvents
'                                    'If (tmSdfCreditExt.sSchStatus <> "C") And (tmSdfCreditExt.sSchStatus <> "H") And (tmSdfCreditExt.sBill <> "Y") Then
'                                    If (tmSdf.sSchStatus <> "C") And (tmSdf.sSchStatus <> "H") And (tmSdf.sBill <> "Y") Then
'                                        'gUnpackDate tmSdfCreditExt.iDate(0), tmSdfCreditExt.iDate(1), slDate
'                                        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
'                                        llSpotDate = gDateValue(slDate)
'                                        'Process spot
'                                        'tmChfSrchKey.lCode = tmSdfCreditExt.lChfCode
'                                        tmChfSrchKey.lCode = tmSdf.lChfCode
'                                        ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'                                        If (ilRet = BTRV_ERR_NONE) Then
'                                            DoEvents
'                                            'If PSA or Promo- bypass
'                                            If (tmChf.sType <> "S") And (tmChf.sType <> "M") Then
'                                                'tmClfSrchKey.lChfCode = tmSdfCreditExt.lChfCode
'                                                'tmClfSrchKey.iLine = tmSdfCreditExt.iLineNo
'                                                tmClfSrchKey.lChfCode = tmSdf.lChfCode
'                                                tmClfSrchKey.iLine = tmSdf.iLineNo
'                                                tmClfSrchKey.iCntRevNo = 32000
'                                                tmClfSrchKey.iPropVer = 32000
'                                                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'                                                Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And (tmClf.sSchStatus = "A")
'                                                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                                                Loop
'                                                If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdfCreditExt.lChfCode) And (tmClf.iLine = tmSdfCreditExt.iLineNo) Then
'                                                    DoEvents
'                                                    ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slSpotPrice)
'                                                    DoEvents
'                                                    If InStr(slSpotPrice, ".") = 0 Then
'                                                        slSpotPrice = ".00"
'                                                    End If
'                                                    'slSpotPrice = ".00"
'                                                    'Select Case tmClf.sPriceType
'                                                    '    Case "T"    'True
'                                                    '        If (tmSdfCreditExt.sPriceType <> "N") And (tmSdfCreditExt.sPriceType <> "P") Then
'                                                    '            gPDNToStr tmClf.sActPrice, 2, slSpotPrice
'                                                    '        End If
'                                                    '    Case "N"    'No Charge
'                                                    '    Case "M"    'MG Line
'                                                    '    Case "B"    'Bonus
'                                                    '    Case "S"    'Spinoff
'                                                    '    Case "P"    'Package
'                                                    '        gPDNToStr tmClf.sActPrice, 2, slSpotPrice
'                                                    '    Case "R"    'Recapturable
'                                                    '    Case "A"    'ADU
'                                                    'End Select
'                                                    'DoEvents
'                                                    For ilLoop = 0 To UBound(tmAdfCredit) Step 1
'                                                        If tmAdfCredit(ilLoop).iCode = -1 Then
'                                                            tmAdfCredit(ilLoop).iCode = tmChf.iAdfCode
'                                                            If llSpotDate <= lmNowDate Then
'                                                                tmAdfCredit(ilLoop).sUnbilled = slSpotPrice
'                                                            Else
'                                                                If llSpotDate - lmNowDate <= 7 Then
'                                                                    tmAdfCredit(ilLoop).s1Wk = slSpotPrice
'                                                                End If
'                                                                If llSpotDate - lmNowDate <= 28 Then
'                                                                    tmAdfCredit(ilLoop).s4Wks = slSpotPrice
'                                                                End If
'                                                                If llSpotDate - lmNowDate <= 91 Then
'                                                                    tmAdfCredit(ilLoop).s13Wks = slSpotPrice
'                                                                End If
'                                                                If llSpotDate - lmNowDate <= 7 * tmSpf.iRNoWks Then
'                                                                    tmAdfCredit(ilLoop).sProjected = slSpotPrice
'                                                                End If
'                                                            End If
'                                                            Exit For
'                                                        End If
'                                                        If tmAdfCredit(ilLoop).iCode = tmChf.iAdfCode Then
'                                                            If llSpotDate <= lmNowDate Then
'                                                                tmAdfCredit(ilLoop).sUnbilled = gAddStr(tmAdfCredit(ilLoop).sUnbilled, slSpotPrice)
'                                                            Else
'                                                                If llSpotDate - lmNowDate <= 7 Then
'                                                                    tmAdfCredit(ilLoop).s1Wk = gAddStr(tmAdfCredit(ilLoop).s1Wk, slSpotPrice)
'                                                                End If
'                                                                If llSpotDate - lmNowDate <= 28 Then
'                                                                    tmAdfCredit(ilLoop).s4Wks = gAddStr(tmAdfCredit(ilLoop).s4Wks, slSpotPrice)
'                                                                End If
'                                                                If llSpotDate - lmNowDate <= 91 Then
'                                                                    tmAdfCredit(ilLoop).s13Wks = gAddStr(tmAdfCredit(ilLoop).s13Wks, slSpotPrice)
'                                                                End If
'                                                                If llSpotDate - lmNowDate <= 7 * tmSpf.iRNoWks Then
'                                                                    tmAdfCredit(ilLoop).sProjected = gAddStr(tmAdfCredit(ilLoop).sProjected, slSpotPrice)
'                                                                End If
'                                                            End If
'                                                            Exit For
'                                                        End If
'                                                        DoEvents
'                                                    Next ilLoop
'                                                    DoEvents
'                                                    If tmChf.iAgfCode > 0 Then
'                                                        For ilLoop = 0 To UBound(tmAgfCredit) Step 1
'                                                            If tmAgfCredit(ilLoop).iCode = -1 Then
'                                                                tmAgfCredit(ilLoop).iCode = tmChf.iAgfCode
'                                                                If llSpotDate <= lmNowDate Then
'                                                                    tmAgfCredit(ilLoop).sUnbilled = slSpotPrice
'                                                                Else
'                                                                    If llSpotDate - lmNowDate <= 7 Then
'                                                                        tmAgfCredit(ilLoop).s1Wk = slSpotPrice
'                                                                    End If
'                                                                    If llSpotDate - lmNowDate <= 28 Then
'                                                                        tmAgfCredit(ilLoop).s4Wks = slSpotPrice
'                                                                    End If
'                                                                    If llSpotDate - lmNowDate <= 91 Then
'                                                                        tmAgfCredit(ilLoop).s13Wks = slSpotPrice
'                                                                    End If
'                                                                    If llSpotDate - lmNowDate <= 7 * tmSpf.iRNoWks Then
'                                                                        tmAgfCredit(ilLoop).sProjected = slSpotPrice
'                                                                    End If
'                                                                End If
'                                                                Exit For
'                                                            End If
'                                                            If tmAgfCredit(ilLoop).iCode = tmChf.iAgfCode Then
'                                                                If llSpotDate <= lmNowDate Then
'                                                                    tmAgfCredit(ilLoop).sUnbilled = gAddStr(tmAgfCredit(ilLoop).sUnbilled, slSpotPrice)
'                                                                Else
'                                                                    If llSpotDate - lmNowDate <= 7 Then
'                                                                        tmAgfCredit(ilLoop).s1Wk = gAddStr(tmAgfCredit(ilLoop).s1Wk, slSpotPrice)
'                                                                    End If
'                                                                    If llSpotDate - lmNowDate <= 28 Then
'                                                                        tmAgfCredit(ilLoop).s4Wks = gAddStr(tmAgfCredit(ilLoop).s4Wks, slSpotPrice)
'                                                                    End If
'                                                                    If llSpotDate - lmNowDate <= 91 Then
'                                                                        tmAgfCredit(ilLoop).s13Wks = gAddStr(tmAgfCredit(ilLoop).s13Wks, slSpotPrice)
'                                                                    End If
'                                                                    If llSpotDate - lmNowDate <= 7 * tmSpf.iRNoWks Then
'                                                                        tmAgfCredit(ilLoop).sProjected = gAddStr(tmAgfCredit(ilLoop).sProjected, slSpotPrice)
'                                                                    End If
'                                                                End If
'                                                                Exit For
'                                                            End If
'                                                            DoEvents
'                                                        Next ilLoop
'                                                    End If
'                                                    DoEvents
'                                                End If
'                                            End If
'                                        End If
'                                    End If
'                                    'ilRet = btrExtGetNext(hmSdf, tmSdfCreditExt, ilExtLen, llRecPos)
'                                    ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
'                                    Do While ilRet = BTRV_ERR_REJECT_COUNT
'                                        'ilRet = btrExtGetNext(hmSdf, tmSdfCreditExt, ilExtLen, llRecPos)
'                                        ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
'                                    Loop
'                                    DoEvents
'                                Loop
'                            End If
'                        End If
'                    End If
'                Else
'                    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tmVef.iCode)
'                        DoEvents
'                        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
'                        llSpotDate = gDateValue(slDate)
'                        DoEvents
'                        If llSpotDate > llEndDate Then
'                            Exit Do
'                        End If
'                        If (tmSdf.sSchStatus <> "C") And (tmSdf.sSchStatus <> "H") And (tmSdf.sBill <> "Y") Then
'                            'Process spot
'                            tmChfSrchKey.lCode = tmSdf.lChfCode
'                            ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'                            If (ilRet = BTRV_ERR_NONE) Then
'                                DoEvents
'                                'If PSA or Promo- bypass
'                                If (tmChf.sType <> "S") And (tmChf.sType <> "M") Then
'                                    tmClfSrchKey.lChfCode = tmSdf.lChfCode
'                                    tmClfSrchKey.iLine = tmSdf.iLineNo
'                                    tmClfSrchKey.iCntRevNo = 32000
'                                    tmClfSrchKey.iPropVer = 32000
'                                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'                                    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And (tmClf.sSchStatus = "A")
'                                        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                                    Loop
'                                    If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
'                                        DoEvents
'                                        ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slSpotPrice)
'                                        DoEvents
'                                        If InStr(slSpotPrice, ".") = 0 Then
'                                            slSpotPrice = ".00"
'                                        End If
'                                        'slSpotPrice = ".00"
'                                        'Select Case tmClf.sPriceType
'                                        '    Case "T"    'True
'                                        '        If (tmSdf.sPriceType <> "N") And (tmSdf.sPriceType <> "P") Then
'                                        '            gPDNToStr tmClf.sActPrice, 2, slSpotPrice
'                                        '        End If
'                                        '    Case "N"    'No Charge
'                                        '    Case "M"    'MG Line
'                                        '    Case "B"    'Bonus
'                                        '    Case "S"    'Spinoff
'                                        '    Case "P"    'Package
'                                        '        gPDNToStr tmClf.sActPrice, 2, slSpotPrice
'                                        '    Case "R"    'Recapturable
'                                        '    Case "A"    'ADU
'                                        'End Select
'                                        ilCounter = 0
'                                        For ilLoop = 0 To UBound(tmAdfCredit) Step 1
'                                            If tmAdfCredit(ilLoop).iCode = -1 Then
'                                                tmAdfCredit(ilLoop).iCode = tmChf.iAdfCode
'                                                If llSpotDate <= lmNowDate Then
'                                                    tmAdfCredit(ilLoop).sUnbilled = slSpotPrice
'                                                Else
'                                                    If llSpotDate - lmNowDate <= 7 Then
'                                                        tmAdfCredit(ilLoop).s1Wk = slSpotPrice
'                                                    End If
'                                                    If llSpotDate - lmNowDate <= 28 Then
'                                                        tmAdfCredit(ilLoop).s4Wks = slSpotPrice
'                                                    End If
'                                                    If llSpotDate - lmNowDate <= 91 Then
'                                                        tmAdfCredit(ilLoop).s13Wks = slSpotPrice
'                                                    End If
'                                                    If llSpotDate - lmNowDate <= 7 * tmSpf.iRNoWks Then
'                                                        tmAdfCredit(ilLoop).sProjected = slSpotPrice
'                                                    End If
'                                                End If
'                                                Exit For
'                                            End If
'                                            DoEvents
'                                            If tmAdfCredit(ilLoop).iCode = tmChf.iAdfCode Then
'                                                If llSpotDate <= lmNowDate Then
'                                                    tmAdfCredit(ilLoop).sUnbilled = gAddStr(tmAdfCredit(ilLoop).sUnbilled, slSpotPrice)
'                                                Else
'                                                    If llSpotDate - lmNowDate <= 7 Then
'                                                        tmAdfCredit(ilLoop).s1Wk = gAddStr(tmAdfCredit(ilLoop).s1Wk, slSpotPrice)
'                                                    End If
'                                                    If llSpotDate - lmNowDate <= 28 Then
'                                                        tmAdfCredit(ilLoop).s4Wks = gAddStr(tmAdfCredit(ilLoop).s4Wks, slSpotPrice)
'                                                    End If
'                                                    If llSpotDate - lmNowDate <= 91 Then
'                                                        tmAdfCredit(ilLoop).s13Wks = gAddStr(tmAdfCredit(ilLoop).s13Wks, slSpotPrice)
'                                                    End If
'                                                    If llSpotDate - lmNowDate <= 7 * tmSpf.iRNoWks Then
'                                                        tmAdfCredit(ilLoop).sProjected = gAddStr(tmAdfCredit(ilLoop).sProjected, slSpotPrice)
'                                                    End If
'                                                End If
'                                                Exit For
'                                            End If
'                                            ilCounter = ilCounter + 1
'                                            If ilCounter > 40 Then
'                                                ilCounter = 0
'                                                DoEvents
'                                            End If
'                                        Next ilLoop
'                                        If tmChf.iAgfCode > 0 Then
'                                            For ilLoop = 0 To UBound(tmAgfCredit) Step 1
'                                                If tmAgfCredit(ilLoop).iCode = -1 Then
'                                                    tmAgfCredit(ilLoop).iCode = tmChf.iAgfCode
'                                                    If llSpotDate <= lmNowDate Then
'                                                        tmAgfCredit(ilLoop).sUnbilled = slSpotPrice
'                                                    Else
'                                                        If llSpotDate - lmNowDate <= 7 Then
'                                                            tmAgfCredit(ilLoop).s1Wk = slSpotPrice
'                                                        End If
'                                                        If llSpotDate - lmNowDate <= 28 Then
'                                                            tmAgfCredit(ilLoop).s4Wks = slSpotPrice
'                                                        End If
'                                                        If llSpotDate - lmNowDate <= 91 Then
'                                                            tmAgfCredit(ilLoop).s13Wks = slSpotPrice
'                                                        End If
'                                                        If llSpotDate - lmNowDate <= 7 * tmSpf.iRNoWks Then
'                                                            tmAgfCredit(ilLoop).sProjected = slSpotPrice
'                                                        End If
'                                                    End If
'                                                    Exit For
'                                                End If
'                                                DoEvents
'                                                If tmAgfCredit(ilLoop).iCode = tmChf.iAgfCode Then
'                                                    If llSpotDate <= lmNowDate Then
'                                                        tmAgfCredit(ilLoop).sUnbilled = gAddStr(tmAgfCredit(ilLoop).sUnbilled, slSpotPrice)
'                                                    Else
'                                                        If llSpotDate - lmNowDate <= 7 Then
'                                                            tmAgfCredit(ilLoop).s1Wk = gAddStr(tmAgfCredit(ilLoop).s1Wk, slSpotPrice)
'                                                        End If
'                                                        If llSpotDate - lmNowDate <= 28 Then
'                                                            tmAgfCredit(ilLoop).s4Wks = gAddStr(tmAgfCredit(ilLoop).s4Wks, slSpotPrice)
'                                                        End If
'                                                        If llSpotDate - lmNowDate <= 91 Then
'                                                            tmAgfCredit(ilLoop).s13Wks = gAddStr(tmAgfCredit(ilLoop).s13Wks, slSpotPrice)
'                                                        End If
'                                                        If llSpotDate - lmNowDate <= 7 * tmSpf.iRNoWks Then
'                                                            tmAgfCredit(ilLoop).sProjected = gAddStr(tmAgfCredit(ilLoop).sProjected, slSpotPrice)
'                                                        End If
'                                                    End If
'                                                    Exit For
'                                                End If
'                                                ilCounter = ilCounter + 1
'                                                If ilCounter > 40 Then
'                                                    ilCounter = 0
'                                                    DoEvents
'                                                End If
'                                            Next ilLoop
'                                        End If
'                                    End If
'                                End If
'                            End If
'                        End If
'                        ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                    Loop
'                End If
'            End If
'            ilRet = btrGetNext(hmVef, tmVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
        slCntrStatus = "HO"
        slCntrType = ""
        ilHOType = 1
        sgCntrForDateStamp = ""
        ilRet = gObtainCntrForDate(SetCrdit, slStartDate, "", slCntrStatus, slCntrType, ilHOType, tmChfAdvtExt())
        If (ilRet = CP_MSG_NOPOPREQ) Or (ilRet = CP_MSG_NONE) Then
            'List of allowed vehicles and dayparts
            For ilChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
                DoEvents
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, tmChfAdvtExt(ilChf).lCode, False, tmChf, tmClfSC(), tmCffSC())
                If ilRet Then
                    slAgyRate = "0.00"
                    If tmChf.iAgfCode > 0 Then
                        tmAgfSrchKey.iCode = tmChf.iAgfCode
                        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            slAgyRate = gIntToStrDec(tmAgf.iComm, 2)
                        End If
                    End If
                    If tmChf.sInstallDefined = "Y" Then
                        'Get Installment
                        tmSbfSrchKey.lChfCode = tmChf.lCode
                        gPackDateLong llStartDate, tmSbfSrchKey.iDate(0), tmSbfSrchKey.iDate(1)
                        tmSbfSrchKey.sTranType = " "
                        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = tmChf.lCode)
                            If (tmSbf.sTranType = "F") And (tmSbf.sBilled <> "Y") Then   'Installment
                                gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
                                If llDate >= llStartDate Then
                                    slPrice = gLongToStrDec(tmSbf.lGross, 2)
                                    If (tmChf.iAgfCode > 0) Then
                                        slPrice = gDivStr(gMulStr(slPrice, gSubStr("100.00", slAgyRate)), "100.00")
                                    End If
                                    mAddPriceToBuckets llDate, slPrice
                                End If
                            End If
                            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    Else
                        For ilClf = LBound(tmClfSC) To UBound(tmClfSC) - 1 Step 1
                            DoEvents
                            gUnpackDateLong tmClfSC(ilClf).ClfRec.iStartDate(0), tmClfSC(ilClf).ClfRec.iStartDate(1), llClfStartDate
                            gUnpackDateLong tmClfSC(ilClf).ClfRec.iEndDate(0), tmClfSC(ilClf).ClfRec.iEndDate(1), llClfEndDate
                            If (llClfStartDate <= llClfEndDate) Then
                                ilCff = tmClfSC(ilClf).iFirstCff
                                Do While ilCff >= 0
                                    DoEvents
                                    gUnpackDateLong tmCffSC(ilCff).CffRec.iStartDate(0), tmCffSC(ilCff).CffRec.iStartDate(1), llCffStartDate
                                    gUnpackDateLong tmCffSC(ilCff).CffRec.iEndDate(0), tmCffSC(ilCff).CffRec.iEndDate(1), llCffEndDate
                                    If llCffEndDate < llCffStartDate Then
                                        Exit Do
                                    End If
                                    Do While gWeekDayLong(llCffStartDate) <> 0
                                        llCffStartDate = llCffStartDate - 1
                                    Loop
                                    For llDate = llCffStartDate To llCffEndDate Step 7
                                        If (llDate >= llStartDate) Then
                                            If tmCffSC(ilCff).CffRec.sPriceType = "T" Then
                                                slSpotPrice = ".00"
                                                If tmCffSC(ilCff).CffRec.sDyWk = "D" Then
                                                    For ilDay = 0 To 6 Step 1
                                                        If llDate + ilDay <= llEndDate Then
                                                            slPrice = gLongToStrDec(tmCffSC(ilCff).CffRec.iDay(ilDay) * tmCffSC(ilCff).CffRec.lActPrice, 2)
                                                            slSpotPrice = gAddStr(slSpotPrice, slPrice)
                                                        End If
                                                    Next ilDay
                                                Else
                                                    slPrice = gLongToStrDec((tmCffSC(ilCff).CffRec.iSpotsWk + tmCffSC(ilCff).CffRec.iXSpotsWk) * tmCffSC(ilCff).CffRec.lActPrice, 2)
                                                    slSpotPrice = gAddStr(slSpotPrice, slPrice)
                                                End If
                                                If tmChf.iAgfCode > 0 Then
                                                    slSpotPrice = gDivStr(gMulStr(slSpotPrice, gSubStr("100.00", slAgyRate)), "100.00")
                                                End If
                                                mAddPriceToBuckets llDate, slSpotPrice
                                            End If
                                        End If
                                    Next llDate
                                    ilCff = tmCffSC(ilCff).iNextCff
                                Loop
                            End If
                        Next ilClf
                        'Add NTR
                        tmSbfSrchKey.lChfCode = tmChf.lCode
                        gPackDateLong llStartDate, tmSbfSrchKey.iDate(0), tmSbfSrchKey.iDate(1)
                        tmSbfSrchKey.sTranType = " "
                        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.lChfCode = tmChf.lCode)
                            If tmSbf.sTranType = "I" Then   'Items Billing
                                gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
                                If llDate >= llStartDate Then
                                    slPrice = gLongToStrDec(tmSbf.lGross, 2)
                                    slPrice = gMulStr(slPrice, Trim$(Str$(tmSbf.iNoItems)))
                                    If (tmChf.iAgfCode > 0) And (tmSbf.sAgyComm = "Y") Then
                                        slPrice = gDivStr(gMulStr(slPrice, gSubStr("100.00", slAgyRate)), "100.00")
                                    End If
                                    mAddPriceToBuckets llDate, slPrice
                                End If
                            End If
                            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                End If
            Next ilChf
        End If
    End If
    'Update Advt
    ilRet = btrGetFirst(hmAdf, tmAdf, imAdfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        ilFound = False
        ilCode = tmAdf.iCode
        ilCounter = 0
        For ilLoop = 0 To UBound(tmAdfCredit) Step 1
            If tmAdfCredit(ilLoop).iCode = ilCode Then
                ilFound = True
                Exit For
            End If
            If tmAdfCredit(ilLoop).iCode = -1 Then
                Exit For
            End If
            ilCounter = ilCounter + 1
            If ilCounter > 40 Then
                ilCounter = 0
                DoEvents
            End If
        Next ilLoop
        ilConflict = False
        Do
            If ilConflict Then
                tmAdfSrchKey.iCode = ilCode
                ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
            End If
            ilConflict = True
            DoEvents
            If ilFound Then
                If (tmAdf.sCreditRestr = "W") Then
                    slAmount = gAddStr(tmAdfCredit(ilLoop).sUnbilled, tmAdfCredit(ilLoop).s1Wk)
                ElseIf (tmAdf.sCreditRestr = "M") Then
                    slAmount = gAddStr(tmAdfCredit(ilLoop).sUnbilled, tmAdfCredit(ilLoop).s4Wks)
                ElseIf (tmAdf.sCreditRestr = "T") Then
                    slAmount = gAddStr(tmAdfCredit(ilLoop).sUnbilled, tmAdfCredit(ilLoop).s13Wks)
                ElseIf (tmAdf.sCreditRestr = "P") Then
                    slAmount = tmAdfCredit(ilLoop).sProjected
                Else
                    If tgSpf.sRUnbilled = "Y" Then
                        slAmount = gAddStr(tmAdfCredit(ilLoop).sUnbilled, tmAdfCredit(ilLoop).sProjected)
                    Else
                        slAmount = tmAdfCredit(ilLoop).sProjected
                    End If
                End If
            Else
                slAmount = ".00"
            End If
            gStrToPDN slAmount, 2, 6, tmAdf.sUnbilled
            ilRet = btrUpdate(hmAdf, tmAdf, imAdfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        ilRet = btrGetNext(hmAdf, tmAdf, imAdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    'Update Agy
    ilRet = btrGetFirst(hmAgf, tmAgf, imAgfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        DoEvents
        ilFound = False
        ilCode = tmAgf.iCode
        ilCounter = 0
        For ilLoop = 0 To UBound(tmAgfCredit) Step 1
            If tmAgfCredit(ilLoop).iCode = ilCode Then
                ilFound = True
                Exit For
            End If
            If tmAgfCredit(ilLoop).iCode = -1 Then
                Exit For
            End If
            ilCounter = ilCounter + 1
            If ilCounter > 40 Then
                ilCounter = 0
                DoEvents
            End If
        Next ilLoop
        ilConflict = False
        Do
            If ilConflict Then
                tmAgfSrchKey.iCode = ilCode
                ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    Exit Do
                End If
            End If
            ilConflict = True
            DoEvents
            If ilFound Then
                'tmAgf.sCreditRestr is defined as
                'N=No Restrictions; L=Credit Limit; W=Cash in Advance each Week;
                'M=Cash in Advance each Month; T=Cash in Advance quarterly;
                'P=Prohibit New Orders
                If (tmAgf.sCreditRestr = "W") Then
                    slAmount = gAddStr(tmAgfCredit(ilLoop).sUnbilled, tmAgfCredit(ilLoop).s1Wk)
                ElseIf (tmAgf.sCreditRestr = "M") Then
                    slAmount = gAddStr(tmAgfCredit(ilLoop).sUnbilled, tmAgfCredit(ilLoop).s4Wks)
                ElseIf (tmAgf.sCreditRestr = "T") Then
                    slAmount = gAddStr(tmAgfCredit(ilLoop).sUnbilled, tmAgfCredit(ilLoop).s13Wks)
                ElseIf (tmAgf.sCreditRestr = "P") Then
                    slAmount = tmAgfCredit(ilLoop).sProjected
                Else
                    If tgSpf.sRUnbilled = "Y" Then
                        slAmount = gAddStr(tmAgfCredit(ilLoop).sUnbilled, tmAgfCredit(ilLoop).sProjected)
                    Else
                        slAmount = tmAgfCredit(ilLoop).sProjected
                    End If
                End If
            Else
                slAmount = ".00"
            End If
            gStrToPDN slAmount, 2, 6, tmAgf.sUnbilled
            ilRet = btrUpdate(hmAgf, tmAgf, imAgfRecLen)
        Loop While ilRet = BTRV_ERR_CONFLICT
        ilRet = btrGetNext(hmAgf, tmAgf, imAgfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    Do
        ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        gPackDate smNowDate, tmSpf.iRCreditDate(0), tmSpf.iRCreditDate(1)
    '    tmSpf.sLkCredit = ""
        ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    'ilRet = btrUnlock(hmUrf, BTRV_UNLOCK_MULTIPLE)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize modular             *
'*                                                     *
'*******************************************************
Private Sub mInit()
'
'   mInit
'   Where:
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slDate As String
    Dim llAdfNoRecs As Integer
    Dim llAgfNoRecs As Integer
    imOkToUnload = True
    imTerminate = False
    
    'Screen.MousePointer = vbHourGlass
    mParseCmmdLine
    'SetCrdit.Height = cmcDone.Top + 5 * cmcDone.Height / 3
    'gCenterStdAlone SetCrdit
    'SetCrdit.Show
    'Screen.MousePointer = vbHourGlass
'    mInitDDE
    If Not gCheckDDFDates() Then
        imTerminate = True
        Exit Sub
    End If
    
    If Not mCheckRecLength() Then
        imTerminate = True
        Exit Sub
    End If
    
    imTerminate = False
    imFirstFocus = True
    smNowDate = Format$(Now, "m/d/yy")
    lmNowDate = gDateValue(smNowDate)
    hmUrf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Urf.Btr)", SetCrdit
    On Error GoTo 0
    imUrfRecLen = Len(tmUrf)
    'ilRet = btrGetFirst(hmUrf, tmUrf, imUrfRecLen, INDEXKEY0, BTRV_LOCK_SINGLE_NOWAIT, SETFORWRITE)   'Get first record as starting point of extend operation
    ilRet = btrGetFirst(hmUrf, tmUrf, imUrfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    'The lock is not working on single user version- requires multi-user
    'btrieve version- using field within Spf to indicate that setcrdit is
    'running- see below
    If (ilRet = BTRV_ERR_REC_LOCKED) Or (ilRet = BTRV_ERR_FILE_LOCKED) Then
        imTerminate = True
        Exit Sub
    End If
    hmAgf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Agf.Btr)", SetCrdit
    On Error GoTo 0
    imAgfRecLen = Len(tmAgf)
    lmAgfNoRecs = btrRecords(hmAgf) - 1
    If lmAgfNoRecs < 0 Then
        imTerminate = True
        Exit Sub
    End If
    hmAdf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Adf.Btr)", SetCrdit
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)
    lmAdfNoRecs = btrRecords(hmAdf) - 1
    If lmAdfNoRecs < 0 Then
        imTerminate = True
        Exit Sub
    End If
    hmSpf = CBtrvTable(TWOHANDLES)    'CBtrvObj()
    ilRet = btrOpen(hmSpf, "", sgDBPath & "Spf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) 'BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Spf.Btr)", SetCrdit
    On Error GoTo 0
    imSpfRecLen = Len(tmSpf)
    ''Until multi-user version of btrieve installed- test field
    'Do
        ilRet = btrGetFirst(hmSpf, tmSpf, imSpfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        'If Asc(tmSpf.sLkCredit) = 255 Then
        '    imTerminate = True
        '    Exit Sub
        'End If
        If smCommandDate = "" Then
            If InStr(1, tgSpf.sGClient, "XYZ Broadcasting", vbTextCompare) > 0 Then
                smNowDate = "12/15/1999"
                lmNowDate = gDateValue(smNowDate)
            End If
        Else
            smNowDate = smCommandDate
            lmNowDate = gDateValue(smNowDate)
        End If
        If Not bmForceRun Then
            gUnpackDate tmSpf.iRCreditDate(0), tmSpf.iRCreditDate(1), slDate
            If slDate <> "" Then
                If lmNowDate < gDateValue(slDate) Then
                    imTerminate = True
                    Exit Sub
                End If
            End If
        End If
        'tmSpf.sLkCredit = Chr(255)
        'ilRet = btrUpdate(hmSpf, tmSpf, imSpfRecLen)
    'Loop While ilRet = BTRV_ERR_CONFLICT
    hmCHF = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Chf.Btr)", SetCrdit
    On Error GoTo 0
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Clf.Btr)", SetCrdit
    On Error GoTo 0
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Cff.Btr)", SetCrdit
    On Error GoTo 0
    imCffRecLen = Len(tmCff)
    hmSbf = CBtrvTable(ONEHANDLE)    'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo mInitErr
    gBtrvErrorMsg ilRet, "mInit (btrOpen: Sbf.Btr)", SetCrdit
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf)
    If imTerminate Then
        Exit Sub
    End If
    'Determine if credit limit should be computed
    If imTerminate Then
        Exit Sub
    End If
    ilRet = gObtainAdvt()
'    gCenterModalForm SetCrdit
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    sgTmfStatus = "E"
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitDDE                        *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Initialize DDE                 *
'*                                                     *
'*******************************************************
Private Sub mInitDDE()
'    Dim ilRet As Integer
'    Dim ilLoop As Integer
'    If Not igStdAloneMode Then
'        'Set up DDE for Branner and help message area within traffic
'        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
'        edcLinkDestHelpMsg.LinkTopic = "Traffic|HelpMsg"
'        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceHelpMsg"
'        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
'    End If
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
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim ilPos As Integer
    Dim ilSpace As Integer
    
    slCommand = Command$
    sgCommandStr = "debug"
    igDirectCall = 0
    sgIniPath = ""
    'slCommand = sgCommandStr  'Command$
    'If (StrComp(slCommand, "Debug", 1) = 0) Or (StrComp(slCommand, "CS", 1) = 0) Then
        igStdAloneMode = True 'Switch from/to stand alone mode
        sgCallAppName = ""
        slStr = "Guide"
        ilTestSystem = False
    'Else
    '    igStdAloneMode = False  'Switch from/to stand alone mode
    '    ilRet = gParseItem(slCommand, 1, "\", slStr)    'Get application name
    '    If Trim$(slStr) = "" Then
    '        MsgBox "Application must be run from the Traffic application", vbCritical, "Program Schedule"
    '        End
    '    End If
    '    ilRet = gParseItem(slStr, 1, "^", sgCallAppName)    'Get application name
    '    ilRet = gParseItem(slStr, 2, "^", slTestSystem)    'Get application name
    '    If StrComp(slTestSystem, "Test", 1) = 0 Then
    '        ilTestSystem = True
    '    Else
    '        ilTestSystem = False
    '    End If
    '    ilRet = gParseItem(slCommand, 2, "\", slStr)    'Get user name
    'End If
    ilPos = InStr(1, slCommand, "/IniLoc:", vbTextCompare)
    If ilPos > 0 Then
        igDirectCall = -1
        ilSpace = InStr(ilPos, slCommand, " ")
        If ilSpace = 0 Then
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8))
        Else
            sgIniPath = Trim$(Mid$(slCommand, ilPos + 8, ilSpace - ilPos - 8))
        End If
        If right$(sgIniPath, 1) <> "\" Then
            sgIniPath = sgIniPath & "\"
        End If
    End If
    gInitStdAlone SetCrdit, slStr, ilTestSystem
    smCommandDate = ""
    ilPos = InStr(1, slCommand, "/D:", 1)
    If ilPos > 0 Then
        ilSpace = InStr(ilPos, slCommand, " ")
        If ilSpace = 0 Then
            smCommandDate = Trim$(Mid$(slCommand, ilPos + 3))
        Else
            smCommandDate = Trim$(Mid$(slCommand, ilPos + 3, ilSpace - ilPos - 3))
        End If
        If gValidDate(smCommandDate) Then
            smCommandDate = gAdjYear(smCommandDate)
        Else
            smCommandDate = ""
        End If
    End If
    ilPos = InStr(1, slCommand, "/ForceRun", 1)
    If ilPos > 0 Then
        bmForceRun = True
    Else
        bmForceRun = False
    End If
    ilRet = gObtainSAF()
    igLogActivityStatus = 32123
    gUserActivityLog "L", "SetCrdit.Frm"
    
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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

    igParentRestarted = False
    If Not igStdAloneMode Then
        'If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
        '    edcLinkDestHelpMsg.LinkExecute "@" & "Done"
        'Else
        '    edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
        '    edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
        '    edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
        '    edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
        '    edcLinkDestHelpMsg.LinkExecute "Done"
        'End If
        'Do While Not igParentRestarted
        '    DoEvents
        'Loop
    End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload Traffic
    Unload SetCrdit
    igManUnload = NO
End Sub

Private Sub tmcSetTime_Timer()
    gUpdateTaskMonitor 0, "SC"
End Sub

Private Sub tmcTimer_Timer()
    Dim ilLoop As Integer
    tmcTimer.Enabled = False
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
    gLogMsg "** Begin Set Credit: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "SetCredit.txt", False
    gUpdateTaskMonitor 1, "SC"
    imOkToUnload = False
    ReDim tmAgfCredit(0 To lmAgfNoRecs) As AGFCREDIT
    For ilLoop = 0 To UBound(tmAgfCredit) Step 1
        tmAgfCredit(ilLoop).iCode = -1  'Set to Unused
        tmAgfCredit(ilLoop).sUnbilled = ".00"
        tmAgfCredit(ilLoop).sProjected = ".00"
        tmAgfCredit(ilLoop).s1Wk = ".00"
        tmAgfCredit(ilLoop).s4Wks = ".00"
        tmAgfCredit(ilLoop).s13Wks = ".00"
    Next ilLoop
    ReDim tmAdfCredit(0 To lmAdfNoRecs) As ADFCREDIT
    For ilLoop = 0 To UBound(tmAdfCredit) Step 1
        tmAdfCredit(ilLoop).iCode = -1  'Set to unused
        tmAdfCredit(ilLoop).sUnbilled = ".00"
        tmAdfCredit(ilLoop).sProjected = ".00"
        tmAdfCredit(ilLoop).s1Wk = ".00"
        tmAdfCredit(ilLoop).s4Wks = ".00"
        tmAdfCredit(ilLoop).s13Wks = ".00"
    Next ilLoop
    mComputeCredit
    gUpdateTaskMonitor 2, "SC"
    imOkToUnload = True
    gLogMsg "** Completed Set Credit: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM"), "SetCredit.txt", False
    mTerminate
End Sub

Private Sub mAddPriceToBuckets(llDate As Long, slPrice As String)
    Dim ilLoop As Integer
    
    For ilLoop = 0 To UBound(tmAdfCredit) Step 1
        If tmAdfCredit(ilLoop).iCode = -1 Then
            tmAdfCredit(ilLoop).iCode = tmChf.iAdfCode
            If llDate <= lmNowDate Then
                tmAdfCredit(ilLoop).sUnbilled = slPrice
            Else
                If llDate - lmNowDate <= 7 Then
                    tmAdfCredit(ilLoop).s1Wk = slPrice
                End If
                If llDate - lmNowDate <= 28 Then
                    tmAdfCredit(ilLoop).s4Wks = slPrice
                End If
                tmAdfCredit(ilLoop).s13Wks = slPrice
                If llDate - lmNowDate <= 7 * tmSpf.iRNoWks Then
                    tmAdfCredit(ilLoop).sProjected = slPrice
                End If
            End If
            Exit For
        End If
        If tmAdfCredit(ilLoop).iCode = tmChf.iAdfCode Then
            If llDate <= lmNowDate Then
                tmAdfCredit(ilLoop).sUnbilled = gAddStr(tmAdfCredit(ilLoop).sUnbilled, slPrice)
            Else
                If llDate - lmNowDate <= 7 Then
                    tmAdfCredit(ilLoop).s1Wk = gAddStr(tmAdfCredit(ilLoop).s1Wk, slPrice)
                End If
                If llDate - lmNowDate <= 28 Then
                    tmAdfCredit(ilLoop).s4Wks = gAddStr(tmAdfCredit(ilLoop).s4Wks, slPrice)
                End If
                tmAdfCredit(ilLoop).s13Wks = gAddStr(tmAdfCredit(ilLoop).s13Wks, slPrice)
                If llDate - lmNowDate <= 7 * tmSpf.iRNoWks Then
                    tmAdfCredit(ilLoop).sProjected = gAddStr(tmAdfCredit(ilLoop).sProjected, slPrice)
                End If
            End If
            Exit For
        End If
        DoEvents
    Next ilLoop
    DoEvents
    If tmChf.iAgfCode > 0 Then
        For ilLoop = 0 To UBound(tmAgfCredit) Step 1
            If tmAgfCredit(ilLoop).iCode = -1 Then
                tmAgfCredit(ilLoop).iCode = tmChf.iAgfCode
                If llDate <= lmNowDate Then
                    tmAgfCredit(ilLoop).sUnbilled = slPrice
                Else
                    If llDate - lmNowDate <= 7 Then
                        tmAgfCredit(ilLoop).s1Wk = slPrice
                    End If
                    If llDate - lmNowDate <= 28 Then
                        tmAgfCredit(ilLoop).s4Wks = slPrice
                    End If
                    tmAgfCredit(ilLoop).s13Wks = slPrice
                    If llDate - lmNowDate <= 7 * tmSpf.iRNoWks Then
                        tmAgfCredit(ilLoop).sProjected = slPrice
                    End If
                End If
                Exit For
            End If
            If tmAgfCredit(ilLoop).iCode = tmChf.iAgfCode Then
                If llDate <= lmNowDate Then
                    tmAgfCredit(ilLoop).sUnbilled = gAddStr(tmAgfCredit(ilLoop).sUnbilled, slPrice)
                Else
                    If llDate - lmNowDate <= 7 Then
                        tmAgfCredit(ilLoop).s1Wk = gAddStr(tmAgfCredit(ilLoop).s1Wk, slPrice)
                    End If
                    If llDate - lmNowDate <= 28 Then
                        tmAgfCredit(ilLoop).s4Wks = gAddStr(tmAgfCredit(ilLoop).s4Wks, slPrice)
                    End If
                    tmAgfCredit(ilLoop).s13Wks = gAddStr(tmAgfCredit(ilLoop).s13Wks, slPrice)
                    If llDate - lmNowDate <= 7 * tmSpf.iRNoWks Then
                        tmAgfCredit(ilLoop).sProjected = gAddStr(tmAgfCredit(ilLoop).sProjected, slPrice)
                    End If
                End If
                Exit For
            End If
            DoEvents
        Next ilLoop
    End If
End Sub

Private Function mCheckRecLength() As Integer
    If Not gRecLengthOk("Adf.Btr", Len(tmAdf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Agf.Btr", Len(tmAgf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Sbf.Btr", Len(tmSbf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Chf.Btr", Len(tmChf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Clf.Btr", Len(tmClf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Cff.Btr", Len(tmCff)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Urf.Btr", Len(tmUrf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    If Not gRecLengthOk("Spf.Btr", Len(tmSpf)) Then
        mCheckRecLength = False
        Exit Function
    End If
    mCheckRecLength = True
End Function

