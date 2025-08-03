Attribute VB_Name = "RPTRTFSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptrtfsubs.bas on Wed 6/17/09 @ 12:56
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptRtfSubs.Bas
'
' Release: 5.3
'
' Description:
'   This file contains the Report selection screen code
Option Explicit
Option Compare Text

Dim tmScr As SCR
Dim hmScr As Integer
Dim imScrRecLen As Integer
' Dan M modify rich text script to be double spaced 9-01-09
Const RTFCOMPARE = "{\rtf1\ansi\ansicpg1252\"
Const TEXTSPACING = 2
Const MAX_TAB_STOPS = 32&
Const EM_SETPARAFORMAT = &H447
Const PFM_LINESPACING = &H100&
Private Type TableInformation
    hODF As Integer
    hSCR As Integer
    hCIF As Integer
    hCSF As Integer
    ODFLength As Integer
    SCRLength As Integer
    CIFLength As Integer
    CSFLength As Integer
End Type
Private Type PARAFORMAT2
    iSize As Integer
    iPad1 As Integer
    lMask As Long
    iNumbering As Integer
    iReserved As Integer
    lStartIndent As Long
    lRightIndent As Long
    lOffset As Long
    iAlignment As Integer
    iTabCount As Integer
    lTabStops(0 To MAX_TAB_STOPS - 1) As Long
    lSpaceBefore As Long          ' Vertical spacing before para
    lSpaceAfter As Long           ' Vertical spacing after para
    lLineSpacing As Long          ' Line spacing depending on Rule
    iStyle As Integer              ' Style handle
    bLineSpacingRule As Byte       ' Rule for line spacing
    bCRC As Byte                   ' Reserved for CRC for rapid searching
    iShadingWeight As Integer      ' Shading in hundredths of a per cent
    iShadingStyle As Integer       ' Nibble 0: style, 1: cfpat, 2: cbpat
    iNumberingStart As Integer     ' Starting value for numbering
    iNumberingStyle As Integer     ' Alignment, roman/arabic, (), ), .,     etc.
    iNumberingTab As Integer       ' Space between 1st indent and 1st-line text
    iBorderSpace As Integer        ' Space between border and text(twips)
    iBorderWidth As Integer        ' Border pen width (twips)
    iBorders As Integer            ' Byte 0: bits specify which borders; Nibble 2: border style; 3: color                                     index*/
End Type


'
'**************************************************************
'*                                                             *
'*      Procedure Name:gGenRTF                              *
'*                                                             *
'*             Created:8/29/05       By:D. Hosaka              *
'*            Modified:              By:                       *
'*                                                             *
'*         Comments: Create RTF file from script input         *
'*                   for previewing from SCR.btr
'*                                                             *
'***************************************************************
Function gGenRTF(slRptName As String, RtfControl As Control) As Integer
    Dim slSelection As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String
    Dim slDate As String
    Dim slTime As String
    Dim ilRet As Integer
    Dim slRichText As String
        'rich text double spaced for rtfPreview
        slRichText = gModifySingleRtfScript(RtfControl.TextRTF)
        gGenRTF = 0
        hmScr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmScr, "", sgDBPath & "Scr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            ilRet = btrClose(hmScr)
            btrDestroy hmScr
            Exit Function
        End If
        imScrRecLen = Len(tmScr)

        'Start Crystal report engine
        ilRet = PEOpenEngine()
        If ilRet = 0 Then
            MsgBox "Unable to open print engine"
            Exit Function
        End If


        If Not gOpenPrtJob(Trim$(slRptName)) Then
            gGenRTF = -1
            Exit Function
        End If

        gCurrDateTime slDate, slTime, slMonth, slDay, slYear        'get current date & time
        'tmScr.iStrLen = Len(RtfControl) 'rtfRichTextBox1)
       ' tmScr.sScript = Trim$(RtfControl) & Chr$(0) '& Chr$(0)           'Trim$(rtfRichTextBox1) & Chr$(0) & Chr$(0)
       'Dan 9-01-09 double spaced alternative
        tmScr.sScript = slRichText
        tmScr.iGenDate(0) = igNowDate(0)
        tmScr.iGenDate(1) = igNowDate(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmScr.lGenTime = lgNowTime
        tmScr.lCode = 0
        ilRet = btrInsert(hmScr, tmScr, Len(tmScr), INDEXKEY0)

        slSelection = "{SCR_Script_Check.scrGenDate} = Date(" & slYear & "," & slMonth & "," & slDay & ")"
        slSelection = slSelection & " And Round({SCR_Script_Check.scrGenTime}) = " & Trim$(str$(CLng(gTimeToCurrency(slTime, False))))

        If Not gSetSelection(slSelection) Then
            gGenRTF = -1
            Exit Function
        End If
    '   Dan M allow rollback report to 8.5 no longer needed 9-01-09
        Report.Show vbModal
'        If Not bgRollback Then
'            Report.Show vbModal
'        Else
'            RollBackReport.Show vbModal
'        End If

        btrClose (hmScr)
        btrDestroy hmScr
        gClearScr               'Date: modified call from gClearRtf to gClearScr    FYM
        PECloseEngine

        gGenRTF = 0

    Exit Function
End Function
Public Sub gRTFModifyFromTable()
Dim ilRet As Integer
Dim slModifiedScript As String
Dim tlOdf As ODF
Dim tlCif As CIF
Dim tlCSF As CSF
Dim tlScr As SCR
Dim tlODFSearch As ODFKEY2
Dim tlTableFacts As TableInformation
tlTableFacts.ODFLength = Len(tlOdf)
tlTableFacts.CIFLength = Len(tlCif)
tlTableFacts.CSFLength = Len(tlCSF)
tlTableFacts.SCRLength = Len(tlScr)
tlODFSearch.lGenTime = lgNowTime
tlODFSearch.iGenDate(0) = igNowDate(0)
tlODFSearch.iGenDate(1) = igNowDate(1)
On Error GoTo ERRORBOX
tlTableFacts = mOpenTables(tlTableFacts)
With tlTableFacts
    ilRet = btrGetEqual(.hODF, tlOdf, .ODFLength, tlODFSearch, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
    Do While mOuterLoopFoundODF(ilRet, tlOdf, tlODFSearch)
        If tlOdf.lCifCode > 0 Then
            ilRet = btrGetEqual(.hCIF, tlCif, .CIFLength, tlOdf.lCifCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If tlCif.lCsfCode > 0 And ilRet = BTRV_ERR_NONE Then
                ilRet = btrGetEqual(.hCSF, tlCSF, .CSFLength, tlCif.lCsfCode, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slModifiedScript = gModifySingleRtfScript(tlCSF.sComment)
                    tlOdf.lFt1CefCode = mWriteModifiedToScr(slModifiedScript, tlTableFacts)   ' use empty ft1CefCode to link to Scr
                    If tlOdf.lFt1CefCode > 0 Then
                         ilRet = btrUpdate(.hODF, tlOdf, .ODFLength)
                    End If
                End If
            End If
        End If
        ilRet = btrGetNext(.hODF, tlOdf, .ODFLength, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(.hSCR)
    btrDestroy .hSCR
    ilRet = btrClose(.hODF)
    btrDestroy .hODF
    ilRet = btrClose(.hCIF)
    btrDestroy .hCIF
    ilRet = btrClose(.hCSF)
    btrDestroy .hCSF
End With
Exit Sub
ERRORBOX:
    gMsgBox Err.Source & ":" & Err.Description, vbOKOnly + vbExclamation, "Error"
End Sub
Private Function mOpenTables(tlTI As TableInformation) As TableInformation
Dim ilRet As Integer
    tlTI.hODF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(tlTI.hODF, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy tlTI.hODF
        Exit Function
    End If
    tlTI.hCIF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(tlTI.hCIF, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy tlTI.hODF
        btrDestroy tlTI.hCIF
        Exit Function
    End If
    tlTI.hSCR = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(tlTI.hSCR, "", sgDBPath & "scr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy tlTI.hODF
        btrDestroy tlTI.hCIF
        btrDestroy tlTI.hSCR
        Exit Function
    End If
    tlTI.hCSF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(tlTI.hCSF, "", sgDBPath & "CSF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy tlTI.hODF
        btrDestroy tlTI.hCIF
        btrDestroy tlTI.hSCR
        btrDestroy tlTI.hCSF
        Exit Function
    End If
    mOpenTables = tlTI
End Function
Private Function mOuterLoopFoundODF(ilRet As Integer, tlOdf As ODF, tlODFSearch As ODFKEY2) As Boolean
    With tlOdf
        If (ilRet = BTRV_ERR_NONE) And (.iGenDate(0) = tlODFSearch.iGenDate(0)) And (.iGenDate(1) = tlODFSearch.iGenDate(1)) And (.lGenTime = tlODFSearch.lGenTime) Then
            mOuterLoopFoundODF = True
        End If
    End With
End Function
Private Function mWriteModifiedToScr(slScript As String, tlTI As TableInformation) As Long
Dim tlScr As SCR
Dim ilRet As Integer
    tlScr.sScript = slScript
    tlScr.lCode = 0
    tlScr.lGenTime = lgNowTime
    tlScr.iGenDate(0) = igNowDate(0)
    tlScr.iGenDate(1) = igNowDate(1)
    ilRet = btrInsert(tlTI.hSCR, tlScr, tlTI.SCRLength, 0)        'watch last number
    If ilRet = BTRV_ERR_NONE Then
        mWriteModifiedToScr = tlScr.lCode
    Else
        Err.Raise 4702, "mWriteModifiedToScr", "Could not insert to table"
    End If
End Function
Public Function gModifySingleRtfScript(slScript As String) As String
    Traffic.RichTextBox1.TextRTF = slScript
    ' Dan M corrupted text(rtf saved as text) crashed Crystal.  Compare text, not textrtf
    If InStr(1, Traffic.RichTextBox1.Text, RTFCOMPARE, vbTextCompare) = 0 Then
        Traffic.RichTextBox1.SelStart = 0   'double spacing on current paragraph...select all
        Traffic.RichTextBox1.SelLength = Len(Traffic.RichTextBox1.Text)
        mSelLineSpacing Traffic.RichTextBox1, TEXTSPACING
        gModifySingleRtfScript = Trim(Traffic.RichTextBox1.TextRTF)
    Else
        gModifySingleRtfScript = "This text cannot be formatted properly as it is too similar to RTF script commands."
    End If
End Function
Private Sub mSelLineSpacing(rtbTarget As RichTextBox, llSpacingRule As Long, Optional llLineSpacing As Long = 20)
    ' SpacingRule
    ' Type of line spacing. To use this member, set the PFM_SPACEAFTER flag in the dwMask member. This member can be one of the following values.
    ' 0 - Single spacing. The dyLineSpacing member is ignored.
    ' 1 - One-and-a-half spacing. The dyLineSpacing member is ignored.
    ' 2 - Double spacing. The dyLineSpacing member is ignored.
    ' 3 - The dyLineSpacing member specifies the spacingfrom one line to the next, in twips. However, if dyLineSpacing specifies a value that is less than single spacing, the control displays single-spaced text.
    ' 4 - The dyLineSpacing member specifies the spacing from one line to the next, in twips. The control uses the exact spacing specified, even if dyLineSpacing specifies a value that is less than single spacing.
    ' 5 - The value of dyLineSpacing / 20 is the spacing, in lines, from one line to the next. Thus, setting dyLineSpacing to 20 produces single-spaced text, 40 is double spaced, 60 is triple spaced, and so on.
    Dim tlPara As PARAFORMAT2
    With tlPara
        .iSize = Len(tlPara)
        .lMask = PFM_LINESPACING
        .bLineSpacingRule = llSpacingRule
        .lLineSpacing = llLineSpacing
    End With
    SendMessage rtbTarget.hwnd, EM_SETPARAFORMAT, ByVal 0&, tlPara
End Sub
Public Sub gClearScr()
'*******************************************************
'*                                                     *
'*      Procedure Name:gClearSCR                       *
'*                                                     *
'*         Created:09/02/09      By:D. Michaelson      *
'*         Modified: copied gClearODR                  *
'*                                                     *
'*         Comments:Clear SCR file by gen date  *
'*                     and time                        *
'*                                                     *
'*******************************************************
    Dim ilRet As Integer
    Dim hlScr As Integer
    Dim tlScrSrchKey1 As SCRKEY1
    Dim tlScr As SCR
    Dim ilScrRecLen As Integer
    hlScr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hlScr, "", sgDBPath & "Scr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlScr)
        btrDestroy hlScr
        Exit Sub
    End If
    ilScrRecLen = Len(tlScr)
    tlScrSrchKey1.iGenDate(0) = igNowDate(0)
    tlScrSrchKey1.iGenDate(1) = igNowDate(1)
    tlScrSrchKey1.lGenTime = lgNowTime
    ilRet = btrGetGreaterOrEqual(hlScr, tlScr, ilScrRecLen, tlScrSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tlScr.iGenDate(0) = igNowDate(0)) And (tlScr.iGenDate(1) = igNowDate(1)) And (tlScr.lGenTime = lgNowTime)
        ilRet = btrDelete(hlScr)
        ilRet = btrGetNext(hlScr, tlScr, ilScrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hlScr)
    btrDestroy hlScr
End Sub
