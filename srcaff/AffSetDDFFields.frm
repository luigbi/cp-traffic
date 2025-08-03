VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSetDDFFields 
   Caption         =   "Set Fields"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   ControlBox      =   0   'False
   Icon            =   "AffSetDDFFields.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmcCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2955
      TabIndex        =   14
      Top             =   3045
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CheckBox ckcTask 
      Caption         =   "Post CP"
      Height          =   195
      Index           =   4
      Left            =   300
      TabIndex        =   9
      Top             =   1935
      Width           =   1800
   End
   Begin VB.CheckBox ckcTask 
      Caption         =   "Agreements"
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   7
      Top             =   1605
      Width           =   1935
   End
   Begin VB.CheckBox ckcTask 
      Caption         =   "Comments"
      Height          =   195
      Index           =   2
      Left            =   300
      TabIndex        =   5
      Top             =   1275
      Width           =   1185
   End
   Begin VB.CheckBox ckcTask 
      Caption         =   "Contacts"
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   3
      Top             =   945
      Width           =   1185
   End
   Begin VB.CheckBox ckcTask 
      Caption         =   "Stations"
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   585
      Width           =   1185
   End
   Begin VB.CommandButton cmcOK 
      Caption         =   "Process"
      Enabled         =   0   'False
      Height          =   375
      Left            =   885
      TabIndex        =   13
      Top             =   3045
      Width           =   1890
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   3240
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   210
      Left            =   900
      TabIndex        =   12
      Top             =   2670
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lacAstCount 
      Height          =   210
      Left            =   105
      TabIndex        =   15
      Top             =   2655
      Width           =   750
   End
   Begin VB.Label lacStatus 
      Height          =   195
      Index           =   5
      Left            =   1845
      TabIndex        =   11
      Top             =   2310
      Width           =   1575
   End
   Begin VB.Label lacStatus 
      Height          =   195
      Index           =   4
      Left            =   2400
      TabIndex        =   10
      Top             =   1935
      Width           =   1575
   End
   Begin VB.Label lacStatus 
      Height          =   195
      Index           =   3
      Left            =   2400
      TabIndex        =   8
      Top             =   1605
      Width           =   1575
   End
   Begin VB.Label lacStatus 
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Top             =   1275
      Width           =   1575
   End
   Begin VB.Label lacStatus 
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   4
      Top             =   945
      Width           =   1575
   End
   Begin VB.Label lacStatus 
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   585
      Width           =   1575
   End
   Begin VB.Label lacSettingFields 
      Alignment       =   2  'Center
      Caption         =   "Setting Fields Status"
      Height          =   270
      Left            =   300
      TabIndex        =   0
      Top             =   150
      Width           =   4530
   End
End
Attribute VB_Name = "frmSetDDFFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmSetDDFFields
'*
'*  Created October,2005 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc. 2005
'*
'******************************************************
Option Explicit
Option Compare Text

'Private tmAstInfo As ASTINFO

Private imMajorVersion As Integer

Private hmAst As Integer
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO

Private lmTotalRecords As Long
Private lmProcessedRecords As Long
Private lmPercent As Long

Private rst_Shtt As ADODB.Recordset
Private rst_att As ADODB.Recordset
Private rst_Cptt As ADODB.Recordset
Private rst_Webl As ADODB.Recordset
Private rst_Ast As ADODB.Recordset
Private rst_Aet As ADODB.Recordset
Private rst_Lst As ADODB.Recordset
Private rst_cct As ADODB.Recordset
Private rst_artt As ADODB.Recordset
Private rst_mgt As ADODB.Recordset
Private rst_emt As ADODB.Recordset
Private rst_rsf As ADODB.Recordset
Private rst_site As ADODB.Recordset
Private rst_mnt As ADODB.Recordset
Private rst_DAT As ADODB.Recordset
Private rst_Ulf As ADODB.Recordset





Private Sub cmcCancel_Click()
    Unload frmSetDDFFields
End Sub

Private Sub cmcOK_Click()
    If sgSetFieldCallSource = "M" Then
        If cmcOK.Caption = "Process" Then
            tmcStart.Enabled = True
            Exit Sub
        End If
    End If
    Unload frmSetDDFFields
End Sub

Private Sub Form_Load()
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
    imMajorVersion = App.Major
    If sgSetFieldCallSource = "S" Then
        For ilLoop = 0 To 4 Step 1
            If (ilLoop = 4) And (imMajorVersion >= 7) Then
                ckcTask(ilLoop).Value = vbUnchecked
            Else
                ckcTask(ilLoop).Value = vbChecked
            End If
        Next ilLoop
        tmcStart.Enabled = True
    Else
        ckcTask(0).Enabled = True
        ckcTask(1).Enabled = False
        ckcTask(2).Enabled = False
        ckcTask(3).Caption = "Program Times"
        ckcTask(4).Caption = "Compliant Counts"
        cmcCancel.Visible = True
        cmcOK.Enabled = True
    End If
    
    ilRet = gOpenMKDFile(hmAst, "Ast.Mkd")
    gCenterStdAlone frmSetDDFFields
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    ilRet = gCloseMKDFile(hmAst, "Ast.Mkd")
    Erase tmAstInfo
    Erase tmCPDat
    rst_Shtt.Close
    rst_att.Close
    rst_Cptt.Close
    rst_Webl.Close
    rst_Ast.Close
    rst_Aet.Close
    rst_Lst.Close
    rst_cct.Close
    rst_artt.Close
    rst_mgt.Close
    rst_emt.Close
    rst_rsf.Close
    rst_site.Close
    rst_mnt.Close
    rst_DAT.Close
    rst_Ulf.Close
    On Error GoTo 0
    Set frmSetDDFFields = Nothing

End Sub


Private Function mSetStations() As Integer
    Dim slAgreementExist As String
    Dim slCommentExist As String
    Dim slHistStartDate As String
    Dim slOnAir As String
    Dim slStationType As String
    Dim llMulticastGroupID As Long
    Dim llCityCode As Long
    Dim llOnCityCode As Long
    Dim llCityLicCode As Long

    On Error GoTo ErrHand
    mSetStations = False
    
    SQLQuery = "SELECT Count(shttCode) FROM SHTT"
    Set rst_Shtt = gSQLSelectCall(SQLQuery)
    If Not rst_Shtt.EOF Then
        lmTotalRecords = rst_Shtt(0).Value
        SQLQuery = "SELECT * FROM SHTT"
        Set rst_Shtt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Shtt.EOF
            'Create City and County names
            llCityCode = mGetCityCode(rst_Shtt!shttCity)
            llOnCityCode = mGetCityCode(rst_Shtt!shttOnCity)
            llCityLicCode = mGetCityCode(rst_Shtt!shttCityLic)
            'Test if Agreements exist
            SQLQuery = "SELECT attCode, attOnAir, attAgreeStart FROM att"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " attShfCode = " & rst_Shtt!shttCode & ")"
            SQLQuery = SQLQuery & " ORDER BY attAgreeStart, attOnAir"
            Set rst_att = gSQLSelectCall(SQLQuery)
            If rst_att.EOF Then
                slAgreementExist = "N"
                slHistStartDate = "1/1/1970"
            Else
                slAgreementExist = "Y"
                If IsNull(rst_att!attAgreeStart) Then
                    slHistStartDate = rst_att!attOnAir
                Else
                    If (DateValue(gAdjYear(rst_att!attAgreeStart)) = DateValue("1/1/1970")) Or (DateValue(gAdjYear(rst_att!attAgreeStart)) = DateValue("1/1/70")) Then    'Placeholder value to prevent using Nulls/outer joins
                        slHistStartDate = rst_att!attOnAir
                    Else
                        slHistStartDate = rst_att!attAgreeStart
                    End If
                End If
            End If
            SQLQuery = "SELECT attCode, attOnAir, attAgreeStart FROM att"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " attShfCode = " & rst_Shtt!shttCode & ")"
            SQLQuery = SQLQuery & " ORDER BY attOnAir"
            Set rst_att = gSQLSelectCall(SQLQuery)
            If Not rst_att.EOF Then
                If (DateValue(gAdjYear(slHistStartDate)) = DateValue("1/1/1970")) Or (DateValue(gAdjYear(slHistStartDate)) = DateValue("1/1/70")) Then    'Placeholder value to prevent using Nulls/outer joins
                    slHistStartDate = rst_att!attOnAir
                Else
                    If (DateValue(gAdjYear(rst_att!attOnAir)) < DateValue(gAdjYear(slHistStartDate))) Then    'Placeholder value to prevent using Nulls/outer joins
                        slHistStartDate = rst_att!attOnAir
                    End If
                End If
            End If
            'Test if Comments exist
            SQLQuery = "SELECT cctCode FROM cct"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " cctShfCode = " & rst_Shtt!shttCode & ")"
            Set rst_cct = gSQLSelectCall(SQLQuery)
            If rst_cct.EOF Then
                slCommentExist = "N"
            Else
                slCommentExist = "Y"
            End If
            'Multicast Group ID
            SQLQuery = "Select mgtGroupID FROM mgt where mgtShfCode = " & rst_Shtt!shttCode
            Set rst_mgt = gSQLSelectCall(SQLQuery)
            If Not rst_mgt.EOF Then
                If IsNull(rst_mgt!mgtGroupID) = True Then
                    llMulticastGroupID = 0
                Else
                    llMulticastGroupID = rst_mgt!mgtGroupID
                End If
            Else
                llMulticastGroupID = 0
            End If
            
            'Set On Air
            slOnAir = "Y"
            'Set Station Type
            slStationType = "C"
            SQLQuery = "Update shtt Set "
            SQLQuery = SQLQuery & "shttAgreementExist = '" & Trim(slAgreementExist) & "', "
            SQLQuery = SQLQuery & "shttCommentExist = '" & Trim(slCommentExist) & "', "
            SQLQuery = SQLQuery & "shttHistStartDate = '" & Format$(slHistStartDate, sgSQLDateForm) & "', "
            If sgSetFieldCallSource = "S" Then
                SQLQuery = SQLQuery & "shttMultiCastGroupID = " & llMulticastGroupID & ", "
                SQLQuery = SQLQuery & "shttOnAir = '" & Trim(slOnAir) & "', "
                '7/25/11:  Clear field as it was not used anylonger in v5.6 or v5.7
                SQLQuery = SQLQuery & "shttEMail = '" & "" & "', "
                SQLQuery = SQLQuery & "shttStationType = '" & Trim(slStationType) & "', "
                SQLQuery = SQLQuery & "shttCityMntCode = " & llCityCode & ", "
                SQLQuery = SQLQuery & "shttOnCityMntCode = " & llOnCityCode & ", "
                SQLQuery = SQLQuery & "shttCityLicMntCode = " & llCityLicCode & " "
            Else
                SQLQuery = SQLQuery & "shttMultiCastGroupID = " & llMulticastGroupID & " "
            End If
            SQLQuery = SQLQuery & " Where shttCode = " & rst_Shtt!shttCode
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SetDDFFields-mSetStations"
                mSetStations = False
                Exit Function
            End If
            mSetGauge
            rst_Shtt.MoveNext
        Loop
        '11/26/17
        gFileChgdUpdate "shtt.mkd", True
    End If
    mSetStations = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetDDFFields-mSetStations"
End Function
Private Function mSetContacts() As Integer
    Dim ilFound As Integer
    Dim ilTest As Integer
    Dim llUpdateArttCode As Long
    Dim ilPos As Integer
    Dim ilStart As Integer
    Dim ilLoop As Integer
    Dim slEmailAddressArray() As String
    Dim slSentToWeb() As String
    Dim ilSeqNo() As Integer
    
    'Merge emt into Artt
    On Error GoTo ErrHand
    mSetContacts = False
    SQLQuery = "SELECT Count(shttCode) FROM SHTT"
    Set rst_Shtt = gSQLSelectCall(SQLQuery)
    If Not rst_Shtt.EOF Then
        lmTotalRecords = rst_Shtt(0).Value
        SQLQuery = "SELECT * FROM SHTT"
        Set rst_Shtt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Shtt.EOF
            ReDim slEMailFromArtt(0 To 0) As String
            ReDim llArttCode(0 To 0) As Long
            SQLQuery = "SELECT arttCode, arttEMail FROM artt"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " arttShttCode = " & rst_Shtt!shttCode & ")"
            Set rst_artt = gSQLSelectCall(SQLQuery)
            Do While Not rst_artt.EOF
                If Not IsNull(rst_artt!arttEmail) Then
                    If Trim$(rst_artt!arttEmail) <> "" Then
                        llArttCode(UBound(llArttCode)) = rst_artt!arttCode
                        ReDim Preserve llArttCode(0 To UBound(llArttCode) + 1) As Long
                        slEMailFromArtt(UBound(slEMailFromArtt)) = UCase$(Trim$(rst_artt!arttEmail))
                        ReDim Preserve slEMailFromArtt(0 To UBound(slEMailFromArtt) + 1) As String
                    End If
                End If
                rst_artt.MoveNext
            Loop
        
            ReDim slEmailAddressArray(0 To 0) As String
            ReDim slSentToWeb(0 To 0) As String
            ReDim ilSeqNo(0 To 0) As Integer
            If igEmailNeedsConv Then
                ilPos = InStr(1, rst_Shtt!shttWebEmail, ",", vbTextCompare)
                If ilPos > 0 Then
                    ilStart = 1
                    ReDim slEmailAddressArray(0 To 0) As String
                    ReDim slSentToWeb(0 To 0) As String
                    ReDim ilSeqNo(0 To 0) As Integer
                    slEmailAddressArray(0) = Trim$(Mid$(rst_Shtt!shttWebEmail, ilStart, ilPos - 1))
                    ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
                    For ilLoop = ilPos To Len(rst_Shtt!shttWebEmail) - 1 Step 1
                        ilStart = ilPos + 1
                        ilPos = InStr(ilStart, rst_Shtt!shttWebEmail, ",", vbTextCompare)
                        If ilPos > 0 Then
                            slEmailAddressArray(UBound(slEmailAddressArray)) = Trim$(Mid$(rst_Shtt!shttWebEmail, ilStart, ilPos - ilStart))
                            slSentToWeb(UBound(slSentToWeb)) = "Y"
                            ilSeqNo(UBound(ilSeqNo)) = UBound(ilSeqNo) + 1
                            ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
                            ReDim Preserve slSentToWeb(0 To UBound(slSentToWeb) + 1) As String
                            ReDim Preserve ilSeqNo(0 To UBound(ilSeqNo) + 1) As Integer
                            ilLoop = ilPos
                        Else
                            slEmailAddressArray(UBound(slEmailAddressArray)) = Trim$(Mid$(rst_Shtt!shttWebEmail, ilStart, Len(rst_Shtt!shttWebEmail) - (ilStart - 1)))
                            slSentToWeb(UBound(slSentToWeb)) = "Y"
                            ilSeqNo(UBound(ilSeqNo)) = UBound(ilSeqNo) + 1
                            ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
                            ReDim Preserve slSentToWeb(0 To UBound(slSentToWeb) + 1) As String
                            ReDim Preserve ilSeqNo(0 To UBound(ilSeqNo) + 1) As Integer
                            Exit For
                        End If
                    Next ilLoop
                Else
                    If Trim$(rst_Shtt!shttWebEmail) <> "" Then
                        ReDim slEmailAddressArray(0 To 1) As String
                        ReDim slSentToWeb(0 To 1) As String
                        ReDim ilSeqNo(0 To 1) As Integer
                        slEmailAddressArray(0) = Trim$(rst_Shtt!shttWebEmail)
                        slSentToWeb(0) = "Y"
                        ilSeqNo(0) = 1
                    End If
                End If
            Else
                SQLQuery = "SELECT * FROM EMT"
                SQLQuery = SQLQuery + " WHERE ("
                SQLQuery = SQLQuery & " emtShttCode = " & rst_Shtt!shttCode & ")"
                Set rst_emt = gSQLSelectCall(SQLQuery)
                Do While Not rst_emt.EOF
                    slEmailAddressArray(UBound(slEmailAddressArray)) = Trim$(rst_emt!emtEmail)
                    slSentToWeb(UBound(slSentToWeb)) = Trim$(rst_emt!emtSendToWeb)
                    ilSeqNo(UBound(ilSeqNo)) = rst_emt!emtSeqNo
                    ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
                    ReDim Preserve slSentToWeb(0 To UBound(slSentToWeb) + 1) As String
                    ReDim Preserve ilSeqNo(0 To UBound(ilSeqNo) + 1) As Integer
                    rst_emt.MoveNext
                Loop
            End If
            For ilLoop = 0 To UBound(slEmailAddressArray) - 1 Step 1
                ilFound = False
                For ilTest = 0 To UBound(slEMailFromArtt) - 1 Step 1
                    If slEMailFromArtt(ilTest) = UCase$(Trim$(slEmailAddressArray(ilLoop))) Then
                        llUpdateArttCode = llArttCode(ilTest)
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    'Add to Artt
                    SQLQuery = "Insert Into artt ( "
                    SQLQuery = SQLQuery & "arttCode, "
                    SQLQuery = SQLQuery & "arttFirstName, "
                    SQLQuery = SQLQuery & "arttLastName, "
                    SQLQuery = SQLQuery & "arttPhone, "
                    SQLQuery = SQLQuery & "arttFax, "
                    SQLQuery = SQLQuery & "arttEmail, "
                    SQLQuery = SQLQuery & "arttEmailRights, "
                    SQLQuery = SQLQuery & "arttState, "
                    SQLQuery = SQLQuery & "arttUsfCode, "
                    SQLQuery = SQLQuery & "arttAddress1, "
                    SQLQuery = SQLQuery & "arttAddress2, "
                    SQLQuery = SQLQuery & "arttCity, "
                    SQLQuery = SQLQuery & "arttAddressState, "
                    SQLQuery = SQLQuery & "arttZip, "
                    SQLQuery = SQLQuery & "arttCountry, "
                    SQLQuery = SQLQuery & "arttType, "
                    SQLQuery = SQLQuery & "arttTntCode, "
                    SQLQuery = SQLQuery & "arttShttCode, "
                    SQLQuery = SQLQuery & "arttAffContact, "
                    SQLQuery = SQLQuery & "arttISCI2Contact, "
                    SQLQuery = SQLQuery & "arttWebEMail, "
                    SQLQuery = SQLQuery & "arttEMailToWeb, "
                    SQLQuery = SQLQuery & "arttWebEMailRefID, "
                    SQLQuery = SQLQuery & "arttUnused "
                    SQLQuery = SQLQuery & ") "
                    SQLQuery = SQLQuery & "Values ( "
                    SQLQuery = SQLQuery & 0 & ", "
                    SQLQuery = SQLQuery & "'" & "" & "', "
                    SQLQuery = SQLQuery & "'" & "" & "', "
                    SQLQuery = SQLQuery & "'" & "" & "', "
                    SQLQuery = SQLQuery & "'" & "" & "', "
                    SQLQuery = SQLQuery & "'" & gFixQuote(slEmailAddressArray(ilLoop)) & "', "
                    SQLQuery = SQLQuery & "'" & "N" & "', "
                    SQLQuery = SQLQuery & 0 & ", "
                    SQLQuery = SQLQuery & igUstCode & ", "
                    SQLQuery = SQLQuery & "'" & "" & "', "
                    SQLQuery = SQLQuery & "'" & "" & "', "
                    SQLQuery = SQLQuery & "'" & "" & "', "
                    SQLQuery = SQLQuery & "'" & "" & "', "
                    SQLQuery = SQLQuery & "'" & "" & "', "
                    SQLQuery = SQLQuery & "'" & "" & "', "
                    SQLQuery = SQLQuery & "'" & "P" & "', "
                    SQLQuery = SQLQuery & 0 & ", "
                    SQLQuery = SQLQuery & rst_Shtt!shttCode & ", "
                    SQLQuery = SQLQuery & "'" & "0" & "', "
                    SQLQuery = SQLQuery & "'" & "0" & "', "
                    SQLQuery = SQLQuery & "'" & "Y" & "', "
                    SQLQuery = SQLQuery & "'" & slSentToWeb(ilLoop) & "', " 'rst_emt!emtSendToWeb & "', "
                    SQLQuery = SQLQuery & ilSeqNo(ilLoop) & ", "
                    SQLQuery = SQLQuery & "'" & "" & "' "
                    SQLQuery = SQLQuery & ") "
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "SetDDFFields-mSetContacts"
                        mSetContacts = False
                        Exit Function
                    End If
                Else
                    SQLQuery = "Update artt Set "
                    SQLQuery = SQLQuery & "arttWebEMail = '" & "Y" & "', "
                    SQLQuery = SQLQuery & "arttEMailToWeb = '" & slSentToWeb(ilLoop) & "', "    'rst_emt!emtSendToWeb & "' "
                    SQLQuery = SQLQuery & "arttWebEMailRefID = " & ilSeqNo(ilLoop)
                    SQLQuery = SQLQuery & " Where arttCode = " & llUpdateArttCode
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "SetDDFFields-mSetContacts"
                        mSetContacts = False
                        Exit Function
                    End If
                End If
            Next ilLoop
            '7/25/11
            'If igEmailNeedsConv Then
            If Not igEmailNeedsConv Then
                SQLQuery = "DELETE FROM EMT"
                SQLQuery = SQLQuery + " WHERE ("
                SQLQuery = SQLQuery & " emtShttCode = " & rst_Shtt!shttCode & ")"
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/12/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "SetDDFFields-mSetContacts"
                    mSetContacts = False
                    Exit Function
                End If
            End If
            mSetGauge
            rst_Shtt.MoveNext
        Loop
    End If
    mSetContacts = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetDDFFields-mSetContacts"
End Function
Private Function mSetComments() As Integer
    Dim ilLoop As Integer
    On Error GoTo ErrHand
    mSetComments = False
    
'    For ilLoop = 1 To 5 Step 1
'
'        SQLQuery = "Insert Into cst ( "
'        SQLQuery = SQLQuery & "cstCode, "
'        SQLQuery = SQLQuery & "cstName, "
'        SQLQuery = SQLQuery & "cstDefault, "
'        SQLQuery = SQLQuery & "cstSortCode, "
'        SQLQuery = SQLQuery & "cstUnused "
'        SQLQuery = SQLQuery & ") "
'        SQLQuery = SQLQuery & "Values ( "
'        SQLQuery = SQLQuery & ilLoop & ", "
'        Select Case ilLoop
'            Case 1
'                SQLQuery = SQLQuery & "'" & gFixQuote("Call: Outgoing") & "', "
'                SQLQuery = SQLQuery & "'" & gFixQuote("Y") & "', "
'            Case 2
'                SQLQuery = SQLQuery & "'" & gFixQuote("Call: Ingoing") & "', "
'                SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
'            Case 3
'                SQLQuery = SQLQuery & "'" & gFixQuote("E-Mail: Outgoing") & "', "
'                SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
'            Case 4
'                SQLQuery = SQLQuery & "'" & gFixQuote("Mass E-Mail: Outgoing") & "', "
'                SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
'            Case 5
'                SQLQuery = SQLQuery & "'" & gFixQuote("E-Mail: Incoming") & "', "
'                SQLQuery = SQLQuery & "'" & gFixQuote("N") & "', "
'        End Select
'        SQLQuery = SQLQuery & ilLoop & ", "
'        SQLQuery = SQLQuery & "'" & "" & "' "
'        SQLQuery = SQLQuery & ") "
'        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'            GoSub ErrHand:
'        End If
'    Next ilLoop
'Move to frmMain
    
    SQLQuery = "Update cct Set "
    SQLQuery = SQLQuery & "cctCstCode = " & 1
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/12/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "SetDDFFields-mSetComments"
        mSetComments = False
        Exit Function
    End If
    
    mSetComments = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetDDFFields-mSetComments"
End Function
Private Function mSetAgreements() As Integer
    Dim ilRet As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilExportType As Integer
    Dim slExportToWeb As String
    Dim slExportToUnivision As String
    Dim slExportToMarketron As String
    Dim slCDStartTime As String
    Dim slPledgeType As String
    Dim ilSeqNo As Integer
    
    On Error GoTo ErrHand
    mSetAgreements = False
    SQLQuery = "SELECT Count(attCode) FROM ATT"
    Set rst_att = gSQLSelectCall(SQLQuery)
    If Not rst_att.EOF Then
        lmTotalRecords = rst_att(0).Value
        SQLQuery = "SELECT * FROM ATT "
        Set rst_att = gSQLSelectCall(SQLQuery)
        Do While Not rst_att.EOF
            ''If IsNull(rst_att!attStartTime) Then
                slCDStartTime = ""
            ''Else
            ''    slCDStartTime = Format$(rst_att!attStartTime, "hh:mmA/P")
            ''End If
            'Run from the menu item if Version 7
            'ilRet = gDetermineAgreementTimes(rst_att!attshfCode, rst_att!attvefCode, Format$(rst_att!attOnAir, "m/d/yy"), Format$(rst_att!attOffAir, "m/d/yy"), Format$(rst_att!attDropDate, "m/d/yy"), slCDStartTime, slStartTime, slEndTime)
            If imMajorVersion < 7 Then
                ilRet = gDetermineAgreementTimes(rst_att!attshfCode, rst_att!attvefCode, Format$(rst_att!attOnAir, "m/d/yy"), Format$(rst_att!attOffAir, "m/d/yy"), Format$(rst_att!attDropDate, "m/d/yy"), slCDStartTime, slStartTime, slEndTime)
            Else
                slStartTime = "12AM"
                slEndTime = "12AM"
            End If
            If sgSetFieldCallSource = "S" Then
                slExportToWeb = "N"
                If rst_att!attExportType = 1 Then
                    slExportToWeb = "Y"
                End If
                slExportToUnivision = "N"
'                If rst_att!attExportType = 2 Then
'                    slExportToUnivision = "Y"
'                End If
                slExportToMarketron = "N"
                '7701
                If gIsVendorWithAgreement(rst_att!attCode, Vendors.NetworkConnect) Then
                    slExportToMarketron = "Y"
                End If
'                If rst_att!attExportToMarketron = "Y" Then
'                    slExportToMarketron = "Y"
'                End If
                slPledgeType = ""
                ilSeqNo = 0
                SQLQuery = "SELECT * "
                SQLQuery = SQLQuery + " FROM dat"
                SQLQuery = SQLQuery + " WHERE (datatfCode= " & rst_att!attCode & ")"
                Set rst_DAT = gSQLSelectCall(SQLQuery)
                Do While Not rst_DAT.EOF
                    Select Case rst_DAT!datDACode
                        Case 0  'Dayprt
                            slPledgeType = "D"
                        Case 1  'Avail
                            slPledgeType = "A"
                        Case 2  'CD or Tape
                            slPledgeType = "C"
                        Case Else
                            slPledgeType = ""
                    End Select
                    SQLQuery = "UPDATE dat Set "
                    SQLQuery = SQLQuery & " datAirPlayNo = " & 1 & ","
                    SQLQuery = SQLQuery & " datEstimatedTime = " & "'N'"
                    SQLQuery = SQLQuery & " Where datCode = " & rst_DAT!datCode
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "SetDDFFields-mSetAgreements"
                        mSetAgreements = False
                        Exit Function
                    End If

'                    SQLQuery = "Insert Into apt ( "
'                    SQLQuery = SQLQuery & "aptCode, "
'                    SQLQuery = SQLQuery & "aptAttCode, "
'                    SQLQuery = SQLQuery & "aptAirPlayNo, "
'                    SQLQuery = SQLQuery & "aptSeqNo, "
'                    SQLQuery = SQLQuery & "aptBreakoutMo, "
'                    SQLQuery = SQLQuery & "aptBreakoutTu, "
'                    SQLQuery = SQLQuery & "aptBreakoutWe, "
'                    SQLQuery = SQLQuery & "aptBreakoutTh, "
'                    SQLQuery = SQLQuery & "aptBreakoutFr, "
'                    SQLQuery = SQLQuery & "aptBreakoutSa, "
'                    SQLQuery = SQLQuery & "aptBreakoutSu, "
'                    SQLQuery = SQLQuery & "aptStartTime, "
'                    SQLQuery = SQLQuery & "aptOffsetDay, "
'                    SQLQuery = SQLQuery & "aptEstimatedTime, "
'                    SQLQuery = SQLQuery & "aptUnused "
'                    SQLQuery = SQLQuery & ") "
'                    SQLQuery = SQLQuery & "Values ( "
'                    SQLQuery = SQLQuery & 0 & ", "
'                    SQLQuery = SQLQuery & rst_att!attCode & ", "
'                    SQLQuery = SQLQuery & 1 & ", "
'                    SQLQuery = SQLQuery & 1 & ", "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    If slPledgeType = "C" Then
'                        SQLQuery = SQLQuery & "'" & Format$(rst_att!attStartTime, sgSQLTimeForm) & "', "
'                    Else
'                        SQLQuery = SQLQuery & "'" & Format$(slStartTime, sgSQLTimeForm) & "', "
'                    End If
'                    SQLQuery = SQLQuery & 0 & ", "
'                    SQLQuery = SQLQuery & "'" & "N" & "', "
'                    SQLQuery = SQLQuery & "'" & "" & "' "
'                    SQLQuery = SQLQuery & ") "
'
'                    ilSeqNo = ilSeqNo + 1
'                    SQLQuery = SQLQuery & "aptCode, "
'                    SQLQuery = SQLQuery & "aptAttCode, "
'                    SQLQuery = SQLQuery & "aptAirPlayNo, "
'                    SQLQuery = SQLQuery & "aptSeqNo, "
'                    SQLQuery = SQLQuery & "aptPledgeType, "
'                    SQLQuery = SQLQuery & "aptFdStatus, "
'                    SQLQuery = SQLQuery & "aptAirMo, "
'                    SQLQuery = SQLQuery & "aptAirTu, "
'                    SQLQuery = SQLQuery & "aptAirWe, "
'                    SQLQuery = SQLQuery & "aptAirTh, "
'                    SQLQuery = SQLQuery & "aptAirFr, "
'                    SQLQuery = SQLQuery & "aptAirSa, "
'                    SQLQuery = SQLQuery & "aptAirSu, "
'                    SQLQuery = SQLQuery & "aptPledgeStartTime, "
'                    SQLQuery = SQLQuery & "aptOffsetDay, "
'                    SQLQuery = SQLQuery & "aptEStimatedTime, "
'                    SQLQuery = SQLQuery & "aptFeedStartTime, "
'                    SQLQuery = SQLQuery & "aptFeedEndTime, "
'                    SQLQuery = SQLQuery & "aptUnused "
'                    SQLQuery = SQLQuery & ") "
'                    SQLQuery = SQLQuery & "Values ( "
'                    SQLQuery = SQLQuery & 0 & ", "
'                    SQLQuery = SQLQuery & rst_att!attCode & ", "
'                    SQLQuery = SQLQuery & 1 & ", "
'                    SQLQuery = SQLQuery & ilSeqNo & ", "
'                    SQLQuery = SQLQuery & "'" & gFixQuote(slPledgeType) & "', "
'                    SQLQuery = SQLQuery & 0 & ", "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    SQLQuery = SQLQuery & "'" & "Y" & "', "
'                    If slPledgeType = "C" Then
'                        SQLQuery = SQLQuery & "'" & Format$(rst_att!attStartTime, sgSQLTimeForm) & "', "
'                    Else
'                        SQLQuery = SQLQuery & "'" & Format$(slStartTime, sgSQLTimeForm) & "', "
'                    End If
'                    SQLQuery = SQLQuery & -1 & ", "
'                    SQLQuery = SQLQuery & "'" & "N" & "', "
'                    SQLQuery = SQLQuery & "'" & Format$("12:00:00AM", sgSQLTimeForm) & "', "
'                    SQLQuery = SQLQuery & "'" & Format$("12:00:00AM", sgSQLTimeForm) & "', "
'                    SQLQuery = SQLQuery & "'" & "" & "' "
'                    SQLQuery = SQLQuery & ") "
'
'                    'cnn.Execute SQLQuery, rdExecDirect
'                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
'                        GoSub ErrHand:
'                    End If
                    rst_DAT.MoveNext
                Loop
                SQLQuery = "Update att Set "
                SQLQuery = SQLQuery & "attVehProgStartTime = '" & Format$(slStartTime, sgSQLTimeForm) & "', "
                SQLQuery = SQLQuery & "attVehProgEndTime = '" & Format$(slEndTime, sgSQLTimeForm) & "', "
                'SQLQuery = SQLQuery & "attExportType = " & ilExportType & ", "
                SQLQuery = SQLQuery & "attExportToWeb = '" & slExportToWeb & "', "
                SQLQuery = SQLQuery & "attExportToUnivision = '" & slExportToUnivision & "', "
                SQLQuery = SQLQuery & "attExportToMarketron = '" & slExportToMarketron & "', "
                SQLQuery = SQLQuery & "attExportToCBS = '" & "N" & "', "
                SQLQuery = SQLQuery & "attExportToClearCh = '" & "N" & "', "
                SQLQuery = SQLQuery & "attNoAirPlays = " & 1 & ", "
                SQLQuery = SQLQuery & "attDesignVersion = " & 1 & ", "
                SQLQuery = SQLQuery & "attPledgeType = '" & slPledgeType & "'"
                'ttp 5270 change manual to export
                If slExportToMarketron = "Y" Or slExportToWeb = "Y" Or slExportToUnivision = "Y" Then
                      SQLQuery = SQLQuery & " , " & "attExportType = " & 1
               End If
            Else
                SQLQuery = "Update att Set "
                SQLQuery = SQLQuery & "attVehProgStartTime = '" & Format$(slStartTime, sgSQLTimeForm) & "', "
                SQLQuery = SQLQuery & "attVehProgEndTime = '" & Format$(slEndTime, sgSQLTimeForm) & "'"
            End If
            SQLQuery = SQLQuery & " Where attCode = " & rst_att!attCode
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SetDDFFields-mSetAgreements"
                mSetAgreements = False
                Exit Function
            End If
            mSetGauge
            rst_att.MoveNext
        Loop
    End If
    
    mSetAgreements = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetDDFFields-mSetAgreements"
End Function
Private Function mSetPostCPs() As Integer
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilPledgeCompliantCount As Integer
    Dim ilAgyCompliantCount As Integer
    Dim llCpttDate As Long
    Dim ilVefCode As Integer
    Dim llAttCode As Long
    Dim llWeekDate As Long
    Dim slWeek1 As String
    Dim slWeek60 As String
    Dim ilAdfCode As Integer
    Dim ilShtt As Integer
    Dim ilAst As Integer
    Dim slMoDate As String
    Dim ilRet As Integer
    Dim slTPdETime As String
    Dim slPledgeEndTime As String
    Dim ilTechnique As Integer
    Dim llPrevAttCode As Long
    Dim llDat As Long
    Dim ilLoop As Integer
    Dim slFeedStartTime As String
    'Dim llFeedStartTime As Long
    Dim slPledgeStartTime As String
    'Dim llPledgeStartTime As Long
    Dim ilPdDay As Integer
    Dim ilDayOk As Integer
    Dim ilFdDay As Integer
    Dim llDatIndex As Long
    Dim slPledgeDays As String
    Dim ilAirStatus As Integer
    Dim slFeedDate As String
    Dim slPledgeDate As String
    Dim llAstCount As Long
    Dim ilPostingType As Integer
    Dim blAttOk As Boolean
    Dim ilSvAirStatus As Integer
    Dim tlAstInfo As ASTINFO
    Dim tlDatPledgeInfo As DATPLEDGEINFO
    Dim tlLST As LST
    
    On Error GoTo ErrHand
    mSetPostCPs = False
    llPrevAttCode = -1
    slWeek1 = Format$(gNow(), sgShowDateForm)   'sgSQLDateForm)
    slWeek1 = gObtainNextSunday(gObtainNextMonday(gObtainNextSunday(slWeek1)))
    slWeek60 = DateAdd("d", -(60 * 7) + 1, slWeek1)
    
    ilRet = gPopAvailNames()
    
    If Not gPopCopy(slWeek60, "Post CPs") Then
        Exit Function
    End If
    
    SQLQuery = "SELECT Count(cpttCode) FROM CPTT WHERE cpttStartDate >= '" & Format(slWeek60, sgSQLDateForm) & "'"
    Set rst_Cptt = gSQLSelectCall(SQLQuery)
    If Not rst_Cptt.EOF Then
        lmTotalRecords = rst_Cptt(0).Value
        SQLQuery = "SELECT * FROM CPTT WHERE cpttStartDate >= '" & Format(slWeek60, sgSQLDateForm) & "'"
        SQLQuery = SQLQuery & " ORDER BY cpttAtfCode"
        Set rst_Cptt = gSQLSelectCall(SQLQuery)
        Do While Not rst_Cptt.EOF
            'Set counts
            blAttOk = True
            llCpttDate = gDateValue(rst_Cptt!CpttStartDate)
            ilVefCode = rst_Cptt!cpttvefcode
            llAttCode = rst_Cptt!cpttatfCode
            If llAttCode <> llPrevAttCode Then
                SQLQuery = "SELECT * FROM att"
                SQLQuery = SQLQuery + " WHERE ("
                SQLQuery = SQLQuery & " attCode = " & llAttCode & ")"
                Set rst_att = gSQLSelectCall(SQLQuery)
                If Not rst_att.EOF Then
                    ilPostingType = rst_att!attPostingType
                Else
                    blAttOk = False
                End If
            End If
            If blAttOk Then
                ilSchdCount = 0
                ilAiredCount = 0
                ilPledgeCompliantCount = 0
                ilAgyCompliantCount = 0
                If ilPostingType = 0 Then  'Receipt
                    ilSchdCount = 0
                    ilAiredCount = 0
                ElseIf ilPostingType = 1 Then  'Count
                    ilSchdCount = rst_Cptt!cpttNoSpotsGen
                    ilAiredCount = rst_Cptt!cpttNoSpotsAired
                    ilPledgeCompliantCount = 0
                    ilAgyCompliantCount = 0
                Else
                    'Spots by date and spots by advertiser
                    llWeekDate = gDateValue(gObtainPrevMonday(gAdjYear(Format$(llCpttDate, "m/d/yy"))))
                    ilTechnique = 1
                    If ilTechnique = 1 Then
                        If llAttCode <> llPrevAttCode Then
                            ReDim tlDat(0 To 30) As DATRST
                            llDat = 0
                            SQLQuery = "SELECT * "
                            SQLQuery = SQLQuery + " FROM dat"
                            SQLQuery = SQLQuery + " WHERE (datatfCode= " & llAttCode & ")"
                            Set rst_DAT = gSQLSelectCall(SQLQuery)
                            Do While Not rst_DAT.EOF
                                gCreateUDTForDat rst_DAT, tlDat(llDat)
                                llDat = llDat + 1
                                If llDat = UBound(tlDat) Then
                                    ReDim Preserve tlDat(0 To UBound(tlDat) + 30) As DATRST
                                End If
                                rst_DAT.MoveNext
                            Loop
                            ReDim Preserve tlDat(0 To llDat) As DATRST
                            llPrevAttCode = llAttCode
                        End If
                        llAstCount = 0
                        ReDim llAstUpdate(0 To 0) As Long
                        SQLQuery = "SELECT * FROM ast"
                        SQLQuery = SQLQuery + " WHERE ("
                        SQLQuery = SQLQuery + " astFeedDate >= '" & Format(llWeekDate, sgSQLDateForm) & "'"
                        SQLQuery = SQLQuery + " AND astFeedDate <= '" & Format(llWeekDate + 6, sgSQLDateForm) & "'"
                        SQLQuery = SQLQuery & " AND astatfCode = " & llAttCode & ")"
                        Set rst_Ast = gSQLSelectCall(SQLQuery)
                        Do While Not rst_Ast.EOF
                            tlDatPledgeInfo.lAttCode = rst_Ast!astAtfCode
                            tlDatPledgeInfo.lDatCode = rst_Ast!astDatCode
                            tlDatPledgeInfo.iVefCode = rst_Ast!astVefCode
                            tlDatPledgeInfo.sFeedDate = Format(rst_Ast!astFeedDate, "m/d/yy")
                            tlDatPledgeInfo.sFeedTime = Format(rst_Ast!astFeedTime, "hh:mm:ssam/pm")
                            ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
                            If tgStatusTypes(gGetAirStatus(tlDatPledgeInfo.iPledgeStatus)).iPledged <> 2 Then
                                llAstCount = llAstCount + 1
                                'Find dat match
                                slFeedStartTime = Format$(rst_Ast!astFeedTime, sgShowTimeWSecForm)
                                ''llFeedStartTime = gTimeToLong(slFeedStartTime, False)
                                'slPledgeStartTime = Format$(rst_Ast!astPledgeStartTime, sgShowTimeWSecForm)
                                ''llPledgeStartTime = gTimeToLong(slPledgeStartTime, False)
                                'If Not IsNull(rst_Ast!astPledgeEndTime) Then
                                '    slPledgeEndTime = Format$(rst_Ast!astPledgeEndTime, sgShowTimeWSecForm)
                                'Else
                                '    slPledgeEndTime = slPledgeStartTime
                                'End If
                                slPledgeStartTime = Format$(tlDatPledgeInfo.sPledgeStartTime, sgShowTimeWSecForm)
                                If Not IsNull(tlDatPledgeInfo.sPledgeEndTime) Then
                                    slPledgeEndTime = Format$(tlDatPledgeInfo.sPledgeEndTime, sgShowTimeWSecForm)
                                Else
                                    slPledgeEndTime = tlDatPledgeInfo.sPledgeStartTime
                                End If
                                slFeedDate = rst_Ast!astFeedDate
                                'slPledgeDate = rst_Ast!astPledgeDate
                                slPledgeDate = tlDatPledgeInfo.sPledgeDate
                                llDatIndex = gMatchAstAndDat(slFeedStartTime, slFeedDate, slPledgeStartTime, slPledgeDate, tlDat())
    '                            For llDat = LBound(tlDat) To UBound(tlDat) - 1 Step 1
    '                                If (gTimeToLong(tlDat(llDat).sFdStTime, False) = llFeedStartTime) And (gTimeToLong(tlDat(llDat).sPdStTime, False) = llPledgeStartTime) Then
    '                                    ilDayOk = False
    '                                    ilFdDay = Weekday(rst_Ast!astFeedDate, vbMonday)
    '                                    Select Case ilFdDay
    '                                        Case 1  'Monday
    '                                            If tlDat(llDat).iFdMon Then
    '                                                ilDayOk = True
    '                                            End If
    '                                        Case 2  'Tuesday
    '                                            If tlDat(llDat).iFdTue Then
    '                                                ilDayOk = True
    '                                            End If
    '                                        Case 3  'Wednesady
    '                                            If tlDat(llDat).iFdWed Then
    '                                                ilDayOk = True
    '                                            End If
    '                                        Case 4  'Thursday
    '                                            If tlDat(llDat).iFdThu Then
    '                                                ilDayOk = True
    '                                            End If
    '                                        Case 5  'Friday
    '                                            If tlDat(llDat).iFdFri Then
    '                                                ilDayOk = True
    '                                            End If
    '                                        Case 6  'Saturday
    '                                            If tlDat(llDat).iFdSat Then
    '                                                ilDayOk = True
    '                                            End If
    '                                        Case 7  'Sunday
    '                                            If tlDat(llDat).iFdSun Then
    '                                                ilDayOk = True
    '                                            End If
    '                                    End Select
    '                                    If ilDayOk Then
    '                                        ilDayOk = False
    '                                        ilPdDay = Weekday(rst_Ast!astPledgeDate, vbMonday)
    '                                        Select Case ilPdDay
    '                                            Case 1  'Monday
    '                                                If tlDat(llDat).iPdMon Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                            Case 2  'Tuesday
    '                                                If tlDat(llDat).iPdTue Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                            Case 3  'Wednesday
    '                                                If tlDat(llDat).iPdWed Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                            Case 4  'Thursday
    '                                                If tlDat(llDat).iPdThu Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                            Case 5  'Friday
    '                                                If tlDat(llDat).iPdFri Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                            Case 6  'Saturday
    '                                                If tlDat(llDat).iPdSat Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                            Case 7  'Sunday
    '                                                If tlDat(llDat).iPdSun Then
    '                                                    ilDayOk = True
    '                                                End If
    '                                        End Select
    '                                    End If
    '                                    If ilDayOk Then
    '                                        llDatIndex = llDat
    '                                        slPledgeDays = String(7, "N")
    '                                        For ilPdDay = 1 To 7
    '                                            Select Case ilPdDay
    '                                                Case 1  'Monday
    '                                                    If tlDat(llDat).iPdMon Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                                Case 2  'Tuesday
    '                                                    If tlDat(llDat).iPdTue Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                                Case 3  'Wednesday
    '                                                    If tlDat(llDat).iPdWed Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                                Case 4  'Thursday
    '                                                    If tlDat(llDat).iPdThu Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                                Case 5  'Friday
    '                                                    If tlDat(llDat).iPdFri Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                                Case 6  'Saturday
    '                                                    If tlDat(llDat).iPdSat Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                                Case 7  'Sunday
    '                                                    If tlDat(llDat).iPdSun Then
    '                                                        Mid(slPledgeDays, ilPdDay, 1) = "Y"
    '                                                    End If
    '                                            End Select
    '                                        Next ilPdDay
    '                                        Exit For
    '                                    End If
    '                                End If
    '                            Next llDat
                                slPledgeDays = String(7, "N")
                                If llDatIndex <> -1 Then
                                    llDat = llDatIndex
                                    For ilPdDay = 1 To 7
                                        Select Case ilPdDay
                                            Case 1  'Monday
                                                If tlDat(llDat).iPdMon Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                            Case 2  'Tuesday
                                                If tlDat(llDat).iPdTue Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                            Case 3  'Wednesday
                                                If tlDat(llDat).iPdWed Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                            Case 4  'Thursday
                                                If tlDat(llDat).iPdThu Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                            Case 5  'Friday
                                                If tlDat(llDat).iPdFri Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                            Case 6  'Saturday
                                                If tlDat(llDat).iPdSat Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                            Case 7  'Sunday
                                                If tlDat(llDat).iPdSun Then
                                                    Mid(slPledgeDays, ilPdDay, 1) = "Y"
                                                End If
                                        End Select
                                    Next ilPdDay
                                Else
                                    Mid(slPledgeDays, Weekday(slPledgeDate, vbMonday), 1) = "Y"
                                End If
                                If (gTimeToLong(slPledgeStartTime, False) = gTimeToLong(slPledgeEndTime, True)) Then
                                    If llDatIndex >= 0 Then
                                        slTPdETime = Format$(gLongToTime(gTimeToLong(Format$(slPledgeStartTime, "h:mm:ssam/pm"), False) + gTimeToLong(tlDat(llDatIndex).sFdEdTime, False) - gTimeToLong(tlDat(llDatIndex).sFdStTime, False)), sgShowTimeWSecForm)
                                    Else
                                        'Add 5 minutes to start time
                                        slTPdETime = Format$(gLongToTime(gTimeToLong(Format$(slPledgeStartTime, "h:mm:ssam/pm"), False) + 300), sgShowTimeWSecForm)
                                    End If
                                Else
                                    slTPdETime = Format$(slPledgeEndTime, "h:mm:ssam/pm")
                                End If
                                ilAirStatus = rst_Ast!astStatus
                                If (gGetAirStatus(ilAirStatus) = 6) Or (gGetAirStatus(ilAirStatus) = 7) Then
                                    ilAirStatus = 1
                                    llAstUpdate(UBound(llAstUpdate)) = rst_Ast!astCode
                                    ReDim Preserve llAstUpdate(0 To UBound(llAstUpdate) + 1) As Long
                                End If
                                ''gIncSpotCounts rst_Ast!astPledgeStatus, ilAirStatus, rst_Ast!astCPStatus, slPledgeDays, Format$(slPledgeDate, "m/d/yy"), Format$(rst_Ast!astAirDate, "m/d/yy"), Format$(slPledgeStartTime, "h:mm:ssAM/PM"), Format$(slTPdETime, "h:mm:ssAM/PM"), Format$(rst_Ast!astAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                                'gIncSpotCounts rst_Ast!astPledgeStatus, ilAirStatus, rst_Ast!astCPStatus, slPledgeDays, Format$(slPledgeDate, "m/d/yy"), Format$(rst_Ast!astAirDate, "m/d/yy"), Format$(slPledgeStartTime, "h:mm:ssAM/PM"), Format$(slTPdETime, "h:mm:ssAM/PM"), Format$(rst_Ast!astAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                                tlAstInfo.lCode = rst_Ast!astCode
                                tlAstInfo.lAttCode = rst_Ast!astAtfCode
                                'tlAstInfo.iPledgeStatus = rst_Ast!astPledgeStatus
                                tlAstInfo.iPledgeStatus = tlDatPledgeInfo.iPledgeStatus
                                tlAstInfo.iStatus = ilAirStatus
                                tlAstInfo.iCPStatus = rst_Ast!astCPStatus
                                tlAstInfo.sTruePledgeDays = slPledgeDays
                                tlAstInfo.sPledgeDate = slPledgeDate
                                tlAstInfo.sAirDate = rst_Ast!astAirDate
                                tlAstInfo.sPledgeStartTime = slPledgeStartTime
                                tlAstInfo.sTruePledgeEndTime = slTPdETime
                                tlAstInfo.sAirTime = Format(rst_Ast!astAirTime, sgShowTimeWSecForm)
                                tlAstInfo.lLstCode = rst_Ast!astLsfCode
                                tlAstInfo.lSdfCode = rst_Ast!astSdfCode
                                SQLQuery = "SELECT *"
                                SQLQuery = SQLQuery & " FROM LST"
                                SQLQuery = SQLQuery & " WHERE lstCode =" & Str(tlAstInfo.lLstCode)
                                Set rst_Lst = gSQLSelectCall(SQLQuery)
                                If Not rst_Lst.EOF Then
                                    gCreateUDTforLST rst_Lst, tlLST
                                    tlAstInfo.lLstBkoutLstCode = tlLST.lBkoutLstCode
                                    tlAstInfo.sLstStartDate = tlLST.sStartDate
                                    tlAstInfo.sLstEndDate = tlLST.sEndDate
                                    tlAstInfo.iLstSpotsWk = tlLST.iSpotsWk
                                    tlAstInfo.iLstMon = tlLST.iMon
                                    tlAstInfo.iLstTue = tlLST.iTue
                                    tlAstInfo.iLstWed = tlLST.iWed
                                    tlAstInfo.iLstThu = tlLST.iThu
                                    tlAstInfo.iLstFri = tlLST.iFri
                                    tlAstInfo.iLstSat = tlLST.iSat
                                    tlAstInfo.iLstSun = tlLST.iSun
                                    tlAstInfo.iLineNo = tlLST.iLineNo
                                    tlAstInfo.iSpotType = tlLST.iSpotType
                                    tlAstInfo.sSplitNet = tlLST.sSplitNetwork
                                    tlAstInfo.iAgfCode = tlLST.iAgfCode
                                    tlAstInfo.sLstLnStartTime = tlLST.sLnStartTime
                                    tlAstInfo.sLstLnEndTime = tlLST.sLnEndTime
                                    gIncSpotCounts tlAstInfo, ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount
                                End If
                            End If
                            rst_Ast.MoveNext
                        Loop
                        If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0 Then
                            lacAstCount.Caption = llAstCount
                            DoEvents
                        End If
                        For ilAst = 0 To UBound(llAstUpdate) - 1 Step 1
                            SQLQuery = "UPDATE ast SET astStatus = 1 WHERE astCode = " & llAstUpdate(ilAst)
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/12/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "SetDDFFields-mSetPostCPs"
                                mSetPostCPs = False
                                Exit Function
                            End If
                        Next ilAst
                    ElseIf ilTechnique = 2 Then
                        ReDim tgCPPosting(0 To 1) As CPPOSTING
                        tgCPPosting(0).lCpttCode = rst_Cptt!cpttCode
                        tgCPPosting(0).iStatus = rst_Cptt!cpttStatus
                        tgCPPosting(0).iPostingStatus = rst_Cptt!cpttPostingStatus
                        tgCPPosting(0).lAttCode = rst_Cptt!cpttatfCode
                        tgCPPosting(0).iAttTimeType = 0 'Not used
                        tgCPPosting(0).iVefCode = rst_Cptt!cpttvefcode  'imVefCode
                        tgCPPosting(0).iShttCode = rst_Cptt!cpttshfcode
                        ilShtt = gBinarySearchStationInfoByCode(tgCPPosting(0).iShttCode)
                        If ilShtt <> -1 Then
                            tgCPPosting(0).sZone = tgStationInfoByCode(ilShtt).sZone
                        Else
                            tgCPPosting(0).sZone = ""
                        End If
                        slMoDate = Format$(llWeekDate, "m/d/yy")
                        tgCPPosting(0).sDate = Format$(slMoDate, sgShowDateForm)
                        tgCPPosting(0).sAstStatus = rst_Cptt!cpttAstStatus
                        igTimes = 1 'By Week
                        ilAdfCode = -1
                        'Dan M 9/26/13  6442
                        ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, True, False, True, False)
                       ' ilRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), ilAdfCode, False, False, True, False)
                        ReDim llAstUpdate(0 To 0) As Long
                        For ilAst = LBound(tmAstInfo) To UBound(tmAstInfo) - 1 Step 1
                            ilAirStatus = tmAstInfo(ilAst).iStatus
                            If (ilAirStatus = 6) Or (ilAirStatus = 7) Then
                                ilAirStatus = 1
                                llAstUpdate(UBound(llAstUpdate)) = rst_Ast!astCode
                                ReDim Preserve llAstUpdate(0 To UBound(llAstUpdate) + 1) As Long
                            End If
                            ''gIncSpotCounts tmAstInfo(ilAst).iPledgeStatus, ilAirStatus, tmAstInfo(ilAst).iCPStatus, tmAstInfo(ilAst).sTruePledgeDays, Format$(tmAstInfo(ilAst).sPledgeDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sAirDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sPledgeStartTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sTruePledgeEndTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                            'gIncSpotCounts tmAstInfo(ilAst).iPledgeStatus, ilAirStatus, tmAstInfo(ilAst).iCPStatus, tmAstInfo(ilAst).sTruePledgeDays, Format$(tmAstInfo(ilAst).sPledgeDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sAirDate, "m/d/yy"), Format$(tmAstInfo(ilAst).sPledgeStartTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sTruePledgeEndTime, "h:mm:ssAM/PM"), Format$(tmAstInfo(ilAst).sAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                            ilSvAirStatus = tmAstInfo(ilAst).iStatus
                            tmAstInfo(ilAst).iStatus = ilAirStatus
                            gIncSpotCounts tmAstInfo(ilAst), ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount
                            tmAstInfo(ilAst).iStatus = ilSvAirStatus
                        Next ilAst
                        For ilAst = 0 To UBound(llAstUpdate) - 1 Step 1
                            SQLQuery = "UPDATE ast SET astStatus = 1 WHERE astCode = " & llAstUpdate(ilAst)
                            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                                '6/12/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "SetDDFFields-mSetPostCPs"
                                mSetPostCPs = False
                                Exit Function
                            End If
                        Next ilAst
                    End If
                    SQLQuery = "Update cptt Set "
                    SQLQuery = SQLQuery & "cpttNoSpotsGen = " & ilSchdCount & ", "
                    SQLQuery = SQLQuery & "cpttNoSpotsAired = " & ilAiredCount & ", "
                    SQLQuery = SQLQuery & "cpttNoCompliant = " & ilPledgeCompliantCount & ", "
                    SQLQuery = SQLQuery & "cpttAgyCompliant = " & ilAgyCompliantCount & " "
                    SQLQuery = SQLQuery & " Where cpttCode = " & rst_Cptt!cpttCode
                    'cnn.Execute SQLQuery, rdExecDirect
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/12/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "SetDDFFields-mSetPostCPs"
                        mSetPostCPs = False
                        Exit Function
                    End If
                End If
            End If
            mSetGauge
            rst_Cptt.MoveNext
        Loop
    End If
    
    mSetPostCPs = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetDDFFields-mSetPostCPs"
End Function
Private Sub tmcStart_Timer()
    Dim ilTask As Integer
    Dim ilRet As Integer
    Dim ilOk As Integer
    
    tmcStart.Enabled = False
    plcGauge.Visible = True
    lmPercent = 0
    ilOk = True
    For ilTask = 0 To 4 Step 1
        plcGauge.Value = 0
        If ckcTask(ilTask).Value = vbChecked Then
            lacStatus(ilTask).Caption = "In Progress"
            lmProcessedRecords = 0
            Select Case ilTask
                Case 0
                    gLogMsg "Set Stations: Start", "SetDDFFields.Txt", False
                    ilRet = mSetStations()
                    If ilRet Then
                        gLogMsg "Set Stations: Completed", "SetDDFFields.Txt", False
                    Else
                        gLogMsg "Set Stations: Stopped", "SetDDFFields.Txt", False
                    End If
                Case 1
                    gLogMsg "Set Contacts: Start", "SetDDFFields.Txt", False
                    ilRet = mSetContacts()
                    If ilRet Then
                        gLogMsg "Set Contacts: Completed", "SetDDFFields.Txt", False
                    Else
                        gLogMsg "Set Contacts: Stopped", "SetDDFFields.Txt", False
                    End If
                Case 2
                    gLogMsg "Set Comments: Start", "SetDDFFields.Txt", False
                    ilRet = mSetComments()
                    If ilRet Then
                        gLogMsg "Set Comments: Completed", "SetDDFFields.Txt", False
                    Else
                        gLogMsg "Set Comments: Stopped", "SetDDFFields.Txt", False
                    End If
                Case 3
                    gLogMsg "Set Agreements: Start", "SetDDFFields.Txt", False
                    ilRet = mSetAgreements()
                    If ilRet Then
                        gLogMsg "Set Agreements: Completed", "SetDDFFields.Txt", False
                    Else
                        gLogMsg "Set Agreements: Stopped", "SetDDFFields.Txt", False
                    End If
                Case 4
                    gLogMsg "Set Post CP: Start", "SetDDFFields.Txt", False
                    ilRet = mSetPostCPs()
                    If ilRet Then
                        gLogMsg "Set Post CP: Completed", "SetDDFFields.Txt", False
                    Else
                        gLogMsg "Set Post CP: Stopped", "SetDDFFields.Txt", False
                    End If
            End Select
            If ilRet Then
                lacStatus(ilTask).Caption = "Completed"
            Else
                ilOk = False
                lacStatus(ilTask).Caption = "Failed"
            End If
        Else
            lacStatus(ilTask).Caption = "Bypassed"
        End If
    Next ilTask
    plcGauge.Visible = False
    If sgSetFieldCallSource = "S" Then
        If ilOk Then
            mSetUserLog
            SQLQuery = "Update att Set "
            SQLQuery = SQLQuery & "attExportType = " & 1 & " "
            SQLQuery = SQLQuery & " Where attExportType > " & 1
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SetDDFFields-tmcStart_Timer"
                Exit Sub
            End If
            SQLQuery = "Update site Set "
            SQLQuery = SQLQuery & "siteDDF092710 = '" & "Y" & "' "
            SQLQuery = SQLQuery & " Where siteCode = " & 1
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SetDDFFields-tmcStart_Timer"
                Exit Sub
            End If
        Else
            SQLQuery = "Update site Set "
            SQLQuery = SQLQuery & "siteDDF092710 = '" & "N" & "' "
            SQLQuery = SQLQuery & " Where siteCode = " & 1
            'cnn.Execute SQLQuery, rdExecDirect
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/12/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "SetDDFFields-tmcStart_Timer"
                Exit Sub
            End If
        End If
    End If
    cmcOK.Caption = "Done"
    cmcOK.Enabled = True
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetDDFFields-tmcStart"
End Sub

Private Sub mSetGauge()
    lmProcessedRecords = lmProcessedRecords + 1
    lmPercent = (lmProcessedRecords * CSng(100)) / lmTotalRecords
    If lmPercent >= 100 Then
        If lmProcessedRecords + 1 < lmTotalRecords Then
            lmPercent = 99
        Else
            lmPercent = 100
        End If
    End If
    If plcGauge.Value <> lmPercent Then
        plcGauge.Value = lmPercent
        DoEvents
    End If
End Sub

Private Function mGetCityCode(slInName As String) As Long
    Dim slName As String
    Dim ilCity As Integer
    Dim llCode As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    slName = UCase$(Trim$(slInName))
    mGetCityCode = 0
    If Trim$(slName) = "" Then
        Exit Function
    End If
    For ilCity = 0 To UBound(tgCityInfo) - 1 Step 1
        If slName = UCase$(Trim$(tgCityInfo(ilCity).sName)) Then
            mGetCityCode = tgCityInfo(ilCity).lCode
            Exit Function
        End If
    Next ilCity
    'Add Name
    Do
        SQLQuery = "SELECT MAX(mntCode) from mnt"
        Set rst_mnt = gSQLSelectCall(SQLQuery)
        If IsNull(rst_mnt(0).Value) Then
            llCode = 1
        Else
            If Not rst_mnt.EOF Then
                llCode = rst_mnt(0).Value + 1
            Else
                llCode = 1
            End If
        End If
        ilRet = 0
        SQLQuery = "Insert Into mnt ( "
        SQLQuery = SQLQuery & "mntCode, "
        SQLQuery = SQLQuery & "mntType, "
        SQLQuery = SQLQuery & "mntName, "
        SQLQuery = SQLQuery & "mntState, "
        SQLQuery = SQLQuery & "mntUnused "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "Values ( "
        SQLQuery = SQLQuery & llCode & ", "
        SQLQuery = SQLQuery & "'" & "C" & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(slInName) & "', "
        SQLQuery = SQLQuery & "'" & "A" & "', "
        SQLQuery = SQLQuery & "'" & "" & "' "
        SQLQuery = SQLQuery & ") "
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            If Not gHandleError4994("AffErrorLog.txt", "SetDDFFileds-mGetCityCode") Then
                mGetCityCode = False
                Exit Function
            End If
            ilRet = 1
        End If
    Loop While ilRet <> 0
    mGetCityCode = llCode
    tgCityInfo(UBound(tgCityInfo)).lCode = llCode
    tgCityInfo(UBound(tgCityInfo)).sName = slInName
    tgCityInfo(UBound(tgCityInfo)).sState = "A"
    ReDim Preserve tgCityInfo(0 To UBound(tgCityInfo) + 1) As MNTINFO
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetDDFFields-mGetCityCode"
End Function

Private Function mGetCountyCode(slInName As String) As Long
    Dim slName As String
    Dim ilCounty As Integer
    Dim llCode As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    slName = UCase$(Trim$(slInName))
    mGetCountyCode = 0
    If Trim$(slName) = "" Then
        Exit Function
    End If
    For ilCounty = 0 To UBound(tgCountyInfo) - 1 Step 1
        If slName = UCase$(Trim$(tgCountyInfo(ilCounty).sName)) Then
            mGetCountyCode = tgCountyInfo(ilCounty).lCode
            Exit Function
        End If
    Next ilCounty
    'Add Name
    Do
        SQLQuery = "SELECT MAX(mntCode) from mnt"
        Set rst_mnt = gSQLSelectCall(SQLQuery)
        If IsNull(rst_mnt(0).Value) Then
            llCode = 1
        Else
            If Not rst_mnt.EOF Then
                llCode = rst_mnt(0).Value + 1
            Else
                llCode = 1
            End If
        End If
        ilRet = 0
        SQLQuery = "Insert Into mnt ( "
        SQLQuery = SQLQuery & "mntCode, "
        SQLQuery = SQLQuery & "mntType, "
        SQLQuery = SQLQuery & "mntName, "
        SQLQuery = SQLQuery & "mntState, "
        SQLQuery = SQLQuery & "mntUnused "
        SQLQuery = SQLQuery & ") "
        SQLQuery = SQLQuery & "Values ( "
        SQLQuery = SQLQuery & llCode & ", "
        SQLQuery = SQLQuery & "'" & "Y" & "', "
        SQLQuery = SQLQuery & "'" & gFixQuote(slInName) & "', "
        SQLQuery = SQLQuery & "'" & "A" & "', "
        SQLQuery = SQLQuery & "'" & "" & "' "
        SQLQuery = SQLQuery & ") "
        'cnn.Execute SQLQuery, rdExecDirect
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            If Not gHandleError4994("AffErrorLog.txt", "SetDDFFileds-mGetCountyCode") Then
                mGetCountyCode = False
                Exit Function
            End If
            ilRet = 1
        End If
    Loop While ilRet <> 0
    mGetCountyCode = llCode
    tgCountyInfo(UBound(tgCountyInfo)).lCode = llCode
    tgCountyInfo(UBound(tgCountyInfo)).sName = slInName
    tgCountyInfo(UBound(tgCountyInfo)).sState = "A"
    ReDim Preserve tgCountyInfo(0 To UBound(tgCountyInfo) + 1) As MNTINFO
    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetDDFFields-mGetCountyCode"
End Function


Private Sub mSetUserLog()
    On Error GoTo ErrHand
    If sgSetFieldCallSource = "S" Then
        SQLQuery = "Update ULF_User_Log Set "
        SQLQuery = SQLQuery & "ulfTrafJobNo = " & -1 & ", "
        SQLQuery = SQLQuery & "ulfAffTaskNo = " & -1
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/12/16: Replaced GoSub
            'GoSub ErrHand:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "SetDDFFields-mSetUserLog"
            Exit Sub
        End If
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmSetDDFFields-mSetUserLog"
End Sub
