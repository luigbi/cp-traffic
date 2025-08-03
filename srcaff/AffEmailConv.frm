VERSION 5.00
Begin VB.Form frmEmailConv 
   Caption         =   "E-Mail Conversion"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5925
   Begin VB.Timer tmcDelay 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4920
      Top             =   3240
   End
   Begin VB.TextBox edcMsg 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin VB.CommandButton BTN_OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmEmailConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub SetMessage(MsgType As Integer, Msg As String)

    Dim ilRet As Integer

    edcMsg.ForeColor = RGB(0, 0, 0)
    BTN_OK.Visible = False
    BTN_OK.Caption = "OK"
    
    'Show Black Text & Button
    If MsgType = 1 Then
        edcMsg.ForeColor = RGB(0, 0, 0)
        BTN_OK.Visible = True
    End If
    
    'Show Red Text & Button
    If MsgType = 2 Then
        edcMsg.ForeColor = RGB(255, 0, 0)
        BTN_OK.Visible = True
    End If
    
    edcMsg = Msg
End Sub

Private Sub BTN_OK_Click()
    Unload Me
End Sub


Public Function mFillEmtFile() As Integer

    'D.S. 11/20/08
    'Create an array of email addresses if needed and provide a basic sanity check
    'Pump all of the station and agreement emails into the EMT file.
    
    Dim ilRet As Integer
    Dim temp_rst As ADODB.Recordset
    Dim att_rst As ADODB.Recordset
    Dim MaxSeq_rst As ADODB.Recordset
    Dim ilPos As Integer
    Dim ilStart As Integer
    Dim ilLoop As Integer
    Dim slEmailAddressArray() As String
    Dim ilNumEmailAddr As Integer
    Dim ilIdx As Integer
    Dim ilSeqNum As Integer
    Dim ilInsertOK As Integer
    Dim ilRowsEffected As Integer
    Dim llStaEmailCount As Long
    Dim llAgreementEmailCount As Long
    Dim llWebEmailCount As Long
    
    On Error GoTo ErrHand
    mFillEmtFile = False
    
   
    llAgreementEmailCount = 0
    llStaEmailCount = 0
    llWebEmailCount = 0
    
    DoEvents
    frmEmailConv.SetMessage 3, "   Upgrading to New E-mail Program..." & vbCrLf & "   This May Take Several Minutes." & vbCrLf & "   *** Do Not Stop Your Computer. ***" & vbCrLf

    DoEvents

    '********** Start Station Emails **********
    SQLQuery = "SELECT shttcode, shttwebemail, shttCallLetters from shtt where shttwebemail <> ''"
    Set temp_rst = gSQLSelectCall(SQLQuery)

    While Not temp_rst.EOF
        'Check to see if the string has multiple email addresses
        ilPos = InStr(1, temp_rst!shttWebEmail, ",", vbTextCompare)
        If ilPos > 0 Then
            ilStart = 1
            ReDim slEmailAddressArray(0 To 0) As String
            slEmailAddressArray(0) = Trim$(Mid$(temp_rst!shttWebEmail, ilStart, ilPos - 1))
            ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
            For ilLoop = ilPos To Len(temp_rst!shttWebEmail) - 1 Step 1
                ilStart = ilPos + 1
                ilPos = InStr(ilStart, temp_rst!shttWebEmail, ",", vbTextCompare)
                If ilPos > 0 Then
                    slEmailAddressArray(UBound(slEmailAddressArray)) = Trim$(Mid$(temp_rst!shttWebEmail, ilStart, ilPos - ilStart))
                    ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
                    ilLoop = ilPos
                Else
                    slEmailAddressArray(UBound(slEmailAddressArray)) = Trim$(Mid$(temp_rst!shttWebEmail, ilStart, Len(temp_rst!shttWebEmail) - (ilStart - 1)))
                    ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
                    Exit For
                End If
            Next ilLoop
        Else
            ReDim slEmailAddressArray(0 To 1) As String
            slEmailAddressArray(0) = Trim$(temp_rst!shttWebEmail)
        End If

        ilNumEmailAddr = UBound(slEmailAddressArray)
        ilSeqNum = 1

        For ilIdx = 0 To ilNumEmailAddr - 1 Step 1

            ilInsertOK = True
            If Not gTestForSingleValidEmailAddress(slEmailAddressArray(ilIdx)) Then
                mFillEmtFile = False
                'gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
                'gLogMsg Trim$(temp_rst!shttCallLetters) & "  contianed an invalid email address: " & slEmailAddressArray(ilIdx) & " that was not saved.", "WebEmailLog.Txt", False
                ilInsertOK = False
            End If

            If ilInsertOK Then
                DoEvents
                llStaEmailCount = llStaEmailCount + 1
                frmEmailConv.SetMessage 0, "   Upgrading to New E-mail Program..." & vbCrLf & "   This May Take Several Minutes." & vbCrLf & "   *** Do Not Stop Your Computer. *** " & vbCrLf & vbCrLf & "   Station Emails " & llStaEmailCount

                SQLQuery = "Insert Into emt ( "
                SQLQuery = SQLQuery & "emtCode, "
                SQLQuery = SQLQuery & "emtShttCode, "
                SQLQuery = SQLQuery & "emtAttCode, "
                SQLQuery = SQLQuery & "emtSeqNo, "
                SQLQuery = SQLQuery & "emtEMail, "
                SQLQuery = SQLQuery & "emtSendToWeb "
                SQLQuery = SQLQuery & ") "
                SQLQuery = SQLQuery & "Values ( "
                SQLQuery = SQLQuery & 0 & ", "
                SQLQuery = SQLQuery & temp_rst!shttCode & ", "
                SQLQuery = SQLQuery & 0 & ", "
                SQLQuery = SQLQuery & ilSeqNum & ", "
                SQLQuery = SQLQuery & "'" & gFixQuote(slEmailAddressArray(ilIdx)) & "', "
                SQLQuery = SQLQuery & "'" & "Y" & "'"
                SQLQuery = SQLQuery & ") "

                If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "EmailConv-mFillEmtFile"
                    mFillEmtFile = False
                    'Close the open record sets - don't error out if the handle was never opened
                    If Not temp_rst Is Nothing Then
                        temp_rst.Close
                    End If
                    If Not att_rst Is Nothing Then
                        att_rst.Close
                    End If
                    If Not MaxSeq_rst Is Nothing Then
                        MaxSeq_rst.Close
                    End If
                    If Not EmailExists_rst Is Nothing Then
                        EmailExists_rst.Close
                    End If
                    Exit Function
                End If
                ilSeqNum = ilSeqNum + 1

            End If
        Next ilIdx
        temp_rst.MoveNext
    Wend
    gLogMsg "Station Emails Found: " & llStaEmailCount, "WebEmailInsertLog.Txt", False

    '********** End Station Emails - Start Agreement Emails **********

    SQLQuery = "SELECT shttcode, shttwebemail, shttCallLetters from shtt"
    Set temp_rst = gSQLSelectCall(SQLQuery)

    While Not temp_rst.EOF
        DoEvents
        SQLQuery = "SELECT attCode, attExportType, attWebEmail from att where attShfCode = " & temp_rst!shttCode & " AND attWebEmail <> ''"
        Set att_rst = gSQLSelectCall(SQLQuery)
        If Not att_rst.EOF Then
            'Check to see if the string has multiple email addresses
            ilPos = InStr(1, att_rst!attWebEmail, ",", vbTextCompare)
            If ilPos > 0 Then
                ilStart = 1
                ReDim slEmailAddressArray(0 To 0) As String
                slEmailAddressArray(0) = Trim$(Mid$(att_rst!attWebEmail, ilStart, ilPos - 1))
                ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
                For ilLoop = ilPos To Len(att_rst!attWebEmail) - 1 Step 1
                    ilStart = ilPos + 1
                    ilPos = InStr(ilStart, att_rst!attWebEmail, ",", vbTextCompare)
                    If ilPos > 0 Then
                        slEmailAddressArray(UBound(slEmailAddressArray)) = Trim$(Mid$(att_rst!attWebEmail, ilStart, ilPos - ilStart))
                        ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
                        ilLoop = ilPos
                    Else
                        slEmailAddressArray(UBound(slEmailAddressArray)) = Trim$(Mid$(att_rst!attWebEmail, ilStart, Len(temp_rst!shttWebEmail) - (ilStart - 1)))
                        ReDim Preserve slEmailAddressArray(0 To UBound(slEmailAddressArray) + 1)
                        Exit For
                    End If
                Next ilLoop
            Else
                ReDim slEmailAddressArray(0 To 1) As String
                slEmailAddressArray(0) = Trim$(att_rst!attWebEmail)
            End If

            ilNumEmailAddr = UBound(slEmailAddressArray)
            ilSeqNum = 1

            For ilIdx = 0 To ilNumEmailAddr - 1 Step 1

                ilInsertOK = True
                If Not gTestForSingleValidEmailAddress(slEmailAddressArray(ilIdx)) Then
                    mFillEmtFile = False
                    'gLogMsg sgErrorMsg, "WebEmailLog.Txt", False
                    'gLogMsg Trim$(temp_rst!shttCallLetters) & "  contianed an invalid email address: " & slEmailAddressArray(ilIdx) & " that was not saved.", "WebEmailLog.Txt", False
                    ilInsertOK = False
                End If

                SQLQuery = "SELECT emtEmail from EMT where emtShttCode = " & temp_rst!shttCode & " And emtEmail = " & "'" & Trim(gFixQuote(slEmailAddressArray(ilIdx))) & "'"
                Set EmailExists_rst = gSQLSelectCall(SQLQuery)
                If EmailExists_rst.EOF Then
                    SQLQuery = "SELECT Max(emtSeqNo) from EMT where emtShttCode = " & temp_rst!shttCode
                    Set MaxSeq_rst = gSQLSelectCall(SQLQuery)
                    If Not MaxSeq_rst.EOF Then
                        If IsNull(MaxSeq_rst(0).Value) Then
                            ilSeqNum = 1
                        Else
                            ilSeqNum = MaxSeq_rst(0).Value + 1
                        End If

                    End If

                    If ilInsertOK Then
                        llAgreementEmailCount = llAgreementEmailCount + 1
                        frmEmailConv.SetMessage 0, "   Upgrading to New E-mail Program..." & vbCrLf & "   This May Take Several Minutes." & vbCrLf & "   *** Do Not Stop Your Computer. *** " & vbCrLf & vbCrLf & "   Station Emails " & llStaEmailCount & vbCrLf & "   Affiliate Emails " & llAgreementEmailCount

                        SQLQuery = "Insert Into emt ( "
                        SQLQuery = SQLQuery & "emtCode, "
                        SQLQuery = SQLQuery & "emtShttCode, "
                        SQLQuery = SQLQuery & "emtAttCode, "
                        SQLQuery = SQLQuery & "emtSeqNo, "
                        SQLQuery = SQLQuery & "emtEMail, "
                        SQLQuery = SQLQuery & "emtSendToWeb "
                        SQLQuery = SQLQuery & ") "
                        SQLQuery = SQLQuery & "Values ( "
                        SQLQuery = SQLQuery & 0 & ", "
                        SQLQuery = SQLQuery & temp_rst!shttCode & ", "
                        SQLQuery = SQLQuery & att_rst!attCode & ", "
                        SQLQuery = SQLQuery & ilSeqNum & ", "
                        SQLQuery = SQLQuery & "'" & gFixQuote(slEmailAddressArray(ilIdx)) & "', "
                        SQLQuery = SQLQuery & "'" & "Y" & "'"
                        SQLQuery = SQLQuery & ") "

                        If gSQLWaitNoMsgBox(SQLQuery, True) <> 0 Then
                                '6/10/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "EmailConv-mFillEmtFile"
                                mFillEmtFile = False
                                'Close the open record sets - don't error out if the handle was never opened
                                If Not temp_rst Is Nothing Then
                                    temp_rst.Close
                                End If
                                If Not att_rst Is Nothing Then
                                    att_rst.Close
                                End If
                                If Not MaxSeq_rst Is Nothing Then
                                    MaxSeq_rst.Close
                                End If
                                If Not EmailExists_rst Is Nothing Then
                                    EmailExists_rst.Close
                                End If
                                Exit Function
                        End If
                    End If
                End If
            Next ilIdx
        End If
        temp_rst.MoveNext
    Wend
    gLogMsg "Agreement Emails Found: " & llAgreementEmailCount, "WebEmailInsertLog.Txt", False
    DoEvents

    '********** End Agreement Emails Start Web Emails **********

    'Don't try to send to the web if they are not a web client
    If gUsingWeb Then
        'Clear out any existing emails
        SQLQuery = "Delete from WebEmt"
        ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
        If ilRowsEffected = -1 Then
            gLogMsg "Delete Failed: " & SQLQuery, "WebEmailInsertLog.Txt", False
        End If

        'Insert into the WebEmt table with every record from the affiliate's EMT table
        SQLQuery = "SELECT * from EMT order by EmtShttCode"
        Set temp_rst = gSQLSelectCall(SQLQuery)
        While Not temp_rst.EOF
            DoEvents

            llWebEmailCount = llWebEmailCount + 1
            frmEmailConv.SetMessage 0, "   Upgrading to New E-mail Program..." & vbCrLf & "   This May Take Several Minutes." & vbCrLf & "   *** Do Not Stop Your Computer. *** " & vbCrLf & vbCrLf & "   Station Emails " & llStaEmailCount & vbCrLf & "   Affiliate Emails " & llAgreementEmailCount & vbCrLf & "   Web Emails " & llWebEmailCount

            If gIsUsingNovelty Then
            
                SQLQuery = "usp_AddStationUser "
                SQLQuery = SQLQuery & temp_rst!emtCode & ", "
                SQLQuery = SQLQuery & "'" & Trim$(gGetCallLettersByShttCode(temp_rst!emtShttCode)) & "', "
                SQLQuery = SQLQuery & temp_rst!emtShttCode & ", "
                SQLQuery = SQLQuery & temp_rst!emtAttCode & " , "
                SQLQuery = SQLQuery & temp_rst!emtSeqNo & ", "
                SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(temp_rst!emtEmail)) & " ', "
                SQLQuery = SQLQuery & "'" & " " & "',"
                SQLQuery = SQLQuery & "' ',"
                SQLQuery = SQLQuery & "' ',"
                SQLQuery = SQLQuery & "' '"
            Else

                SQLQuery = "Insert Into WebEmt ( "
                SQLQuery = SQLQuery & "Code, "
                SQLQuery = SQLQuery & "CallLetters, "
                SQLQuery = SQLQuery & "ShttCode, "
                SQLQuery = SQLQuery & "AttCode, "
                SQLQuery = SQLQuery & "SeqNo, "
                SQLQuery = SQLQuery & "EMail, "
                SQLQuery = SQLQuery & "Status, "
                SQLQuery = SQLQuery & "DateModified "
                SQLQuery = SQLQuery & ") "
                SQLQuery = SQLQuery & "Values ( "
                SQLQuery = SQLQuery & temp_rst!emtCode & ", "
                SQLQuery = SQLQuery & "'" & Trim$(gGetCallLettersByShttCode(temp_rst!emtShttCode)) & "', "
                SQLQuery = SQLQuery & temp_rst!emtShttCode & ", "
                SQLQuery = SQLQuery & temp_rst!emtAttCode & ", "
                SQLQuery = SQLQuery & temp_rst!emtSeqNo & ", "
                SQLQuery = SQLQuery & "'" & gFixQuote(Trim$(temp_rst!emtEmail)) & "', "
                SQLQuery = SQLQuery & "'" & " " & "',"
                SQLQuery = SQLQuery & "'" & Format(Now, "ddddd ttttt") & "' "
                SQLQuery = SQLQuery & ") "
            End If
            
            ilRowsEffected = gExecWebSQLWithRowsEffected(SQLQuery)
            If ilRowsEffected = -1 Then
                gLogMsg "Insert Failed: " & SQLQuery, "WebEmailInsertLog.Txt", False
                llWebEmailCount = llWebEmailCount - 1
            End If

            temp_rst.MoveNext
        Wend
    End If
    gLogMsg "   Web Emails Found: " & llWebEmailCount, "WebEmailInsertLog.Txt", False
    If llWebEmailCount = llAgreementEmailCount + llStaEmailCount Then
        frmEmailConv.SetMessage 2, "   Upgrading to New E-mail Program..." & vbCrLf & "   This May Take Several Minutes." & vbCrLf & "   *** Do Not Stop Your Computer. *** " & vbCrLf & vbCrLf & "   Station Emails " & llStaEmailCount & vbCrLf & "   Affiliate Emails " & llAgreementEmailCount & vbCrLf & "   Web Emails " & llWebEmailCount & vbCrLf & "   Upgrade was Successful"
        gLogMsg "Upgrade was Successful", "WebEmailInsertLog.Txt", False
    Else
        gLogMsg "Upgrade Failed, Counts Don't Match.", "WebEmailInsertLog.Txt", False
        frmEmailConv.SetMessage 1, "   Upgrading to New E-mail Program..." & vbCrLf & "   This May Take Several Minutes." & vbCrLf & "   *** Do Not Stop Your Computer. *** " & vbCrLf & vbCrLf & "   Station Emails " & llStaEmailCount & vbCrLf & "   Affiliate Emails " & llAgreementEmailCount & vbCrLf & "   Web Emails " & llWebEmailCount & vbCrLf & "   Upgrade Failed - Counts Don't Match"
    End If
    mFillEmtFile = True

    'Close the open record sets - don't error out if the handle was never opened

    'Even this is causing some errors.  I'm taking it out, it only gets ran once anyway.
'    If Not temp_rst Is Nothing Then
'        temp_rst.Close
'    End If
'    If Not att_rst Is Nothing Then
'        att_rst.Close
'    End If
'    If Not MaxSeq_rst Is Nothing Then
'        MaxSeq_rst.Close
'    End If
'    If Not EmailExists_rst Is Nothing Then
'        EmailExists_rst.Close
'    End If

    gLogMsg "   Total Station & Agreement Emails Found: " & llStaEmailCount + llAgreementEmailCount, "WebEmailInsertLog.Txt", False
    gLogMsg "   Total Emails Sent To Web: " & llWebEmailCount, "WebEmailInsertLog.Txt", False
    'Unload frmEmailConv

Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmEmailConv-mFillEmtFile"
    'Close the open record sets - don't error out if the handle was never opened
    If Not temp_rst Is Nothing Then
        temp_rst.Close
    End If
    If Not att_rst Is Nothing Then
        att_rst.Close
    End If
    If Not MaxSeq_rst Is Nothing Then
        MaxSeq_rst.Close
    End If
    If Not EmailExists_rst Is Nothing Then
        EmailExists_rst.Close
    End If
End Function

Private Sub Form_Initialize()
    gCenterForm frmEmailConv
End Sub

Private Sub Form_Load()

    Dim ilRet As Integer
    
    tmcDelay.Enabled = True
    
End Sub

Private Sub tmcDelay_Timer()

    Dim ilRet As Integer
    
    tmcDelay.Enabled = False
    ilRet = mFillEmtFile()

End Sub
