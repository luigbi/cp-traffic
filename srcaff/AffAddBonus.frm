VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmAddBonus 
   Caption         =   "Add Bonus"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "AffAddBonus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   9615
   Begin VB.TextBox txtSpotLength 
      Height          =   360
      Left            =   1515
      TabIndex        =   5
      Top             =   1410
      Width           =   1320
   End
   Begin VB.ListBox lbcContract 
      Height          =   3375
      ItemData        =   "AffAddBonus.frx":08CA
      Left            =   6585
      List            =   "AffAddBonus.frx":08CC
      TabIndex        =   9
      Top             =   450
      Width           =   2820
   End
   Begin VB.ListBox lbcAdvertiser 
      Height          =   3375
      ItemData        =   "AffAddBonus.frx":08CE
      Left            =   3300
      List            =   "AffAddBonus.frx":08D0
      TabIndex        =   7
      Top             =   450
      Width           =   2820
   End
   Begin VB.TextBox txtTime 
      Height          =   360
      Left            =   1530
      TabIndex        =   3
      Top             =   915
      Width           =   1320
   End
   Begin VB.TextBox txtDate 
      Height          =   360
      Left            =   1530
      TabIndex        =   1
      Top             =   450
      Width           =   1320
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   1965
      Top             =   3990
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4785
      FormDesignWidth =   9615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Bonus"
      Height          =   375
      Left            =   5820
      TabIndex        =   10
      Top             =   4290
      Width           =   1890
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7860
      TabIndex        =   11
      Top             =   4290
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Spot Length"
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   1455
      Width           =   1620
   End
   Begin VB.Label Label4 
      Caption         =   "Bonus Time"
      Height          =   255
      Left            =   105
      TabIndex        =   2
      Top             =   960
      Width           =   1620
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Contracts"
      Height          =   255
      Left            =   7035
      TabIndex        =   8
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label lacTitle1 
      Alignment       =   2  'Center
      Caption         =   "Advertisers"
      Height          =   255
      Left            =   3270
      TabIndex        =   6
      Top             =   135
      Width           =   2850
   End
   Begin VB.Label Label1 
      Caption         =   "Bonus Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   495
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddBonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmContact - allows for selection of station/vehicle/advertiser for contact information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

Private imAdfCode As Integer
Private smDate As String     'Bonus Date
Private smTime As String     'Bonus Time



Private Sub cmdAdd_Click()
    Dim sAirDate As String
    Dim sAirTime As String
    Dim lCntrNo As Long
    Dim lLstCode As Long
    Dim sZone As String
    Dim iLoop As Integer
    Dim iSet As Integer
    Dim ilLen As Integer
    
    If txtDate.Text = "" Then
        Beep
        gMsgBox "Date must be specified.", vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If
    If gIsDate(txtDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        txtDate.SetFocus
        Exit Sub
    Else
        sAirDate = Format(txtDate.Text, sgShowDateForm)
    End If
    If lbcContract.ListIndex < 0 Then
        Beep
        gMsgBox "Contract Must be Specified.", vbCritical
        Exit Sub
    End If
    lCntrNo = CLng(lbcContract.List(lbcContract.ListIndex))
    If Trim$(txtTime.Text) = "" Then
        Beep
        gMsgBox "Time must be specified.", vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If
    sAirTime = txtTime.Text
    If gIsTime(sAirTime) = False Then   'Time not valid.
        Beep
        gMsgBox "Please enter a valid Time (hh:mm:ss am 0r pm).", vbCritical
        txtTime.SetFocus
        Exit Sub
    Else
        sAirTime = Format$(gConvertTime(sAirTime), "hh:mm:ss")
    End If
    If txtSpotLength.Text = "" Then
        Beep
        gMsgBox "Spot Length must be specified.", vbOKOnly
        txtSpotLength.SetFocus
        Exit Sub
    End If
    ilLen = Val(txtSpotLength.Text)
    'SQLQuery = "SELECT shtt.shttTimeZone FROM shtt WHERE (shttCode = " & tgCPPosting(0).iShttCode & ")"
    'Set rst = gSQLSelectCall(SQLQuery)
    'If rst.EOF = True Then
    '    sZone = ""
    'Else
    '    sZone = Trim$(rst!shttTimeZone)
    'End If
    sZone = ""
    For iLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        If tgCPPosting(0).iVefCode = tgVehicleInfo(iLoop).iCode Then
            For iSet = LBound(tgVehicleInfo(iLoop).sZone) To UBound(tgVehicleInfo(iLoop).sZone) Step 1
                If tgVehicleInfo(iLoop).sZone(iSet) = tgCPPosting(0).sZone Then
                    If tgVehicleInfo(iLoop).sFed(iSet) = "*" Then
                        sZone = tgVehicleInfo(iLoop).sZone(iSet)
                    Else
                        sZone = tgVehicleInfo(iLoop).sFed(iSet) & "ST"
                    End If
                End If
            Next iSet
        End If
    Next iLoop
'    On Error GoTo ErrHand
'    SQLQuery = "INSERT INTO lst (lstType, lstSdfCode, lstCntrNo, "
'    SQLQuery = SQLQuery & "lstAdfCode, lstAgfCode, lstProd, "
'    SQLQuery = SQLQuery & "lstLineNo, lstLnVefCode, lstStartDate,"
'    SQLQuery = SQLQuery & "lstEndDate, lstMon, lstTue, "
'    SQLQuery = SQLQuery & "lstWed, lstThu, lstFri, "
'    SQLQuery = SQLQuery & "lstSat, lstSun, lstSpotsWk, "
'    SQLQuery = SQLQuery & "lstPriceType, lstPrice, lstSpotType, "
'    SQLQuery = SQLQuery & "lstLogVefCode, lstLogDate, lstLogTime, "
'    SQLQuery = SQLQuery & "lstDemo, lstAud, lstISCI, "
'    SQLQuery = SQLQuery & "lstWkNo, lstBreakNo, lstPositionNo, "
'    SQLQuery = SQLQuery & "lstSeqNo, lstZone, lstCart, "
'    SQLQuery = SQLQuery & "lstCpfCode, lstCrfCsfCode, lstStatus, "
'    SQLQuery = SQLQuery & "lstLen, lstUnits, lstCifCode, "
'    SQLQuery = SQLQuery & "lstAnfCode)"
'    SQLQuery = SQLQuery & " VALUES (" & 2 & ", " & 0 & ", " & lCntrNo & ", "
'    SQLQuery = SQLQuery & imAdfCode & ", " & 0 & ", '" & "" & "', "
'    SQLQuery = SQLQuery & 0 & ", " & 0 & ", '" & Format$("1/1/1970", sgSQLDateForm) & "', "
'    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', " & 0 & ", " & 0 & ", "
'    SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
'    SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
'    SQLQuery = SQLQuery & 1 & ", " & 0 & ", " & 5 & ", "
'    SQLQuery = SQLQuery & tgCPPosting(0).iVefCode & ", '" & Format$(sAirDate, sgSQLDateForm) & "', '" & Format$(sAirTime, sgSQLTimeForm) & "', "
'    SQLQuery = SQLQuery & "'" & "" & "', " & 0 & ", '" & "" & "', "
'    SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
'    SQLQuery = SQLQuery & 0 & ", '" & sZone & "', '" & "" & "', "
'    SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 21 & ", "
'    SQLQuery = SQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
'    SQLQuery = SQLQuery & 0 & ")"
'    cnn.BeginTrans
'    cnn.Execute SQLQuery, rdExecDirect
'    cnn.CommitTrans
'    SQLQuery = "Select MAX(lstCode) from lst"
'    Set rst = gSQLSelectCall(SQLQuery)
'    lLstCode = rst(0).Value
'    SQLQuery = "INSERT INTO ast"
'    SQLQuery = SQLQuery + "(astAtfCode, astShfCode, astVefCode, "
'    SQLQuery = SQLQuery + "astSdfCode, astLsfCode, astAirDate, astAirTime, "
'    SQLQuery = SQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, astPledgeStartTime,astPledgeEndTime)"
'    SQLQuery = SQLQuery + " VALUES "
'    SQLQuery = SQLQuery + "(" & tgCPPosting(0).lAttCode & ", " & tgCPPosting(0).iShttCode & ", "
'    SQLQuery = SQLQuery & tgCPPosting(0).iVefCode & ", " & 0 & ", " & lLstCode & ", "
'    SQLQuery = SQLQuery + "'" & Format$(sAirDate, sgSQLDateForm) & "', '" & Format$(sAirTime, sgSQLTimeForm) & "', "
'    SQLQuery = SQLQuery & 21 & ", " & "1" & ", '" & Format$(sAirDate, sgSQLDateForm) & "', "
'    SQLQuery = SQLQuery & "'" & Format$(sAirTime, sgSQLTimeForm) & "', '" & Format$(sAirDate, sgSQLDateForm) & "', '" & Format$(sAirTime, sgSQLTimeForm) & "', '" & Format$(sAirTime, sgSQLTimeForm) & "')"
'    cnn.BeginTrans
'    cnn.Execute SQLQuery, rdExecDirect
'    cnn.CommitTrans
'    On Error GoTo 0
    If gAddBonusSpot(lCntrNo, imAdfCode, tgCPPosting(0).iVefCode, sAirDate, sAirTime, sZone, tgCPPosting(0).lAttCode, tgCPPosting(0).iShttCode, "", "", "", ilLen) Then
        igUpdateDTGrid = True
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "frmAddBonus-cmdAdd"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
    txtDate.Text = ""
    Unload frmAddBonus
End Sub


Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.05
    Me.Height = Screen.Height / 1.7
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()

    frmAddBonus.Caption = "Add Bonus - " & sgClientName
    Screen.MousePointer = vbHourglass
    'Me.Width = Screen.Width / 1.5
    'Me.Height = Screen.Height / 1.7
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    smDate = ""
    smTime = ""
    
    mFillAdvt
    Screen.MousePointer = vbDefault
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set frmAddBonus = Nothing
End Sub





Private Sub lbcAdvertiser_Click()
    On Error GoTo ErrHand
    
    lbcContract.Clear
    If lbcAdvertiser.ListIndex < 0 Then
        Exit Sub
    End If
    If txtDate.Text = "" Then
        gMsgBox "Date must be specified.", vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If
    If gIsDate(txtDate.Text) = False Then
        Beep
        gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
        txtDate.SetFocus
        Exit Sub
    Else
        smDate = Format(txtDate.Text, sgShowDateForm)
    End If
    Screen.MousePointer = vbHourglass
    imAdfCode = lbcAdvertiser.ItemData(lbcAdvertiser.ListIndex)
    
    'SQLQuery = "SELECT DISTINCT chf.chfCntrNo from ADF_Advertisers adf, CHF_Contract_Header chf"
    SQLQuery = "SELECT DISTINCT chfCntrNo"
    SQLQuery = SQLQuery + " FROM ADF_Advertisers, "
    SQLQuery = SQLQuery & "CHF_Contract_Header"
    SQLQuery = SQLQuery + " WHERE (adfCode = chfAdfCode"
    SQLQuery = SQLQuery + " AND adfCode = " & imAdfCode
    SQLQuery = SQLQuery + " AND chfStartDate <= '" & Format$(smDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery + " AND chfEndDate >= '" & Format$(smDate, sgSQLDateForm) & "'"
    SQLQuery = SQLQuery + " AND (chfStatus = 'O'"
    SQLQuery = SQLQuery + " OR chfStatus = 'H'))"
    SQLQuery = SQLQuery + " ORDER BY chfCntrNo"
    
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        lbcContract.AddItem rst!chfCntrNo  ', " & rst(1).Value & ""
        rst.MoveNext
    Wend
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Add Bonus-lbcAdvertiser"
End Sub

Private Sub txtTime_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub txtDate_Change()
    imAdfCode = -1
    lbcAdvertiser.ListIndex = -1
    lbcContract.Clear
End Sub

Private Sub txtDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub mFillAdvt()
    Dim iFound As Integer
    Dim iLoop As Integer
    On Error GoTo ErrHand
    
    lbcAdvertiser.Clear
    lbcContract.Clear
    
    'SQLQuery = "SELECT adf.adfName, adf.adfCode from ADF_Advertisers adf"
    SQLQuery = "SELECT adfName, adfCode"
    SQLQuery = SQLQuery + " FROM ADF_Advertisers"
    SQLQuery = SQLQuery + " ORDER BY adfName"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        iFound = False
        If Not iFound Then
            lbcAdvertiser.AddItem rst!adfName '& ", " & rst(1).Value
            lbcAdvertiser.ItemData(lbcAdvertiser.NewIndex) = rst!adfCode
        End If
        rst.MoveNext
    Wend
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "Add Bonus-mFillAdvt"
End Sub

