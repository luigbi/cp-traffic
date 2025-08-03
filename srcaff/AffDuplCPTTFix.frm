VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDuplCPTTFix 
   Caption         =   "Duplicate CPTT Fix"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   Icon            =   "AffDuplCPTTFix.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7350
   Begin VB.ListBox lbcResults 
      Height          =   2985
      ItemData        =   "AffDuplCPTTFix.frx":08CA
      Left            =   210
      List            =   "AffDuplCPTTFix.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   420
      Width           =   6870
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4035
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4680
      FormDesignWidth =   7350
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3975
      TabIndex        =   1
      Top             =   4065
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Fix"
      Height          =   375
      Left            =   2025
      TabIndex        =   0
      Top             =   4065
      Width           =   1335
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   210
      Left            =   1500
      TabIndex        =   4
      Top             =   3690
      Visible         =   0   'False
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label labCount 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   315
      Width           =   3240
   End
End
Attribute VB_Name = "frmDuplCPTTFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmDuplCPTTFix - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private rst_Cptt As ADODB.Recordset




Private Sub cmdCancel_Click()
    Unload frmDuplCPTTFix
End Sub

Private Sub cmdOk_Click()
    Screen.MousePointer = vbHourglass
    plcGauge.Value = 0
    plcGauge.Visible = True
    lbcResults.Clear
    mDuplCPTTFix
    cmdOK.Enabled = False
    cmdCancel.Caption = "Done"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / 3
    Me.Height = (Screen.Height) / 3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_Cptt.Close
    Set frmDuplCPTTFix = Nothing
End Sub

Private Sub mDuplCPTTFix()
    Dim llUpper As Long
    Dim llLoop As Long
    Dim llPrevCpttCode As Long
    Dim llPrevAttCode As Long
    Dim llPrevStartDate As Long
    Dim ilPrevPostingStatus As Integer
    Dim llNoCPTTToCheck As Long
    Dim llNoChecked As Long
    Dim llPercent As Long
    Dim llNoDeleted As Long
    ReDim llCpttToDelete(0 To 0) As Long
    
    On Error GoTo ErrHand
    SQLQuery = "SELECT Count(cpttCode) FROM cptt "
    Set rst = gSQLSelectCall(SQLQuery)
    If rst(0).Value > 0 Then
        llNoCPTTToCheck = rst(0).Value
        gLogMsg "Checking for duplicate CPTT on " & llNoCPTTToCheck & " records", "DuplCPTTFix.Txt", False
        llNoChecked = 0
        lbcResults.AddItem "Checking for duplicate CPTT's"
        llPrevAttCode = -1
        SQLQuery = "SELECT cpttCode, cpttAtfCode, cpttStartDate, cpttPostingStatus FROM CPTT ORDER BY cpttAtfCode, cpttStartDate"
        Set rst_Cptt = gSQLSelectCall(SQLQuery)
        llUpper = 0
        While Not rst_Cptt.EOF
            If llPrevAttCode = rst_Cptt!cpttatfCode Then
                If llPrevStartDate = DateValue(Format(rst_Cptt!CpttStartDate, "m/d/yy")) Then
                    If ilPrevPostingStatus >= rst_Cptt!cpttPostingStatus Then
                        llCpttToDelete(llUpper) = rst_Cptt!cpttCode
                    Else
                        llCpttToDelete(llUpper) = llPrevCpttCode
                        llPrevCpttCode = rst_Cptt!cpttCode
                        llPrevAttCode = rst_Cptt!cpttatfCode
                        llPrevStartDate = DateValue(Format$(rst_Cptt!CpttStartDate, "m/d/yy"))
                        ilPrevPostingStatus = rst_Cptt!cpttPostingStatus
                    End If
                    llUpper = llUpper + 1
                    ReDim Preserve llCpttToDelete(0 To llUpper) As Long
                Else
                    llPrevCpttCode = rst_Cptt!cpttCode
                    llPrevAttCode = rst_Cptt!cpttatfCode
                    llPrevStartDate = DateValue(Format$(rst_Cptt!CpttStartDate, "m/d/yy"))
                    ilPrevPostingStatus = rst_Cptt!cpttPostingStatus
                End If
            Else
                llPrevCpttCode = rst_Cptt!cpttCode
                llPrevAttCode = rst_Cptt!cpttatfCode
                llPrevStartDate = DateValue(Format$(rst_Cptt!CpttStartDate, "m/d/yy"))
                ilPrevPostingStatus = rst_Cptt!cpttPostingStatus
            End If
            llNoChecked = llNoChecked + 1
            llPercent = (llNoChecked * CSng(100)) / llNoCPTTToCheck
            If llPercent >= 100 Then
                llPercent = 100
            End If
            If plcGauge.Value <> llPercent Then
                plcGauge.Value = llPercent
                DoEvents
            End If
            rst_Cptt.MoveNext
        Wend
        rst_Cptt.Close
        If llUpper > 0 Then
            plcGauge.Value = 0
            llNoDeleted = 0
            gLogMsg "Number of duplicate CPTT found " & llUpper, "DuplCPTTFix.Txt", False
            lbcResults.AddItem "Deleting " & llUpper & " CPTT's"
            For llLoop = 0 To llUpper - 1 Step 1
                SQLQuery = "DELETE From cptt " & " WHERE (cpttCode = " & llCpttToDelete(llLoop) & ")"
                If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                    '6/10/16: Replaced GoSub
                    'GoSub ErrHand:
                    Screen.MousePointer = vbDefault
                    gHandleError "AffErrorLog.txt", "DuplCPTTFix-mDuplCPTTFix"
                    Exit Sub
                End If
                llNoDeleted = llNoDeleted + 1
                llPercent = (llNoDeleted * CSng(100)) / llUpper
                If llPercent >= 100 Then
                    llPercent = 100
                End If
                If plcGauge.Value <> llPercent Then
                    plcGauge.Value = llPercent
                    DoEvents
                End If
            Next llLoop
            gLogMsg "Deleting Duplicate CPTT Finished", "DuplCPTTFix.Txt", False
        Else
            lbcResults.AddItem "No duplicate CPTT found"
            gLogMsg "No duplicate CPTT found", "DuplCPTTFix.Txt", False
        End If
    Else
        lbcResults.AddItem "No CPTT exist"
        gLogMsg "No CPTT exist", "DuplCPTTFix.Txt", False
    End If
    On Error GoTo 0
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "DuplCPTTFix-mDuplCPTTFix"
End Sub

