VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDuplSHTTFix 
   Caption         =   "Clean-up Station Info"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   Icon            =   "AffDuplSHTTFix.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   7350
   Begin VB.ListBox lbcResults 
      Height          =   2595
      ItemData        =   "AffDuplSHTTFix.frx":08CA
      Left            =   210
      List            =   "AffDuplSHTTFix.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   900
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
      FormDesignHeight=   4830
      FormDesignWidth =   7350
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3975
      TabIndex        =   1
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Fix"
      Height          =   375
      Left            =   2025
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin ComctlLib.ProgressBar plcGauge 
      Height          =   210
      Left            =   1500
      TabIndex        =   3
      Top             =   3945
      Visible         =   0   'False
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   370
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Remove Unused Market Names and Formats"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   255
      TabIndex        =   6
      Top             =   285
      Width           =   6795
   End
   Begin VB.Label lacDuplCheck 
      Alignment       =   2  'Center
      Caption         =   "Checking for Duplicates with Market Name, Owner Name,  Format and Territory"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   270
      TabIndex        =   5
      Top             =   60
      Width           =   6795
   End
   Begin VB.Label lacTitle2 
      Alignment       =   2  'Center
      Caption         =   "Results"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   5790
   End
End
Attribute VB_Name = "frmDuplSHTTFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmDuplSHTTFix - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imDuplCheck As Integer
Private imTerminate As Integer
Private lmNoChecked As Long
Private lmTotalToCheck As Long
Private lmPercent As Long
Private rst_Raf As ADODB.Recordset





Private Sub cmdCancel_Click()
    If imDuplCheck Then
        imTerminate = True
        Exit Sub
    End If
    Unload frmDuplSHTTFix
End Sub

Private Sub cmdOk_Click()
    Dim ilRet As Integer
    
    Screen.MousePointer = vbHourglass
    gLogMsg "Clean-up of Station Information started", "CleanupStationInfoLog.Txt", False
    imDuplCheck = True
    lmNoChecked = 0
    
    If Not gPopMarkets() Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "Unable to Load Existing Market Names."
        Exit Sub
    End If
    lmTotalToCheck = 2 * UBound(tgMarketInfo)
    If Not gPopOwnerNames() Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "Unable to Load Existing Owner Names."
        Exit Sub
    End If
    lmTotalToCheck = lmTotalToCheck + UBound(tgOwnerInfo)

    If Not gPopFormats() Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "Unable to Load Existing Format Names."
        Exit Sub
    End If
    lmTotalToCheck = lmTotalToCheck + 2 * UBound(tgFormatInfo)
    
    If Not gPopMntInfo("T", tgTerritoryInfo()) Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "Unable to Load Existing Territory Names."
        Exit Sub
    End If
    lmTotalToCheck = lmTotalToCheck + UBound(tgTerritoryInfo)
    
    plcGauge.Value = 0
    plcGauge.Visible = True
    lbcResults.Clear
    ilRet = mRemoveDuplicateMarketNames()
    If ilRet = 2 Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "User terminated Duplicate Station Info."
        Exit Sub
    End If
    If ilRet = 3 Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "Unable to Remove Duplicate Market Names."
        Exit Sub
    End If
    If ilRet = 1 Then
        '11/26/17
        gFileChgdUpdate "shtt.mkd", True
        If Not gPopMarkets() Then
            Screen.MousePointer = vbDefault
            imDuplCheck = False
            mFailDuplStationInfo "Unable to Load Existing Market Names."
            Exit Sub
        End If
    End If
    ilRet = mRemoveDuplicateOwnerNames()
    If ilRet = 2 Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "User terminated Duplicate Station Info."
        Exit Sub
    End If
    If ilRet = 3 Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "Unable to Remove Duplicate Owner Names."
        Exit Sub
    End If
    If ilRet = 1 Then
        '11/26/17
        gFileChgdUpdate "shtt.mkd", True
        If Not gPopOwnerNames() Then
            Screen.MousePointer = vbDefault
            imDuplCheck = False
            mFailDuplStationInfo "Unable to Load Existing Owners Names."
            Exit Sub
        End If
    End If
    ilRet = mRemoveDuplicateFormatNames()
    If ilRet = 2 Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "User terminated Duplicate Station Info."
        Exit Sub
    End If
    If ilRet = 3 Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "Unable to Remove Dupliacte Format Names."
        Exit Sub
    End If
    If ilRet = 1 Then
        '11/26/17
        gFileChgdUpdate "shtt.mkd", True
        If Not gPopFormats() Then
            Screen.MousePointer = vbDefault
            imDuplCheck = False
            mFailDuplStationInfo "Unable to Load Existing Format Names."
            Exit Sub
        End If
    End If
    ilRet = mRemoveDuplicateTerritoryNames()
    If ilRet = 2 Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "User terminated Duplicate Station Info."
        Exit Sub
    End If
    If ilRet = 3 Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "Unable to Remove Duplicate Territory Names."
        Exit Sub
    End If
    If ilRet = 1 Then
        '11/26/17
        gFileChgdUpdate "shtt.mkd", True
        If Not gPopMntInfo("T", tgTerritoryInfo()) Then
            Screen.MousePointer = vbDefault
            imDuplCheck = False
            mFailDuplStationInfo "Unable to Load Existing Territory Names."
            Exit Sub
        End If
    End If
    'Remove unused formats and markets
    ilRet = mRemoveUnsedShttItems()
    If ilRet = 2 Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "User terminated Duplicate Station Info."
        Exit Sub
    End If
    If ilRet = 3 Then
        Screen.MousePointer = vbDefault
        imDuplCheck = False
        mFailDuplStationInfo "Unable to Remove Unused Formats and Markets."
        Exit Sub
    End If
    If ilRet = 1 Then
        '11/26/17
        gFileChgdUpdate "shtt.mkd", True
        If Not gPopFormats() Then
            Screen.MousePointer = vbDefault
            imDuplCheck = False
            mFailDuplStationInfo "Unable to Load Existing Format Names."
            Exit Sub
        End If
        If Not gPopMarkets() Then
            Screen.MousePointer = vbDefault
            imDuplCheck = False
            mFailDuplStationInfo "Unable to Load Existing Market Names."
            Exit Sub
        End If
    End If
    gLogMsg "Clean-up of Station Information completed", "CleanupStationInfoLog.Txt", False
    plcGauge.Value = 100
    imDuplCheck = False
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
    imDuplCheck = False
    imTerminate = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_Raf.Close
    Set frmDuplSHTTFix = Nothing
End Sub



Private Function mRemoveDuplicateMarketNames() As Integer
    Dim llOutsideLoop As Long
    Dim llInsideLoop As Long
    Dim ilRet As Integer
    Dim llDuplRemoved As Long
    
    mRemoveDuplicateMarketNames = 0
    On Error GoTo ErrHand:
    mSetResults "Checking and removing duplicated Market Names...", RGB(0, 0, 0)
    For llOutsideLoop = LBound(tgMarketInfo) To UBound(tgMarketInfo) - 1 Step 1
        DoEvents
        If imTerminate Then
            mRemoveDuplicateMarketNames = 2
            Exit Function
        End If
        If tgMarketInfo(llOutsideLoop).lCode > 0 Then
            llDuplRemoved = 0
            For llInsideLoop = llOutsideLoop + 1 To UBound(tgMarketInfo) - 1 Step 1
                If StrComp(tgMarketInfo(llInsideLoop).sName, tgMarketInfo(llOutsideLoop).sName, vbTextCompare) = 0 Then
                    llDuplRemoved = llDuplRemoved + 1
                    mRemoveDuplicateMarketNames = 1
                    SQLQuery = "Update shtt Set shttMktCode = " & tgMarketInfo(llOutsideLoop).lCode & " Where shttMktCode = " & tgMarketInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveDuplicateMarketNames"
                        mRemoveDuplicateMarketNames = 3
                        Exit Function
                    End If
                    SQLQuery = "Update mat Set matMktCode = " & tgMarketInfo(llOutsideLoop).lCode & " Where matMktCode = " & tgMarketInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveDuplicateMarketNames"
                        mRemoveDuplicateMarketNames = 3
                        Exit Function
                    End If
                    SQLQuery = "Update mgt Set mgtMktCode = " & tgMarketInfo(llOutsideLoop).lCode & " Where mgtMktCode = " & tgMarketInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveDuplicateMarketNames"
                        mRemoveDuplicateMarketNames = 3
                        Exit Function
                    End If
                    ilRet = gUpdateRegions("M", tgMarketInfo(llInsideLoop).lCode, tgMarketInfo(llOutsideLoop).lCode, "CleanupStationInfoLog.Txt")
                    If Not ilRet Then
                        mRemoveDuplicateMarketNames = 3
                        Exit Function
                    End If
                    SQLQuery = "DELETE FROM mkt WHERE mktCode = " & tgMarketInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveDuplicateMarketNames"
                        mRemoveDuplicateMarketNames = 3
                        Exit Function
                    End If
                    tgMarketInfo(llInsideLoop).lCode = -1
                End If
            Next llInsideLoop
            If llDuplRemoved > 0 Then
                mSetResults "Market: " & Trim$(tgMarketInfo(llOutsideLoop).sName) & " removed " & llDuplRemoved, vbRed  'RGB(0, 0, 0)
            End If
        End If
        lmNoChecked = lmNoChecked + 1
        lmPercent = (lmNoChecked * CSng(100)) / lmTotalToCheck
        If lmPercent >= 100 Then
            lmPercent = 100
        End If
        If plcGauge.Value <> lmPercent Then
            plcGauge.Value = lmPercent
            DoEvents
        End If
    Next llOutsideLoop
    
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmDuplSHTTFix-mRemoveDuplicateMarketNames"
     mRemoveDuplicateMarketNames = 3
    Exit Function
End Function

Private Function mRemoveDuplicateFormatNames() As Integer
    Dim llOutsideLoop As Long
    Dim llInsideLoop As Long
    Dim ilRet As Integer
    Dim llDuplRemoved As Long
    
    mRemoveDuplicateFormatNames = 0
    On Error GoTo ErrHand:
    mSetResults "Checking and removing duplicated Format Names...", RGB(0, 0, 0)
    For llOutsideLoop = LBound(tgFormatInfo) To UBound(tgFormatInfo) - 1 Step 1
        DoEvents
        If imTerminate Then
            mRemoveDuplicateFormatNames = 2
            Exit Function
        End If
        If tgFormatInfo(llOutsideLoop).lCode > 0 Then
            llDuplRemoved = 0
            For llInsideLoop = llOutsideLoop + 1 To UBound(tgFormatInfo) - 1 Step 1
                If StrComp(tgFormatInfo(llInsideLoop).sName, tgFormatInfo(llOutsideLoop).sName, vbTextCompare) = 0 Then
                    llDuplRemoved = llDuplRemoved + 1
                    mRemoveDuplicateFormatNames = 1
                    SQLQuery = "Update shtt Set shttFmtCode = " & tgFormatInfo(llOutsideLoop).lCode & " Where shttFmtCode = " & tgFormatInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveDuplicateFormatNames"
                        mRemoveDuplicateFormatNames = 3
                        Exit Function
                    End If
                    ilRet = gUpdateRegions("F", tgFormatInfo(llInsideLoop).lCode, tgFormatInfo(llOutsideLoop).lCode, "CleanupStationInfoLog.Txt")
                    If Not ilRet Then
                        mRemoveDuplicateFormatNames = 3
                        Exit Function
                    End If
                    SQLQuery = "DELETE FROM FMT_Station_Format WHERE fmtCode = " & tgFormatInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveDuplicateFormatNames"
                        mRemoveDuplicateFormatNames = 3
                        Exit Function
                    End If
                    tgFormatInfo(llInsideLoop).lCode = -1
                End If
            Next llInsideLoop
            If llDuplRemoved > 0 Then
                mSetResults "Format: " & Trim$(tgFormatInfo(llOutsideLoop).sName) & " removed " & llDuplRemoved, vbRed  'RGB(0, 0, 0)
            End If
        End If
        lmNoChecked = lmNoChecked + 1
        lmPercent = (lmNoChecked * CSng(100)) / lmTotalToCheck
        If lmPercent >= 100 Then
            lmPercent = 100
        End If
        If plcGauge.Value <> lmPercent Then
            plcGauge.Value = lmPercent
            DoEvents
        End If
    Next llOutsideLoop
    
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmDuplSHTTFix-mRemoveDuplicateFormatNames"
    mRemoveDuplicateFormatNames = 3
    Exit Function
End Function

Private Function mRemoveDuplicateOwnerNames() As Integer
    Dim llOutsideLoop As Long
    Dim llInsideLoop As Long
    Dim ilRet As Integer
    Dim llDuplRemoved As Long
    
    mRemoveDuplicateOwnerNames = 0
    On Error GoTo ErrHand:
    mSetResults "Checking and removing duplicated Owner Names...", RGB(0, 0, 0)
    For llOutsideLoop = LBound(tgOwnerInfo) To UBound(tgOwnerInfo) - 1 Step 1
        DoEvents
        If imTerminate Then
            mRemoveDuplicateOwnerNames = 2
            Exit Function
        End If
        If tgOwnerInfo(llOutsideLoop).lCode > 0 Then
            llDuplRemoved = 0
            For llInsideLoop = llOutsideLoop + 1 To UBound(tgOwnerInfo) - 1 Step 1
                If StrComp(tgOwnerInfo(llInsideLoop).sName, tgOwnerInfo(llOutsideLoop).sName, vbTextCompare) = 0 Then
                    llDuplRemoved = llDuplRemoved + 1
                    mRemoveDuplicateOwnerNames = 1
                    'SQLQuery = "Update shtt Set shttOwnerArttCode = " & tgOwnerInfo(llOutsideLoop).lCode & " Where shttFmtCode = " & tgOwnerInfo(llInsideLoop).lCode
                    SQLQuery = "Update shtt Set shttOwnerArttCode = " & tgOwnerInfo(llOutsideLoop).lCode & " Where shttOwnerArttCode = " & tgOwnerInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveDuplicateOwnerNames"
                        mRemoveDuplicateOwnerNames = 3
                        Exit Function
                    End If
                    ilRet = gUpdateRegions("O", tgOwnerInfo(llInsideLoop).lCode, tgOwnerInfo(llOutsideLoop).lCode, "CleanupStationInfoLog.Txt")
                    If Not ilRet Then
                        mRemoveDuplicateOwnerNames = 3
                        Exit Function
                    End If
                    SQLQuery = "DELETE FROM artt WHERE arttCode = " & tgOwnerInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveDuplicateOwnerNames"
                        mRemoveDuplicateOwnerNames = 3
                        Exit Function
                    End If
                    tgOwnerInfo(llInsideLoop).lCode = -1
                End If
            Next llInsideLoop
            If llDuplRemoved > 0 Then
                mSetResults "Owner: " & Trim$(tgOwnerInfo(llOutsideLoop).sName) & " removed " & llDuplRemoved, vbRed    'RGB(0, 0, 0)
            End If
        End If
        lmNoChecked = lmNoChecked + 1
        lmPercent = (lmNoChecked * CSng(100)) / lmTotalToCheck
        If lmPercent >= 100 Then
            lmPercent = 100
        End If
        If plcGauge.Value <> lmPercent Then
            plcGauge.Value = lmPercent
            DoEvents
        End If
    Next llOutsideLoop
    
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmDuplSHTTFix-mRemoveDuplicateOwnerNames"
    mRemoveDuplicateOwnerNames = 3
    Exit Function
End Function

Private Function mRemoveDuplicateTerritoryNames() As Integer
    Dim llOutsideLoop As Long
    Dim llInsideLoop As Long
    Dim ilRet As Integer
    Dim llDuplRemoved As Long
    
    mRemoveDuplicateTerritoryNames = 0
    On Error GoTo ErrHand:
    mSetResults "Checking and removing duplicated Territory Names...", RGB(0, 0, 0)
    For llOutsideLoop = LBound(tgTerritoryInfo) To UBound(tgTerritoryInfo) - 1 Step 1
        DoEvents
        If imTerminate Then
            mRemoveDuplicateTerritoryNames = 2
            Exit Function
        End If
        If tgTerritoryInfo(llOutsideLoop).lCode > 0 Then
            llDuplRemoved = 0
            For llInsideLoop = llOutsideLoop + 1 To UBound(tgTerritoryInfo) - 1 Step 1
                If StrComp(tgTerritoryInfo(llInsideLoop).sName, tgTerritoryInfo(llOutsideLoop).sName, vbTextCompare) = 0 Then
                    llDuplRemoved = llDuplRemoved + 1
                    mRemoveDuplicateTerritoryNames = 1
                    SQLQuery = "Update shtt Set shttMntCode = " & tgTerritoryInfo(llOutsideLoop).lCode & " Where shttMntCode = " & tgTerritoryInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveDuplicateTerritoryNames"
                        mRemoveDuplicateTerritoryNames = 3
                        Exit Function
                    End If
                    SQLQuery = "Update mat Set matMntCode = " & tgTerritoryInfo(llOutsideLoop).lCode & " Where matMntCode = " & tgTerritoryInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveDuplicateTerritoryNames"
                        mRemoveDuplicateTerritoryNames = 3
                        Exit Function
                    End If
                    SQLQuery = "DELETE FROM Mnt WHERE mntCode = " & tgTerritoryInfo(llInsideLoop).lCode
                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                        '6/10/16: Replaced GoSub
                        'GoSub ErrHand:
                        Screen.MousePointer = vbDefault
                        gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveDuplicateTerritoryNames"
                        mRemoveDuplicateTerritoryNames = 3
                        Exit Function
                    End If
                    tgTerritoryInfo(llInsideLoop).lCode = -1
                End If
            Next llInsideLoop
            If llDuplRemoved > 0 Then
                mSetResults "Territory: " & Trim$(tgTerritoryInfo(llOutsideLoop).sName) & " removed " & llDuplRemoved, vbRed    'RGB(0, 0, 0)
            End If
        End If
        lmNoChecked = lmNoChecked + 1
        lmPercent = (lmNoChecked * CSng(100)) / lmTotalToCheck
        If lmPercent >= 100 Then
            lmPercent = 100
        End If
        If plcGauge.Value <> lmPercent Then
            plcGauge.Value = lmPercent
            DoEvents
        End If
    Next llOutsideLoop
    
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmDuplSHTTFix-mRemoveDuplicateTerritoryNames"
    mRemoveDuplicateTerritoryNames = 3
    Exit Function
End Function

'***************************************************************************
'
'***************************************************************************
Private Sub mSetResults(Msg As String, FGC As Long)
    gLogMsg Msg, "CleanupStationInfoLog.Txt", False
    lbcResults.AddItem Msg
    lbcResults.ListIndex = lbcResults.ListCount - 1
    lbcResults.ForeColor = FGC
    DoEvents
End Sub

'***************************************************************************
'
'***************************************************************************
Private Sub mFailDuplStationInfo(sMsg As String)
    mSetResults sMsg, RGB(255, 0, 0)
    imDuplCheck = False
    cmdCancel.Caption = "&Done"
    cmdCancel.SetFocus
    Screen.MousePointer = vbDefault
End Sub


Private Function mRemoveUnsedShttItems() As Integer
    Dim llFmt As Long
    Dim llMkt As Long
    Dim llShtt As Long
    Dim ilFound As Integer
    Dim ilRet As Integer
    
    mRemoveUnsedShttItems = 0
    On Error GoTo ErrHand:
    mSetResults "Checking and removing Unused Markets and Formats...", RGB(0, 0, 0)
    
    For llFmt = 0 To UBound(tgFormatInfo) - 1 Step 1
        ilFound = False
        For llShtt = 0 To UBound(tgStationInfo) - 1 Step 1
            DoEvents
            If imTerminate Then
                mRemoveUnsedShttItems = 2
                Exit Function
            End If
            If tgFormatInfo(llFmt).lCode = tgStationInfo(llShtt).iFormatCode Then
                ilFound = True
                Exit For
            End If
        Next llShtt
        If Not ilFound Then
            SQLQuery = "SELECT rafCode FROM raf_region_area, sef_Split_Entity WHERE sefRafCode = rafCode AND ((rafCategory = 'F') and (rafType = 'C' OR rafType = 'N')) AND sefIntCode = " & tgFormatInfo(llFmt).lCode
            Set rst_Raf = gSQLSelectCall(SQLQuery)
            If Not rst_Raf.EOF Then
                ilFound = True
            Else
                SQLQuery = "SELECT rafCode FROM raf_region_area, sef_Split_Entity  WHERE sefRafCode = rafCode AND ((rafCategory = '" & " " & "')  AND (sefCategory = 'F') and (rafType = 'C' OR rafType = 'N')) AND sefIntCode = " & tgFormatInfo(llFmt).lCode
                Set rst_Raf = gSQLSelectCall(SQLQuery)
                If Not rst_Raf.EOF Then
                    ilFound = True
                End If
            End If
        End If
        If Not ilFound Then
            SQLQuery = "DELETE FROM FMT_Station_Format WHERE fmtCode = " & tgFormatInfo(llFmt).lCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveUnsedShttItems"
                mRemoveUnsedShttItems = 3
                Exit Function
            End If
            tgFormatInfo(llFmt).lCode = -1
            mRemoveUnsedShttItems = 1
            mSetResults "Format: " & Trim$(tgFormatInfo(llFmt).sName) & " removed ", vbRed  'RGB(0, 0, 0)
        End If
        lmNoChecked = lmNoChecked + 1
        lmPercent = (lmNoChecked * CSng(100)) / lmTotalToCheck
        If lmPercent >= 100 Then
            lmPercent = 100
        End If
        If plcGauge.Value <> lmPercent Then
            plcGauge.Value = lmPercent
            DoEvents
        End If
    Next llFmt
    For llMkt = 0 To UBound(tgMarketInfo) - 1 Step 1
        ilFound = False
        For llShtt = 0 To UBound(tgStationInfo) - 1 Step 1
            DoEvents
            If imTerminate Then
                mRemoveUnsedShttItems = 2
                Exit Function
            End If
            If tgMarketInfo(llMkt).lCode = tgStationInfo(llShtt).iMktCode Then
                ilFound = True
                Exit For
            End If
        Next llShtt
        If Not ilFound Then
            SQLQuery = "SELECT rafCode FROM raf_region_area, sef_Split_Entity WHERE sefRafCode = rafCode AND ((rafCategory = 'M') and (rafType = 'C' OR rafType = 'N')) AND sefIntCode = " & tgMarketInfo(llMkt).lCode
            Set rst_Raf = gSQLSelectCall(SQLQuery)
            If Not rst_Raf.EOF Then
                ilFound = True
            Else
                SQLQuery = "SELECT rafCode FROM raf_region_area, sef_Split_Entity  WHERE sefRafCode = rafCode AND ((rafCategory = '" & " " & "')  AND (sefCategory = 'M') and (rafType = 'C' OR rafType = 'N')) AND sefIntCode = " & tgMarketInfo(llMkt).lCode
                Set rst_Raf = gSQLSelectCall(SQLQuery)
                If Not rst_Raf.EOF Then
                    ilFound = True
                End If
            End If
        End If
        If Not ilFound Then
            SQLQuery = "DELETE FROM mkt WHERE mktCode = " & tgMarketInfo(llMkt).lCode
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/10/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "DuplSHTTFix-mRemoveUnsedShttItems"
                mRemoveUnsedShttItems = 3
                Exit Function
            End If
            tgMarketInfo(llMkt).lCode = -1
            mRemoveUnsedShttItems = 1
            mSetResults "Market: " & Trim$(tgMarketInfo(llMkt).sName) & " removed ", vbRed  'RGB(0, 0, 0)
        End If
        lmNoChecked = lmNoChecked + 1
        lmPercent = (lmNoChecked * CSng(100)) / lmTotalToCheck
        If lmPercent >= 100 Then
            lmPercent = 100
        End If
        If plcGauge.Value <> lmPercent Then
            plcGauge.Value = lmPercent
            DoEvents
        End If
    Next llMkt
    On Error GoTo 0
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmDuplSHTTFix-mRemoveUnsedShttItems"
     mRemoveUnsedShttItems = 3
    Exit Function
End Function
