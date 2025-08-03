VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmModel 
   Caption         =   "Model"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   Icon            =   "AffModel.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7350
   Begin VB.Timer tmcPrt 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7155
      Top             =   3330
   End
   Begin VB.PictureBox pbcPrinting 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   1560
      ScaleHeight     =   1200
      ScaleWidth      =   3825
      TabIndex        =   7
      Top             =   1245
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ListBox lbcResult 
      Height          =   3570
      ItemData        =   "AffModel.frx":08CA
      Left            =   210
      List            =   "AffModel.frx":08CC
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   6870
   End
   Begin VB.ListBox lbcExport 
      Height          =   3570
      ItemData        =   "AffModel.frx":08CE
      Left            =   765
      List            =   "AffModel.frx":08D0
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   180
      Width           =   5730
   End
   Begin VB.ListBox lbcVehicles 
      Height          =   3570
      ItemData        =   "AffModel.frx":08D2
      Left            =   210
      List            =   "AffModel.frx":08D4
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   180
      Width           =   6870
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6915
      Top             =   3960
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
      TabIndex        =   2
      Top             =   4050
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2025
      TabIndex        =   1
      Top             =   4065
      Width           =   1335
   End
   Begin VB.Image imcPrt 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   6300
      Picture         =   "AffModel.frx":08D6
      Top             =   4005
      Width           =   480
   End
   Begin VB.Label labCount 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   315
      Width           =   3240
   End
   Begin VB.Label labCount 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   3240
   End
End
Attribute VB_Name = "frmModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmModel - enters affiliate representative information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private rst_rht As ADODB.Recordset




Private Sub cmdCancel_Click()
    igModelReturn = False
    lgModelFromCode = -1
    Unload frmModel
End Sub

Private Sub cmdOk_Click()
    
    'On Error GoTo ErrHand
    If igModelType = 1 Then
        If lbcVehicles.ListIndex < 0 Then
            igModelReturn = False
            lgModelFromCode = -1
            Unload frmModel
            Exit Sub
        End If
        If (sgUstWin(12) <> "I") Then
            igModelReturn = False
            lgModelFromCode = -1
            Unload frmModel
            Exit Sub
        End If
    ElseIf igModelType = 2 Then
        If lbcExport.ListIndex < 0 Then
            igModelReturn = False
            lgModelFromCode = -1
            Unload frmModel
            Exit Sub
        End If
    ElseIf igModelType = 3 Then
        igModelReturn = True
        lgModelFromCode = -1
        Unload frmModel
        Exit Sub
    End If
    If igModelType = 1 Then
        lgModelFromCode = lbcVehicles.ItemData(lbcVehicles.ListIndex)
    ElseIf igModelType = 2 Then
        lgModelFromCode = lbcExport.ItemData(lbcExport.ListIndex)
    End If
    igModelReturn = True
    Unload frmModel
    Exit Sub
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / 3
    Me.Height = (Screen.Height) / 2 '3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    igModelReturn = False
    mPopulate
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_rht.Close
    Set frmModel = Nothing
End Sub

Private Sub mPopulate()
    Dim ilLoop As Integer
    Dim blInclude As Boolean
    '8156
    Dim ilvehicle As Vendors
    
    On Error GoTo ErrHand
    If igModelType = 1 Then
        frmModel.Caption = "Model"
        lbcExport.Visible = False
        lbcResult.Visible = False
        imcPrt.Visible = False
        lbcVehicles.Clear
        For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
            'Test if any radar program schedule created
            SQLQuery = "SELECT * FROM rht WHERE (rhtVefCode = " & tgVehicleInfo(ilLoop).iCode & ")"
            Set rst_rht = gSQLSelectCall(SQLQuery)
            Do While Not rst_rht.EOF
                lbcVehicles.AddItem Trim$(tgVehicleInfo(ilLoop).sVehicle) & ": " & Trim$(rst_rht!rhtRadarNetCode) & "-" & Trim$(rst_rht!rhtRadarVehCode)
                lbcVehicles.ItemData(lbcVehicles.NewIndex) = rst_rht!rhtCode
                rst_rht.MoveNext
            Loop
        Next ilLoop
        rst_rht.Close
        lbcVehicles.ListIndex = -1
    ElseIf igModelType = 2 Then
        frmModel.Caption = "Custom"
        lbcVehicles.Visible = False
        lbcResult.Visible = False
        imcPrt.Visible = False
        lbcExport.Clear
        For ilLoop = 0 To UBound(tgSpecInfo) Step 1
            If tgSpecInfo(ilLoop).sType = "A" Then
                '8156
                If gAdjustAllowedExportsImports(Vendors.NetworkConnect, False) Then
                    lbcExport.AddItem "Marketron"
                    lbcExport.ItemData(lbcExport.NewIndex) = Asc("1")
                End If
                If gUsingUnivision Then
                    lbcExport.AddItem "Univision Scheduled Station Spots"
                    lbcExport.ItemData(lbcExport.NewIndex) = Asc("2")
                End If
                If gUsingWeb Then
                    lbcExport.AddItem "Counterpoint Affidavit"
                    lbcExport.ItemData(lbcExport.NewIndex) = Asc("3")
                End If
            Else
                blInclude = True
                If (gISCIExport = False) And (tgSpecInfo(ilLoop).sType = "I") Then
                    blInclude = False
                End If
                If (sgRCSExportCart4 = "N") And (tgSpecInfo(ilLoop).sType = "4") Then
                    blInclude = False
                End If
                If (sgRCSExportCart5 = "N") And (tgSpecInfo(ilLoop).sType = "5") Then
                    blInclude = False
                End If
                If (gUsingXDigital = False) And (tgSpecInfo(ilLoop).sType = "X") Then
                    blInclude = False
                End If
                If (gWegenerExport = False) And (tgSpecInfo(ilLoop).sType = "W") Then
                    blInclude = False
                End If
                '8156
                If blInclude Then
                    Select Case tgSpecInfo(ilLoop).sType
                        Case "X"
                            ilvehicle = XDS_Break
                        Case "W"
                            ilvehicle = Wegener_Compel
                        Case "P"
                           ilvehicle = Vendors.Wegener_IPump
                        Case "D"
                           ilvehicle = Vendors.iDc
                        Case Else
                            ilvehicle = Vendors.None
                    End Select
                    If ilvehicle > Vendors.None Then
                        blInclude = gAdjustAllowedExportsImports(ilvehicle, False)
                        If Not blInclude And ilvehicle = Vendors.XDS_Break Then
                            blInclude = gAdjustAllowedExportsImports(XDS_ISCI, False)
                        End If
                    End If
                End If
                If blInclude Then
                    lbcExport.AddItem Trim$(tgSpecInfo(ilLoop).sFullName)
                    lbcExport.ItemData(lbcExport.NewIndex) = Asc(tgSpecInfo(ilLoop).sType)
                End If
            End If
        Next ilLoop
    ElseIf igModelType = 3 Then
        frmModel.Caption = "Result List"
        lbcExport.Visible = False
        lbcVehicles.Visible = False
        lbcResult.Visible = True
        cmdCancel.Visible = False
        imcPrt.Visible = True
        imcPrt.Picture = frmDirectory!imcPrinter.Picture
        cmdOK.Left = (frmModel.Width - cmdOK.Width) / 2
        cmdOK.Caption = "Done"
        pbcPrinting.Move lbcResult.Left + (lbcResult.Width - pbcPrinting.Width) / 2, lbcResult.Top + (lbcResult.Height - pbcPrinting.Height) / 2
        mDisplay sgResultFileName
    End If
    On Error GoTo 0
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErorLog.txt", "frmModel-mPopulate"
End Sub

Private Sub mDisplay(slFileName As String)

    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim llRet As Long
    Dim slTemp As String
    Dim slRetString As String
    Dim llMaxWidth As Long
    Dim llValue As Long
    Dim llRg As Long
    Dim slCurDir As String
    Dim blFdStart As Boolean
    Dim blHeaderFd As Boolean
    Dim slSearchString As String
    Dim ilPos As Integer
    Dim slDateTime As String
    
    slCurDir = CurDir
    'Make Sure we start out each time without a horizontal scroll bar
    slSearchString = "Result List, Started:"
    lbcResult.Clear
    llValue = 0
    llRg = 0
    llRet = SendMessageByNum(lbcResult.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    llMaxWidth = 0
    If fs.FILEEXISTS(sgMsgDirectory & slFileName) Then
        Set tlTxtStream = fs.OpenTextFile(sgMsgDirectory & slFileName, ForReading, False)
    Else
        lbcResult.AddItem "** No Data Available **"
        Exit Sub
    End If
    slTemp = ""
    blFdStart = False
    blHeaderFd = False
    Do While tlTxtStream.AtEndOfStream <> True
        slRetString = tlTxtStream.ReadLine
        If Not blFdStart Then
            ilPos = InStr(1, slRetString, slSearchString, vbBinaryCompare)
            If ilPos > 0 Then
                blHeaderFd = True
                ilPos = InStr(1, slRetString, ":", vbBinaryCompare)
                If ilPos > 0 Then
                    slDateTime = Trim$(Mid(slRetString, ilPos + 1))
                    If gDateValue(Format(slDateTime, sgShowDateForm)) >= gDateValue(Format(gNow(), sgShowDateForm)) - 1 Then
                        blFdStart = True
                    End If
                End If
            End If
        End If
        If blFdStart Then
            lbcResult.AddItem slRetString
            If (frmMessages.pbcArial.TextWidth(slRetString)) > llMaxWidth Then
                llMaxWidth = (frmMessages.pbcArial.TextWidth(slRetString))
            End If
        End If
    Loop
    If blFdStart = False Then
        If blHeaderFd = False Then
            lbcResult.AddItem "** 'Result List, Started:' header line not found **"
        Else
            lbcResult.AddItem "** Start date on Yesterday, Today or in the Future not found **"
        End If
    End If
    'Show a horzontal scroll bar if needed
    If llMaxWidth > lbcResult.Width Then
        llValue = llMaxWidth / 15 + 120
        llRg = 0
        llRet = SendMessageByNum(lbcResult.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    End If
    tlTxtStream.Close
    ChDir slCurDir
    
End Sub

Private Sub imcPrt_Click()
    Dim ilCurrentLineNo As Integer
    Dim ilLinesPerPage As Integer
    Dim slRecord As String
    Dim slHeading As String
    Dim ilLoop As Integer
    Dim ilRet As Integer
    If lbcResult.ListCount <= 0 Then
        Exit Sub
    End If
    pbcPrinting.Visible = True
    DoEvents
    ilCurrentLineNo = 0
    ilLinesPerPage = (Printer.Height - 1440) / Printer.TextHeight("TEST") - 1
    ilRet = 0
    On Error GoTo imcPrtErr:
    slHeading = "Printing content from " & Format(gNow(), "m/d/yy") & " of " & sgResultFileName & " for " & Trim$(sgUserName) & " on " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    '6/12/16: Replaced GoSub
    'GoSub mHeading1
    mHeader1 slHeading, ilCurrentLineNo, ilRet
    If ilRet <> 0 Then
        Printer.EndDoc
        On Error GoTo 0
        pbcPrinting.Visible = False
        Exit Sub
    End If
    'Output Information
    For ilLoop = 0 To lbcResult.ListCount - 1 Step 1
        slRecord = "    " & lbcResult.List(ilLoop)
        '6/12/16: Replaced GoSub
        'GoSub mLineOutput
        mLineOutput slHeading, slRecord, ilCurrentLineNo, ilLinesPerPage, ilRet
        If ilRet <> 0 Then
            Printer.EndDoc
            On Error GoTo 0
            pbcPrinting.Visible = False
            Exit Sub
        End If
    Next ilLoop
    Printer.EndDoc
    On Error GoTo 0
    'pbcPrinting.Visible = False
    tmcPrt.Enabled = True
    Exit Sub
'mHeading1:  'Output file name and date
'    Printer.Print slHeading
'    If ilRet <> 0 Then
'        Return
'    End If
'    ilCurrentLineNo = ilCurrentLineNo + 1
'    Printer.Print " "
'    ilCurrentLineNo = ilCurrentLineNo + 1
'    Return
'mLineOutput:
'    If ilCurrentLineNo >= ilLinesPerPage Then
'        Printer.NewPage
'        If ilRet <> 0 Then
'            Return
'        End If
'        ilCurrentLineNo = 0
'        GoSub mHeading1
'        If ilRet <> 0 Then
'            Return
'        End If
'    End If
'    Printer.Print slRecord
'    ilCurrentLineNo = ilCurrentLineNo + 1
'    Return
imcPrtErr:
    ilRet = Err.Number
        gMsgBox "Printing Error #  " & Str$(ilRet), vbCritical
    Resume Next
End Sub

Private Sub tmcPrt_Timer()

    tmcPrt.Enabled = False
    pbcPrinting.Visible = False

End Sub

Private Sub mLineOutput(slHeading As String, slRecord As String, ilCurrentLineNo As Integer, ilLinesPerPage As Integer, ilRet As Integer)
    On Error GoTo imcPrtErr:
    If ilCurrentLineNo >= ilLinesPerPage Then
        Printer.NewPage
        If ilRet <> 0 Then
            'Return
            Exit Sub
        End If
        ilCurrentLineNo = 0
        mHeader1 slHeading, ilCurrentLineNo, ilRet
        If ilRet <> 0 Then
            'Return
            Exit Sub
        End If
    End If
    Printer.Print slRecord
    ilCurrentLineNo = ilCurrentLineNo + 1
    Exit Sub
imcPrtErr:
    ilRet = Err.Number
    gMsgBox "Printing Error #  " & Str$(ilRet), vbCritical
    Resume Next
End Sub

Private Sub mHeader1(slHeading As String, ilCurrentLineNo As Integer, ilRet As Integer)
    On Error GoTo imcPrtErr:
    Printer.Print slHeading
    If ilRet <> 0 Then
        'Return
        Exit Sub
    End If
    ilCurrentLineNo = ilCurrentLineNo + 1
    Printer.Print " "
    ilCurrentLineNo = ilCurrentLineNo + 1
    Exit Sub
imcPrtErr:
    ilRet = Err.Number
    gMsgBox "Printing Error #  " & Str$(ilRet), vbCritical
    Resume Next
End Sub

