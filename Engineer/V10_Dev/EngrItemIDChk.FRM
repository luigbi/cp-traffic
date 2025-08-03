VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "Mshflxgd.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MsComm32.ocx"
Begin VB.Form EngrItemIDChk 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6585
   ClientLeft      =   4485
   ClientTop       =   1470
   ClientWidth     =   9330
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6585
   ScaleWidth      =   9330
   Begin VB.CheckBox ckcShow 
      Caption         =   "Show Discrepancies Only"
      Height          =   210
      Left            =   210
      TabIndex        =   26
      Top             =   1410
      Value           =   1  'Checked
      Width           =   3780
   End
   Begin MSCommLib.MSComm spcItemID 
      Left            =   8940
      Top             =   5730
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame frcSendCommands 
      Caption         =   "Send Commands"
      Height          =   1005
      Left            =   195
      TabIndex        =   9
      Top             =   1635
      Visible         =   0   'False
      Width           =   8805
      Begin VB.TextBox edcReply 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   6945
         TabIndex        =   21
         Top             =   630
         Width           =   1695
      End
      Begin VB.TextBox edcReply 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   6945
         TabIndex        =   15
         Top             =   285
         Width           =   1695
      End
      Begin VB.TextBox edcQuery 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1620
         TabIndex        =   17
         Top             =   630
         Width           =   2160
      End
      Begin VB.TextBox edcChecksum 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   5355
         TabIndex        =   19
         Top             =   630
         Width           =   525
      End
      Begin VB.TextBox edcQuery 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1620
         TabIndex        =   11
         Top             =   285
         Width           =   2160
      End
      Begin VB.TextBox edcChecksum 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   5355
         TabIndex        =   13
         Top             =   285
         Width           =   525
      End
      Begin VB.Label lacReply 
         Caption         =   "Reply"
         Height          =   270
         Index           =   1
         Left            =   6270
         TabIndex        =   20
         Top             =   630
         Width           =   735
      End
      Begin VB.Label lacReply 
         Caption         =   "Reply"
         Height          =   270
         Index           =   0
         Left            =   6270
         TabIndex        =   14
         Top             =   285
         Width           =   735
      End
      Begin VB.Label lacQueryPri 
         Caption         =   "Protection- Send"
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label lacCheckSum 
         Caption         =   "Checksum"
         Height          =   270
         Index           =   1
         Left            =   4215
         TabIndex        =   18
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label lacQueryPri 
         Caption         =   "Primary- Send"
         Height          =   270
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   285
         Width           =   1365
      End
      Begin VB.Label lacCheckSum 
         Caption         =   "Checksum"
         Height          =   270
         Index           =   0
         Left            =   4215
         TabIndex        =   12
         Top             =   285
         Width           =   1065
      End
   End
   Begin VB.CheckBox ckcProt 
      Caption         =   "Check Protection Item ID"
      Height          =   210
      Left            =   4680
      TabIndex        =   2
      Top             =   195
      Value           =   1  'Checked
      Width           =   2520
   End
   Begin VB.CheckBox ckcPri 
      Caption         =   "Check Primary Item ID"
      Height          =   210
      Left            =   2355
      TabIndex        =   1
      Top             =   195
      Value           =   1  'Checked
      Width           =   2250
   End
   Begin VB.Frame frcItemID 
      Caption         =   "Check Item ID"
      Height          =   900
      Left            =   4500
      TabIndex        =   6
      Top             =   465
      Width           =   4020
      Begin VB.TextBox edcItemID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   930
         TabIndex        =   8
         Top             =   345
         Width           =   2085
      End
      Begin VB.Label lacItemID 
         Caption         =   "Item ID"
         Height          =   270
         Left            =   150
         TabIndex        =   7
         Top             =   345
         Width           =   735
      End
   End
   Begin VB.Frame frcCheckDate 
      Caption         =   "Check Date"
      Height          =   900
      Left            =   705
      TabIndex        =   3
      Top             =   465
      Width           =   3045
      Begin VB.TextBox edcDate 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   930
         TabIndex        =   5
         Top             =   405
         Width           =   1335
      End
      Begin VB.Label lacDate 
         Caption         =   "Date"
         Height          =   270
         Left            =   150
         TabIndex        =   4
         Top             =   405
         Width           =   735
      End
   End
   Begin VB.CommandButton cmcTest 
      Caption         =   "&Check"
      Height          =   375
      Left            =   5190
      TabIndex        =   24
      Top             =   6000
      Width           =   1200
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3105
      TabIndex        =   23
      Top             =   6000
      Width           =   1200
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   165
      Left            =   -15
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2310
      Width           =   120
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   9135
      Top             =   5175
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   6585
      FormDesignWidth =   9330
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdResults 
      Height          =   4035
      Left            =   210
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1725
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   7117
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   8070
      Picture         =   "EngrItemIDChk.frx":0000
      Top             =   5970
      Width           =   480
   End
   Begin VB.Label lacScreen 
      Caption         =   "Item ID Check"
      Height          =   270
      Left            =   105
      TabIndex        =   0
      Top             =   30
      Width           =   1425
   End
End
Attribute VB_Name = "EngrItemIDChk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: EngrItemIDChk.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Contract revision number increment screen code
Option Explicit
Option Compare Text

Private hmSEE As Integer

Dim imFirstTime As Integer
Dim imFirstActivate As Integer
Dim imLastColSorted As Integer
Dim imLastSort As Integer

Dim smEndCodes() As String

Dim tmSHE As SHE
Dim smCurrSEEStamp As String
Dim tmCurrSEE() As SEE
Dim tmARE As ARE

Dim tmItemIDChk() As ITEMIDCHK

Const ITEMIDINDEX = 0
Const CHECKTITLEINDEX = 1
Const CHECKLENINDEX = 2
Const TYPEINDEX = 3
Const RESULTTITLEINDEX = 4
Const RESULTLENINDEX = 5





Private Sub ckcPri_Click()
    ckcShow_Click
End Sub

Private Sub ckcProt_Click()
    ckcShow_Click
End Sub

Private Sub ckcShow_Click()
    Dim llRow As Long
    Dim slStr As String
    Dim ilLoop As Integer
    Dim slCart As String
    Dim slMsg As String
    
    llRow = grdResults.FixedRows
    slStr = Trim$(grdResults.TextMatrix(llRow, ITEMIDINDEX))
    If slStr <> "" Then
        gGrid_Clear grdResults, True
        For ilLoop = 0 To UBound(tmItemIDChk) - 1 Step 1
            slCart = Trim$(tmItemIDChk(ilLoop).sItemID)
            If (tmItemIDChk(ilLoop).sAudioStatus = "F") Or (tmItemIDChk(ilLoop).sAudioStatus = "O") Then
                If (tmItemIDChk(ilLoop).sAudioStatus = "F") Or (ckcShow.Value = vbUnchecked) Then
                    If llRow + 1 > grdResults.Rows Then
                        grdResults.AddItem ""
                    End If
                    grdResults.Row = llRow
                    grdResults.TextMatrix(llRow, ITEMIDINDEX) = slCart
                    grdResults.TextMatrix(llRow, CHECKTITLEINDEX) = Trim$(tmItemIDChk(ilLoop).sTitle)
                    grdResults.TextMatrix(llRow, TYPEINDEX) = "Pri"
                    grdResults.TextMatrix(llRow, RESULTTITLEINDEX) = Trim$(tmItemIDChk(ilLoop).sPriResult)
                    llRow = llRow + 1
                End If
            End If
            If (tmItemIDChk(ilLoop).sProtStatus = "F") Or (tmItemIDChk(ilLoop).sProtStatus = "O") Then
                If (tmItemIDChk(ilLoop).sProtStatus = "F") Or (ckcShow.Value = vbUnchecked) Then
                    If llRow + 1 > grdResults.Rows Then
                        grdResults.AddItem ""
                    End If
                    grdResults.Row = llRow
                    grdResults.TextMatrix(llRow, ITEMIDINDEX) = slCart
                    grdResults.TextMatrix(llRow, CHECKTITLEINDEX) = Trim$(tmItemIDChk(ilLoop).sTitle)
                    grdResults.TextMatrix(llRow, TYPEINDEX) = "Prot"
                    grdResults.TextMatrix(llRow, RESULTTITLEINDEX) = Trim$(tmItemIDChk(ilLoop).sProtResult)
                    llRow = llRow + 1
                End If
            End If
        Next ilLoop
    End If
End Sub

Private Sub cmcDone_Click()
    igReturnCallStatus = CALLCANCELLED
    Unload EngrItemIDChk
End Sub

Private Sub cmcTest_Click()
    Dim slStr As String
    Dim ilITE As Integer
    Dim ilRet As Integer
    Dim slQuery As String
    Dim slResultTitle As String
    Dim slResultLength As String
    Dim llRow As Long
    Dim slReplyChar As String
    Dim slPriQuery As String
    Dim slProtQuery As String
    Dim slCart As String
    Dim ilLoop As Integer
    Dim ilDoConnectTest As Integer
    
    'Screen.MousePointer = vbHourglass
    gSetMousePointer grdResults, grdResults, vbHourglass
    gGrid_Clear grdResults, True
    llRow = grdResults.FixedRows
    
    ilDoConnectTest = True
    slStr = edcDate.text
    If slStr = "" Then
        slStr = edcItemID.text
        If slStr = "" Then
            'Screen.MousePointer = vbDefault
            gSetMousePointer grdResults, grdResults, vbDefault
            Exit Sub
        Else
            If ckcPri.Value = vbChecked Then
                For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
                    If tgCurrITE(ilITE).sType = "P" Then
                        ''If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or _
                        ''   ((StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0) And (sgSpecialPassword <> "") And (Len(sgSpecialPassword) = 4)) Then
                        ''    slQuery = edcQuery(0).text
                        ''    slPriQuery = edcReply(0).text
                        ''    slCart = Mid$(slQuery, Len(slPriQuery) + 1)
                        ''Else
                            slCart = edcItemID.text
                        '    gBuildItemIDQuery slCart, tgCurrITE(ilITE), slQuery, slPriQuery
                        ''End If
                        'ilRet = gTestItemID(spcItemID, tgCurrITE(ilITE), slQuery, slPriQuery, slResult)
                        ilRet = gTestItemID(spcItemID, tgCurrITE(ilITE), slCart, ilDoConnectTest, slResultTitle, slResultLength)
                        If llRow + 1 > grdResults.Rows Then
                            grdResults.AddItem ""
                        End If
                        grdResults.Row = llRow
                        grdResults.TextMatrix(llRow, ITEMIDINDEX) = slCart
                        grdResults.TextMatrix(llRow, CHECKTITLEINDEX) = ""
                        grdResults.TextMatrix(llRow, CHECKLENINDEX) = ""
                        grdResults.TextMatrix(llRow, TYPEINDEX) = "Pri"
                        'slResult = Mid$(slResult, Len(slPriQuery) + 1)
                        grdResults.TextMatrix(llRow, RESULTTITLEINDEX) = slResultTitle
                        grdResults.TextMatrix(llRow, RESULTLENINDEX) = Left$(slResultLength, Len(slResultLength) - 3) & "." & Right$(slResultLength, 3)
                        llRow = llRow + 1
                        Exit For
                    End If
                Next ilITE
            End If
            If ckcProt.Value = vbChecked Then
                For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
                    If tgCurrITE(ilITE).sType = "S" Then
                        ''If StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0 Or _
                        ''   ((StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0) And (sgSpecialPassword <> "") And (Len(sgSpecialPassword) = 4)) Then
                        ''    slQuery = edcQuery(1).text
                        ''    slPriQuery = edcReply(1).text
                        ''    slCart = Mid$(slQuery, Len(slPriQuery) + 1)
                        ''Else
                            slCart = edcItemID.text
                        '    gBuildItemIDQuery slCart, tgCurrITE(ilITE), slQuery, slProtQuery
                        ''End If
                        'ilRet = gTestItemID(spcItemID, tgCurrITE(ilITE), slQuery, slProtQuery, slResult)
                        ilRet = gTestItemID(spcItemID, tgCurrITE(ilITE), slCart, ilDoConnectTest, slResultTitle, slResultLength)
                        If llRow + 1 > grdResults.Rows Then
                            grdResults.AddItem ""
                        End If
                        grdResults.Row = llRow
                        grdResults.TextMatrix(llRow, ITEMIDINDEX) = slCart
                        grdResults.TextMatrix(llRow, CHECKTITLEINDEX) = ""
                        grdResults.TextMatrix(llRow, CHECKLENINDEX) = ""
                        grdResults.TextMatrix(llRow, TYPEINDEX) = "Prot"
                        'slResult = Mid$(slResult, Len(slProtQuery) + 1)
                        grdResults.TextMatrix(llRow, RESULTTITLEINDEX) = slResultTitle
                        grdResults.TextMatrix(llRow, RESULTLENINDEX) = Left$(slResultLength, Len(slResultLength) - 3) & "." & Right$(slResultLength, 3)
                        llRow = llRow + 1
                        Exit For
                    End If
                Next ilITE
            End If
        End If
    Else
        If gIsDate(slStr) Then
            If sgItemIDDate <> "" Then
                If gDateValue(slStr) = gDateValue(sgItemIDDate) Then
                    mCheckItemIDArray tgItemIDChk()
                    'retain results so that the grid can be shown if user changes which item to show
                    ReDim tmItemIDChk(0 To UBound(tgItemIDChk)) As ITEMIDCHK
                    For ilLoop = 0 To UBound(tgItemIDChk) - 1 Step 1
                        LSet tmItemIDChk(ilLoop) = tgItemIDChk(ilLoop)
                    Next ilLoop
                Else
                    'Build SEE for date and check
                    mBuildItemIDbyDate
                    mCheckItemIDArray tmItemIDChk()
                    mUpdateSEE tmItemIDChk()
                End If
            Else
                'Build SEE for date and check
                mBuildItemIDbyDate
                mCheckItemIDArray tmItemIDChk()
                mUpdateSEE tmItemIDChk()
            End If
        End If
    End If
    'Screen.MousePointer = vbDefault
    gSetMousePointer grdResults, grdResults, vbDefault
End Sub

Private Sub edcChecksum_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDate_Change()
    If edcDate.text <> "" Then
        edcItemID.text = ""
        edcQuery(0).text = ""
        edcChecksum(0).text = ""
        edcQuery(1).text = ""
        edcChecksum(1).text = ""
        ckcShow.Visible = True
    End If
    mSetGridHeight
End Sub

Private Sub edcDate_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcItemID_Change()
    Dim slQuery As String
    Dim ilITE As Integer
    Dim slMessageNo As String
    Dim slMaxNo As String
    Dim slHex As String
    
    If edcItemID.text <> "" Then
        edcDate.text = ""
        ckcShow.Visible = False
    End If
    For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
        If tgCurrITE(ilITE).sType = "P" Then
            slQuery = ""
            slQuery = slQuery & Trim$(tgCurrITE(ilITE).sStartCode)
            slQuery = slQuery & Trim$(tgCurrITE(ilITE).sMachineID)
            slMessageNo = Trim$(Str$(tgCurrITE(ilITE).iCurrMgsID + 1))
            slMaxNo = Trim$(Str$(tgCurrITE(ilITE).iMaxMgsID))
            If Val(slMessageNo) > Val(slMaxNo) Then
                slMessageNo = Trim$(Str$(tgCurrITE(ilITE).iMinMgsID))
            End If
            Do While Len(slMessageNo) < Len(slMaxNo)
                slMessageNo = "0" & slMessageNo
            Loop
            slQuery = slQuery & slMessageNo
            slQuery = slQuery & Trim$(tgCurrITE(ilITE).sMgsType)
            edcQuery(0).text = slQuery & edcItemID.text
            slQuery = ""
            slQuery = slQuery & Trim$(tgCurrITE(ilITE).sReplyCode)
            slQuery = slQuery & Trim$(tgCurrITE(ilITE).sMachineID)
            slQuery = slQuery & slMessageNo
            slQuery = slQuery & Trim$(tgCurrITE(ilITE).sMgsType)
            edcReply(0).text = slQuery
        End If
        If tgCurrITE(ilITE).sType = "S" Then
            slQuery = ""
            slQuery = slQuery & Trim$(tgCurrITE(ilITE).sStartCode)
            slQuery = slQuery & Trim$(tgCurrITE(ilITE).sMachineID)
            slMessageNo = Trim$(Str$(tgCurrITE(ilITE).iCurrMgsID + 1))
            slMaxNo = Trim$(Str$(tgCurrITE(ilITE).iMaxMgsID))
            If Val(slMessageNo) > Val(slMaxNo) Then
                slMessageNo = Trim$(Str$(tgCurrITE(ilITE).iMinMgsID))
            End If
            Do While Len(slMessageNo) < Len(slMaxNo)
                slMessageNo = "0" & slMessageNo
            Loop
            slQuery = slQuery & slMessageNo
            slQuery = slQuery & Trim$(tgCurrITE(ilITE).sMgsType)
            edcQuery(1).text = slQuery & edcItemID.text
            slQuery = ""
            slQuery = slQuery & Trim$(tgCurrITE(ilITE).sReplyCode)
            slQuery = slQuery & Trim$(tgCurrITE(ilITE).sMachineID)
            slQuery = slQuery & slMessageNo
            slQuery = slQuery & Trim$(tgCurrITE(ilITE).sMgsType)
            edcReply(1).text = slQuery
        End If
    Next ilITE
    mSetGridHeight
End Sub

Private Sub edcItemID_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcQuery_Change(Index As Integer)
    Dim slQuery As String
    Dim slHex As String
    
    slQuery = edcQuery(Index).text
    slHex = gIntToHex(gCheckSum(slQuery), 2)
    edcChecksum(Index).text = slHex
End Sub

Private Sub edcQuery_GotFocus(Index As Integer)
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        Exit Sub
    End If
    imFirstActivate = False
    If imFirstTime Then
        imFirstTime = True
    End If
    Me.Refresh
End Sub
Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    gSetFonts EngrItemIDChk
    gCenterFormModal EngrItemIDChk
End Sub

Private Sub Form_Load()
    'Screen.MousePointer = vbHourglass
    gSetMousePointer grdResults, grdResults, vbHourglass
    mGridColumns
    mInit
    'Screen.MousePointer = vbDefault
    gSetMousePointer grdResults, grdResults, vbDefault
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
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
    Dim ilNoButtons As Integer
    Dim ilLoop As Integer
    Dim ilITE As Integer
    Dim ilRet As Integer
    
    imFirstActivate = True
    imFirstTime = True
    imLastColSorted = -1
    imLastSort = -1
    edcDate.text = sgItemIDDate
    hmSEE = CBtrvTable(TWOHANDLES) 'CBtrvObj()
    ilRet = btrOpen(hmSEE, "", sgDBPath & "SEE.Eng", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    mPopETE
    For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
        If tgCurrITE(ilITE).sType = "P" Then
            If Trim$(tgCurrITE(ilITE).sName) = "" Then
                ckcPri.Value = vbUnchecked
                ckcPri.Enabled = False
            End If
            Exit For
        End If
    Next ilITE
    For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
        If tgCurrITE(ilITE).sType = "S" Then
            If Trim$(tgCurrITE(ilITE).sName) = "" Then
                ckcProt.Value = vbUnchecked
                ckcProt.Enabled = False
            End If
            Exit For
        End If
    Next ilITE
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mTerminate                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
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
    'Screen.MousePointer = vbDefault
    gSetMousePointer grdResults, grdResults, vbDefault
    Unload EngrItemIDChk
    Set EngrItemIDChk = Nothing   'Remove data segment
End Sub

Private Sub Form_Resize()
    mGridColumnWidth
    mSetGridHeight
    gGrid_FillWithRows grdResults
End Sub

Private Sub Form_Unload(Cancel As Integer)
    btrDestroy hmSEE
    Erase tmCurrSEE
    Erase tmItemIDChk
    Erase smEndCodes
    
    Set EngrItemIDChk = Nothing
End Sub

Private Sub mGridColumns()
    Dim ilCol As Integer
    Dim ilRow As Integer
    
    gGrid_AlignAllColsLeft grdResults
    mGridColumnWidth
    'Set Titles
    grdResults.TextMatrix(0, ITEMIDINDEX) = "Item ID"
    grdResults.TextMatrix(0, CHECKTITLEINDEX) = "Title"
    grdResults.TextMatrix(0, CHECKLENINDEX) = "Length"
    grdResults.TextMatrix(0, TYPEINDEX) = "Pri/Prot"
    grdResults.TextMatrix(0, RESULTTITLEINDEX) = "Returned Title"
    grdResults.TextMatrix(0, RESULTLENINDEX) = "Returned Length"
    grdResults.Row = 1
    For ilCol = 0 To grdResults.Cols - 1 Step 1
        grdResults.Col = ilCol
        grdResults.CellAlignment = flexAlignLeftCenter
    Next ilCol
    mSetGridHeight
    gGrid_Clear grdResults, True
    grdResults.Row = grdResults.FixedRows
End Sub

Private Sub mGridColumnWidth()
    Dim ilCol As Integer
    
    grdResults.Width = EngrItemIDChk.Width - 2 * grdResults.Left
    grdResults.ColWidth(ITEMIDINDEX) = grdResults.Width / 8
    grdResults.ColWidth(CHECKTITLEINDEX) = grdResults.Width / 3
    grdResults.ColWidth(CHECKLENINDEX) = grdResults.Width / 16
    grdResults.ColWidth(TYPEINDEX) = grdResults.Width / 8
    grdResults.ColWidth(RESULTLENINDEX) = grdResults.Width / 8
    grdResults.ColWidth(RESULTTITLEINDEX) = grdResults.Width - GRIDSCROLLWIDTH
    For ilCol = ITEMIDINDEX To RESULTLENINDEX Step 1
        If ilCol <> RESULTTITLEINDEX Then
            If grdResults.ColWidth(RESULTTITLEINDEX) > grdResults.ColWidth(ilCol) Then
                grdResults.ColWidth(RESULTTITLEINDEX) = grdResults.ColWidth(RESULTTITLEINDEX) - grdResults.ColWidth(ilCol)
            Else
                Exit For
            End If
        End If
    Next ilCol
End Sub


Private Sub mSetGridHeight()
    If ((StrComp("COUNTERPOINT", StrConv(sgUserName, vbUpperCase)) = 0) And (Len(sgSpecialPassword) = 5)) Or _
       (((StrComp("GUIDE", StrConv(sgUserName, vbUpperCase)) = 0) And (Len(sgSpecialPassword) = 4))) Then
        'If edcItemID.text <> "" Then
        '    frcSendCommands.Visible = True
        '    frcSendCommands.Top = frcCheckDate.Top + frcCheckDate.Height + 120
        '    ckcShow.Top = frcSendCommands.Top + frcSendCommands.Height + 120
        '    grdResults.Height = cmcDone.Top - (ckcShow.Top + ckcShow.Height) - 240
        '    grdResults.Top = ckcShow.Top + ckcShow.Height + 120
        'Else
        '    frcSendCommands.Visible = False
            ckcShow.Top = frcCheckDate.Top + frcCheckDate.Height + 120
            grdResults.Height = cmcDone.Top - (ckcShow.Top + ckcShow.Height) - 240
            grdResults.Top = ckcShow.Top + ckcShow.Height + 120
        'End If
    Else
        'frcSendCommands.Visible = False
        ckcShow.Top = frcCheckDate.Top + frcCheckDate.Height + 120
        grdResults.Height = cmcDone.Top - (ckcShow.Top + ckcShow.Height) - 240
        grdResults.Top = ckcShow.Top + ckcShow.Height + 120
    End If
    gGrid_IntegralHeight grdResults

End Sub





Private Sub mBuildItemIDbyDate()
    Dim ilRet As Integer
    Dim llRow As Long
    Dim ilUpper As Integer
    Dim slItemID As String
    Dim slTitle As String
    Dim ilFound As Integer
    Dim slAirDate As String
    Dim ilLoop As Integer
    Dim ilSpot As Integer
    Dim ilETE As Integer
    Dim slTestItemID As String
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim ilATE As Integer
    
    ReDim tmItemIDChk(0 To 0) As ITEMIDCHK
    slAirDate = edcDate.text
    ilRet = gGetRec_SHE_ScheduleHeaderByDate(slAirDate, "EngrItemID Test-Get Schedule by Date", tmSHE)
    If ilRet Then
        ilRet = gGetRecs_SEE_ScheduleEventsAPI(hmSEE, smCurrSEEStamp, -1, tmSHE.lCode, "EngrItemIDChk-Get Events", tmCurrSEE())
        For llRow = 0 To UBound(tmCurrSEE) - 1 Step 1
            ilSpot = False
            For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                If tmCurrSEE(llRow).iEteCode = tgCurrETE(ilETE).iCode Then
                    If tgCurrETE(ilETE).sCategory = "S" Then
                        ilSpot = True
                    End If
                    Exit For
                End If
            Next ilETE
            If ilSpot Then
                slTestItemID = ""
                slTitle = ""
                'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                '    If tmCurrSEE(llRow).iAudioAseCode = tgCurrASE(ilASE).iCode Then
                    ilASE = gBinarySearchASE(tmCurrSEE(llRow).iAudioAseCode, tgCurrASE())
                    If ilASE <> -1 Then
                        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                        '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                            ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                            If ilANE <> -1 Then
                                For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
                                    If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                                        slTestItemID = tgCurrATE(ilATE).sTestItemID
                                        Exit For
                                    End If
                                Next ilATE
                        '        If slTestItemID <> "" Then
                        '            Exit For
                        '        End If
                            End If
                        'Next ilANE
                '        If slTestItemID <> "" Then
                '            Exit For
                '        End If
                    End If
                'Next ilASE
                If slTestItemID = "Y" Then
                    ilRet = gGetRec_ARE_AdvertiserRefer(tmCurrSEE(llRow).lAreCode, "EngrItemIDChk-mBuildItemIDbyDate: Advertiser", tmARE)
                    If ilRet Then
                        slTitle = Trim$(tmARE.sName)
                    End If
                End If
                If slTestItemID = "Y" Then
                    slItemID = Trim$(tmCurrSEE(llRow).sAudioItemID)
                    If slItemID <> "" Then
                        ilFound = False
                        For ilLoop = 0 To UBound(tmItemIDChk) - 1 Step 1
                            If (StrComp(slItemID, Trim$(tmItemIDChk(ilLoop).sItemID), vbTextCompare) = 0) And (tmItemIDChk(ilLoop).lLength = 100 * tmCurrSEE(llRow).lDuration) Then
                                ilFound = True
                                tmItemIDChk(ilLoop).sAudioStatus = "U"
                                Exit For
                            End If
                        Next ilLoop
                        If Not ilFound Then
                            ilUpper = UBound(tmItemIDChk)
                            tmItemIDChk(ilUpper).sItemID = slItemID
                            tmItemIDChk(ilUpper).sAudioStatus = "U"
                            tmItemIDChk(ilUpper).sProtStatus = "U"
                            tmItemIDChk(ilUpper).sTitle = slTitle
                            tmItemIDChk(ilUpper).lLength = 100 * tmCurrSEE(llRow).lDuration 'Return is in mil sec (30000)
                            tmItemIDChk(ilUpper).sPriResult = ""
                            tmItemIDChk(ilUpper).sPriLen = ""
                            tmItemIDChk(ilUpper).sProtResult = ""
                            tmItemIDChk(ilUpper).sProtLen = ""
                            tmItemIDChk(ilUpper).lSeeCode = tmCurrSEE(llRow).lCode
                            ReDim Preserve tmItemIDChk(0 To ilUpper + 1) As ITEMIDCHK
                        End If
                    End If
                End If
                slTestItemID = ""
                'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                '    If tmCurrSEE(llRow).iProtAneCode = tgCurrANE(ilANE).iCode Then
                    ilANE = gBinarySearchANE(tmCurrSEE(llRow).iProtAneCode, tgCurrANE())
                    If ilANE <> -1 Then
                        For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
                            If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                                slTestItemID = tgCurrATE(ilATE).sTestItemID
                                Exit For
                            End If
                        Next ilATE
                '        If slTestItemID <> "" Then
                '            Exit For
                '        End If
                    End If
                'Next ilANE
                If (slTestItemID = "Y") And (slTitle = "") Then
                    ilRet = gGetRec_ARE_AdvertiserRefer(tmCurrSEE(llRow).lAreCode, "EngrItemIDChk-mBuildItemIDbyDate: Advertiser", tmARE)
                    If ilRet Then
                        slTitle = Trim$(tmARE.sName)
                    End If
                End If
                If slTestItemID = "Y" Then
                    slItemID = Trim$(tmCurrSEE(llRow).sProtItemID)
                    If slItemID <> "" Then
                        ilFound = False
                        For ilLoop = 0 To UBound(tmItemIDChk) - 1 Step 1
                            If (StrComp(slItemID, Trim$(tmItemIDChk(ilLoop).sItemID), vbTextCompare) = 0) And (tmItemIDChk(ilLoop).lLength = 100 * tmCurrSEE(llRow).lDuration) Then
                                ilFound = True
                                tmItemIDChk(ilLoop).sProtStatus = "U"
                                Exit For
                            End If
                        Next ilLoop
                        If Not ilFound Then
                            ilUpper = UBound(tmItemIDChk)
                            tmItemIDChk(ilUpper).sItemID = slItemID
                            tmItemIDChk(ilUpper).sAudioStatus = ""
                            tmItemIDChk(ilUpper).sProtStatus = "U"
                            tmItemIDChk(ilUpper).sTitle = slTitle
                            tmItemIDChk(ilUpper).lLength = 100 * tmCurrSEE(llRow).lDuration 'Return is in mil sec (30000)
                            tmItemIDChk(ilUpper).sPriResult = ""
                            tmItemIDChk(ilUpper).sPriLen = ""
                            tmItemIDChk(ilUpper).sProtResult = ""
                            tmItemIDChk(ilUpper).sProtLen = ""
                            tmItemIDChk(ilUpper).lSeeCode = tmCurrSEE(llRow).lCode
                            ReDim Preserve tmItemIDChk(0 To ilUpper + 1) As ITEMIDCHK
                        End If
                    End If
                End If
            End If
        Next llRow
    End If

End Sub

Private Sub mCheckItemIDArray(tlItemIDChk() As ITEMIDCHK)
    Dim ilLoop As Integer
    Dim ilITE As Integer
    Dim slCart As String
    Dim slResultTitle As String
    Dim slResultLength As String
    Dim slTitle As String
    Dim llRow As Long
    Dim slPriQuery As String
    Dim slProtQuery As String
    Dim slQuery As String
    Dim ilRet As Integer
    Dim ilTestPort As Integer
    Dim ilDoConnectTest As Integer
    Dim llResultLength As Long
    Dim tlPriITE As ITE
    Dim tlSecITE As ITE
    
    tlPriITE.sType = ""
    For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
        If tgCurrITE(ilITE).sType = "P" Then
            LSet tlPriITE = tgCurrITE(ilITE)
            Exit For
        End If
    Next ilITE
    tlSecITE.sType = ""
    For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
        If tgCurrITE(ilITE).sType = "S" Then
            LSet tlSecITE = tgCurrITE(ilITE)
            Exit For
        End If
    Next ilITE
    ilTestPort = True
    ilDoConnectTest = True
    llRow = grdResults.FixedRows
    For ilLoop = 0 To UBound(tlItemIDChk) - 1 Step 1
        slTitle = gFileNameFilter(Trim$(tlItemIDChk(ilLoop).sTitle))    'Trim$(tlItemIDChk(ilLoop).sTitle)
        If (tlItemIDChk(ilLoop).sAudioStatus = "U") Then
            If (tlPriITE.sType = "P") And (ilTestPort) Then
                slCart = Trim$(tlItemIDChk(ilLoop).sItemID)
                'gBuildItemIDQuery slCart, tlPriITE, slQuery, slPriQuery
                'ilRet = gTestItemID(spcItemID, tlPriITE, slQuery, slPriQuery, slResult)
                ilRet = gTestItemID(spcItemID, tlPriITE, slCart, ilDoConnectTest, slResultTitle, slResultLength)
                If ilRet Then
                    'slResult = Mid$(slResult, Len(slPriQuery) + 1)
                    llResultLength = Val(slResultLength)
                    tlItemIDChk(ilLoop).sPriResult = slResultTitle
                    tlItemIDChk(ilLoop).sPriLen = slResultLength
                    If (StrComp(Trim$(slTitle), Trim$(slResultTitle), vbTextCompare) = 0) And ((tlItemIDChk(ilLoop).lLength - tgSOE.iSpotItemIDWindow <= llResultLength) And ((tlItemIDChk(ilLoop).lLength + tgSOE.iSpotItemIDWindow >= llResultLength))) Then
                        tlItemIDChk(ilLoop).sAudioStatus = "O"
                    Else
                        tlItemIDChk(ilLoop).sAudioStatus = "F"
                    End If
                Else
                    If StrComp(slResultTitle, "Failed", vbTextCompare) = 0 Then
                        ilTestPort = False
                        tlItemIDChk(ilLoop).sAudioStatus = "N"
                    Else
                        tlItemIDChk(ilLoop).sAudioStatus = "F"
                    End If
                End If
                If (ckcPri.Value = vbChecked) Then
                    If (tlItemIDChk(ilLoop).sAudioStatus = "F") Or (ckcShow.Value = vbUnchecked) Then
                        If llRow + 1 > grdResults.Rows Then
                            grdResults.AddItem ""
                        End If
                        grdResults.Row = llRow
                        grdResults.TextMatrix(llRow, ITEMIDINDEX) = slCart
                        grdResults.TextMatrix(llRow, CHECKTITLEINDEX) = Trim$(tlItemIDChk(ilLoop).sTitle) '& " " & slTitle 'Trim$(tlItemIDChk(ilLoop).sTitle)
                        grdResults.TextMatrix(llRow, CHECKLENINDEX) = Trim$(Str$(tlItemIDChk(ilLoop).lLength \ 1000))
                        grdResults.TextMatrix(llRow, TYPEINDEX) = "Pri"
                        grdResults.TextMatrix(llRow, RESULTTITLEINDEX) = slResultTitle
                        grdResults.TextMatrix(llRow, RESULTLENINDEX) = Left$(slResultLength, Len(slResultLength) - 3) & "." & Right$(slResultLength, 3)
                        llRow = llRow + 1
                    End If
                End If
                ilDoConnectTest = False
            End If
        End If
        If (tlItemIDChk(ilLoop).sProtStatus = "U") Then
            If (tlPriITE.sType = "P") And (ilTestPort) Then
                slCart = Trim$(tlItemIDChk(ilLoop).sItemID)
                'gBuildItemIDQuery slCart, tlSecITE, slQuery, slProtQuery
                'ilRet = gTestItemID(spcItemID, tlSecITE, slQuery, slProtQuery, slResult)
                ilRet = gTestItemID(spcItemID, tlPriITE, slCart, ilDoConnectTest, slResultTitle, slResultLength)
                If ilRet Then
                '    slResult = Mid$(slResult, Len(slProtQuery) + 1)
                    llResultLength = Val(slResultLength)
                    tlItemIDChk(ilLoop).sProtResult = slResultTitle
                    tlItemIDChk(ilLoop).sProtLen = slResultLength
                    If (StrComp(Trim$(slTitle), Trim$(slResultTitle), vbTextCompare) = 0) And ((tlItemIDChk(ilLoop).lLength - tgSOE.iSpotItemIDWindow <= llResultLength) And ((tlItemIDChk(ilLoop).lLength + tgSOE.iSpotItemIDWindow >= llResultLength))) Then
                        tlItemIDChk(ilLoop).sProtStatus = "O"
                    Else
                        tlItemIDChk(ilLoop).sProtStatus = "F"
                    End If
                Else
                    If StrComp(slResultTitle, "Failed", vbTextCompare) = 0 Then
                        ilTestPort = False
                        tlItemIDChk(ilLoop).sProtStatus = "N"
                    Else
                        tlItemIDChk(ilLoop).sProtStatus = "F"
                    End If
                End If
                If (ckcProt.Value = vbChecked) Then
                    If (tlItemIDChk(ilLoop).sProtStatus = "F") Or (ckcShow.Value = vbUnchecked) Then
                        If llRow + 1 > grdResults.Rows Then
                            grdResults.AddItem ""
                        End If
                        grdResults.Row = llRow
                        grdResults.TextMatrix(llRow, ITEMIDINDEX) = slCart
                        grdResults.TextMatrix(llRow, CHECKTITLEINDEX) = Trim$(tlItemIDChk(ilLoop).sTitle) '& " " & slTitle 'Trim$(tlItemIDChk(ilLoop).sTitle)
                        grdResults.TextMatrix(llRow, CHECKLENINDEX) = Trim$(Str$(tlItemIDChk(ilLoop).lLength \ 1000))
                        grdResults.TextMatrix(llRow, TYPEINDEX) = "Prot"
                        grdResults.TextMatrix(llRow, RESULTTITLEINDEX) = slResultTitle
                        grdResults.TextMatrix(llRow, RESULTLENINDEX) = Left$(slResultLength, Len(slResultLength) - 3) & "." & Right$(slResultLength, 3)
                        llRow = llRow + 1
                    End If
                End If
                ilDoConnectTest = False
            End If
        End If
    Next ilLoop
End Sub

Private Sub mPopETE()
    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "EngrLibETE-mPopETE Event Types", tgCurrETE())
    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrItemIDChk-mPopETE Event Properties", tgCurrEPE())
End Sub

Private Sub mPopASE()
    Dim ilRet As Integer

    mPopANE
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrItemIDChk-mPopASE Audio Source", tgCurrASE())
End Sub

Private Sub mPopANE()

    Dim ilRet As Integer

    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrItemIDChk-mPopASE Audio Audio Names", tgCurrANE())
End Sub


Private Sub mUpdateSEE(tlItemIDChk() As ITEMIDCHK)
    Dim ilLoop As Integer
    Dim slAudioItemIDChk As String
    Dim slProtItemIDChk As String
    Dim ilRet As Integer
    
    For ilLoop = 0 To UBound(tlItemIDChk) - 1 Step 1
        slAudioItemIDChk = ""
        If (tlItemIDChk(ilLoop).sAudioStatus = "O") Or (tlItemIDChk(ilLoop).sAudioStatus = "F") Then
            slAudioItemIDChk = tlItemIDChk(ilLoop).sAudioStatus
        End If
        slProtItemIDChk = ""
        If (tlItemIDChk(ilLoop).sProtStatus = "O") Or (tlItemIDChk(ilLoop).sProtStatus = "F") Then
            slProtItemIDChk = tlItemIDChk(ilLoop).sProtStatus
        End If
        If (slAudioItemIDChk <> "") Or (slProtItemIDChk <> "") Then
            ilRet = gPutUpdate_SEE_ItemIDCheck(tlItemIDChk(ilLoop).lSeeCode, slAudioItemIDChk, slProtItemIDChk, "EngrItemIDChk-Update SEE Item ID Status")
        End If
    Next ilLoop

End Sub

Private Sub grdResults_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If y < grdResults.RowHeight(0) Then
        mSortCol grdResults.Col
        Exit Sub
    End If
End Sub

Private Sub imcPrint_Click()
Dim ilRptDest As Integer            'disply, print, save as file
Dim slRptName As String
Dim slExportName As String
Dim slRptType As String
Dim llResult As Long
Dim ilExportType As Integer
Dim llGridRow As Integer
Dim slStr As String

    igRptIndex = ITEMIDCHK_RPT
    ilRptDest = 1                   'force to display for debugging, else force to Print
    slExportName = ""               'no export for now
    slRptType = ""
    
    Set rstItemIDRpt = New Recordset
    gGenerateItemIDRst     'generate the ddfs for report
    
    rstItemIDRpt.Open
    'build the data definition (.ttx) file in the database path for crystal to access
    llResult = CreateFieldDefFile(rstItemIDRpt, sgDBPath & "\ItemID.ttx", True)
    
    'loop thru the ItemID grid to print whats shown on the screen
    For llGridRow = grdResults.FixedRows To grdResults.Rows - 1
        slStr = Trim$(grdResults.TextMatrix(llGridRow, ITEMIDINDEX))
        If slStr = "" Then
            Exit For
        Else
            rstItemIDRpt.AddNew
            rstItemIDRpt.Fields("ItemID") = grdResults.TextMatrix(llGridRow, ITEMIDINDEX)
            rstItemIDRpt.Fields("Title") = grdResults.TextMatrix(llGridRow, CHECKTITLEINDEX)
            rstItemIDRpt.Fields("PrimProt") = grdResults.TextMatrix(llGridRow, TYPEINDEX)
            rstItemIDRpt.Fields("ReturnTitle") = grdResults.TextMatrix(llGridRow, RESULTTITLEINDEX)
        End If
    Next llGridRow
    
    'gObtainReportforCrystal slRptName, slExportName     'determine which .rpt to call and setup an export name is user selected output to export
    igRptSource = vbModal
    slRptName = "ItemID.rpt"      'concatenate the crystal report name plus extension
    
    'set up the formulas to pass to Crystal
    'Selected by date or ItemID
    slStr = Trim(edcDate.text)
    If slStr = "" Then
        'by Item ID
        sgCrystlFormula2 = "'" & Trim$(edcItemID) & "'"
    Else
        'slStr = "Date(" + Format$(slStr, "yyyy") + "," + Format$(slStr, "mm") + "," + Format$(slStr, "dd") + ")"
        sgCrystlFormula2 = "'" & Format$(slStr, "ddddd") & "'"
    End If
    If ckcShow.Value = vbChecked Then           'discrep only
        sgCrystlFormula3 = "'Y'"
    Else
        sgCrystlFormula3 = "'N'"
    End If
    
    EngrCrystal.gActiveCrystalReports ilExportType, ilRptDest, Trim$(slRptName) & Trim$(slRptType), slExportName, rstItemIDRpt
    
    'Screen.MousePointer = vbDefault
    gSetMousePointer grdResults, grdResults, vbDefault
    
    Set rstItemIDRpt = Nothing
    'If igRptSource = vbModal Then
    '    Unload EngrItemIDChk
    'End If
    Exit Sub

End Sub
Private Sub mSortCol(ilCol As Integer)
    gGrid_SortByCol grdResults, ITEMIDINDEX, ilCol, imLastColSorted, imLastSort
End Sub
