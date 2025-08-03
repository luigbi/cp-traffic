VERSION 5.00
Begin VB.Form PostCpmManualCntrl 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5040
   ClientLeft      =   10230
   ClientTop       =   7470
   ClientWidth     =   9390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbcEnd 
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   8
      Top             =   4320
      Width           =   15
   End
   Begin VB.TextBox edcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.PictureBox pbcPostManualCntrl 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   200
      Picture         =   "PostCpmManualCntr.frx":0000
      ScaleHeight     =   3495
      ScaleWidth      =   8535
      TabIndex        =   11
      Top             =   560
      Width           =   8535
      Begin VB.Label lacFrame 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         TabIndex        =   13
         Top             =   500
         Visible         =   0   'False
         Width           =   8535
      End
   End
   Begin VB.PictureBox plcPostManualCntrl 
      Height          =   3640
      Left            =   120
      ScaleHeight     =   3541.102
      ScaleMode       =   0  'User
      ScaleWidth      =   8895
      TabIndex        =   12
      Top             =   480
      Width           =   8960
      Begin VB.VScrollBar vbcPostManualCntrl 
         Height          =   3604
         LargeChange     =   12
         Left            =   8640
         Min             =   1
         TabIndex        =   10
         Top             =   0
         Value           =   1
         Width           =   240
      End
   End
   Begin VB.PictureBox pbcArrow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "PostCpmManualCntr.frx":817C8
      ScaleHeight     =   180
      ScaleWidth      =   105
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1100
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   240
      Top             =   4200
   End
   Begin VB.CommandButton cmcSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5820
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   -30
      ScaleHeight     =   45
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   285
      Width           =   60
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   60
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   4860
      Width           =   45
   End
   Begin VB.CommandButton cmcCancel 
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
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.PictureBox pbcClickFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   45
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4725
      Width           =   45
   End
   Begin VB.CommandButton cmcDone 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label plcMegaphoneLastUpdated 
      Caption         =   "Last Megaphone Import Date:  12/12/12"
      Height          =   225
      Left            =   7500
      TabIndex        =   14
      Top             =   4530
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Label plcScreen 
      Caption         =   "Ad Server Impression Posting for:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   200
      TabIndex        =   0
      Top             =   30
      Width           =   2970
   End
End
Attribute VB_Name = "PostCpmManualCntrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of PostCpmManualCntrl.FRM on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  lmRowSelected                 tmGhfSrchKey0                 tmGsfSrchKey0             *
'*  tmIhfSrchKey1                 tmIhfSrchKey2                 tmIsfSrchKey0             *
'*  tmIsfSrchKey1                 tmIsfSrchKey2                                           *
'******************************************************************************************

'******************************************************
'*  PostCpmManualCntrl - displays missed spots to be changed to Makegoods
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFirstTime As Integer
Private imBSMode As Integer
Private imMouseDown As Integer
Private imTerminate As Integer
Private lmScrollTop As Long
Private lmEnableRow As Long
Private lmEnableCol As Long

Private imChg As Integer
Private imManualPostColSorted As Integer

Private imManualPostSort As Integer
Private imFromArrow As Integer
Private lmTopRow As Long

Dim imBoxNo As Integer   'Current event name Box
Dim imRowNo As Integer
Dim fmAdjFactorW As Single  'Width adjustment factor
Dim fmAdjFactorH As Single  'Width adjustment factor
Dim imLBCtrls As Integer
Dim imLBPost As Integer
Dim imLoaded As Integer

Dim lServerVendor As Long
Dim sContractNumber As String
Dim sAdvertiserCode As String
Dim iYear As Single
Dim iMonth As Single
Dim sMonthName As String
Dim daysInMonth As Integer
Dim imCurSort As Integer    '0=Sort by Contract Number; 1=Sort by Advertiser Name
Dim imSortOrder As Integer  '0 = Ascending; 1 = Descending

Dim tgCPMPostLog() As PostCpmManual
Dim tgTempCPMPostLog() As PostCpmManual
'Rate Card Field Areas
Dim tmPostCtrls(0 To 14)  As FIELDAREA

Const CONTRACTINDEX = 1
Const ADVERTISERINDEX = 2
Const IDINDEX = 3
Const VEHICLEINDEX = 4
Const ADLOCATIONINDEX = 5
Const LENGTHINDEX = 6 ' Added 11-20-2023 JJB for Booster
Const DATERANGEINDEX = 7
Const PRICETYPEINDEX = 8
Const CPMINDEX = 9
'Impressions
Const ORDEREDINDEX = 10
Const BILLEDINDEX = 11
Const TOTALINVOICEDINDEX = 12
Const UNBILLEDINDEX = 13 'Posted Index
Const MONTHCOSTINDEX = 14

Dim smPLShow() As String * 40
Dim imPlSave() As Integer

Private Sub cmcCancel_Click()
    mTerminate
End Sub

Private Sub cmcCancel_GotFocus()
    If imLoaded Then
        mSetShow imBoxNo
        imRowNo = UBound(tgCPMPostLog)
    End If
    gCtrlGotFocus cmcCancel
    imBoxNo = -1
End Sub

Private Sub cmcDone_Click()
    If Not imUpdateAllowed Then
        cmcCancel_Click
        Exit Sub
    End If
    
    If mSaveRecChg(True) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        Exit Sub
    End If
    mTerminate
End Sub

Private Sub cmcDone_GotFocus()
    If imLoaded Then
        mSetShow imBoxNo
        imRowNo = UBound(tgCPMPostLog) - 1
    End If
    gCtrlGotFocus cmcCancel
    imBoxNo = -1
End Sub


Function mBilledMonthCount(iContractNumber As Long, iCPM_ID As Integer) As Long
'******************************************************************************************
'* Note: This returns the number of months already billed so that a true calculation      *
'*       of month cost can be made.                                                       *
'*                                                                                        *
'* Created 2023.02.08 - JJB                                                               *
'* TTP 10579                                                                              *
'******************************************************************************************

    Dim rs As ADODB.Recordset
    
    SQLQuery = ""
    SQLQuery = SQLQuery & "select "
    SQLQuery = SQLQuery & "     count(*) as iCount "
    SQLQuery = SQLQuery & "from "
    SQLQuery = SQLQuery & "     ibf_Impression_Bill "
    SQLQuery = SQLQuery & "where "
    SQLQuery = SQLQuery & "         ibfBilled   = 'Y' "
    SQLQuery = SQLQuery & "     and ibfCntrNo   = " & str(iContractNumber)
    SQLQuery = SQLQuery & "     and ibfPodCPMID = " & str(iCPM_ID)
   
    Set rs = gSQLSelectCall(SQLQuery)
    
    If Not rs.EOF Then
        mBilledMonthCount = rs!iCount
    Else
        mBilledMonthCount = 0
    End If
    
    rs.Close
    Set rs = Nothing
   
End Function


'*******************************************************
'*                                                     *
'*      Procedure Name:mSaveRecChg                     *
'*                                                     *
'*             Created:05/03/21       By:L.Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if record altered and*
'*                      requires updating              *
'*                                                     *
'*******************************************************
Private Function mSaveRecChg(ilAsk As Integer) As Integer
'
'   iAsk = True
'   iRet = mSaveRecChg(iAsk)
'   Where:
'       iAsk (I)- True = Ask if changed records should be updated;
'                 False= Update record if required without asking user
'       iRet (O)- True if updated or added, False if not updated or added
'
    Dim ilRes As Integer
    Dim slMess As String
    Dim ilAltered As Integer
    
    ilAltered = gAnyLogChgd()
    If ilAltered = YES Then
            If ilAsk Then
                slMess = "Save Changes Logs for " & sMonthName & " " & iYear
                ilRes = MsgBox(slMess, vbYesNoCancel + vbQuestion, "Update")
                If ilRes = vbCancel Then
                    mSaveRecChg = False
                    Exit Function
                End If
                If ilRes = vbYes Then
                    ilRes = mSaveRec()
                    mSaveRecChg = ilRes
                    Exit Function
                End If
            Else
                ilRes = mSaveRec()
                mSaveRecChg = ilRes
                Exit Function
            End If
    End If
    mSaveRecChg = True
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gAnyLogChgd                     *
'*                                                     *
'*             Created:05/03/21       By:L.Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Check any row log changed      *
'*                                                     *
'*******************************************************
Public Function gAnyLogChgd() As Integer
  Dim llRowNo As Long
    For llRowNo = imLBPost To UBound(tgCPMPostLog) Step 1
        If Not imPlSave(1, llRowNo) Then
            If imPlSave(0, llRowNo) Then
                gAnyLogChgd = True
                Exit Function
            End If
        End If
    Next llRowNo
    gAnyLogChgd = False
End Function

Private Sub cmcSave_Click()
    Dim ilRet As Integer
    Screen.MousePointer = vbHourglass
     If mSaveRecChg(False) = False Then
        If imTerminate Then
            cmcCancel_Click
            Exit Sub
        End If
        Exit Sub
    End If
    imBoxNo = -1
    mPopCpmLog
    pbcPostManualCntrl.Cls
    pbcPostManualCntrl_Paint
    mSetCommands
    pbcSTab.SetFocus
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmcSave_GotFocus()
    If imLoaded Then
        mSetShow imBoxNo
        imRowNo = UBound(tgCPMPostLog)
    End If
    gCtrlGotFocus cmcSave
    imBoxNo = -1
End Sub

Private Sub edcDropDown_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcDropDown_KeyPress(KeyAscii As Integer)
    Dim slStr As String
    Dim slComp As String
        'Filter characters (allow only BackSpace, numbers 0 thru 9, Decimal point (1 only)
        If (KeyAscii <> KEYBACKSPACE) And ((KeyAscii < KEY0) Or (KeyAscii > KEY9)) Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
        slStr = edcDropDown.Text
        slStr = Left$(slStr, edcDropDown.SelStart) & Chr$(KeyAscii) & right$(slStr, Len(slStr) - edcDropDown.SelStart - edcDropDown.SelLength)
        slComp = "999999999"
        If gCompNumberStr(slStr, slComp) > 0 Then
            Beep
            KeyAscii = 0
            Exit Sub
        End If
    
End Sub

Private Sub Form_Activate()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCol                                                                                 *
'******************************************************************************************
    If Not imFirstActivate Then
        DoEvents    'Process events so pending keys are not sent to this
                    'form when keypreview turn on
        gShowBranner imUpdateAllowed
        Me.KeyPreview = True
        Exit Sub
    End If
    
    imFirstActivate = False
    
    If (igWinStatus(POSTLOGSJOB) = 1) And (Trim$(tgUrf(0).sName) <> sgCPName) And (Trim$(tgUrf(0).sName) <> sgSUName) Then
        imUpdateAllowed = False
    Else
        imUpdateAllowed = True
    End If
    mPopCpmLog
End Sub

Private Sub Form_Click()
    pbcClickFocus.SetFocus
End Sub

Private Sub Form_Initialize()
    'Me.Width = (CLng(75) * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
    'Me.Height = (CLng(75) * ((Screen.Height) / (480 * 15 / Me.Height))) / 100
    'gCenterStdAlone PostCpmManualCntrl
    'DoEvents
    'mSetControls
End Sub

Private Sub Form_Load()
   If Screen.Width * 15 = 640 Then
        fmAdjFactorW = 1#
        fmAdjFactorH = 1#
    Else
        fmAdjFactorW = ((lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100) / Me.Width
        Me.Width = (lgPercentAdjW * ((Screen.Width) / (640 * 15 / Me.Width))) / 100
        fmAdjFactorH = ((lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100) / Me.height
        Me.height = (lgPercentAdjH * ((Screen.height) / (480 * 15 / Me.height))) / 100
    End If
    mInit
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    plcMegaphoneLastUpdated.Left = plcPostManualCntrl.Left + plcPostManualCntrl.Width - plcMegaphoneLastUpdated.Width
    plcMegaphoneLastUpdated.Top = cmcSave.Top
End Sub

Private Sub Form_Terminate()
    Dim ilRet As Integer
    On Error Resume Next
    sgManualPostCpmFilter = ""
    Set PostCpmManualCntrl = Nothing   'Remove data segment
    'igManUnload = NO
    
End Sub

Private Sub mInit()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llRow As Long
    Dim slStr As String
    
    imTerminate = False
    imFirstActivate = True
    imMouseDown = False
    imFirstTime = True
    imBSMode = False
    imFromArrow = False
    imManualPostColSorted = -1
    imManualPostSort = -1
    imLBCtrls = 1
    imBoxNo = -1
    imLoaded = 0
    imLBPost = 1
    imChg = False
    imCurSort = -1
    'ReDim tgCPMPostLog(0) As PostCpmManual
    
    ilRet = gParseItem(sgManualPostCpmFilter, 1, "\", slStr)
    lServerVendor = CLng(slStr)
    ilRet = gParseItem(sgManualPostCpmFilter, 2, "\", slStr)
    iYear = CSng(slStr)
    ilRet = gParseItem(sgManualPostCpmFilter, 3, "\", slStr)
    iMonth = CSng(slStr)
    ilRet = gParseItem(sgManualPostCpmFilter, 4, "\", slStr)
    sMonthName = slStr
    ilRet = gParseItem(sgManualPostCpmFilter, 5, "\", slStr)
    sAdvertiserCode = slStr
    ilRet = gParseItem(sgManualPostCpmFilter, 6, "\", slStr)
    sContractNumber = slStr
    
    If lServerVendor = 1 Then
        plcMegaphoneLastUpdated = "Last Megaphone Import Date: " & mGetMegaphoneLastImportDate
        plcMegaphoneLastUpdated.Visible = True
    End If
        
    mInitBox
    gCenterForm Me
    
    If sgManualPostCpmFilter = "" Then
        imTerminate = True
        Exit Sub
    End If
    
    mSetCommands
    Screen.MousePointer = vbDefault
    'gSetMousePointer grdManualPost, grdManualPost, vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Screen.MousePointer = vbDefault
    'gSetMousePointer grdManualPost, grdManualPost, vbDefault
    Exit Sub


End Sub

Private Function mGetMegaphoneLastImportDate() As String
   'JJB 03/15/2024
   
    Dim rs As ADODB.Recordset
    
    SQLQuery = ""
    SQLQuery = SQLQuery & " Select "
    SQLQuery = SQLQuery & "     avfLastImportDate "
    SQLQuery = SQLQuery & " from "
    SQLQuery = SQLQuery & "     avf_AdVendor "
    SQLQuery = SQLQuery & " where "
    SQLQuery = SQLQuery & "     avfCode = 1 "
    
    Set rs = gSQLSelectCall(SQLQuery)
    
    If Not rs.EOF Then
        If IsNull(rs!avfLastImportDate) Then
            mGetMegaphoneLastImportDate = "NEVER"
        Else
            mGetMegaphoneLastImportDate = rs!avfLastImportDate
        End If
    End If

End Function

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
    Dim ilRet As Integer
    Screen.MousePointer = vbDefault
    Erase tgCPMPostLog
    Erase tmPostCtrls
    Erase smPLShow
    Erase imPlSave
    imBoxNo = -1
    imRowNo = 1
    imCurSort = -1
    imSortOrder = 0
    sgManualPostCpmFilter = ""
    imSettingValue = False
    Unload PostCpmManualCntrl
    
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mPopCpmLog                        *
'*                                                     *
'*             Created:2/16/21       By: L.Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'***************************************************
Private Sub mPopCpmLog()

 Dim ilRet As Integer
 Dim llLoop As Long
 Dim slStr As String
 Dim slStr1 As String
 
 ilRet = mReadRec()
 
 For llLoop = imLBPost To UBound(tgTempCPMPostLog)
    ilRet = gBinarySearchAdf(tgTempCPMPostLog(llLoop).iAdfCode)
    If ilRet <> -1 Then
        tgTempCPMPostLog(llLoop).sAdfName = tgCommAdf(ilRet).sName
    End If
    
    ilRet = gBinarySearchVef(tgTempCPMPostLog(llLoop).iVehCode)
    If ilRet <> -1 Then
        tgTempCPMPostLog(llLoop).sVehName = tgMVef(ilRet).sName
    End If
    
    ilRet = gBinarySearchRdf(tgTempCPMPostLog(llLoop).iRdfCode)
    If ilRet <> -1 Then
        tgTempCPMPostLog(llLoop).sAdLocation = tgMRdf(ilRet).sName
    End If
    
    slStr = tgTempCPMPostLog(llLoop).lContractNumber
    Do While Len(slStr) < 5
        slStr = "0" & slStr
    Loop
    
    slStr1 = tgTempCPMPostLog(llLoop).iID
    Do While Len(slStr1) < 5
        slStr1 = "0" & slStr1
    Loop
    tgTempCPMPostLog(llLoop).sKey = slStr & "|" & slStr1
    tgTempCPMPostLog(llLoop).sKey = tgTempCPMPostLog(llLoop).lContractNumber & "|" & tgTempCPMPostLog(llLoop).iID
    
 Next llLoop
 
 imCurSort = 0
 imSortOrder = 0
 If UBound(tgTempCPMPostLog) > 0 Then
    ArraySortTyp fnAV(tgTempCPMPostLog(), 1), UBound(tgTempCPMPostLog), 0, LenB(tgTempCPMPostLog(1)), 0, LenB(tgTempCPMPostLog(1).sKey), 0
 End If
 
 ReDim tgCPMPostLog(0 To UBound(tgTempCPMPostLog)) As PostCpmManual
 For llLoop = imLBPost To UBound(tgTempCPMPostLog)
    tgCPMPostLog(llLoop) = tgTempCPMPostLog(llLoop)
 Next llLoop
 
 mInitPostCpmCtrls
 mMoveRecToCtrl
 vbcPostManualCntrl.Value = vbcPostManualCntrl.Min
 If UBound(tgCPMPostLog) <= vbcPostManualCntrl.LargeChange + 1 Then
        vbcPostManualCntrl.Max = vbcPostManualCntrl.Min
 Else
        vbcPostManualCntrl.Max = UBound(tgCPMPostLog) - vbcPostManualCntrl.LargeChange + 1    'Show one extra line
 End If
End Sub

Private Sub mInitPostCpmCtrls()
    Dim llUpperBound As Long
    Dim ilLoop As Integer
    Dim llIndex As Long
    
    llUpperBound = UBound(tgCPMPostLog)
    ReDim smPLShow(0 To MONTHCOSTINDEX, 0 To llUpperBound) As String * 40 'Values shown in program area
    ReDim imPlSave(0 To 1, 0 To llUpperBound) As Integer 'Values saved (program name) in program area 0 = ValueCHanged, 1 = ReadOnly Status
    For ilLoop = LBound(smPLShow, 1) To UBound(smPLShow, 1) Step 1
        For llIndex = LBound(smPLShow, 2) To UBound(smPLShow, 2) Step 1
            smPLShow(ilLoop, llIndex) = ""
        Next llIndex
    Next ilLoop
End Sub

Private Sub mMoveRecToCtrl()
    Dim llRowNo As Long
    Dim slStr As String
    Dim monthCost As Currency
    Dim ilLoop As Integer
    Dim llPosted As Long
    
   
    For llRowNo = imLBPost To UBound(tgCPMPostLog) Step 1
    
        imPlSave(1, llRowNo) = mReadonlyRow(llRowNo)
        If Not imPlSave(1, llRowNo) Then
            imPlSave(0, llRowNo) = 0
        End If
        
        slStr = tgCPMPostLog(llRowNo).lContractNumber
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(CONTRACTINDEX)
        smPLShow(CONTRACTINDEX, llRowNo) = tmPostCtrls(CONTRACTINDEX).sShow
        
        slStr = tgCPMPostLog(llRowNo).sAdfName
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(ADVERTISERINDEX)
        smPLShow(ADVERTISERINDEX, llRowNo) = tmPostCtrls(ADVERTISERINDEX).sShow
        
        slStr = tgCPMPostLog(llRowNo).sVehName
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(VEHICLEINDEX)
        smPLShow(VEHICLEINDEX, llRowNo) = tmPostCtrls(VEHICLEINDEX).sShow
        
        slStr = tgCPMPostLog(llRowNo).sAdLocation
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(ADLOCATIONINDEX)
        smPLShow(ADLOCATIONINDEX, llRowNo) = tmPostCtrls(ADLOCATIONINDEX).sShow
          
        ' Added 11-20-2023 JJB for Booster
        slStr = tgCPMPostLog(llRowNo).Length ' ??
        slStr = IIF(slStr = "0", "", slStr) ' If zero then display blank
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(LENGTHINDEX)
        smPLShow(LENGTHINDEX, llRowNo) = tmPostCtrls(LENGTHINDEX).sShow
        '''
        
        slStr = tgCPMPostLog(llRowNo).iID
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(IDINDEX)
        smPLShow(IDINDEX, llRowNo) = tmPostCtrls(IDINDEX).sShow
        
        slStr = Format$(tgCPMPostLog(llRowNo).StartDate, "m/d/yy") & "-" & Format$(tgCPMPostLog(llRowNo).EndDate, "m/d/yy")
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(DATERANGEINDEX)
        smPLShow(DATERANGEINDEX, llRowNo) = tmPostCtrls(DATERANGEINDEX).sShow
        
        If tgCPMPostLog(llRowNo).sPriceType = "C" Then
            slStr = "CPM"
        ElseIf tgCPMPostLog(llRowNo).sPriceType = "F" Then
            slStr = "Flat Rate"
        Else
            slStr = ""
        End If
        
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(PRICETYPEINDEX)
        smPLShow(PRICETYPEINDEX, llRowNo) = tmPostCtrls(PRICETYPEINDEX).sShow
        
        If tgCPMPostLog(llRowNo).lCPM > 0 Then
            slStr = gLongToStrDec(tgCPMPostLog(llRowNo).lCPM, 2)
        Else
            slStr = ""
        End If
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(CPMINDEX)
        smPLShow(CPMINDEX, llRowNo) = tmPostCtrls(CPMINDEX).sShow
        
        slStr = tgCPMPostLog(llRowNo).lOrdered
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(ORDEREDINDEX)
        smPLShow(ORDEREDINDEX, llRowNo) = tmPostCtrls(ORDEREDINDEX).sShow
        
        slStr = tgCPMPostLog(llRowNo).lTotalBilled
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(BILLEDINDEX)
        smPLShow(BILLEDINDEX, llRowNo) = tmPostCtrls(BILLEDINDEX).sShow
        
        slStr = tgCPMPostLog(llRowNo).lTotalInvoiced
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(TOTALINVOICEDINDEX)
        smPLShow(TOTALINVOICEDINDEX, llRowNo) = tmPostCtrls(TOTALINVOICEDINDEX).sShow
        
        If imPlSave(1, llRowNo) Then
            llPosted = tgCPMPostLog(llRowNo).lBilled
        Else
            llPosted = tgCPMPostLog(llRowNo).lUnBilled
        End If
        
        slStr = llPosted
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(UNBILLEDINDEX)
        smPLShow(UNBILLEDINDEX, llRowNo) = tmPostCtrls(UNBILLEDINDEX).sShow
        
        If tgCPMPostLog(llRowNo).sPriceType = "C" And llPosted > 0 Then
            If tgCPMPostLog(llRowNo).lTotalBilled > tgCPMPostLog(llRowNo).lOrdered Then
                monthCost = ((tgCPMPostLog(llRowNo).lCPM / 100) * (llPosted - (tgCPMPostLog(llRowNo).lTotalBilled - tgCPMPostLog(llRowNo).lOrdered))) / 1000
            Else
                monthCost = ((tgCPMPostLog(llRowNo).lCPM / 100) * llPosted) / 1000
            End If
            gFormatStr CStr(monthCost), 0, 2, slStr
        ElseIf tgCPMPostLog(llRowNo).sPriceType = "F" Then
            ' BOOSTER PHASE 2:  LINE ITEM 10 - Update "Month Cost" for flat rate lines to use new daily or existing monthly formula depending on Site setting 2023-11-20 JJB
            If tgSpfx.iLineCostType = 1 Then
                ' Daily Costs
                slStr = Format(mCalulatePeriodAmountByDaily(llRowNo), "#.##")
            Else
                ' Monthly Costs
                slStr = mCalculateFlatMonthlyCost(llRowNo)
            End If
            '''''''''''''''''''''''''''' END
            slStr = gDivStr(gMulStr(slStr, "1000.00"), "1000.00")
        End If
        gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(MONTHCOSTINDEX)
        smPLShow(MONTHCOSTINDEX, llRowNo) = tmPostCtrls(MONTHCOSTINDEX).sShow
    Next llRowNo
End Sub

'Boostr Phase 2 Line 10: Update "Month Cost" for flat rate lines to use new daily or existing monthly formula depending on Site setting
'2023-11-20 JJB & Joel

'mCalulatePeriodAmountByDaily
'   Inputs:
'       slPeriodStart     - Start date of the month being reported (string Date MM/DD/YYYY)
'       slPeriodEnd       - End date of the month being reported   (string Date MM/DD/YYYY)
'       slLineStartDate   - Start Date of the PCF Line             (string Date MM/DD/YYYY)
'       slLineEndDate     - End Date of the PCF Line               (string Date MM/DD/YYYY)
'       slRemainingAmount - The Remaining amount (Line Total minus whats already been Invoiced) (string Amount ####.##)
'   Output:
'       Month Amount (double precision number ####.##)
Public Function mCalulatePeriodAmountByDaily(llRowNo As Long) As Double

    Dim slLineStartDate         As String
    Dim slPeriodStart           As String
    Dim slPeriodEnd             As String
    Dim slLineEndDate           As String
    Dim slLastBilledDate        As String
    Dim llBilledAmount          As Long
    Dim llTotalAmount           As Long
    
    Dim dlDailyAmount           As Double   'The daily $ Amount
    Dim ilNumberOfDaysRemaining As Integer  'How many days from Invoice Start Date to Line EndDate
    Dim ilNumberOfDaysInPeriod  As Integer  'How many days are in this period
    Dim dStartDate              As Date     'Temp Start Date
    Dim dEndDate                As Date      'Temp End
       
    ' Assign need values
    If tgCPMPostLog(llRowNo).sChfBillCycle = "C" Then
        'Normal Calendar
        slLineStartDate = Format$(Format$(tgCPMPostLog(llRowNo).StartDate), "m/d/yyyy")
        slPeriodStart = gObtainStartCal(DateValue(sMonthName & " 15, " & iYear))
        slPeriodEnd = gObtainEndCal(DateValue(sMonthName & " 15, " & iYear))
        slLineEndDate = Format$(Format$(tgCPMPostLog(llRowNo).EndDate), "m/d/yyyy")
    Else
        'Standard Calendar
        slLineStartDate = Format$(Format$(tgCPMPostLog(llRowNo).StartDate), "m/d/yyyy")
        slPeriodStart = gObtainStartStd(DateValue(sMonthName & " 15, " & iYear))
        slPeriodEnd = gObtainEndStd(DateValue(sMonthName & " 15, " & iYear))
        slLineEndDate = Format$(Format$(tgCPMPostLog(llRowNo).EndDate), "m/d/yyyy")
    End If
    
    llBilledAmount = tgCPMPostLog(llRowNo).lPhfTotalGross + tgCPMPostLog(llRowNo).lRvfTotalGross
    llTotalAmount = tgCPMPostLog(llRowNo).lPcfTotalCost
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slLastBilledDate       'convert last bdcst billing date to string
    
    'Determine how many days remain of this line (beyond what's been billed)
    If DateValue(slLastBilledDate) + 1 > DateValue(slLineStartDate) Then
        dStartDate = DateValue(slLastBilledDate) + 1
    Else
        dStartDate = DateValue(slLineStartDate)
    End If
    
    dEndDate = DateValue(slLineEndDate)
    ilNumberOfDaysRemaining = DateDiff("d", dStartDate, dEndDate) + 1
    If ilNumberOfDaysRemaining <= 0 Then Exit Function
    
    'Determine how many days of this Line are being invoiced
    dStartDate = IIF(DateValue(slLineStartDate) > gDateValue(slPeriodStart), gDateValue(slLineStartDate), gDateValue(slPeriodStart))
    dEndDate = IIF(DateValue(slLineEndDate) < gDateValue(slPeriodEnd), gDateValue(slLineEndDate), gDateValue(slPeriodEnd))
    ilNumberOfDaysInPeriod = DateDiff("d", dStartDate, dEndDate) + 1
    If ilNumberOfDaysInPeriod <= 0 Then Exit Function
    
    'Determine the daily amount
    dlDailyAmount = ((llTotalAmount - llBilledAmount) / 100) / ilNumberOfDaysRemaining
    
    'Determine the amount to apply to this Period (slPeriodStart - slPeriodEnd)
    mCalulatePeriodAmountByDaily = dlDailyAmount * ilNumberOfDaysInPeriod
    
    Debug.Print "mCalulatePeriodAmountByDaily: "
    Debug.Print " -> Line Dates: " & slLineStartDate & " to " & slLineEndDate
    Debug.Print " -> RemainingAmount: " & Format((llTotalAmount - llBilledAmount) / 100, "#.00")
    Debug.Print " -> NumberOfDaysRemaining: " & ilNumberOfDaysRemaining
    Debug.Print " -> Period: " & slPeriodStart & " to " & slPeriodEnd & " = " & ilNumberOfDaysInPeriod
    Debug.Print " -> DailyAmount: " & dlDailyAmount
    Debug.Print " -> Month Amount: " & mCalulatePeriodAmountByDaily
    
End Function

Private Sub pbcEnd_GotFocus()
    imBoxNo = -1
    vbcPostManualCntrl.Value = 1
    pbcSTab.SetFocus
End Sub

Private Sub pbcPostManualCntrl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilBox As Integer
    Dim ilMaxRow As Integer
    Dim ilCompRow As Integer
    Dim ilRow As Integer
    Dim ilRowNo As Integer
    Dim ilLoop As Long
    Dim ilAltered As Integer
    
    For ilBox = imLBCtrls To UBound(tmPostCtrls) Step 1
        If ((X >= tmPostCtrls(ilBox).fBoxX) And (X <= (tmPostCtrls(ilBox).fBoxX + tmPostCtrls(ilBox).fBoxW))) Then
            If (Y < tmPostCtrls(ilBox).fBoxY - 15) Then
                If ilBox = CONTRACTINDEX Or ilBox = ADVERTISERINDEX Then
                    If imCurSort = 1 And ilBox <> ADVERTISERINDEX Then
                        imCurSort = 0
                        imSortOrder = 0
                    ElseIf imCurSort = 0 And ilBox <> CONTRACTINDEX Then
                        imCurSort = 1
                        imSortOrder = 0
                    Else
                        If imSortOrder = 0 Then
                            imSortOrder = 1
                        Else
                            imSortOrder = 0
                        End If
                    End If
                    mSetShow imBoxNo
                End If
                imBoxNo = -1
                imRowNo = -1
                imChg = False
                mSetCommands
                mResortLog
                pbcSTab.SetFocus
                Exit Sub
            End If
        End If
    Next ilBox
    
    ilCompRow = vbcPostManualCntrl.LargeChange + 1
    If UBound(tgCPMPostLog) >= ilCompRow Then
        ilMaxRow = ilCompRow
    Else
        ilMaxRow = UBound(tgCPMPostLog)
    End If
    For ilRow = 1 To ilMaxRow Step 1
        For ilBox = imLBCtrls To UBound(tmPostCtrls) Step 1
            If (X >= tmPostCtrls(ilBox).fBoxX) And (X <= (tmPostCtrls(ilBox).fBoxX + tmPostCtrls(ilBox).fBoxW)) Then
                If (Y >= ((ilRow - 1) * (fgBoxGridH + 15) + tmPostCtrls(ilBox).fBoxY)) And (Y <= ((ilRow - 1) * (fgBoxGridH + 15) + tmPostCtrls(ilBox).fBoxY + tmPostCtrls(ilBox).fBoxH)) Then
                    ilRowNo = ilRow + vbcPostManualCntrl.Value - 1
                    If ilRowNo > UBound(tgCPMPostLog) Then
                        Beep
                        mSetFocus imBoxNo
                        Exit Sub
                    End If
                    
                    If imPlSave(1, ilRowNo) Then
                        Beep
                        If ilMaxRow = ilRow Then
                            cmcCancel.SetFocus
                            Exit Sub
                        End If
                        
                        For ilLoop = ilRow To UBound(tgCPMPostLog) Step 1
                            If Not imPlSave(1, ilLoop) Then
                                imRowNo = ilLoop
                                ilBox = UNBILLEDINDEX
                                mEnableBox ilBox
                                Exit Sub
                            End If
                        Next ilLoop
                        Exit Sub
                    End If
                    
                    Select Case ilBox
                        Case UNBILLEDINDEX
                            If imRowNo > 0 Then
                                mSetShow imBoxNo
                            End If
                            imRowNo = ilRow + vbcPostManualCntrl.Value - 1
                            imBoxNo = ilBox
                            mEnableBox ilBox
                            Exit Sub
                        Case Else
                            Beep
                            imRowNo = ilRow + vbcPostManualCntrl.Value - 1
                            imBoxNo = UNBILLEDINDEX
                            mEnableBox imBoxNo
                            Exit Sub
                    End Select
                End If
            End If
        Next ilBox
    Next ilRow
    If ilMaxRow > 0 Then
        imBoxNo = -1
        pbcTab.SetFocus
    Else
        cmcCancel.SetFocus
    End If
End Sub


Private Sub vbcPostManualCntrl_Change()

    If imFirstTime Then
        pbcPostManualCntrl.Cls
        pbcPostManualCntrl_Paint
        imFirstTime = False
        Exit Sub
    End If

    If imSettingValue = True Then
        pbcPostManualCntrl.Cls
        pbcPostManualCntrl_Paint
        imSettingValue = False
    Else
        mSetShow imBoxNo
        pbcPostManualCntrl.Cls
        pbcPostManualCntrl_Paint
        
        If vbcPostManualCntrl.Value > imRowNo Then
            pbcTab.SetFocus
        ElseIf vbcPostManualCntrl.Value < imRowNo Then
            pbcSTab.SetFocus
        Else
            imBoxNo = -1
            pbcSTab.SetFocus
        End If
    End If
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mReadRec                        *
'*                                                     *
'*             Created:2/16/21       By: L.Bianchi     *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Private Function mReadRec() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Labels (Marked)                                                                  *
'*  mReadRecErr                                                                           *
'******************************************************************************************

'
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    'grdManualPost.Visible = False
    
    Dim rs As ADODB.Recordset
    
    Dim llUpper As Long
    Dim llLoop As Long
    Dim slStamp As String
    Dim llMax As Long
    
    Dim StartDate_CAL As String
    Dim EndDate_CAL As String
    
    Dim StartDate_STD As String
    Dim EndDate_STD As String
    
    If lServerVendor <= 0 Then
      mReadRec = False
      Exit Function
    End If
    
    ' Calendar Dates
    StartDate_CAL = gObtainStartCal(iYear & "-" & iMonth & "-" & "15")
    EndDate_CAL = gObtainEndCal(iYear & "-" & iMonth & "-" & "15")
    ' Standard Dates
    StartDate_STD = gObtainStartStd(iYear & "-" & iMonth & "-" & "15")
    EndDate_STD = gObtainEndStd(iYear & "-" & iMonth & "-" & "15")
    
    ReDim tgTempCPMPostLog(0) As PostCpmManual
          
    SQLQuery = "                     SELECT "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.chfBillCycle, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.VefCode, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.chfadfCode, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.ibfCode, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.ContractID, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.ID, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.rdfCode, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.Length, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.StartDate, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.EndDate, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.PriceType,"
    SQLQuery = SQLQuery & vbCrLf & "    pl1.CPM, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.Ordered, "
    SQLQuery = SQLQuery & vbCrLf & "    pl4.TotalBilled, "
    SQLQuery = SQLQuery & vbCrLf & "    pl4.TotalInvoiced, "
    SQLQuery = SQLQuery & vbCrLf & "    pl5.Billed, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.Unbilled, "
    SQLQuery = SQLQuery & vbCrLf & "    pl1.pcfTotalCost, "
    SQLQuery = SQLQuery & vbCrLf & "    pl2.RvfTotalGross, "
    SQLQuery = SQLQuery & vbCrLf & "    pl2.RvfTranDate, "
    SQLQuery = SQLQuery & vbCrLf & "    pl3.phfTotalGross, "
    SQLQuery = SQLQuery & vbCrLf & "    pl3.PhfTranDate"
    SQLQuery = SQLQuery & vbCrLf & " FROM"
    SQLQuery = SQLQuery & vbCrLf & "    ("
    SQLQuery = SQLQuery & vbCrLf & "      select "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfBillCycle, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfadfCode, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfVefCode        as VefCode, "
    SQLQuery = SQLQuery & vbCrLf & "        ibf2.ibfCode, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfCntrNo         as ContractID, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPodCPMID       as ID, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfRdfCode        as rdfCode, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfLen            as Length, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfStartDate      as StartDate,"
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfEndDate        as EndDate, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPriceType      as PriceType,"
    SQLQuery = SQLQuery & vbCrLf & "        pcf.PcfPodCPM         as CPM, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfImpressionGoal as Ordered, "
    SQLQuery = SQLQuery & vbCrLf & "        ibf2.ibfImpressions   as UnBilled, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfTotalCost "
    SQLQuery = SQLQuery & vbCrLf & "     from "
    SQLQuery = SQLQuery & vbCrLf & "        pcf_Pod_CPM_Cntr as pcf "
    SQLQuery = SQLQuery & vbCrLf & "        inner join CHF_Contract_Header      as chf  on pcf.pcfChfCode    = chf.chfCode"
    SQLQuery = SQLQuery & vbCrLf & "        inner join VFF_Vehicle_Features     as vff  on vff.vffVefCode    = pcf.pcfVefCode"
    SQLQuery = SQLQuery & vbCrLf & "        left Outer Join ibf_Impression_Bill as ibf2 on ibf2.ibfCntrNo    = chf.chfCntrNo    and "
    SQLQuery = SQLQuery & vbCrLf & "                                                       ibf2.ibfPodCPMId  = pcf.pcfPodCPMID  and "
    SQLQuery = SQLQuery & vbCrLf & "                                                       pcf.pcfVefCode    = ibf2.ibfVefCode  and "
    SQLQuery = SQLQuery & vbCrLf & "                                                       ibf2.ibfBilled    in ('N','')        and "
    SQLQuery = SQLQuery & vbCrLf & "                                                       ibf2.ibfBillYear  = " & iYear & "    and "
    SQLQuery = SQLQuery & vbCrLf & "                                                       ibf2.ibfBillMonth = " & iMonth
    SQLQuery = SQLQuery & vbCrLf & "     WHERE "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfDelete    = 'N' and "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfSchStatus = 'F' and "
    SQLQuery = SQLQuery & vbCrLf & "        vff.vffAvfCode   = " & lServerVendor & vbCrLf
    If sContractNumber <> "" Then
      SQLQuery = SQLQuery & vbCrLf & "      and chf.chfCntrNo  in " & sContractNumber
    ElseIf sAdvertiserCode <> "" Then
        SQLQuery = SQLQuery & vbCrLf & "    and chf.chfadfCode in " & sAdvertiserCode
    End If
    'TTP 10481 - Post Log Manual Ad Server Posting screen: ad server lines active in future months are appearing in months prior to their start date
    
    SQLQuery = SQLQuery & vbCrLf & "        and pcfStartDate <= '" & Format(EndDate_STD, "yyyy-mm-dd") & "' "
    SQLQuery = SQLQuery & vbCrLf & "        and pcfEndDate   >= '" & Format(StartDate_STD, "yyyy-mm-dd") & "' "
    SQLQuery = SQLQuery & vbCrLf & "    ) pl1 "
  
'******************************************************************************************
'* TTP 10579 - JJB
'* BEGIN
'******************************************************************************************

'--------------------
    SQLQuery = SQLQuery & vbCrLf & "    Left Outer Join "
    SQLQuery = SQLQuery & vbCrLf & "    ( "
    SQLQuery = SQLQuery & vbCrLf & "     Select  "
    SQLQuery = SQLQuery & vbCrLf & "        sum(rvfGross) As RvfTotalGross, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfCntrNo, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPodCPMID, "
    SQLQuery = SQLQuery & vbCrLf & "        max(rvfTranDate) as RvfTranDate "
    SQLQuery = SQLQuery & vbCrLf & "    From  "
    SQLQuery = SQLQuery & vbCrLf & "        CHF_Contract_Header chf "
    SQLQuery = SQLQuery & vbCrLf & "            inner join pcf_Pod_CPM_Cntr     as pcf on pcf.pcfChfCode = chf.chfCode "
    SQLQuery = SQLQuery & vbCrLf & "            inner join VFF_Vehicle_Features as vff on vff.vffVefCode = pcf.pcfVefCode "
    SQLQuery = SQLQuery & vbCrLf & "            inner join RVF_Receivables      as rvf on rvf.rvfCntrNo  = chf.chfCntrNo                   And "
    SQLQuery = SQLQuery & vbCrLf & "                                                     rvf.rvfPcfCode  = pcf.pcfCode                     And "
    SQLQuery = SQLQuery & vbCrLf & "                                                    (rvf.rvfTranType = 'IN' or rvf.rvfTranType = 'HI') "
    SQLQuery = SQLQuery & vbCrLf & "    Group by "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfCntrNo, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPodCPMID, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfDelete,  "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfSchStatus, "
    SQLQuery = SQLQuery & vbCrLf & "        pcfStartDate,  "
    SQLQuery = SQLQuery & vbCrLf & "        pcfEndDate, "
    SQLQuery = SQLQuery & vbCrLf & "        vff.vffAvfCode, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPriceType "
    If sAdvertiserCode <> "" Then
        SQLQuery = SQLQuery & vbCrLf & "    ,chf.chfadfCode "
    End If
    SQLQuery = SQLQuery & vbCrLf & "    Having "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfSchStatus = 'F' and "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPriceType = 'F' and "
    SQLQuery = SQLQuery & vbCrLf & "        vff.vffAvfCode   = " & lServerVendor
    If sContractNumber <> "" Then
        SQLQuery = SQLQuery & vbCrLf & "    and chf.chfCntrNo  in " & sContractNumber
    ElseIf sAdvertiserCode <> "" Then
        SQLQuery = SQLQuery & vbCrLf & "    and chf.chfadfCode in " & sAdvertiserCode
    End If
    SQLQuery = SQLQuery & vbCrLf & "    ) pl2 on pl1.ContractID = pl2.chfCntrNo and "
    SQLQuery = SQLQuery & vbCrLf & "             pl1.ID         = pl2.pcfPodCPMID "
    
'--------------------
    SQLQuery = SQLQuery & vbCrLf & "    Left Outer Join "
    SQLQuery = SQLQuery & vbCrLf & "    ( "
    SQLQuery = SQLQuery & vbCrLf & "     Select "
    SQLQuery = SQLQuery & vbCrLf & "        sum(phfGross)    as phfTotalGross, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfCntrNo,  "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPodCPMID, "
    SQLQuery = SQLQuery & vbCrLf & "        max(phfTranDate) as PhfTranDate "
    SQLQuery = SQLQuery & vbCrLf & "     From  "
    SQLQuery = SQLQuery & vbCrLf & "        CHF_Contract_Header chf  "
    SQLQuery = SQLQuery & vbCrLf & "            inner join pcf_Pod_CPM_Cntr     as pcf on pcf.pcfChfCode  = chf.chfCode  "
    SQLQuery = SQLQuery & vbCrLf & "            inner join VFF_Vehicle_Features as vff on vff.vffVefCode  = pcf.pcfVefCode "
    SQLQuery = SQLQuery & vbCrLf & "            inner join phf_Payment_History  as phf on phf.phfCntrNo   = chf.chfCntrNo   and "
    SQLQuery = SQLQuery & vbCrLf & "                                                      phf.phfPcfCode  = PCF.pcfCode     and "
    SQLQuery = SQLQuery & vbCrLf & "                                                     (phf.phfTranType = 'IN' or phf.phfTranType = 'HI') "
    SQLQuery = SQLQuery & vbCrLf & "     Group by "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfCntrNo, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPodCPMID, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfDelete, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfSchStatus, "
    SQLQuery = SQLQuery & vbCrLf & "        pcfStartDate, "
    SQLQuery = SQLQuery & vbCrLf & "        pcfEndDate, "
    SQLQuery = SQLQuery & vbCrLf & "        vff.vffAvfCode, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPriceType "
    If sAdvertiserCode <> "" Then
        SQLQuery = SQLQuery & vbCrLf & "    ,chf.chfadfCode "
    End If
    SQLQuery = SQLQuery & vbCrLf & "     Having "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfSchStatus = 'F' and "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPriceType = 'F' and "
    SQLQuery = SQLQuery & vbCrLf & "        vff.vffAvfCode   = " & lServerVendor
    If sContractNumber <> "" Then
        SQLQuery = SQLQuery & vbCrLf & "    and chf.chfCntrNo  in " & sContractNumber
    ElseIf sAdvertiserCode <> "" Then
        SQLQuery = SQLQuery & vbCrLf & "    and chf.chfadfCode in " & sAdvertiserCode
    End If
    SQLQuery = SQLQuery & vbCrLf & "    ) pl3 on pl1.ContractID = pl3.chfCntrNo and "
    SQLQuery = SQLQuery & vbCrLf & "             pl1.ID         = pl3.pcfPodCPMID "
   
'******************************************************************************************
'* TTP 10579 - JJB
'* END
'******************************************************************************************

'--------------------
    SQLQuery = SQLQuery & vbCrLf & "    Left Outer Join "
    SQLQuery = SQLQuery & vbCrLf & "    ( "
    SQLQuery = SQLQuery & vbCrLf & "     Select "
    SQLQuery = SQLQuery & vbCrLf & "        sum(ibfImpressions)      as TotalBilled, "
    SQLQuery = SQLQuery & vbCrLf & "        sum(ibfBilledImpression) as TotalInvoiced, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfCntrNo, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPodCPMID "
    SQLQuery = SQLQuery & vbCrLf & "     From "
    SQLQuery = SQLQuery & vbCrLf & "        CHF_Contract_Header chf "
    SQLQuery = SQLQuery & vbCrLf & "            inner join pcf_Pod_CPM_Cntr     as pcf  on pcf.pcfChfCode   = chf.chfCode "
    SQLQuery = SQLQuery & vbCrLf & "            inner join VFF_Vehicle_Features as vff  on vff.vffVefCode   = pcf.pcfVefCode "
    SQLQuery = SQLQuery & vbCrLf & "            inner join ibf_Impression_Bill  as ibf2 on ibf2.ibfCntrNo   = chf.chfCntrNo and "
    SQLQuery = SQLQuery & vbCrLf & "                                                       ibf2.ibfPodCPMId = pcf.pcfPodCPMID and "
    SQLQuery = SQLQuery & vbCrLf & "                                                       pcf.pcfVefCode   = ibf2.ibfVefCode and "
    SQLQuery = SQLQuery & vbCrLf & "                                                       ibf2.ibfBilled   in ('Y')"
    SQLQuery = SQLQuery & vbCrLf & "     Group by "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfCntrNo, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPodCPMID, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfDelete, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfSchStatus, "
    SQLQuery = SQLQuery & vbCrLf & "        pcfStartDate, "
    SQLQuery = SQLQuery & vbCrLf & "        pcfEndDate, "
    SQLQuery = SQLQuery & vbCrLf & "        vff.vffAvfCode, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPriceType "
    If sAdvertiserCode <> "" Then
        SQLQuery = SQLQuery & vbCrLf & "    ,chf.chfadfCode"
    End If
    SQLQuery = SQLQuery & vbCrLf & "     Having "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfDelete    = 'N' and "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfSchStatus = 'F' and "
    SQLQuery = SQLQuery & vbCrLf & "        vff.vffAvfCode   = " & lServerVendor
    If sContractNumber <> "" Then
      SQLQuery = SQLQuery & vbCrLf & "      and chf.chfCntrNo in " & sContractNumber
    ElseIf sAdvertiserCode <> "" Then
        SQLQuery = SQLQuery & vbCrLf & "    and chf.chfadfCode in " & sAdvertiserCode
    End If
    SQLQuery = SQLQuery & vbCrLf & "    ) pl4 on pl1.ContractID = pl4.chfCntrNo and "
    SQLQuery = SQLQuery & vbCrLf & "             pl1.ID         = pl4.pcfPodCPMID "
 

'--------------------
    SQLQuery = SQLQuery & vbCrLf & "    Left Outer Join "
    SQLQuery = SQLQuery & vbCrLf & "    ( "
    SQLQuery = SQLQuery & vbCrLf & "     Select "
    SQLQuery = SQLQuery & vbCrLf & "        sum(ibfImpressions) As Billed, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfCntrNo, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPodCPMID "
    SQLQuery = SQLQuery & vbCrLf & "     From "
    SQLQuery = SQLQuery & vbCrLf & "        CHF_Contract_Header chf "
    SQLQuery = SQLQuery & vbCrLf & "            inner join pcf_Pod_CPM_Cntr as pcf on pcf.pcfChfCode        = chf.chfCode"
    SQLQuery = SQLQuery & vbCrLf & "            inner join VFF_Vehicle_Features as vff on vff.vffVefCode    = pcf.pcfVefCode"
    SQLQuery = SQLQuery & vbCrLf & "            inner join ibf_Impression_Bill as ibf2 on ibf2.ibfCntrNo    = chf.chfCntrNo and "
    SQLQuery = SQLQuery & vbCrLf & "                                                      ibf2.ibfPodCPMId  = pcf.pcfPodCPMID and "
    SQLQuery = SQLQuery & vbCrLf & "                                                      pcf.pcfVefCode    = ibf2.ibfVefCode and "
    SQLQuery = SQLQuery & vbCrLf & "                                                      ibf2.ibfBilled    in ('Y') and "
    SQLQuery = SQLQuery & vbCrLf & "                                                      ibf2.ibfBilled    in ('Y') and "
    SQLQuery = SQLQuery & vbCrLf & "                                                      ibf2.ibfBillYear  = " & iYear & " and "
    SQLQuery = SQLQuery & vbCrLf & "                                                      ibf2.ibfBillMonth = " & iMonth
    SQLQuery = SQLQuery & vbCrLf & "     Group by "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfCntrNo, "
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPodCPMID, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfDelete, "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfSchStatus, "
    SQLQuery = SQLQuery & vbCrLf & "        pcfStartDate, "
    SQLQuery = SQLQuery & vbCrLf & "        pcfEndDate,"
    SQLQuery = SQLQuery & vbCrLf & "        vff.vffAvfCode,"
    SQLQuery = SQLQuery & vbCrLf & "        pcf.pcfPriceType,"
    SQLQuery = SQLQuery & vbCrLf & "        ibf2.ibfBillYear,"
    SQLQuery = SQLQuery & vbCrLf & "        ibf2.ibfBillMonth"
    If sAdvertiserCode <> "" Then
        SQLQuery = SQLQuery & vbCrLf & "    ,chf.chfadfCode"
    End If
    SQLQuery = SQLQuery & vbCrLf & "     Having "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfDelete    = 'N' and "
    SQLQuery = SQLQuery & vbCrLf & "        chf.chfSchStatus = 'F' and "
    SQLQuery = SQLQuery & vbCrLf & "        vff.vffAvfCode = " & lServerVendor
    If sContractNumber <> "" Then
      SQLQuery = SQLQuery & vbCrLf & "      and chf.chfCntrNo in " & sContractNumber
    ElseIf sAdvertiserCode <> "" Then
        SQLQuery = SQLQuery & vbCrLf & "    and chf.chfadfCode in " & sAdvertiserCode
    End If
    SQLQuery = SQLQuery & vbCrLf & "    ) pl5 on pl1.ContractID = pl5.chfCntrNo and "
    SQLQuery = SQLQuery & vbCrLf & "             pl1.ID         = pl5.pcfPodCPMID "
    SQLQuery = SQLQuery & vbCrLf & "             and ( "
    SQLQuery = SQLQuery & vbCrLf & "                    ( '" & Format(StartDate_STD, "yyyy-mm-dd") & "' between StartDate and EndDate )"
    SQLQuery = SQLQuery & vbCrLf & "                    or "
    SQLQuery = SQLQuery & vbCrLf & "                    ('" & Format(EndDate_STD, "yyyy-mm-dd") & "'    between StartDate and EndDate ) "
    SQLQuery = SQLQuery & vbCrLf & "                    or "
    SQLQuery = SQLQuery & vbCrLf & "                    ('" & Format(StartDate_STD, "yyyy-mm-dd") & "' < StartDate and '" & Format(EndDate_STD, "yyyy-mm-dd") & "' > EndDate)"
    SQLQuery = SQLQuery & vbCrLf & "                ) "
    
    Set rs = gSQLSelectCall(SQLQuery)
    
    llUpper = imLBPost
    While Not rs.EOF
        ReDim Preserve tgTempCPMPostLog(llUpper) As PostCpmManual
        
        With tgTempCPMPostLog(llUpper)
            .lContractNumber = rs!ContractID
            .iID = rs!id
            .iRdfCode = rs!rdfCode
            .Length = rs!Length
            .StartDate = rs!StartDate
            .EndDate = rs!EndDate
            .sPriceType = rs!PriceType
            .lCPM = rs!CPM
            .lOrdered = rs!Ordered
            .iVehCode = rs!vefCode
            .sChfBillCycle = rs!chfBillCycle
            .lPcfTotalCost = rs!pcfTotalCost
            .iAdfCode = rs!chfadfCode
        
            .lBilled = IIF(IsNull(rs("Billed")), -1, rs!Billed)
            .ilbfUnbilledCode = IIF(IsNull(rs("ibfCode")), .ilbfUnbilledCode, rs!ibfCode)
            .lUnBilled = IIF(IsNull(rs("Unbilled")), .lUnBilled, rs!Unbilled)
            .lTotalBilled = IIF(IsNull(rs("TotalBilled")), .lTotalBilled, rs!TotalBilled)
            .lTotalInvoiced = IIF(IsNull(rs("TotalInvoiced")), .lTotalInvoiced, rs!TotalInvoiced)
            .lRvfTotalGross = IIF(IsNull(rs("RvfTotalGross")), .lRvfTotalGross, rs!RvfTotalGross) * 100 'Joey ' Booster v2 issue 6 fix
            .lPhfTotalGross = IIF(IsNull(rs("PhfTotalGross")), .lPhfTotalGross, rs!phfTotalGross) * 100 'Joey ' Booster v2 issue 6 fix
            .rvfTranDate = IIF(IsNull(rs("rvfTranDate")), .rvfTranDate, rs!rvfTranDate)
            .phfTranDate = IIF(IsNull(rs("phfTranDate")), .phfTranDate, rs!phfTranDate)
        End With
        
        llUpper = llUpper + 1
        rs.MoveNext
    Wend
    If llUpper > 1 Then
         imLoaded = True
    End If
    
    mReadRec = True
    Exit Function
    
mReadRecErr: 'VBC NR
    On Error GoTo 0
    mReadRec = False
    Exit Function
End Function


Private Sub pbcClickFocus_GotFocus()
    mSetShow imBoxNo
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilPrev As Integer
    Dim ilLoop As Integer
    Dim ilBox As Integer
    Dim ilRowNo As Integer
    
    If GetFocus() <> pbcSTab.HWnd Then
        Exit Sub
    End If
    If imLoaded <> -1 Then
        Exit Sub
    End If
    ilBox = imBoxNo
    Select Case ilBox
        Case -1
            For ilLoop = imLBPost To UBound(tgCPMPostLog) Step 1
                If Not imPlSave(1, ilLoop) Then
                    imRowNo = ilLoop
                    ilBox = UNBILLEDINDEX
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            Next ilLoop
        Case UNBILLEDINDEX
            If imRowNo > 0 Then
                mSetShow (ilBox)
            End If
             If imLBPost <= imRowNo - 1 Then
                 ilRowNo = imRowNo - 1
                If ilRowNo < vbcPostManualCntrl.Value Then
                    imSettingValue = True
                    vbcPostManualCntrl.Value = vbcPostManualCntrl.Value - 1
                    imSettingValue = False
                End If
                 For ilLoop = ilRowNo To imLBPost Step -1
                    If Not imPlSave(1, ilLoop) Then
                        imRowNo = ilLoop
                        mEnableBox ilBox
                        Exit Sub
                    End If
                Next ilLoop
                mEnableBox ilBox
                Exit Sub
            Else
                mEnableBox ilBox
            End If
    End Select
End Sub

Private Sub pbcTab_GotFocus()
    Dim ilBox As Integer
    Dim ilLoop As Integer
    Dim ilRowNo As Integer
    
    If GetFocus() <> pbcTab.HWnd Then
        Exit Sub
    End If
    
    If imLoaded <> -1 Then
        cmcCancel.SetFocus
        Exit Sub
    End If
    
        ilBox = imBoxNo
    Select Case ilBox
        Case -1
            For ilLoop = UBound(tgCPMPostLog) To 1 Step -1
                If Not imPlSave(1, ilLoop) Then
                    imRowNo = ilLoop
                    ilBox = UNBILLEDINDEX
                    imBoxNo = ilBox
                    mEnableBox ilBox
                    Exit Sub
                End If
            Next ilLoop
        Case UNBILLEDINDEX
             If imRowNo > 0 Then
                mSetShow (ilBox)
             End If
             If UBound(tgCPMPostLog) >= imRowNo + 1 Then
                 ilRowNo = imRowNo + 1
                If ilRowNo > vbcPostManualCntrl.Value + vbcPostManualCntrl.LargeChange Then ' + 1 Then
                    imSettingValue = True
                    vbcPostManualCntrl.Value = vbcPostManualCntrl.Value + 1
                    imSettingValue = False
                End If
                 For ilLoop = ilRowNo To UBound(tgCPMPostLog) Step 1
                    If Not imPlSave(1, ilLoop) Then
                        imRowNo = ilLoop
                        mEnableBox ilBox
                        Exit Sub
                    End If
                Next ilLoop
            End If
    End Select
    If cmcSave.Enabled Then
        cmcSave.SetFocus
    Else
        cmcCancel.SetFocus
    End If
    
    imRowNo = imRowNo - 1
End Sub

Private Sub mEnableBox(ilBoxNo As Integer)
'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    pbcArrow.Move pbcArrow.Left, plcPostManualCntrl.Top + tmPostCtrls(ilBoxNo).fBoxY + (imRowNo - vbcPostManualCntrl.Value) * (fgBoxGridH + 15) + 45
    pbcArrow.Visible = True
    lacFrame.Move 0, tmPostCtrls(CONTRACTINDEX).fBoxY + (imRowNo - vbcPostManualCntrl.Value) * (fgBoxGridH + 15) - 30
    lacFrame.Visible = True
    Select Case ilBoxNo
        Case UNBILLEDINDEX
          edcDropDown.Width = tmPostCtrls(ilBoxNo).fBoxW '- cmcDropDown.Width
          edcDropDown.MaxLength = 10
          gMoveTableCtrl plcPostManualCntrl, edcDropDown, tmPostCtrls(ilBoxNo).fBoxX + 30, tmPostCtrls(ilBoxNo).fBoxY + (imRowNo - vbcPostManualCntrl.Value) * (fgBoxGridH + 15) + 45
          edcDropDown.Visible = True
          edcDropDown.SetFocus
          edcDropDown.Text = Trim$(smPLShow(UNBILLEDINDEX, imRowNo))
    End Select
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetShow                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Format user input for a control*
'*                      to be displayed on the form    *
'*                                                     *
'*******************************************************
Private Sub mSetShow(ilBoxNo As Integer)
    Dim slStr As String
    pbcArrow.Visible = False
    lacFrame.Visible = False
    If (ilBoxNo < imLBCtrls) Or (ilBoxNo > UBound(tmPostCtrls)) Then
        Exit Sub
    End If
    Select Case ilBoxNo
        Case UNBILLEDINDEX
            edcDropDown.Visible = False
            slStr = edcDropDown.Text
            If slStr = "" Then
                Exit Sub
            End If
            
            If StrComp(Trim$(smPLShow(UNBILLEDINDEX, imRowNo)), slStr, vbTextCompare) <> 0 Then
                smPLShow(UNBILLEDINDEX, imRowNo) = slStr
                tgCPMPostLog(imRowNo).lUnBilled = CLng(slStr)
                imPlSave(0, imRowNo) = True
                imChg = True
            End If
            pbcPostManualCntrl.CurrentX = tmPostCtrls(ilBoxNo).fBoxX + fgBoxInsetX
            pbcPostManualCntrl.CurrentY = tmPostCtrls(ilBoxNo).fBoxY + (imRowNo - vbcPostManualCntrl.Value) * (fgBoxGridH + 15)
            gSetShow pbcPostManualCntrl, smPLShow(UNBILLEDINDEX, imRowNo), tmPostCtrls(ilBoxNo)
            pbcPostManualCntrl.Print tmPostCtrls(ilBoxNo).sShow
            edcDropDown.Text = ""
    End Select
    mSetCommands
End Sub

Private Function mSaveRec() As Integer
'
'   iRet = mSaveRec()
'   Where:
'       iRet (O)- True if updated or added, False if not updated or added
    Dim ilRet As Integer
    If Not imUpdateAllowed Then
        mSaveRec = False
        Exit Function
    End If
    
    If imChg Then
        ilRet = mUpdateIbf()
    End If
    imChg = False
    mSaveRec = True
    Exit Function
mSaveRecErr:
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    'gSetMousePointer grdManualPost, grdManualPost, vbDefault
    imTerminate = True
    mSaveRec = False
    Exit Function
End Function
Private Sub mSetCommands()
'
'   mSetCommands
'   Where:
'
    If Not imUpdateAllowed Then
        cmcSave.Enabled = False
        cmcDone.Enabled = False
    Else
        cmcDone.Enabled = True
        If imChg Then
            cmcSave.Enabled = True
        Else
            cmcSave.Enabled = False
        End If
    End If
End Sub

Private Sub tmcStart_Timer()
    
    tmcStart.Enabled = False
    If imTerminate Then
        mTerminate
        Exit Sub
    End If
End Sub

Private Function mUpdateIbf() As Integer
    Dim llRowNo As Long
    Dim slQuery As String
    Dim llCount As Long
    For llRowNo = imLBPost To UBound(tgCPMPostLog) Step 1
        If Not imPlSave(1, llRowNo) Then
            If imPlSave(0, llRowNo) Then
                If tgCPMPostLog(llRowNo).ilbfUnbilledCode = 0 Then
                    slQuery = "Insert into ibf_Impression_Bill (ibfCntrNo, ibfPodCPMId, ibfVefCode, ibfBillCycle,ibfBillYear, ibfBillMonth, ibfImpressions, ibfBilled, ibfUrfCode, ibfSource) Values (" & vbCrLf
                    slQuery = slQuery & tgCPMPostLog(llRowNo).lContractNumber & "," & vbCrLf
                    slQuery = slQuery & tgCPMPostLog(llRowNo).iID & "," & vbCrLf
                    slQuery = slQuery & tgCPMPostLog(llRowNo).iVehCode & "," & vbCrLf
                    slQuery = slQuery & "'" & tgCPMPostLog(llRowNo).sChfBillCycle & "'," & vbCrLf
                    slQuery = slQuery & iYear & "," & vbCrLf
                    slQuery = slQuery & iMonth & "," & vbCrLf
                    
                    slQuery = slQuery & tgCPMPostLog(llRowNo).lUnBilled & "," & vbCrLf
                    slQuery = slQuery & "'N'," & vbCrLf
                    slQuery = slQuery & tgUrf(0).iCode & "," & vbCrLf
                    slQuery = slQuery & "'M'" & vbCrLf
                    slQuery = slQuery & ")"
                Else
                    slQuery = "Update ibf_Impression_Bill set ibfImpressions = " & tgCPMPostLog(llRowNo).lUnBilled & vbCrLf
                    slQuery = slQuery & ",ibfUrfCode = " & tgUrf(0).iCode & vbCrLf
                    slQuery = slQuery & ",ibfBillCycle = '" & tgCPMPostLog(llRowNo).sChfBillCycle & "'" & vbCrLf
                    slQuery = slQuery & "Where ibfCode = " & tgCPMPostLog(llRowNo).ilbfUnbilledCode
                End If
                If gSQLAndReturn(slQuery, False, llCount) <> 0 Then
                    gHandleError "TrafficErrors.txt", "PostCpmManualCtrl-mUpdateIbf"
                End If
                mUpdateIbf = True
            End If
        End If
    Next llRowNo
End Function


Private Sub mInitBox()
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
    Dim ilSpaceBetweenButtons As Integer
    Dim llAdjTop As Long
    Dim llMax As Long
    
    On Error GoTo mInitBoxErr
    
    plcScreen.Caption = "Ad Server Impression Posting for: " & sMonthName & " " & iYear
    plcScreen.Width = Len(plcScreen.Caption) * 156
    plcScreen.Move 135, fgBevelY
    
    
    cmcDone.Move 2415, 5655
    cmcCancel.Move cmcDone.Left + cmcDone.Width + 120, cmcDone.Top
    cmcSave.Move cmcCancel.Left + cmcCancel.Width + 120, cmcDone.Top
    plcPostManualCntrl.Move 135, 500
    
    'ContractID
    gSetCtrl tmPostCtrls(CONTRACTINDEX), 10, 375, 375, fgBoxGridH
    'Advertiser
    gSetCtrl tmPostCtrls(ADVERTISERINDEX), 400, tmPostCtrls(CONTRACTINDEX).fBoxY, 950, fgBoxGridH
    'ID
    gSetCtrl tmPostCtrls(IDINDEX), 1365, tmPostCtrls(CONTRACTINDEX).fBoxY, 200, fgBoxGridH
    'Vechicle
    gSetCtrl tmPostCtrls(VEHICLEINDEX), 1580, tmPostCtrls(CONTRACTINDEX).fBoxY, 950, fgBoxGridH
    'Ad Location
    gSetCtrl tmPostCtrls(ADLOCATIONINDEX), 2545, tmPostCtrls(CONTRACTINDEX).fBoxY, 800, fgBoxGridH
    'Length
    gSetCtrl tmPostCtrls(LENGTHINDEX), 3360, tmPostCtrls(CONTRACTINDEX).fBoxY, 400, fgBoxGridH
    'Date Range
    gSetCtrl tmPostCtrls(DATERANGEINDEX), 3775, tmPostCtrls(CONTRACTINDEX).fBoxY, 650, fgBoxGridH
    'Price Type
    gSetCtrl tmPostCtrls(PRICETYPEINDEX), 4440, tmPostCtrls(CONTRACTINDEX).fBoxY, 400, fgBoxGridH
    'CPM
    gSetCtrl tmPostCtrls(CPMINDEX), 4855, tmPostCtrls(CONTRACTINDEX).fBoxY, 400, fgBoxGridH
    'Ordered
    gSetCtrl tmPostCtrls(ORDEREDINDEX), 5270, tmPostCtrls(CONTRACTINDEX).fBoxY, 550, fgBoxGridH
    'Billed
    gSetCtrl tmPostCtrls(BILLEDINDEX), 5835, tmPostCtrls(CONTRACTINDEX).fBoxY, 550, fgBoxGridH
    'TotalInvoiced
    gSetCtrl tmPostCtrls(TOTALINVOICEDINDEX), 6400, tmPostCtrls(CONTRACTINDEX).fBoxY, 550, fgBoxGridH
    'Posted
    gSetCtrl tmPostCtrls(UNBILLEDINDEX), 6965, tmPostCtrls(CONTRACTINDEX).fBoxY, 550, fgBoxGridH
    'MonthCost
    gSetCtrl tmPostCtrls(MONTHCOSTINDEX), 7530, tmPostCtrls(CONTRACTINDEX).fBoxY, 550, fgBoxGridH
    
    llMax = 0
    For ilLoop = imLBCtrls To UBound(tmPostCtrls) Step 1
        tmPostCtrls(ilLoop).fBoxW = CLng(fmAdjFactorW * tmPostCtrls(ilLoop).fBoxW)
        Do While (tmPostCtrls(ilLoop).fBoxW Mod 15) <> 0
            tmPostCtrls(ilLoop).fBoxW = tmPostCtrls(ilLoop).fBoxW + 1
        Loop
        If ilLoop > 0 Then
            tmPostCtrls(ilLoop).fBoxX = CLng(fmAdjFactorW * tmPostCtrls(ilLoop).fBoxX)
            Do While (tmPostCtrls(ilLoop).fBoxX Mod 15) <> 0
                tmPostCtrls(ilLoop).fBoxX = tmPostCtrls(ilLoop).fBoxX + 1
            Loop
            If (tmPostCtrls(ilLoop).fBoxX > 90) Then
                Do
                    If tmPostCtrls(ilLoop - 1).fBoxX + tmPostCtrls(ilLoop - 1).fBoxW + 15 < tmPostCtrls(ilLoop).fBoxX Then
                        tmPostCtrls(ilLoop - 1).fBoxW = tmPostCtrls(ilLoop - 1).fBoxW + 15
                    ElseIf tmPostCtrls(ilLoop - 1).fBoxX + tmPostCtrls(ilLoop - 1).fBoxW + 15 > tmPostCtrls(ilLoop).fBoxX Then
                        tmPostCtrls(ilLoop - 1).fBoxW = tmPostCtrls(ilLoop - 1).fBoxW - 15
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
        If tmPostCtrls(ilLoop).fBoxX + tmPostCtrls(ilLoop).fBoxW + 15 > llMax Then
            llMax = tmPostCtrls(ilLoop).fBoxX + tmPostCtrls(ilLoop).fBoxW + 15
        End If
    Next ilLoop
    
    pbcPostManualCntrl.Width = llMax
    plcPostManualCntrl.Width = llMax + vbcPostManualCntrl.Width + 2 * fgBevelX + 15
    
    pbcPostManualCntrl.Left = plcPostManualCntrl.Left + fgBevelX
    pbcPostManualCntrl.Top = plcPostManualCntrl.Top + fgBevelY
    
    vbcPostManualCntrl.Left = plcPostManualCntrl.Width - vbcPostManualCntrl.Width - 2 * fgBevelX
    vbcPostManualCntrl.Top = fgBevelY
    PostCpmManualCntrl.Width = plcPostManualCntrl.Left + plcPostManualCntrl.Width + plcPostManualCntrl.Left * 2
    
    ilSpaceBetweenButtons = fmAdjFactorW * (cmcCancel.Left - (cmcDone.Left + cmcDone.Width))
    Do While ilSpaceBetweenButtons Mod 15 <> 0
        ilSpaceBetweenButtons = ilSpaceBetweenButtons + 1
    Loop
    cmcDone.Left = (PostCpmManualCntrl.Width - cmcDone.Width - cmcCancel.Width - cmcSave.Width - 4 * ilSpaceBetweenButtons) / 2
    cmcCancel.Left = cmcDone.Left + cmcDone.Width + ilSpaceBetweenButtons
    cmcSave.Left = cmcCancel.Left + cmcCancel.Width + ilSpaceBetweenButtons
    cmcDone.Top = PostCpmManualCntrl.height - (3 * cmcDone.height) / 2
    cmcCancel.Top = cmcDone.Top
    cmcSave.Top = cmcDone.Top
    llAdjTop = cmcDone.Top - plcPostManualCntrl.Top - 120 - tmPostCtrls(1).fBoxH
    If llAdjTop < 0 Then
        llAdjTop = 0
    End If
    Do While (llAdjTop Mod 15) <> 0
        llAdjTop = llAdjTop + 1
    Loop
    Do While ((llAdjTop Mod (CInt(fgBoxGridH) + 15))) <> 0
        llAdjTop = llAdjTop - 1
    Loop
    Do While plcPostManualCntrl.Top + llAdjTop + 2 * fgBevelY + 240 < cmcCancel.Top
        llAdjTop = llAdjTop + CInt(fgBoxGridH) + 15
    Loop
    
    plcPostManualCntrl.height = llAdjTop + 2 * fgBevelY
    pbcPostManualCntrl.height = plcPostManualCntrl.height - 2 * fgBevelY
    vbcPostManualCntrl.height = pbcPostManualCntrl.height - 100
    vbcPostManualCntrl.Top = 0
    plcPostManualCntrl.Width = llMax + vbcPostManualCntrl.Width + 2 * fgBevelX + 15
    lacFrame.Width = llMax - 15
    
    Exit Sub
mInitBoxErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub pbcPostManualCntrl_Paint()
    Dim ilBox As Integer
    Dim ilRow As Long
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim slFont As String
    Dim ilRet As Integer
    Dim ilPRowSameContract As Integer
    Dim llForeColor As Long
    Dim ilMonthCost As Currency
    Dim llPosted As Long
    
    mPaintPostTitle
    
    If imLoaded <> -1 Then
        Exit Sub
    End If
    ilStartRow = vbcPostManualCntrl.Value
    ilEndRow = vbcPostManualCntrl.Value + vbcPostManualCntrl.LargeChange
    ilRow = imLBPost
    If ilEndRow - 1 >= UBound(tgCPMPostLog) Then
        ilEndRow = UBound(tgCPMPostLog)
    End If
    
    'slFont = pbcPostManualCntrl.FontName
    For ilRow = ilStartRow To ilEndRow Step 1
        If ilRow > 1 Then
            If tgCPMPostLog(ilRow).lContractNumber = tgCPMPostLog(ilRow - 1).lContractNumber Then
                ilPRowSameContract = True
            Else
                ilPRowSameContract = False
            End If
        End If
        
        For ilBox = imLBCtrls To UBound(tmPostCtrls) Step 1
            pbcPostManualCntrl.CurrentX = tmPostCtrls(ilBox).fBoxX + fgBoxInsetX
            pbcPostManualCntrl.CurrentY = tmPostCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15)
            slStr = smPLShow(ilBox, ilRow)
            Select Case ilBox
                Case CONTRACTINDEX
                    If ilPRowSameContract Then
                        slStr = ""
                    End If
                Case ADVERTISERINDEX
                    If ilPRowSameContract Then
                        slStr = ""
                    End If
                Case MONTHCOSTINDEX
                    
                    If tgCPMPostLog(ilRow).sPriceType = "C" Then
                        If imPlSave(1, ilRow) Then
                            llPosted = tgCPMPostLog(ilRow).lBilled
                        Else
                            llPosted = tgCPMPostLog(ilRow).lUnBilled
                        End If
                        If llPosted > 0 Then
                            
                              slStr = llPosted
                              gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(UNBILLEDINDEX)
                              smPLShow(UNBILLEDINDEX, ilRow) = tmPostCtrls(UNBILLEDINDEX).sShow
                              If tgCPMPostLog(ilRow).lTotalBilled > tgCPMPostLog(ilRow).lOrdered Then
                                  ilMonthCost = ((tgCPMPostLog(ilRow).lCPM / 100) * (llPosted - (tgCPMPostLog(ilRow).lTotalBilled - tgCPMPostLog(ilRow).lOrdered))) / 1000
                              Else
                                  ilMonthCost = ((tgCPMPostLog(ilRow).lCPM / 100) * llPosted) / 1000
                              End If
                              gFormatStr CStr(ilMonthCost), 0, 2, slStr
                        End If
                    End If
            End Select
            
            If ilBox = UNBILLEDINDEX Then
                
                If imPlSave(1, ilRow) And tgCPMPostLog(ilRow).sPriceType <> "C" And tgCPMPostLog(ilRow).sPriceType <> "F" Then
                    gPaintArea pbcPostManualCntrl, tmPostCtrls(ilBox).fBoxX, tmPostCtrls(ilBox).fBoxY + (ilRow - ilStartRow) * (fgBoxGridH + 15), tmPostCtrls(ilBox).fBoxW - 15, tmPostCtrls(ilBox).fBoxH - 15, LIGHTYELLOW
                ElseIf imPlSave(1, ilRow) Then
                    llForeColor = pbcPostManualCntrl.ForeColor
                    pbcPostManualCntrl.ForeColor = DARKGREEN
                    gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(ilBox)
                    pbcPostManualCntrl.Print tmPostCtrls(ilBox).sShow
                    pbcPostManualCntrl.ForeColor = llForeColor
                Else
                    gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(ilBox)
                    pbcPostManualCntrl.Print tmPostCtrls(ilBox).sShow
                End If
            Else
                gSetShow pbcPostManualCntrl, slStr, tmPostCtrls(ilBox)
                pbcPostManualCntrl.Print tmPostCtrls(ilBox).sShow
            End If
        Next ilBox
    Next ilRow
End Sub

Private Sub mPaintPostTitle()
    Dim llColor As Long
    Dim slFontName As String
    Dim flFontSize As Single
    Dim ilLoop As Integer
    Dim llTop As Long
    Dim ilFillStyle As Integer
    Dim llFillColor As Long
    Dim ilLineCount As Integer
    Dim ilHalfY As Integer

    llColor = pbcPostManualCntrl.ForeColor
    slFontName = pbcPostManualCntrl.FontName
    flFontSize = pbcPostManualCntrl.FontSize
    ilFillStyle = pbcPostManualCntrl.FillStyle
    llFillColor = pbcPostManualCntrl.FillColor
    pbcPostManualCntrl.ForeColor = BLUE
    pbcPostManualCntrl.FontBold = False
    pbcPostManualCntrl.FontSize = 7
    pbcPostManualCntrl.FontName = "Arial"
    pbcPostManualCntrl.FontSize = 7  'Font size done twice as indicated in FontSize property area in manual

    
    pbcPostManualCntrl.Line (tmPostCtrls(CONTRACTINDEX).fBoxX - 15, 15)-Step(tmPostCtrls(CONTRACTINDEX).fBoxW + 15, tmPostCtrls(CONTRACTINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(CONTRACTINDEX).fBoxX, 30)-Step(tmPostCtrls(CONTRACTINDEX).fBoxW - 15, tmPostCtrls(CONTRACTINDEX).fBoxY - 60), LIGHTBLUE, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(CONTRACTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPostManualCntrl.Print "Contract"
    
    pbcPostManualCntrl.Line (tmPostCtrls(ADVERTISERINDEX).fBoxX - 15, 15)-Step(tmPostCtrls(ADVERTISERINDEX).fBoxW + 15, tmPostCtrls(ADVERTISERINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(ADVERTISERINDEX).fBoxX, 30)-Step(tmPostCtrls(ADVERTISERINDEX).fBoxW - 15, tmPostCtrls(ADVERTISERINDEX).fBoxY - 60), LIGHTBLUE, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(ADVERTISERINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPostManualCntrl.Print "Advertiser"
    
    pbcPostManualCntrl.Line (tmPostCtrls(IDINDEX).fBoxX - 15, 15)-Step(tmPostCtrls(IDINDEX).fBoxW + 15, tmPostCtrls(IDINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(IDINDEX).fBoxX, 30)-Step(tmPostCtrls(IDINDEX).fBoxW - 15, tmPostCtrls(IDINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(IDINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = 15
    pbcPostManualCntrl.Print "ID"
    
    pbcPostManualCntrl.Line (tmPostCtrls(VEHICLEINDEX).fBoxX - 15, 15)-Step(tmPostCtrls(VEHICLEINDEX).fBoxW + 15, tmPostCtrls(VEHICLEINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(VEHICLEINDEX).fBoxX, 30)-Step(tmPostCtrls(VEHICLEINDEX).fBoxW - 15, tmPostCtrls(VEHICLEINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(VEHICLEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = 15
    pbcPostManualCntrl.Print "Vehicle"
    
    pbcPostManualCntrl.Line (tmPostCtrls(ADLOCATIONINDEX).fBoxX - 15, 15)-Step(tmPostCtrls(ADLOCATIONINDEX).fBoxW + 15, tmPostCtrls(ADLOCATIONINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(ADLOCATIONINDEX).fBoxX, 30)-Step(tmPostCtrls(ADLOCATIONINDEX).fBoxW - 15, tmPostCtrls(ADLOCATIONINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(ADLOCATIONINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = 15
    pbcPostManualCntrl.Print "Ad"
    pbcPostManualCntrl.CurrentX = tmPostCtrls(ADLOCATIONINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = tmPostCtrls(ADLOCATIONINDEX).fBoxY / 2 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPostManualCntrl.Print "Location"
    
    pbcPostManualCntrl.Line (tmPostCtrls(LENGTHINDEX).fBoxX - 15, 15)-Step(tmPostCtrls(LENGTHINDEX).fBoxW + 15, tmPostCtrls(LENGTHINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(LENGTHINDEX).fBoxX, 30)-Step(tmPostCtrls(LENGTHINDEX).fBoxW - 15, tmPostCtrls(LENGTHINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(LENGTHINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = 15
    pbcPostManualCntrl.Print "Length"
    
    pbcPostManualCntrl.Line (tmPostCtrls(DATERANGEINDEX).fBoxX - 15, 15)-Step(tmPostCtrls(DATERANGEINDEX).fBoxW + 15, tmPostCtrls(DATERANGEINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(DATERANGEINDEX).fBoxX, 30)-Step(tmPostCtrls(DATERANGEINDEX).fBoxW - 15, tmPostCtrls(DATERANGEINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(DATERANGEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = 15
    pbcPostManualCntrl.Print "Date Range"
    
    pbcPostManualCntrl.Line (tmPostCtrls(PRICETYPEINDEX).fBoxX - 15, 15)-Step(tmPostCtrls(PRICETYPEINDEX).fBoxW + 15, tmPostCtrls(PRICETYPEINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(PRICETYPEINDEX).fBoxX, 30)-Step(tmPostCtrls(PRICETYPEINDEX).fBoxW - 15, tmPostCtrls(PRICETYPEINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(PRICETYPEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = 15
    pbcPostManualCntrl.Print "Price"
    pbcPostManualCntrl.CurrentX = tmPostCtrls(PRICETYPEINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = tmPostCtrls(PRICETYPEINDEX).fBoxY / 2 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPostManualCntrl.Print "Type"
    
'  BEGIN BOOSTER:  Change "CPM" label to "Rate"  2023-11-20 JJB
    pbcPostManualCntrl.Line (tmPostCtrls(CPMINDEX).fBoxX - 15, 15)-Step(tmPostCtrls(CPMINDEX).fBoxW + 15, tmPostCtrls(CPMINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(CPMINDEX).fBoxX, 30)-Step(tmPostCtrls(CPMINDEX).fBoxW - 15, tmPostCtrls(CPMINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(CPMINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = 15
    'pbcPostManualCntrl.Print "CPM"
    pbcPostManualCntrl.Print "Rate"
' END
    
    ilHalfY = tmPostCtrls(ORDEREDINDEX).fBoxY / 2
    Do While ilHalfY Mod 15 <> 0
        ilHalfY = ilHalfY - 1
    Loop
    
'  BEGIN BOOSTER:  Change "Impressions" label to "Quantity"  2023-11-20 JJB
    pbcPostManualCntrl.Line (tmPostCtrls(ORDEREDINDEX).fBoxX - 15, 15)-Step(tmPostCtrls(ORDEREDINDEX).fBoxW + tmPostCtrls(BILLEDINDEX).fBoxW + tmPostCtrls(TOTALINVOICEDINDEX).fBoxW + tmPostCtrls(UNBILLEDINDEX).fBoxW + 60, ilHalfY), BLUE, B
    'pbcPostManualCntrl.CurrentX = tmPostCtrls(ORDEREDINDEX).fBoxX + tmPostCtrls(ORDEREDINDEX).fBoxW / 2 + tmPostCtrls(UNBILLEDINDEX).fBoxW / 2 + tmPostCtrls(TOTALINVOICEDINDEX).fBoxW / 2 + tmPostCtrls(BILLEDINDEX).fBoxW / 2 - pbcPostManualCntrl.TextWidth("Impressions") / 2 + 15 'fgBoxInsetX
    pbcPostManualCntrl.CurrentX = tmPostCtrls(ORDEREDINDEX).fBoxX + tmPostCtrls(ORDEREDINDEX).fBoxW / 2 + tmPostCtrls(UNBILLEDINDEX).fBoxW / 2 + tmPostCtrls(TOTALINVOICEDINDEX).fBoxW / 2 + tmPostCtrls(BILLEDINDEX).fBoxW / 2 - pbcPostManualCntrl.TextWidth("Quantity") / 2 + 15 'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = 15 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    'pbcPostManualCntrl.Print "Impressions"
    pbcPostManualCntrl.Print "Quantity"
' END
    pbcPostManualCntrl.Line (tmPostCtrls(ORDEREDINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmPostCtrls(ORDEREDINDEX).fBoxW + 15, tmPostCtrls(ORDEREDINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(ORDEREDINDEX).fBoxX, ilHalfY + 30)-Step(tmPostCtrls(ORDEREDINDEX).fBoxW - 15, tmPostCtrls(ORDEREDINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(ORDEREDINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = ilHalfY + 30
    pbcPostManualCntrl.Print "Ordered"
    
    pbcPostManualCntrl.Line (tmPostCtrls(BILLEDINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmPostCtrls(BILLEDINDEX).fBoxW + 15, tmPostCtrls(BILLEDINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(BILLEDINDEX).fBoxX, ilHalfY + 30)-Step(tmPostCtrls(BILLEDINDEX).fBoxW - 15, tmPostCtrls(BILLEDINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(BILLEDINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = ilHalfY + 30
    pbcPostManualCntrl.Print "Total Billed"
    
    pbcPostManualCntrl.Line (tmPostCtrls(TOTALINVOICEDINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmPostCtrls(TOTALINVOICEDINDEX).fBoxW + 15, tmPostCtrls(TOTALINVOICEDINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(TOTALINVOICEDINDEX).fBoxX, ilHalfY + 30)-Step(tmPostCtrls(TOTALINVOICEDINDEX).fBoxW - 15, tmPostCtrls(TOTALINVOICEDINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(TOTALINVOICEDINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = ilHalfY + 30
    pbcPostManualCntrl.Print "Total Invoiced"
    
    pbcPostManualCntrl.Line (tmPostCtrls(UNBILLEDINDEX).fBoxX - 15, ilHalfY + 15)-Step(tmPostCtrls(UNBILLEDINDEX).fBoxW + 15, tmPostCtrls(UNBILLEDINDEX).fBoxY - 15), BLUE, B
    pbcPostManualCntrl.CurrentX = tmPostCtrls(UNBILLEDINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = ilHalfY + 30
    pbcPostManualCntrl.Print "Posted"
    
    pbcPostManualCntrl.Line (tmPostCtrls(MONTHCOSTINDEX).fBoxX - 15, 15)-Step(tmPostCtrls(MONTHCOSTINDEX).fBoxW + 15, tmPostCtrls(MONTHCOSTINDEX).fBoxY - 30), BLUE, B
    pbcPostManualCntrl.Line (tmPostCtrls(MONTHCOSTINDEX).fBoxX, 30)-Step(tmPostCtrls(MONTHCOSTINDEX).fBoxW - 15, tmPostCtrls(MONTHCOSTINDEX).fBoxY - 60), LIGHTYELLOW, BF
    pbcPostManualCntrl.CurrentX = tmPostCtrls(MONTHCOSTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = 15
    pbcPostManualCntrl.Print "Month"
    pbcPostManualCntrl.CurrentX = tmPostCtrls(MONTHCOSTINDEX).fBoxX + 15  'fgBoxInsetX
    pbcPostManualCntrl.CurrentY = tmPostCtrls(MONTHCOSTINDEX).fBoxY / 2 '+ (ilRow - 1) * (fgBoxGridH + 15) '+ fgBoxInsetY
    pbcPostManualCntrl.Print "Cost"
    
    
    ilLineCount = 0
    llTop = tmPostCtrls(1).fBoxY
    Do
        For ilLoop = imLBCtrls To UBound(tmPostCtrls) Step 1
        
            Select Case ilLoop
                Case UNBILLEDINDEX
                Case Else
                    pbcPostManualCntrl.FillStyle = 0 'Solid
                    pbcPostManualCntrl.FillColor = LIGHTYELLOW
            End Select
            
            pbcPostManualCntrl.Line (tmPostCtrls(ilLoop).fBoxX - 15, llTop - 15)-Step(tmPostCtrls(ilLoop).fBoxW + 15, tmPostCtrls(ilLoop).fBoxH + 15), BLUE, B
            Select Case ilLoop
                Case UNBILLEDINDEX
                Case Else
                    pbcPostManualCntrl.FillStyle = ilFillStyle
                    pbcPostManualCntrl.FillColor = llFillColor
            End Select
        Next ilLoop
        ilLineCount = ilLineCount + 1
        llTop = llTop + tmPostCtrls(1).fBoxH + 15
    Loop While llTop + tmPostCtrls(1).fBoxH < pbcPostManualCntrl.height
    vbcPostManualCntrl.LargeChange = ilLineCount - 1
    pbcPostManualCntrl.FontSize = flFontSize
    pbcPostManualCntrl.FontName = slFontName
    pbcPostManualCntrl.FontSize = flFontSize
    pbcPostManualCntrl.ForeColor = llColor
    pbcPostManualCntrl.FontBold = True
End Sub


Private Sub mSetFocus(ilBoxNo As Integer)
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    If ilBoxNo < imLBCtrls Or ilBoxNo > UBound(tmPostCtrls) Then
        Exit Sub
    End If
    
    Select Case ilBoxNo 'Branch on box type (control)
        Case UNBILLEDINDEX
            edcDropDown.Visible = True
            edcDropDown.SetFocus
    End Select
End Sub

Private Function mNumberOfMonths(slStartDate As String, slEndDate As String) As Integer
    Dim ilStartMonth As Integer
    Dim ilStartYear As Integer
    Dim ilEndMonth As Integer
    Dim ilEndYear As Integer
          
    ilStartMonth = month(slStartDate)
    ilStartYear = year(slStartDate)
    ilEndMonth = month(slEndDate)
    ilEndYear = year(slEndDate)
       
    If ilStartYear = ilEndYear Then
        mNumberOfMonths = (ilEndMonth - ilStartMonth) + 1 '10/18/2021 - Luigi: Added (+1) for bug ttp# 10308
    Else
        mNumberOfMonths = (13 - ilStartMonth) + ilEndMonth + 12 * (ilEndYear - ilStartYear - 1)
        If ((ilEndYear - ilStartYear - 1) > 0) Then
           mNumberOfMonths = mNumberOfMonths - 1
        End If
    End If
    
    If mNumberOfMonths = 0 Then
        mNumberOfMonths = 1
    End If
End Function

Private Function mCalculateFlatMonthlyCost(llRowNo As Long)
    Dim slStartDate As String
    Dim slRvfTranDate As String
    Dim slPhfTranDate As String
    Dim slEndDate As String
    Dim ilRet As Integer
    Dim lDate As Date
    Dim llDate2 As Date
    Dim llAnotherDate
    Dim ilNoWeeks As Integer
    Dim ilNoMonths As Integer
    Dim slStr As String
    
    If tgCPMPostLog(llRowNo).sChfBillCycle = "C" Then
        slStartDate = Format$(gObtainStartCal(Format$(tgCPMPostLog(llRowNo).StartDate)), "m/d/yyyy")
        slRvfTranDate = Format$(gObtainStartCal(Format$(tgCPMPostLog(llRowNo).rvfTranDate)), "m/d/yyyy")
        slPhfTranDate = Format$(gObtainStartCal(Format$(tgCPMPostLog(llRowNo).phfTranDate)), "m/d/yyyy")
        slEndDate = Format$(gObtainEndCal(Format$(tgCPMPostLog(llRowNo).EndDate)), "m/d/yyyy")
    Else
        slStartDate = Format$(gObtainStartStd(Format$(tgCPMPostLog(llRowNo).StartDate)), "m/d/yyyy")
        slRvfTranDate = Format$(gObtainStartStd(Format$(tgCPMPostLog(llRowNo).rvfTranDate)), "m/d/yyyy")
        slPhfTranDate = Format$(gObtainStartStd(Format$(tgCPMPostLog(llRowNo).phfTranDate)), "m/d/yyyy")
        slEndDate = Format$(gObtainEndStd(Format$(tgCPMPostLog(llRowNo).EndDate)), "m/d/yyyy")
    End If
    
    If gDateValue(slRvfTranDate) > gDateValue(slStartDate) Then
        slStartDate = slRvfTranDate
    End If
    
    If gDateValue(slPhfTranDate) > gDateValue(slStartDate) Then
        slStartDate = slPhfTranDate
    End If
        
    '10/21/2021 - Luigi: fix to calculate date difference in number of months for broadcast calendar (ttp# 10308)
    If tgCPMPostLog(llRowNo).sChfBillCycle = "C" Then
        ilRet = (1 + DateDiff("m", slStartDate, slEndDate))
    Else
        lDate = gDateValue(slStartDate)
        llDate2 = gDateValue(slEndDate)
        ilNoWeeks = (llDate2 - lDate) / 7 + 1
        'determine # std bdcst months from start to end
        llAnotherDate = lDate
        ilNoMonths = 0
        Do While llAnotherDate < llDate2
            slStr = gObtainEndStd(Format$(llAnotherDate, "m/d/yy"))
            llAnotherDate = gDateValue(slStr) + 1
            ilNoMonths = ilNoMonths + 1
        Loop
        ilRet = ilNoMonths
    End If
    
    ' The ilRet value is now modified to return the total number of months but subtracting the number of months that have already been invioced/billed
    ilRet = ilRet - mBilledMonthCount(tgCPMPostLog(llRowNo).lContractNumber, tgCPMPostLog(llRowNo).iID) ' TTP 10579 JJB
    
    If ilRet > 0 Then
        mCalculateFlatMonthlyCost = ((tgCPMPostLog(llRowNo).lPcfTotalCost / 100) - ((tgCPMPostLog(llRowNo).lRvfTotalGross / 100) + (tgCPMPostLog(llRowNo).lPhfTotalGross / 100))) / ilRet ' Booster v2 issue 6 fix
    Else
        mCalculateFlatMonthlyCost = ((tgCPMPostLog(llRowNo).lPcfTotalCost / 100) - ((tgCPMPostLog(llRowNo).lRvfTotalGross / 100) + (tgCPMPostLog(llRowNo).lPhfTotalGross / 100))) ' Booster v2 issue 6 fix
    End If
    
    mCalculateFlatMonthlyCost = Round(mCalculateFlatMonthlyCost, 2)
    
End Function


Private Function mReadonlyRow(llRowNo As Long) As Integer
    
    If tgCPMPostLog(llRowNo).sPriceType <> "C" And tgCPMPostLog(llRowNo).sPriceType <> "F" Then
        mReadonlyRow = True
        Exit Function
    End If
    
    If tgCPMPostLog(llRowNo).lBilled > -1 And Not tgCPMPostLog(llRowNo).lUnBilled > 0 Then
        mReadonlyRow = True
        Exit Function
    End If
    
    mReadonlyRow = False
End Function


Private Sub mResortLog()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slStr1 As String
    Dim ilRet As Integer

    pbcPostManualCntrl.Cls
    ilRet = 0
    On Error GoTo mResortLogErr
    
    If imCurSort > -1 Then
        If imCurSort = 0 Then
            If imSortOrder = 0 Then
               For ilLoop = imLBPost To UBound(tgTempCPMPostLog)
                    slStr = tgTempCPMPostLog(ilLoop).lContractNumber
                    Do While Len(slStr) < 5
                        slStr = "0" & slStr
                    Loop
                    
                    slStr1 = tgTempCPMPostLog(ilLoop).iID
                    Do While Len(slStr1) < 5
                        slStr1 = "0" & slStr1
                    Loop
                    tgTempCPMPostLog(ilLoop).sKey = slStr & "|" & slStr1
               Next ilLoop
               ArraySortTyp fnAV(tgTempCPMPostLog(), 1), UBound(tgTempCPMPostLog), 0, LenB(tgTempCPMPostLog(1)), 0, LenB(tgTempCPMPostLog(1).sKey), 0
               
            Else
                For ilLoop = imLBPost To UBound(tgTempCPMPostLog)
                    slStr = tgTempCPMPostLog(ilLoop).lContractNumber
                    Do While Len(slStr) < 5
                        slStr = "#" & slStr
                    Loop
                    
                    slStr1 = CLng(10000 / tgTempCPMPostLog(ilLoop).iID)
                    Do While Len(slStr1) < 5
                        slStr1 = "#" & slStr1
                    Loop
                    tgTempCPMPostLog(ilLoop).sKey = slStr & "|" & slStr1
               Next ilLoop
               ArraySortTyp fnAV(tgTempCPMPostLog(), 1), UBound(tgTempCPMPostLog), 1, LenB(tgTempCPMPostLog(1)), 0, LenB(tgTempCPMPostLog(1).sKey), 0
            End If
            
        End If
        If imCurSort = 1 Then
            If imSortOrder = 0 Then
               For ilLoop = imLBPost To UBound(tgTempCPMPostLog)
                    slStr = tgTempCPMPostLog(ilLoop).lContractNumber
                    Do While Len(slStr) < 5
                        slStr = "0" & slStr
                    Loop
                    
                    slStr1 = tgTempCPMPostLog(ilLoop).iID
                    Do While Len(slStr1) < 5
                        slStr1 = "0" & slStr1
                    Loop
                    tgTempCPMPostLog(ilLoop).sKey = tgTempCPMPostLog(ilLoop).sAdfName & "|" & slStr & "|" & slStr1
               Next ilLoop
               ArraySortTyp fnAV(tgTempCPMPostLog(), 1), UBound(tgTempCPMPostLog), 0, LenB(tgTempCPMPostLog(1)), 0, LenB(tgTempCPMPostLog(1).sKey), 0
            Else
               For ilLoop = imLBPost To UBound(tgTempCPMPostLog)
                    slStr = CLng(10000 / tgTempCPMPostLog(ilLoop).lContractNumber)
                    Do While Len(slStr) < 5
                        slStr = "#" & slStr
                    Loop
                    
                    slStr1 = CLng(10000 / tgTempCPMPostLog(ilLoop).iID)
                    Do While Len(slStr1) < 5
                        slStr1 = "#" & slStr1
                    Loop
                    tgTempCPMPostLog(ilLoop).sKey = tgTempCPMPostLog(ilLoop).sAdfName & "|" & slStr & "|" & slStr1
               Next ilLoop
               ArraySortTyp fnAV(tgTempCPMPostLog(), 1), UBound(tgTempCPMPostLog), 1, LenB(tgTempCPMPostLog(1)), 0, LenB(tgTempCPMPostLog(1).sKey), 0
            End If
       End If
    End If
    For ilLoop = imLBPost To UBound(tgTempCPMPostLog)
        tgCPMPostLog(ilLoop) = tgTempCPMPostLog(ilLoop)
    Next ilLoop
    mMoveRecToCtrl
    pbcPostManualCntrl_Paint
    Exit Sub
mResortLogErr:
    ilRet = 1
    Resume Next
End Sub


