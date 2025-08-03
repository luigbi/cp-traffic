VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmAffDP 
   Caption         =   "Dayparts"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   Icon            =   "AffDP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   4230
   Visible         =   0   'False
   Begin VB.Frame frcStatus 
      Caption         =   "Pledge Status"
      Height          =   1425
      Left            =   405
      TabIndex        =   6
      Top             =   2595
      Width           =   3390
      Begin VB.OptionButton rbcStatus 
         Caption         =   "Air Cmml Only"
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   9
         Top             =   1035
         Width           =   2400
      End
      Begin VB.OptionButton rbcStatus 
         Caption         =   "Delay Cmml/Prg"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   8
         Top             =   675
         Width           =   2295
      End
      Begin VB.OptionButton rbcStatus 
         Caption         =   "Air in Daypart"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Top             =   315
         Value           =   -1  'True
         Width           =   2145
      End
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All Dayparts"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   345
      TabIndex        =   4
      Top             =   45
      Width           =   2010
   End
   Begin VB.ListBox lbcDayParts 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      ItemData        =   "AffDP.frx":08CA
      Left            =   345
      List            =   "AffDP.frx":08D1
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   570
      Width           =   3465
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3705
      Top             =   90
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   4845
      FormDesignWidth =   4230
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2340
      TabIndex        =   1
      Top             =   4290
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Label lblDayPart 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   285
      Width           =   3450
   End
   Begin VB.Label lblIndex 
      Height          =   255
      Left            =   105
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmAffDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmAffDP - Affiliate Daypart Information
'*
'*  Created August,2001 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer
Private imState As Integer
Private imDPNum As Integer   'total number of elements in the tmAllDayParts array + 1
Private rst_daypart As ADODB.Recordset
Private Enum eWeekDays
                MON = 0
                TUE = 1
                WED = 2
                THU = 3
                FRI = 4
                SAT = 5
                SUN = 6
            End Enum
        
Private Type ALLDAYPARTS
            sName As String * 20
            sInOut As String * 1
            sAnfCode As String * 2
            lStartTime As Long
            lEndTime As Long
            eWeekDays(MON To SUN) As String * 1  'The days that the daypart does or does not run
            iShowDP As Boolean                   'True show daypart in listbox, False hide daypart
        End Type

Private tmAllDayParts() As ALLDAYPARTS

Private Sub ClearControls()
    imState = 0
    imFieldChgd = False
End Sub

Private Sub BindControls()
    'lblIndex.Caption = rst(0).Value
    'txtFName.Text = Trim$(rst(1).Value)
    'txtLName.Text = Trim$(rst(2).Value)
    'txtPhone.Text = Trim$(rst(3).Value)
    'txtFax.Text = Trim$(rst(4).Value)
    'txtEMail.Text = Trim$(rst(5).Value)
    'imState = rst(6).Value
    'imFieldChgd = False
End Sub

Private Sub chkAll_Click()
    Dim lRet As Long
    Dim lRg As Long
    Dim iValue As Integer
    
    If chkAll.Value = 1 Then
        iValue = True
    Else
        iValue = False
    End If
    If lbcDayParts.ListCount > 0 Then
        lRg = CLng(lbcDayParts.ListCount - 1) * &H10000 Or 0
        lRet = SendMessageByNum(lbcDayParts.hwnd, LB_SELITEMRANGE, iValue, lRg)
    End If

End Sub

Private Sub cmdCancel_Click()
    
    'If the user cancels out of the daypart screen reset the radio buttons
    'back to all false
    'frmAgmnt!optTimeType(0).Value = False
    'frmAgmnt!optTimeType(1).Value = False
    'frmAgmnt!optTimeType(2).Value = False
    igReload = False
    igReturnPledgeStatus = 0

    Unload frmAffDP
End Sub

Private Sub cmdOk_Click()

    Dim ilRet As Integer

    On Error GoTo ErrHand
    
    'If Not frmAgmnt!optExAll(1).Value Then
        'If UBound(tgDat) > LBound(tgDat) Then
        If igPledgeExist Then
            ilRet = gMsgBox("Warning: Avails Will Be Removed Prior To Adding Dayparts!", vbOKCancel)
            If ilRet = vbCancel Then
                Exit Sub
            End If
        End If
    'End If
    If rbcStatus(1).Value Then
        igReturnPledgeStatus = 1
    ElseIf rbcStatus(2).Value Then
        igReturnPledgeStatus = 2
    Else
        igReturnPledgeStatus = 0
    End If
    mFillGridWithDayParts
    Unload frmAffDP
    
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "frmAffDP-cmdOk"
End Sub

Private Sub Form_Load()
   
    Dim ilIdx As Integer
    Dim ilRet As Integer
    
    
    Screen.MousePointer = vbHourglass
    frmAffDP.Caption = "Dayparts - " & sgClientName
    Me.Width = (Screen.Width) / 1.55
    Me.Height = (Screen.Height) / 2.4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
     
    On Error GoTo ErrHand
    
    mGetAllDayparts
    mAdjustTime
    mShowDPList

    Screen.MousePointer = vbDefault
    Exit Sub

ErrHand:
    gMsg = ""
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmAfDP-Form-Load: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Erase tmAllDayParts
    Set frmAffDP = Nothing
End Sub

Private Sub optState_Click(Index As Integer)
    imFieldChgd = True
End Sub

Private Sub mGetAllDayparts()

    Dim ilWeekDay As Integer
    Dim llSTime As Long
    Dim llETime As Long
    Dim ilDone As Integer
    Dim ilIdx As Integer
    Dim slTempDP As String
   
    On Error GoTo ErrHand
    
    imDPNum = 0
    ReDim tmAllDayParts(0 To 0) As ALLDAYPARTS
    For ilIdx = 0 To (UBound(tgRdfCodes) - 1) Step 1
        'get the dayparts out of rdf
        SQLQuery = "SELECT * "
        SQLQuery = SQLQuery & " FROM RDF_Standard_Daypart"
        SQLQuery = SQLQuery + " WHERE (rdfCode = " & tgRdfCodes(ilIdx) & ")"
        SQLQuery = SQLQuery + " ORDER BY rdfName"
        
        Set rst_daypart = gSQLSelectCall(SQLQuery)
        If rst_daypart.EOF Then
            Exit Sub
        End If
        
        While Not rst_daypart.EOF
            tmAllDayParts(imDPNum).iShowDP = True 'Set default to True
        
            'This is the 7th element of 7.  A daypart can have up to 7 different times associated
            'with it.  Exp. A daypart named "Drive Time" might have times of 6a-10a and 3p-7p.
            'When dayparts are created in RDF the first time is in the 7th element, the second time
            'is in the 6th element and so on. So we know that at least the 7th element must be a daypart
            ilDone = False
            tmAllDayParts(imDPNum).sName = Trim$(rst_daypart!rdfName)
            tmAllDayParts(imDPNum).sInOut = rst_daypart!rdfInOut
            tmAllDayParts(imDPNum).sAnfCode = rst_daypart!rdfAnfCode
            
            tmAllDayParts(imDPNum).lStartTime = gTimeToLong(rst_daypart!rdfStartTime7, False)
            tmAllDayParts(imDPNum).lEndTime = gTimeToLong(rst_daypart!rdfEndTime7, False)
            tmAllDayParts(imDPNum).eWeekDays(MON) = rst_daypart!rdfMo7
            tmAllDayParts(imDPNum).eWeekDays(TUE) = rst_daypart!rdfTu7
            tmAllDayParts(imDPNum).eWeekDays(WED) = rst_daypart!rdfWe7
            tmAllDayParts(imDPNum).eWeekDays(THU) = rst_daypart!rdfTh7
            tmAllDayParts(imDPNum).eWeekDays(FRI) = rst_daypart!rdfFr7
            tmAllDayParts(imDPNum).eWeekDays(SAT) = rst_daypart!rdfSa7
            tmAllDayParts(imDPNum).eWeekDays(SUN) = rst_daypart!rdfSu7
            
            imDPNum = imDPNum + 1
            ReDim Preserve tmAllDayParts(0 To imDPNum) As ALLDAYPARTS
    
            '***************************************************************************
            
            llSTime = gTimeToLong(rst_daypart!rdfStartTime6, False)
            llETime = gTimeToLong(rst_daypart!rdfEndTime6, False)
            'For the 6th thru the 1st elements we test to see if start and end time are 0.
            'If both are 0 (12m - 12m) then we assume that this is not a daypart. This is
            'the default times in RDF before a daypart is entered.  Per Dick Levine
            If llSTime = 0 And llETime = 0 Then
                ilDone = True
            Else
                tmAllDayParts(imDPNum).sName = Trim$(rst_daypart!rdfName)
                tmAllDayParts(imDPNum).sInOut = rst_daypart!rdfInOut
                tmAllDayParts(imDPNum).sAnfCode = rst_daypart!rdfAnfCode
                tmAllDayParts(imDPNum).lStartTime = llSTime
                tmAllDayParts(imDPNum).lEndTime = llETime
                
                tmAllDayParts(imDPNum).eWeekDays(MON) = rst_daypart!rdfMo6
                tmAllDayParts(imDPNum).eWeekDays(TUE) = rst_daypart!rdfTu6
                tmAllDayParts(imDPNum).eWeekDays(WED) = rst_daypart!rdfWe6
                tmAllDayParts(imDPNum).eWeekDays(THU) = rst_daypart!rdfTh6
                tmAllDayParts(imDPNum).eWeekDays(FRI) = rst_daypart!rdfFr6
                tmAllDayParts(imDPNum).eWeekDays(SAT) = rst_daypart!rdfSa6
                tmAllDayParts(imDPNum).eWeekDays(SUN) = rst_daypart!rdfSu6
                
                imDPNum = imDPNum + 1
                ReDim Preserve tmAllDayParts(0 To imDPNum) As ALLDAYPARTS
            End If
            
            '***************************************************************************
            
            If Not ilDone Then
                llSTime = gTimeToLong(rst_daypart!rdfStartTime5, False)
                llETime = gTimeToLong(rst_daypart!rdfEndTime5, False)
                If llSTime = 0 And llETime = 0 Then
                    ilDone = True
                Else
                    tmAllDayParts(imDPNum).sName = Trim$(rst_daypart!rdfName)
                    tmAllDayParts(imDPNum).sInOut = rst_daypart!rdfInOut
                    tmAllDayParts(imDPNum).sAnfCode = rst_daypart!rdfAnfCode
                    tmAllDayParts(imDPNum).lStartTime = llSTime
                    tmAllDayParts(imDPNum).lEndTime = llETime
                    
                    tmAllDayParts(imDPNum).eWeekDays(MON) = rst_daypart!rdfMo5
                    tmAllDayParts(imDPNum).eWeekDays(TUE) = rst_daypart!rdfTu5
                    tmAllDayParts(imDPNum).eWeekDays(WED) = rst_daypart!rdfWe5
                    tmAllDayParts(imDPNum).eWeekDays(THU) = rst_daypart!rdfTh5
                    tmAllDayParts(imDPNum).eWeekDays(FRI) = rst_daypart!rdfFr5
                    tmAllDayParts(imDPNum).eWeekDays(SAT) = rst_daypart!rdfSa5
                    tmAllDayParts(imDPNum).eWeekDays(SUN) = rst_daypart!rdfSu5
                    
                    imDPNum = imDPNum + 1
                    ReDim Preserve tmAllDayParts(0 To imDPNum) As ALLDAYPARTS
                End If
            End If
            
            '***************************************************************************
            
            If Not ilDone Then
                llSTime = gTimeToLong(rst_daypart!rdfStartTime4, False)
                llETime = gTimeToLong(rst_daypart!rdfEndTime4, False)
                If llSTime = 0 And llETime = 0 Then
                    ilDone = True
                Else
                    tmAllDayParts(imDPNum).sName = Trim$(rst_daypart!rdfName)
                    tmAllDayParts(imDPNum).sInOut = rst_daypart!rdfInOut
                    tmAllDayParts(imDPNum).sAnfCode = rst_daypart!rdfAnfCode
                    tmAllDayParts(imDPNum).lStartTime = llSTime
                    tmAllDayParts(imDPNum).lEndTime = llETime
                    
                    tmAllDayParts(imDPNum).eWeekDays(MON) = rst_daypart!rdfMo4
                    tmAllDayParts(imDPNum).eWeekDays(TUE) = rst_daypart!rdfTu4
                    tmAllDayParts(imDPNum).eWeekDays(WED) = rst_daypart!rdfWe4
                    tmAllDayParts(imDPNum).eWeekDays(THU) = rst_daypart!rdfTh4
                    tmAllDayParts(imDPNum).eWeekDays(FRI) = rst_daypart!rdfFr4
                    tmAllDayParts(imDPNum).eWeekDays(SAT) = rst_daypart!rdfSa4
                    tmAllDayParts(imDPNum).eWeekDays(SUN) = rst_daypart!rdfSu4
                    
                    imDPNum = imDPNum + 1
                    ReDim Preserve tmAllDayParts(0 To imDPNum) As ALLDAYPARTS
                End If
            End If
            
            '***************************************************************************
            
            If Not ilDone Then
                llSTime = gTimeToLong(rst_daypart!rdfStartTime3, False)
                llETime = gTimeToLong(rst_daypart!rdfEndTime3, False)
                If llSTime = 0 And llETime = 0 Then
                    ilDone = True
                Else
                    tmAllDayParts(imDPNum).sName = Trim$(rst_daypart!rdfName)
                    tmAllDayParts(imDPNum).sInOut = rst_daypart!rdfInOut
                    tmAllDayParts(imDPNum).sAnfCode = rst_daypart!rdfAnfCode
                    tmAllDayParts(imDPNum).lStartTime = llSTime
                    tmAllDayParts(imDPNum).lEndTime = llETime
                    
                    tmAllDayParts(imDPNum).eWeekDays(MON) = rst_daypart!rdfMo3
                    tmAllDayParts(imDPNum).eWeekDays(TUE) = rst_daypart!rdfTu3
                    tmAllDayParts(imDPNum).eWeekDays(WED) = rst_daypart!rdfWe3
                    tmAllDayParts(imDPNum).eWeekDays(THU) = rst_daypart!rdfTh3
                    tmAllDayParts(imDPNum).eWeekDays(FRI) = rst_daypart!rdfFr3
                    tmAllDayParts(imDPNum).eWeekDays(SAT) = rst_daypart!rdfSa3
                    tmAllDayParts(imDPNum).eWeekDays(SUN) = rst_daypart!rdfSu3
                    
                    imDPNum = imDPNum + 1
                    ReDim Preserve tmAllDayParts(0 To imDPNum) As ALLDAYPARTS
                End If
            End If
            
            '***************************************************************************
            
            If Not ilDone Then
                llSTime = gTimeToLong(rst_daypart!rdfStartTime2, False)
                llETime = gTimeToLong(rst_daypart!rdfEndTime2, False)
                If llSTime = 0 And llETime = 0 Then
                    ilDone = True
                Else
                    tmAllDayParts(imDPNum).sName = Trim$(rst_daypart!rdfName)
                    tmAllDayParts(imDPNum).sInOut = rst_daypart!rdfInOut
                    tmAllDayParts(imDPNum).sAnfCode = rst_daypart!rdfAnfCode
                    tmAllDayParts(imDPNum).lStartTime = llSTime
                    tmAllDayParts(imDPNum).lEndTime = llETime
                    
                    tmAllDayParts(imDPNum).eWeekDays(MON) = rst_daypart!rdfMo2
                    tmAllDayParts(imDPNum).eWeekDays(TUE) = rst_daypart!rdfTu2
                    tmAllDayParts(imDPNum).eWeekDays(WED) = rst_daypart!rdfWe2
                    tmAllDayParts(imDPNum).eWeekDays(THU) = rst_daypart!rdfTh2
                    tmAllDayParts(imDPNum).eWeekDays(FRI) = rst_daypart!rdfFr2
                    tmAllDayParts(imDPNum).eWeekDays(SAT) = rst_daypart!rdfSa2
                    tmAllDayParts(imDPNum).eWeekDays(SUN) = rst_daypart!rdfSu2
                    
                    imDPNum = imDPNum + 1
                    ReDim Preserve tmAllDayParts(0 To imDPNum) As ALLDAYPARTS
                End If
            End If
            
            '***************************************************************************
            
            If Not ilDone Then
                llSTime = gTimeToLong(rst_daypart!rdfStartTime1, False)
                llETime = gTimeToLong(rst_daypart!rdfEndTime1, False)
                If llSTime = 0 And llETime = 0 Then
                    ilDone = True
                Else
                    tmAllDayParts(imDPNum).sName = Trim$(rst_daypart!rdfName)
                    tmAllDayParts(imDPNum).sInOut = rst_daypart!rdfInOut
                    tmAllDayParts(imDPNum).sAnfCode = rst_daypart!rdfAnfCode
                    tmAllDayParts(imDPNum).lStartTime = llSTime
                    tmAllDayParts(imDPNum).lEndTime = llETime
                    
                    tmAllDayParts(imDPNum).eWeekDays(MON) = rst_daypart!rdfMo1
                    tmAllDayParts(imDPNum).eWeekDays(TUE) = rst_daypart!rdfTu1
                    tmAllDayParts(imDPNum).eWeekDays(WED) = rst_daypart!rdfWe1
                    tmAllDayParts(imDPNum).eWeekDays(THU) = rst_daypart!rdfTh1
                    tmAllDayParts(imDPNum).eWeekDays(FRI) = rst_daypart!rdfFr1
                    tmAllDayParts(imDPNum).eWeekDays(SAT) = rst_daypart!rdfSa1
                    tmAllDayParts(imDPNum).eWeekDays(SUN) = rst_daypart!rdfSu1
                    
                    imDPNum = imDPNum + 1
                    ReDim Preserve tmAllDayParts(0 To imDPNum) As ALLDAYPARTS
                End If
            End If
                      
            rst_daypart.MoveNext
        Wend
    Next ilIdx
    
    Exit Sub

ErrHand:
    gHandleError "AffErrorLog.txt", "frmAfDP-General-mGetAllDayParts"
End Sub

Private Sub mShowDPList()
    
    Dim ilIdx As Integer
    Dim slTempDP As String
    Dim slDaysOfWeek As String
    
    On Error GoTo ErrHand
    
    lbcDayParts.Clear
    For ilIdx = 0 To imDPNum - 1 Step 1
        If tmAllDayParts(ilIdx).iShowDP <> False Then
            slTempDP = Trim$(tmAllDayParts(ilIdx).sName)
            Do While Len(slTempDP) < 20
                slTempDP = slTempDP & " "
            Loop
            'make sure that the columns never run together
            slTempDP = slTempDP & " "
            slDaysOfWeek = ""
            
            If tmAllDayParts(ilIdx).eWeekDays(MON) = "Y" Then
                slDaysOfWeek = slDaysOfWeek & " Mo"
            Else
                slDaysOfWeek = slDaysOfWeek & " --"
            End If
            
            If tmAllDayParts(ilIdx).eWeekDays(TUE) = "Y" Then
                slDaysOfWeek = slDaysOfWeek & " Tu"
            Else
                slDaysOfWeek = slDaysOfWeek & " --"
            End If
            
            If tmAllDayParts(ilIdx).eWeekDays(WED) = "Y" Then
                slDaysOfWeek = slDaysOfWeek & " We"
            Else
                slDaysOfWeek = slDaysOfWeek & " --"
            End If
            
            If tmAllDayParts(ilIdx).eWeekDays(THU) = "Y" Then
                slDaysOfWeek = slDaysOfWeek & " Th"
            Else
                slDaysOfWeek = slDaysOfWeek & " --"
            End If
            
            If tmAllDayParts(ilIdx).eWeekDays(FRI) = "Y" Then
                slDaysOfWeek = slDaysOfWeek & " Fr"
            Else
                slDaysOfWeek = slDaysOfWeek & " --"
            End If
            
            If tmAllDayParts(ilIdx).eWeekDays(SAT) = "Y" Then
                slDaysOfWeek = slDaysOfWeek & " Sa"
            Else
                slDaysOfWeek = slDaysOfWeek & " --"
            End If
            
            If tmAllDayParts(ilIdx).eWeekDays(SUN) = "Y" Then
                slDaysOfWeek = slDaysOfWeek & " Su"
            Else
                slDaysOfWeek = slDaysOfWeek & " --"
            End If
            
            Do While Len(slDaysOfWeek) < 25
                slDaysOfWeek = slDaysOfWeek & " "
            Loop
            
            slTempDP = slTempDP & slDaysOfWeek
            'make sure that the columns never run together
            slTempDP = slTempDP & " "
            slTempDP = slTempDP & Format$(gLongToTime(tmAllDayParts(ilIdx).lStartTime), sgShowTimeWOSecForm)
            slTempDP = slTempDP & " - "
            slTempDP = slTempDP & Format$(gLongToTime(tmAllDayParts(ilIdx).lEndTime), sgShowTimeWOSecForm)
            
            'If igLiveDayPart Then 'Live Daypart
            '    lblDayPart.Caption = "   Daypart                 Feed Days            Times"
            'ElseIf igCDTapeDayPart Then  'CD/Tape Daypart
            '    lblDayPart.Caption = "   Daypart                 Sold Days            Times"
            'Else
            '    lblDayPart.Caption = "   Daypart                 Unknown              Times"
            'End If
            
            If igLiveDayPart Then 'Live Daypart
                lblDayPart.Caption = "   Daypart                 Feed Days            Times"
            ElseIf igCDTapeDayPart Then  'CD/Tape Daypart
                lblDayPart.Caption = "   Daypart                 Sold Days            Times"
            Else
                lblDayPart.Caption = "   Daypart                 Unknown              Times"
            End If
            
            lbcDayParts.AddItem slTempDP
            lbcDayParts.ItemData(lbcDayParts.NewIndex) = ilIdx
        End If
    Next ilIdx
   
    Exit Sub

ErrHand:
    gMsg = ""
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmAfDP-General-mShowDPList: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub mFillGridWithDayParts()
 
    Dim ilIdx As Integer
    Dim ilWkDayIdx As Integer
    Dim ilOnOff As Integer
    
    On Error GoTo ErrHand
    ReDim tgDat(0 To 0)
    
    'Populate the tgDat array with the user selected dayparts. tgDat then populates the Affiliate -
    'Pledge screen.
    For ilIdx = 0 To (lbcDayParts.ListCount - 1)
        If lbcDayParts.Selected(ilIdx) Then
            'Set the feed and plegde times both the same as Live
            tgDat(UBound(tgDat)).sFdSTime = Format$(gLongToTime(tmAllDayParts(lbcDayParts.ItemData(ilIdx)).lStartTime), sgShowTimeWOSecForm)
            tgDat(UBound(tgDat)).sFdETime = Format$(gLongToTime(tmAllDayParts(lbcDayParts.ItemData(ilIdx)).lEndTime), sgShowTimeWOSecForm)
            tgDat(UBound(tgDat)).sPdSTime = Format$(gLongToTime(tmAllDayParts(lbcDayParts.ItemData(ilIdx)).lStartTime), sgShowTimeWOSecForm)
            tgDat(UBound(tgDat)).sPdETime = Format$(gLongToTime(tmAllDayParts(lbcDayParts.ItemData(ilIdx)).lEndTime), sgShowTimeWOSecForm)
            
            ''Set the type of daypart - live = 0; cd/tape = 2; 1 = avails - will not be seen here
            'If igLiveDayPart = True Then
            '    tgDat(UBound(tgDat)).iDACode = 0
            'ElseIf igCDTapeDayPart Then
            '    tgDat(UBound(tgDat)).iDACode = 2
            'Else 'Avails - this should will never happen here
            '    tgDat(UBound(tgDat)).iDACode = 1
            'End If
                           
            'Set the days of the week
            For ilWkDayIdx = MON To SUN Step 1
                ilOnOff = 0
                If tmAllDayParts(lbcDayParts.ItemData(ilIdx)).eWeekDays(ilWkDayIdx) = "Y" Then
                    ilOnOff = 1
                End If
                tgDat(UBound(tgDat)).iFdDay(ilWkDayIdx) = ilOnOff
                tgDat(UBound(tgDat)).iPdDay(ilWkDayIdx) = ilOnOff
            Next ilWkDayIdx
            tgDat(UBound(tgDat)).sPdDayFed = ""
            ReDim Preserve tgDat(0 To (UBound(tgDat) + 1))
        End If
    Next ilIdx
    
    Exit Sub

ErrHand:
    gMsg = ""
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmAfDP-General-mFillGridWithDayParts: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub

Private Sub mAdjustTime()

    
    Dim ilTimeAdj As Integer
    Dim ilLoop As Integer
    Dim ilVefIdx As Integer
    Dim ilZoneIdx As Integer
    Dim llSTime As Long
    Dim llETime As Long
            
    On Error GoTo ErrHand
    
    'We only adjust times for Live Day Parts not CD/Tape Day Parts
    If Not igLiveDayPart Then
        Exit Sub
    End If
    
    ilTimeAdj = 0
    If (igDayPartShttCode > 0) And (igDayPartVefCode > 0) Then
        For ilLoop = LBound(tgStationInfo) To UBound(tgStationInfo) - 1 Step 1
            If tgStationInfo(ilLoop).iCode = igDayPartShttCode Then
                For ilVefIdx = LBound(tgVehicleInfo) To UBound(tgVehicleInfo) - 1 Step 1
                    If tgVehicleInfo(ilVefIdx).iCode = igDayPartVefCode Then
                        For ilZoneIdx = LBound(tgVehicleInfo(ilVefIdx).sZone) To UBound(tgVehicleInfo(ilVefIdx).sZone) Step 1
                            If StrComp(tgStationInfo(ilLoop).sZone, tgVehicleInfo(ilVefIdx).sZone(ilZoneIdx), 1) = 0 Then
                                ilTimeAdj = tgVehicleInfo(ilVefIdx).iVehLocalAdj(ilZoneIdx)
                                Exit For
                            End If
                        Next ilZoneIdx
                        Exit For
                    End If
                Next ilVefIdx
                Exit For
            End If
        Next ilLoop
    End If
    
    For ilLoop = 0 To UBound(tmAllDayParts) - 1 Step 1
        'For now it does not make sense to adjust 12m - 12m times
        If Not ((tmAllDayParts(ilLoop).lStartTime = 0) And (tmAllDayParts(ilLoop).lEndTime = 0)) Then
            llSTime = tmAllDayParts(ilLoop).lStartTime + 3600 * ilTimeAdj
            'If the end time is 12m then we need to add a day (86400). Dayparts ending at 12m with
            'negative time zone corrections would be thrown out otherwise.  Exp. 7p-12m would
            'be (12m = zero) minus the time offset.
            If tmAllDayParts(ilLoop).lEndTime = 0 Then
                tmAllDayParts(ilLoop).lEndTime = tmAllDayParts(ilLoop).lEndTime + 86400
            End If
            llETime = tmAllDayParts(ilLoop).lEndTime + 3600 * ilTimeAdj
            
            If (llSTime < 0) Or (llSTime > 86400) Or (llETime < 0) Or (llETime > 86400) Then
                tmAllDayParts(ilLoop).iShowDP = False 'Default is True
            Else
                tmAllDayParts(ilLoop).lStartTime = llSTime
                tmAllDayParts(ilLoop).lEndTime = llETime
           End If
        End If
    Next ilLoop
    
    Exit Sub

ErrHand:
    gMsg = ""
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmAffDP-General-mAdjustTime: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If

End Sub
