VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmContactEMail 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Station E-Mail"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7230
   ScaleWidth      =   9480
   Begin VB.TextBox edcFromName 
      Height          =   315
      Left            =   1290
      TabIndex        =   25
      Top             =   3045
      Width           =   2295
   End
   Begin VB.PictureBox pbcTextWidth 
      Height          =   255
      Left            =   285
      ScaleHeight     =   195
      ScaleWidth      =   495
      TabIndex        =   23
      Top             =   4425
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox cbcConcern 
      Height          =   315
      ItemData        =   "AffContactEMail.frx":0000
      Left            =   4800
      List            =   "AffContactEMail.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   2670
      Width           =   4380
   End
   Begin VB.PictureBox pbcTitle 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   1665
      TabIndex        =   18
      Top             =   450
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.ListBox lbcStations 
      Height          =   1815
      ItemData        =   "AffContactEMail.frx":0004
      Left            =   225
      List            =   "AffContactEMail.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   465
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.ListBox lbcTitle 
      Height          =   1815
      ItemData        =   "AffContactEMail.frx":0008
      Left            =   2160
      List            =   "AffContactEMail.frx":000A
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   465
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.CommandButton cmcBrowse 
      Caption         =   "Browse..."
      Height          =   300
      Left            =   7965
      TabIndex        =   12
      Top             =   5310
      Width           =   1425
   End
   Begin VB.TextBox edcAttachment 
      Height          =   315
      Left            =   1545
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5295
      Width           =   6345
   End
   Begin VB.PictureBox pbcSTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   0
      Width           =   45
   End
   Begin VB.PictureBox pbcTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   75
      ScaleHeight     =   60
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   1650
      Width           =   45
   End
   Begin VB.CheckBox ckcEMail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   5085
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "A check indicates that Monday is an airing day for a weekly buy"
      Top             =   1635
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Timer tmcStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8010
      Top             =   5820
   End
   Begin VB.CommandButton cmcSend 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2910
      TabIndex        =   13
      Top             =   5745
      Width           =   2010
   End
   Begin VB.TextBox edcMessage 
      Height          =   1230
      Left            =   1245
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3990
      Width           =   8130
   End
   Begin VB.TextBox edcSubject 
      Height          =   315
      Left            =   1245
      TabIndex        =   7
      Top             =   3540
      Width           =   8130
   End
   Begin VB.CommandButton cmcExit 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5250
      TabIndex        =   0
      Top             =   5745
      Width           =   2010
   End
   Begin VB.ListBox lbcResults 
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   645
      ItemData        =   "AffContactEMail.frx":000C
      Left            =   960
      List            =   "AffContactEMail.frx":0013
      TabIndex        =   15
      Top             =   6300
      Visible         =   0   'False
      Width           =   8445
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdContact 
      Height          =   1005
      Left            =   2160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   255
      Visible         =   0   'False
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   1773
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483634
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8715
      Top             =   5760
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   7230
      FormDesignWidth =   9480
   End
   Begin MSComDlg.CommonDialog cdcBrowse 
      Left            =   9210
      Top             =   5715
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.TabStrip tbcEMail 
      Height          =   2310
      Left            =   90
      TabIndex        =   16
      Top             =   60
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   4075
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Mass E-Mails"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Individual E-Mails"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin V81Affiliate.CSI_Calendar cbcFollowUp 
      Height          =   255
      Left            =   1470
      TabIndex        =   19
      Top             =   2670
      Width           =   1710
      _ExtentX        =   2699
      _ExtentY        =   450
      Text            =   "8/26/2020"
      BorderStyle     =   1
      CSI_ShowDropDownOnFocus=   -1  'True
      CSI_InputBoxBoxAlignment=   0
      CSI_CalBackColor=   16777130
      CSI_CalDateFormat=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_DayNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CSI_MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CSI_CurDayBackColor=   16777215
      CSI_CurDayForeColor=   0
      CSI_ForceMondaySelectionOnly=   0   'False
      CSI_AllowBlankDate=   -1  'True
      CSI_AllowTFN    =   -1  'True
      CSI_DefaultDateType=   1
   End
   Begin VB.Label lbcFrom 
      Caption         =   "From Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   24
      Top             =   3060
      Width           =   1095
   End
   Begin VB.Label lacConcern 
      Caption         =   "Concerning:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3630
      TabIndex        =   22
      Top             =   2670
      Width           =   1365
   End
   Begin VB.Label lacFollowup 
      Caption         =   "Follow-up Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   21
      Top             =   2670
      Width           =   1440
   End
   Begin VB.Image imcSpellCheck 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   825
      Picture         =   "AffContactEMail.frx":0023
      ToolTipText     =   "Check Spelling"
      Top             =   4770
      Width           =   360
   End
   Begin VB.Label lacAttachment 
      Caption         =   "Attachment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   165
      TabIndex        =   10
      Top             =   5295
      Width           =   1185
   End
   Begin VB.Label lacSubject 
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   6
      Top             =   3540
      Width           =   855
   End
   Begin VB.Label lacResults 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Results:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   165
      TabIndex        =   14
      Top             =   6330
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lacMessage 
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
End
Attribute VB_Name = "frmContactEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private imFirstTime As Integer
Private imShttCode As Integer
Private imMktRepUstCode As Integer
Private imServRepUstCode As Integer
Private smUserType As String
Private imInChg As Integer
Private imBSMode As Integer

Dim imCtrlVisible As Integer
Dim lmEnableRow As Long
Dim lmEnableCol As Long
Dim lmTopRow As Long
'ttp5352
Dim smMessagePath As String
'Email sent
Private rst_artt As ADODB.Recordset
Private rst_Ust As ADODB.Recordset
Private rst_tnt As ADODB.Recordset
Private rst_cef As ADODB.Recordset
Private rst_dnt As ADODB.Recordset
Private rstTitles As ADODB.Recordset
'Dim smUserName As String
Dim bmSendToUser As Boolean
Dim imTabIndex As Integer
'Personnel Contact Grid- grdContact
Const PTOINDEX = 0
Const PCCINDEX = 1
Const PBCINDEX = 2
Const PCNAMEINDEX = 3
Const PCTITLEINDEX = 4
Const PCEMAILINDEX = 5
Const PCARTTCODEINDEX = 6
Const PUSTCODEINDEX = 7

Const TABMASS = 1
Const TABINDIVIDUAL = 2

Const SCALLLETTERINDEX = 1     'If changed, change frmStation
Const SSHTTCODEINDEX = 19   'Value must match how it is defined in frmStationSearch.  Constant defined in frmStation and here
Private Const LOGFILE As String = ""
Private Const FORMNAME As String = "Contact E-Mail"


Private Sub cbcConcern_Change()
    Dim iLoop As Integer
    Dim sName As String
    Dim iLen As Integer
    Dim iSel As Integer
    Dim lRow As Long
    Dim iZone As Integer
    
    If imInChg Then
        Exit Sub
    End If
    imInChg = True
    Screen.MousePointer = vbHourglass
    sName = LTrim$(cbcConcern.Text)
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    lRow = SendMessageByString(cbcConcern.hwnd, CB_FINDSTRING, -1, sName)
    If lRow >= 0 Then
        cbcConcern.ListIndex = lRow
        cbcConcern.SelStart = iLen
        cbcConcern.SelLength = Len(cbcConcern.Text)
    End If
    Screen.MousePointer = vbDefault
    imInChg = False
End Sub

Private Sub cbcConcern_Click()
    cbcConcern_Change
End Sub

Private Sub cbcConcern_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
End Sub

Private Sub cbcConcern_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If cbcConcern.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
End Sub

Private Sub cbcFollowUp_GotFocus()
    cbcFollowUp.ZOrder
End Sub

Private Sub ckcEMail_Click()
    Dim slStr As String
    Dim llRow As Long
    If ckcEMail.Value = vbChecked Then
        If lmEnableCol = PTOINDEX Then
            grdContact.Col = lmEnableCol
            grdContact.CellFontName = "Monotype Sorts"
            grdContact.TextMatrix(lmEnableRow, lmEnableCol) = "4"
            grdContact.TextMatrix(lmEnableRow, PCCINDEX) = ""
            grdContact.TextMatrix(lmEnableRow, PBCINDEX) = ""
            slStr = grdContact.TextMatrix(grdContact.Row, PUSTCODEINDEX)
            'Only allow one user to be defined with the To
            If slStr <> "" Then
                If Val(slStr) > 0 Then
                    For llRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
                        If lmEnableRow <> llRow Then
                            If grdContact.TextMatrix(llRow, PCNAMEINDEX) <> "" Then
                                If grdContact.TextMatrix(llRow, PTOINDEX) = "4" Then
                                    slStr = grdContact.TextMatrix(llRow, PUSTCODEINDEX)
                                    If (slStr <> "") Then
                                        If Val(slStr) > 0 Then
                                            grdContact.TextMatrix(llRow, PTOINDEX) = ""
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next llRow
                End If
            End If
        ElseIf lmEnableCol = PCCINDEX Then
            grdContact.Col = lmEnableCol
            grdContact.CellFontName = "Monotype Sorts"
            grdContact.TextMatrix(lmEnableRow, lmEnableCol) = "4"
            grdContact.TextMatrix(lmEnableRow, PTOINDEX) = ""
            grdContact.TextMatrix(lmEnableRow, PBCINDEX) = ""
        ElseIf lmEnableCol = PBCINDEX Then
            grdContact.Col = lmEnableCol
            grdContact.CellFontName = "Monotype Sorts"
            grdContact.TextMatrix(lmEnableRow, lmEnableCol) = "4"
            grdContact.TextMatrix(lmEnableRow, PTOINDEX) = ""
            grdContact.TextMatrix(lmEnableRow, PCCINDEX) = ""
        End If
    Else
        If lmEnableCol = PTOINDEX Then
            grdContact.TextMatrix(lmEnableRow, lmEnableCol) = ""
        ElseIf lmEnableCol = PCCINDEX Then
            grdContact.TextMatrix(lmEnableRow, lmEnableCol) = ""
        ElseIf lmEnableCol = PBCINDEX Then
            grdContact.TextMatrix(lmEnableRow, lmEnableCol) = ""
        End If
    End If
End Sub

Private Sub cmcBrowse_Click()
    Dim ilLoop As Integer
    cdcBrowse.Filter = "All Files|*.*"    'Setup the CommonDialog
    cdcBrowse.ShowOpen 'Show the Open Dialog
    If cdcBrowse.fileName <> "" Then
        If LenB(Trim(edcAttachment)) > 0 Then
            edcAttachment.Text = edcAttachment.Text & ";" & cdcBrowse.fileName
        Else
            edcAttachment.Text = cdcBrowse.fileName
        End If
    End If
    gChDrDir
End Sub

Private Sub edcAttachment_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub edcMessage_Change()
    mSetCommands
End Sub

Private Sub edcSubject_Change()
    mSetCommands
End Sub

Private Sub edcSubject_GotFocus()
    gCtrlGotFocus ActiveControl
End Sub

Private Sub Form_Activate()
    If imFirstTime Then
        mMousePointer vbHourglass
        tmcStart.Enabled = True
    End If
End Sub

Private Sub Form_Initialize()
    Me.Width = Screen.Width / 1.5
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gSetFonts frmContactEMail
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    
    On Error Resume Next
    ilRet = mCloseCEFFile()
    rst_Ust.Close
    rst_artt.Close
    rst_tnt.Close
    rst_dnt.Close
    
    Set frmContactEMail = Nothing
    
End Sub

Private Sub cmcExit_Click()
    Unload frmContactEMail
End Sub
Private Function mAddAddressTestUser(llRow As Long, slList As String) As String
    Dim slNewList As String
    Dim slAddress As String
    
    slAddress = grdContact.TextMatrix(llRow, PCEMAILINDEX)
    If LenB(slList) = 0 Then
        slNewList = slAddress
    Else
        slNewList = slList & "," & slAddress
    End If
    mSetUserFlag llRow
    mAddAddressTestUser = slNewList

End Function
Private Function mIsUser(llRow As Long) As Boolean
    If grdContact.TextMatrix(llRow, PCNAMEINDEX) <> "" Then
        If grdContact.TextMatrix(llRow, PUSTCODEINDEX) > 0 Then
            mIsUser = True
        End If
    End If
End Function
Private Sub mSetUserFlag(llRow As Long)
'Send to user false?  then find if current selected is a user.
    If Not bmSendToUser Then
        bmSendToUser = mIsUser(llRow)
    End If
End Sub
Private Function mPrepRecordset() As ADODB.Recordset
    Dim myRs As ADODB.Recordset
    
    Set myRs = New ADODB.Recordset
        With myRs.Fields
            .Append "Station", adChar, 10
            .Append "LastName", adChar, 40
            .Append "FirstName", adChar, 40
            '7917 from 40 to 70
            .Append "Address", adChar, 70
        End With
    myRs.Open
    Set mPrepRecordset = myRs
End Function
Private Function mGetStation(ilStation As Integer) As String
    Dim myRst As ADODB.Recordset
    Dim slSql As String
    
    slSql = "SELECT  shttCallLetters from Shtt where shttcode = " & ilStation
    On Error GoTo errbox:
    Set myRst = gSQLSelectCall(slSql)
    If Not myRst.EOF Then
        mGetStation = Trim$(myRst.Fields("shttCallLetters").Value)
    End If
    myRst.Close
    Set myRst = Nothing
    Exit Function
errbox:
    gHandleError LOGFILE, FORMNAME & "-mGetStation"
End Function
Private Sub cmcSend_Click()

'    Dim tlEmailInfo As EmailInformation 'Dan M. to send to global procedure
    Dim llRow As Long
'    Dim slToAddresses As String
 '   Dim slCCAddresses As String
    Dim slBCCAddresses As String
    Dim blRet As Boolean
    Dim blCommentAdded As Boolean
    Dim slStr As String
    Dim ilShtt As Integer
    Dim llPos As Long
'    Dim slPrevEMail As String
    Dim ilShttIndex As Integer
    Dim blFound As Boolean
    Dim blFoundOneBCC As Boolean
    'ttp 5352
    Dim slFailedValidation As String
    Dim c As Integer
    Dim slFailures() As String
    Dim ilUpperBound As Integer
    Dim myFileSys As FileSystemObject
    Dim myStream As TextStream
    Dim myFile As file
    Const MAXVALUE As Integer = 32765
    Const INVALIDFILE As String = "EmailFormatImproper.txt"
    'ttp 5351
    Dim myAddresses As ADODB.Recordset
    Dim slStation As String
    Dim slTitle As String

    ilUpperBound = 0
    ReDim slFailures(0) As String
    lbcResults.Clear
    lbcResults.ForeColor = vbBlack
    Screen.MousePointer = vbHourglass
    bmSendToUser = False
    'dan m 10/14/11 send button now enabled, but email not sent if subject/message not defined. to be more consistent with rest of email in system.
    If Trim$(edcSubject.Text) = "" Then
        Screen.MousePointer = vbDefault
        MsgBox "Subject must be defined! ", vbInformation
        Exit Sub
    End If
    If Trim$(edcMessage.Text) = "" Then
        Screen.MousePointer = vbDefault
        MsgBox "Message must be defined! ", vbInformation
        Exit Sub
    End If
On Error GoTo errbox:
    Set ogEmailer = New CEmail
On Error GoTo ErrHand:
    If Len(ogEmailer.ErrorMessage) > 0 Then
        Screen.MousePointer = vbDefault
        gMsgBox "There is a problem with emailing: " & ogEmailer.ErrorMessage, vbInformation, "Email Error"
        Exit Sub
    End If
    blCommentAdded = False
'    slToAddresses = ""
'    slCCAddresses = ""
    slBCCAddresses = ""
 '   slPrevEMail = ""
    Set myAddresses = mPrepRecordset()
    If imTabIndex = TABINDIVIDUAL Then
        For llRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
            If grdContact.TextMatrix(llRow, PCNAMEINDEX) <> "" Then
                If grdContact.TextMatrix(llRow, PTOINDEX) = "4" Then
                    'slToAddresses = mAddAddressTestUser(llRow, slToAddresses)
                    If ogEmailer.TestAddress(grdContact.TextMatrix(llRow, PCEMAILINDEX)) Then
                        ogEmailer.AddTOAddress grdContact.TextMatrix(llRow, PCEMAILINDEX), grdContact.TextMatrix(llRow, PCNAMEINDEX)
                        mSetUserFlag llRow
                        grdContact.Col = PCNAMEINDEX
                        slStr = grdContact.TextMatrix(llRow, PUSTCODEINDEX)
                        If (slStr <> "") And (Not blCommentAdded) Then
                            If Val(slStr) > 0 Then
                                'dan M 10/17/11 ttp 4901, use imShttCode instead of igContactEmailshttcode
                                'blRet = mAddComment(igContactEmailShttCode, Val(slStr))
                                blRet = mAddComment(imShttCode, Val(slStr))
                                blCommentAdded = True
                            End If
                        End If
                    Else
                        'ttp 5352
                        slFailedValidation = mFailedValidation(imShttCode, grdContact.TextMatrix(llRow, PCNAMEINDEX)) & ":" & grdContact.TextMatrix(llRow, PCEMAILINDEX)
                        If ilUpperBound < MAXVALUE Then
                            ilUpperBound = ilUpperBound + 1
                            ReDim Preserve slFailures(ilUpperBound)
                            slFailures(ilUpperBound - 1) = slFailedValidation
                        Else
                            'too many failures, don't write any more
                        End If
                    End If 'valid email?
                ElseIf grdContact.TextMatrix(llRow, PCCINDEX) = "4" Then
                    'slCCAddresses = mAddAddressTestUser(llRow, slCCAddresses)
                    If ogEmailer.TestAddress(grdContact.TextMatrix(llRow, PCEMAILINDEX)) Then
                        ogEmailer.AddCCAddress grdContact.TextMatrix(llRow, PCEMAILINDEX), grdContact.TextMatrix(llRow, PCNAMEINDEX)
                        mSetUserFlag llRow
                    Else
                        'ttp 5352
                        slFailedValidation = mFailedValidation(imShttCode, grdContact.TextMatrix(llRow, PCNAMEINDEX)) & ":" & grdContact.TextMatrix(llRow, PCEMAILINDEX)
                        If ilUpperBound < MAXVALUE Then
                            ilUpperBound = ilUpperBound + 1
                            ReDim Preserve slFailures(ilUpperBound)
                            slFailures(ilUpperBound - 1) = slFailedValidation
                        Else
                            'too many failures, don't write any more
                        End If
                    End If
                ElseIf grdContact.TextMatrix(llRow, PBCINDEX) = "4" Then
                    'slBCCAddresses = mAddAddressTestUser(llRow, slBCCAddresses)
                    If ogEmailer.TestAddress(grdContact.TextMatrix(llRow, PCEMAILINDEX)) Then
                        ogEmailer.AddBCCAddress grdContact.TextMatrix(llRow, PCEMAILINDEX)
                        mSetUserFlag llRow
                    Else
                        'ttp 5352
                        slFailedValidation = mFailedValidation(imShttCode, grdContact.TextMatrix(llRow, PCNAMEINDEX)) & ":" & grdContact.TextMatrix(llRow, PCEMAILINDEX)
                        If ilUpperBound < MAXVALUE Then
                            ilUpperBound = ilUpperBound + 1
                            ReDim Preserve slFailures(ilUpperBound)
                            slFailures(ilUpperBound - 1) = slFailedValidation
                        Else
                            'too many failures, don't write any more
                        End If
                    End If
                End If ' TO, CC OR BCC?
            End If 'name exists
        Next llRow
'        If LenB(slToAddresses) = 0 Then
'            MsgBox "A 'to address' must be included. ", vbInformation
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
    Else
        For ilShtt = 0 To UBound(igCommentShttCode) - 1 Step 1
            blCommentAdded = False
            SQLQuery = "SELECT * FROM artt"
            SQLQuery = SQLQuery + " WHERE ("
            SQLQuery = SQLQuery & " arttType = 'P'"
            SQLQuery = SQLQuery & " AND arttShttCode = " & igCommentShttCode(ilShtt) & ")"
            SQLQuery = SQLQuery & " ORDER BY arttFirstName, arttLastName"
            Set rst_artt = gSQLSelectCall(SQLQuery)
            Do While Not rst_artt.EOF
'               Dan M 10/14/11 first 3 options don't have to have a title
               ' If (rst_artt!arttTntCode > 0) And (Trim$(rst_artt!arttEmail) <> "") Then
                If Trim$(rst_artt!arttEmail) <> "" Then
                    'ttp 5352
                    If ogEmailer.TestAddress(Trim$(rst_artt!arttEmail)) Then
                        'dan M 10/17/11 blCommentAdded not needed here for mass emails
                            'blCommentAdded = False
                        For llRow = 0 To lbcTitle.ListCount - 1 Step 1
                            If lbcTitle.Selected(llRow) Then
                                blFound = False
                                ' a user defined label? make sure title in table matches what user selected
                                If lbcTitle.ItemData(llRow) > 0 Then
                                    If rst_artt!arttTntCode = lbcTitle.ItemData(llRow) Then
                                        blFound = True
                                    End If
                                ' one of 3 preset labels. don't worry about title in table, use fields below instead
                                Else
                                    If ilShttIndex <> -1 Then
                                        If lbcTitle.ItemData(llRow) = -1 Then   'Affiliate Label
                                            If rst_artt!arttAffContact = "1" Then
                                                blFound = True
                                            End If
                                        ElseIf lbcTitle.ItemData(llRow) = -2 Then   'ISCI Export
                                            If rst_artt!arttISCI2Contact = "1" Then
                                                blFound = True
                                            End If
                                        ElseIf lbcTitle.ItemData(llRow) = -3 Then   'E-Mail
                                            If rst_artt!arttWebEMail = "Y" Then
                                                blFound = True
                                            End If
                                        End If
                                    End If
                                End If
                                If blFound And (Trim$(rst_artt!arttEmail) <> "") Then
                                    blFoundOneBCC = True
                                    If Not blCommentAdded Then
                                        blRet = mAddComment(igCommentShttCode(ilShtt), 0)
                                        blCommentAdded = True
                                    End If
                                    slStr = Trim$(rst_artt!arttEmail)
                                    llPos = InStr(1, UCase(slBCCAddresses), "," & UCase(slStr), vbBinaryCompare)
                                    If llPos <= 0 Then
                                        slBCCAddresses = slBCCAddresses & "," & slStr
                                        ogEmailer.AddBCCAddress slStr
                                    End If
                                    '7575 moved
                                    'ttp 5351
                                    'slTitle = lbcTitle.List(llRow)
                                    slStation = mGetStation(rst_artt("arttshttcode").Value)
                                    myAddresses.AddNew Array("Station", "FirstName", "LastName", "Address"), Array(slStation, Trim$(rst_artt.Fields("arttFirstName").Value), Trim$(rst_artt.Fields("arttLastName").Value), slStr)
                                    'dan m 10/14/11 changed to above. testing if email already being used.
    '                                If LenB(slBCCAddresses) = 0 Then
    '                                    slBCCAddresses = slStr
    '                                Else
    '                                    If StrComp(UCase(slPrevEMail), UCase(slStr), vbBinaryCompare) <> 0 Then
    '                                        llPos = InStr(1, UCase(slBCCAddresses), UCase(slStr) & ",", vbBinaryCompare)
    '                                        If llPos <= 0 Then
    '                                            slBCCAddresses = slBCCAddresses & "," & slStr
    '                                        End If
    '                                    Else
    '                                        slBCCAddresses = slBCCAddresses & "," & slStr
    '                                    End If
    '                                End If
     '                               slPrevEMail = slStr
                                End If
                            End If
                        Next llRow
                    Else
                        'ttp 5352
                        slFailedValidation = mFailedValidation(rst_artt!arttShttCode, Trim$(rst_artt!arttFirstName) & " " & Trim$(rst_artt!arttLastName)) & ":" & Trim$(rst_artt!arttEmail)
                        If ilUpperBound < MAXVALUE Then
                            ilUpperBound = ilUpperBound + 1
                            ReDim Preserve slFailures(ilUpperBound)
                            slFailures(ilUpperBound - 1) = slFailedValidation
                        Else
                            'too many failures, don't write any mroe
                        End If
                    End If 'valid email?
                End If  'email blank?
                rst_artt.MoveNext
            Loop
        Next ilShtt
    End If
    If (Not blCommentAdded) And (imTabIndex = TABINDIVIDUAL) Then
        'dan M 10/17/11 ttp 4901, use imShttCode instead of igContactEmailshttcode
       ' blRet = mAddComment(igContactEmailShttCode, 0)
        blRet = mAddComment(imShttCode, 0)
        blCommentAdded = True
    End If
'    With tlEmailInfo
'        .sSubject = edcSubject.Text
'        .sMessage = edcMessage.Text
'        .sAttachment = edcAttachment.Text
'        .sFromAddress = sgEMail
'        .sFromName = smUserName
'        .sToMultiple = slToAddresses
'        .sBCCMulitple = slBCCAddresses
'        .sCCMultiple = slCCAddresses
'    End With
    'Dan M sending only bccs. need to add 'to' address
    If blFoundOneBCC Then
        'ttp 5352
        If ogEmailer.TestAddress(sgEMail) Then
            'Dan 4/01/13
            'ogEmailer.AddTOAddress sgEMail, smUserName
        If Len(Trim$(edcFromName.Text)) > 0 Then
            ogEmailer.AddTOAddress sgEMail, Trim$(edcFromName.Text)
        Else
            ogEmailer.AddTOAddress sgEMail
        End If
        Else
            mShowResults True
            lbcResults.AddItem "Cannot send: your email address is not valid."
            lbcResults.ForeColor = vbRed
            Screen.MousePointer = vbDefault
            Set ogEmailer = Nothing
            Exit Sub
        End If
    End If
    With ogEmailer
        .Subject = edcSubject.Text
        .Message = edcMessage.Text
        .Attachment = edcAttachment.Text
        .FromAddress = sgEMail
        '6050 instead of smUserName
        If Len(Trim$(edcFromName.Text)) > 0 Then
            .FromName = Trim$(edcFromName.Text)
        End If
'        .ToMultiple = slToAddresses
'        .BCCMultiple = slBCCAddresses
'        .CCMultiple = slCCAddresses
    End With
    mShowResults True
    If Len(sgSpecialPassword) = 4 Then
        MsgBox "Guide, testing email. No email sent"
    Else
        ogEmailer.Send lbcResults
    End If
    'ttp 5352 one email? don't make text file.  Otherwise, append list of stations and usernames/addresses.  If many emails, don't write them all out.
    If Len(slFailedValidation) > 0 Then
        Set myFileSys = New FileSystemObject
    On Error GoTo ERRFILE
        If ilUpperBound = 1 Then
              gAddMsgToListBox Me, lbcResults.Width, " An email could not be sent due to invalid email addresses: " & slFailedValidation, lbcResults
        Else
            If myFileSys.FILEEXISTS(smMessagePath & INVALIDFILE) Then
            'older than a week? delete file
                Set myFile = myFileSys.GetFile(smMessagePath & INVALIDFILE)
                If DateDiff("d", myFile.DateLastModified, Date) > 6 Then
                    myFile.Delete
                End If
                Set myFile = Nothing
            End If
            Set myStream = myFileSys.OpenTextFile(smMessagePath & INVALIDFILE, ForAppending, True)
            myStream.WriteLine (vbCrLf & "User: " & sgUserName & " List of improperly formatted email addresses during mass emailing on " & gNow())
            If ilUpperBound > MAXVALUE - 3 Then
                'Dan 8/26/20 wrong address
               ' gAddMsgToListBox Me, lbcResults.Width, "more than " & ilUpperBound - 1 & " emails could not be sent due to improperly formatted email addresses.   For the list of improper addresses, see " & sgExportDirectory & INVALIDFILE, lbcResults
                gAddMsgToListBox Me, lbcResults.Width, "more than " & ilUpperBound - 1 & " emails could not be sent due to improperly formatted email addresses.   For the partial list of improper addresses, see " & smMessagePath & INVALIDFILE, lbcResults
            Else
                gAddMsgToListBox Me, lbcResults.Width, ilUpperBound & " emails could not be sent due to improperly formatted email addresses.   For the list of improper addresses, see " & smMessagePath & INVALIDFILE, lbcResults
            End If
            myStream.WriteBlankLines (1)
            For c = 0 To UBound(slFailures) - 1
                myStream.WriteLine (slFailures(c))
            Next c
            myStream.Close
        End If
    On Error GoTo 0
        lbcResults.ForeColor = vbRed
    End If
    '7575
    For llRow = 0 To lbcTitle.ListCount - 1 Step 1
        If lbcTitle.Selected(llRow) Then
            slTitle = slTitle & lbcTitle.List(llRow) & ","
        End If
    Next llRow
    slTitle = mLoseLastLetterIfComma(slTitle)
    'ttp 5351
    If myAddresses.RecordCount > 0 Then
        mLogEmail myAddresses, slTitle
    End If
    cmcExit.Caption = "Done"
    Erase slFailures
    Set myStream = Nothing
    Set myFileSys = Nothing
    Set ogEmailer = Nothing
    If Not myAddresses Is Nothing Then
        If (myAddresses.State And adStateOpen) <> 0 Then
            myAddresses.Close
        End If
        Set myAddresses = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ERRFILE:
    Screen.MousePointer = vbDefault
    lbcResults.AddItem ("Error creating " & INVALIDFILE)
    Exit Sub
ErrHand:
    'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError LOGFILE, FORMNAME & "-cmcSend"
    grdContact.Redraw = True
    Exit Sub
errbox:
    gMsgBox "Could not send email:  Couldn't find csiNetUtiliites", vbInformation, "missing dll"
    Screen.MousePointer = vbDefault
End Sub
Function mLogEmail(myRs As ADODB.Recordset, slTitle As String) As Boolean
    Dim myFileSys As FileSystemObject
    Dim myStream As TextStream
    Dim slGoodAddresses As String
    Dim myFile As file
    Dim myCreateDate As Date
    Dim llDays As Long
    Dim blStreamOpen As Boolean
    Dim blWriteFirstLine As Boolean
    '8346
    Dim c As Integer
    Const MYTAB As String = "   "
    Const LOGFILE As String = "EmailWeeklyLog.txt"
    Const OLDLOGFILE As String = "EmailPreviousWeekly_Log.txt"
    
    With myRs
        .Sort = "station,LastName"
        .MoveFirst
        Do While Not .EOF
            slGoodAddresses = slGoodAddresses & "," & Trim$(.Fields("Station").Value) & "(" & Trim(.Fields("Address").Value) & ")"
            myRs.MoveNext
        Loop
    End With
On Error GoTo ErrHandler:
    Set myFileSys = New FileSystemObject
    'exists? then if created over a week ago, save as 'oldLogFile'.
    If myFileSys.FILEEXISTS(smMessagePath & LOGFILE) Then
        Set myFile = myFileSys.GetFile(smMessagePath & LOGFILE)
        '8348
        myCreateDate = myFile.DateLastModified
        '7575  this wasn't here before.  Set the date back to monday before finding difference
        myCreateDate = DateAdd("d", -Weekday(myCreateDate, vbMonday) + 1, myCreateDate)
        llDays = DateDiff("d", myCreateDate, Date, vbMonday)
        If llDays > 6 Then
            If myFileSys.FILEEXISTS(smMessagePath & OLDLOGFILE) Then
                myFileSys.DeleteFile (smMessagePath & OLDLOGFILE)
            End If
            myFile.Move smMessagePath & OLDLOGFILE
           'Dan M 6/5/15 wrong
            'myFile.Delete
            blWriteFirstLine = True
        End If
    Else
        blWriteFirstLine = True
    End If
    'append or create
    Set myStream = myFileSys.OpenTextFile(smMessagePath & LOGFILE, ForAppending, True)
    blStreamOpen = True
    With myStream
        If blWriteFirstLine Then
            .WriteLine "Mass Emails for week of " & Format(mMondayOfWeek(), "mm/dd/yyyy") & ":"
        End If
        .WriteLine MYTAB & "Email- to: '" & slTitle & "' on " & Date
        .WriteLine MYTAB & MYTAB & "subject-" & ogEmailer.Subject
        .WriteLine MYTAB & MYTAB & "message-" & ogEmailer.Message
        .WriteLine MYTAB & MYTAB & "Sent to: " & Mid(slGoodAddresses, 2)
        '8346
        If lbcResults.ListCount > 0 Then
            .WriteLine MYTAB & MYTAB & "Result:"
            For c = 0 To lbcResults.ListCount - 1
                .WriteLine MYTAB & MYTAB & MYTAB & lbcResults.List(c)
            Next c
        Else
            .WriteLine MYTAB & MYTAB & "Result: Not sent!"
        End If
    End With
Cleanup:
    Set myFileSys = Nothing
    Set myFile = Nothing
    If blStreamOpen Then
        myStream.Close
    End If
    Set myStream = Nothing
    Exit Function
ErrHandler:
    gHandleError LOGFILE, FORMNAME & "-mLogEmail"
    'gLogMsg "Problem with " & FORMNAME & " -mLogEmail: " & Err.Description, "AffErrorLog.txt", False
    mLogEmail = False
    GoTo Cleanup
End Function
Public Function mMondayOfWeek(Optional dlDate As Date) As Date

    If dlDate = #12:00:00 AM# Then
        dlDate = Date
    End If
   Select Case Weekday(dlDate)
      Case 1 ' Sunday
         mMondayOfWeek = dlDate - 6
      Case 2 ' Monday
         mMondayOfWeek = dlDate
      Case 3 To 7 'Tuesday-Saturday
         mMondayOfWeek = dlDate - Weekday(dlDate) + 2
    End Select

End Function
Function mFailedValidation(ilStationCode As Integer, slFullName As String) As String
    Dim slStation As String
    
    mFailedValidation = slFullName
    slStation = mGetStation(ilStationCode)
    If Len(slStation) > 0 Then
        mFailedValidation = "station " & slStation & " " & slFullName
    End If

'    Dim myRst As ADODB.Recordset
'    Dim slSql As String
'
'    mFailedValidation = slFullName
'    slSql = "SELECT  shttCallLetters from Shtt where shttcode = " & ilStationCode
'    On Error GoTo errbox:
'    Set myRst = gSQLSelectCall(slSql)
'    If Not myRst.EOF Then
'        mFailedValidation = "station " & Trim$(myRst!shttCallLetters) & " " & slFullName
'    End If
'    Set rst = Nothing
'    Exit Function
'errbox:
'    gHandleError LOGFILE, FORMNAME & "-mFailedValidation"
End Function
Sub mShowResults(flag As Boolean)
    lbcResults.Visible = flag
    lacResults.Visible = flag
End Sub

Private Sub Form_Load()
    gCenterForm frmContactEMail
    mInit
End Sub


Private Sub mInit()
   
    Dim ilRet As Integer
    '6050 moved slUserName here
    Dim slUserName As String
    
    imFirstTime = True
    lmEnableRow = -1
    pbcSTab.Move -100, -100
    pbcTab.Move -100, -100
    If Trim$(sgReportName) <> "" Then
        slUserName = sgReportName
    Else
        slUserName = sgUserName
    End If
    '6050 slUserName
    edcFromName.Text = slUserName
    mPopTitles
    If igContactEmailShttCode > 0 Then
        'Sendkeys "%I", False
        tbcEMail.Tabs(TABINDIVIDUAL).Selected = True
    Else
        'SendKeys "%M", False
        tbcEMail_Click
    End If
    imInChg = False
    imBSMode = False
    mPopVehicle
    mGetUserInfo
    ilRet = mOpenCEFFile()
    'dan M 4/11/12  give 'csi' a fake email...for testing
    If Len(sgSpecialPassword) = 4 Then
        sgEMail = "testOnly@counterpoint.net"
    End If
    'dan M 10/14/11  no email? let user know won't be able to send email
    If (sgEMail = "") Then
        MsgBox "This user must define an email address before being able to send email.", vbOKOnly, "No 'from' email address"
    End If
    'dan M 10/17/11 don't allow mass emails if there are no stations in the array
    If UBound(igCommentShttCode) = 0 Then
        tbcEMail.Tabs(TABINDIVIDUAL).Selected = True
        tbcEMail.Enabled = False
        lbcTitle.Visible = False
        pbcTitle.Visible = False
    Else
        tbcEMail.Tabs(TABMASS).Caption = "&Mass E-Mails (Max # of Stations: " & UBound(igCommentShttCode) & ")"

    End If
    smMessagePath = sgDBPath & "Messages\"
  '  mSetCommands
End Sub

Private Sub mSetCommands()
    Dim llRow As Long
    
    cmcSend.Enabled = False
    'Dan M while 'csi' can get to this page, he cannot send email, as his sgEmail is blank.
'    If (sgEMail <> "") Then
'        For llRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
'            If grdContact.TextMatrix(llRow, PCNAMEINDEX) <> "" Then
'                If grdContact.TextMatrix(llRow, PTOINDEX) = "4" Then
'                    cmcSend.Enabled = True
'                    Exit Sub
'                End If
'            End If
'        Next llRow
'    End If
    'Dan M 10/14/11 to be consistent with rest of email throughout system, enable button but don't send
'    If Trim$(edcSubject.Text) = "" Then
'        Exit Sub
'    End If
'    If Trim$(edcMessage.Text) = "" Then
'        Exit Sub
'    End If
    If (sgEMail = "") Then
        Exit Sub
    End If
    If imTabIndex = TABMASS Then
        For llRow = 0 To lbcTitle.ListCount - 1 Step 1
            If lbcTitle.Selected(llRow) Then
                cmcSend.Enabled = True
                Exit Sub
            End If
        Next llRow
    ElseIf imTabIndex = TABINDIVIDUAL Then
        For llRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
            If grdContact.TextMatrix(llRow, PCEMAILINDEX) <> "" Then
                cmcSend.Enabled = True
                Exit Sub
            End If
        Next llRow
    End If

End Sub
Private Sub mClearEmailFields()
    edcMessage.Text = ""
    edcSubject.Text = ""
    edcAttachment.Text = ""
    cbcFollowUp.Text = ""
    cbcConcern.Text = ""
End Sub

Private Sub mSetGridColumns()
    Dim ilCol As Integer
    
        
    
    grdContact.Width = tbcEMail.Width - lbcStations.Left - lbcStations.Width - 240
    grdContact.Left = lbcStations.Left + lbcStations.Width + 120
    'grdContact.Height = edcSubject.Top - (frcBy.Top + frcBy.Height) - 2 * (cbcFollowUp.Top - edcSubject.Top - edcSubject.Height)
    lbcTitle.Height = tbcEMail.Height - 600
    lbcStations.Height = lbcTitle.Height
    grdContact.Height = lbcTitle.Height + 90 'edcSubject.Top - frcBy.Top - 2 * frcBy.Height
    grdContact.Top = lbcTitle.Top
    gGrid_IntegralHeight grdContact
    grdContact.Height = grdContact.Height + 30
    gGrid_FillWithRows grdContact
    'grdContact.Move lacSubject.Left, frcBy.Height + (edcSubject.Top - (frcBy.Top + frcBy.Height) - grdContact.Height) \ 2
    'grdContact.Move lacSubject.Left, frcBy.Height + (edcSubject.Top - grdContact.Height) \ 2
    grdContact.ColWidth(PCARTTCODEINDEX) = 0
    grdContact.ColWidth(PUSTCODEINDEX) = 0
    grdContact.ColWidth(PTOINDEX) = grdContact.Width * 0.06
    grdContact.ColWidth(PCCINDEX) = grdContact.Width * 0.06
    grdContact.ColWidth(PBCINDEX) = grdContact.Width * 0.06
    grdContact.ColWidth(PCNAMEINDEX) = grdContact.Width * 0.2
    grdContact.ColWidth(PCTITLEINDEX) = grdContact.Width * 0.2
           
    grdContact.ColWidth(PCEMAILINDEX) = grdContact.Width - GRIDSCROLLWIDTH - 15
    For ilCol = 0 To PCEMAILINDEX Step 1
        If ilCol <> PCEMAILINDEX Then
            grdContact.ColWidth(PCEMAILINDEX) = grdContact.ColWidth(PCEMAILINDEX) - grdContact.ColWidth(ilCol)
        End If
    Next ilCol
    'Align columns to left
    gGrid_AlignAllColsLeft grdContact
    
End Sub
Private Sub mSetGridTitles()
    grdContact.TextMatrix(0, PTOINDEX) = "To:"
    grdContact.TextMatrix(0, PCCINDEX) = "CC:"
    grdContact.TextMatrix(0, PBCINDEX) = "BC:"
    grdContact.TextMatrix(0, PCNAMEINDEX) = "Name"
    grdContact.TextMatrix(0, PCTITLEINDEX) = "Title"
    grdContact.TextMatrix(0, PCEMAILINDEX) = "E-Mail Address"
End Sub

Private Sub grdContact_EnterCell()
    mSetShow
End Sub

Private Sub grdContact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'grdContact.ToolTipText = ""
    If (grdContact.MouseRow >= grdContact.FixedRows) And (grdContact.TextMatrix(grdContact.MouseRow, grdContact.MouseCol)) <> "" Then
        grdContact.ToolTipText = grdContact.TextMatrix(grdContact.MouseRow, grdContact.MouseCol)
    Else
        grdContact.ToolTipText = ""
    End If
End Sub

Private Sub grdContact_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilRow As Integer
    Dim ilCol As Integer

    'Determine row and col mouse up onto
    On Error GoTo grdContactErr
    ilCol = grdContact.MouseCol
    ilRow = grdContact.MouseRow
    If ilCol < grdContact.FixedCols Then
        grdContact.Redraw = True
        Exit Sub
    End If
    If ilRow < grdContact.FixedRows Then
        grdContact.Redraw = True
        Exit Sub
    End If
    lmTopRow = grdContact.TopRow
    DoEvents
    If grdContact.TextMatrix(ilRow, PCNAMEINDEX) = "" Then
        grdContact.Redraw = True
        cmcExit.SetFocus
        Exit Sub
    End If
    grdContact.Col = ilCol
    grdContact.Row = ilRow
    If Not mColOk() Then
        grdContact.Redraw = True
        Exit Sub
    End If
    grdContact.Redraw = True
    mEnableBox
    On Error GoTo 0
    Exit Sub
grdContactErr:
    On Error GoTo 0
    cmcExit.SetFocus
    grdContact.Redraw = False
    grdContact.Redraw = True
    Exit Sub
End Sub

Private Sub grdContact_Scroll()
    mSetShow
End Sub

Private Sub imcSpellCheck_Click()
    'dan m 10/13/11 ttp 4877
    If Len(edcMessage.Text) > 0 Then
        gSpellCheckUsingMSWord edcMessage
    End If
End Sub

Private Sub lbcStations_Click()
    mClearGrid
    If lbcStations.ListIndex >= 0 Then
        imShttCode = lbcStations.ItemData(lbcStations.ListIndex)
        mPopContactGrid
    End If
End Sub

Private Sub lbcTitle_Click()
    mSetCommands
End Sub

Private Sub pbcSTab_GotFocus()
    Dim ilNext As Integer
    Dim ilIndex As Integer

    If GetFocus() <> pbcSTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        Do
            ilNext = False
            Select Case grdContact.Col
                Case PTOINDEX
                    cmcExit.SetFocus
                    Exit Sub
                Case Else
                    grdContact.Col = grdContact.Col - 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
        mSetShow
    Else
        cmcExit.SetFocus
        Exit Sub
    End If
    mEnableBox
End Sub

Private Sub pbcTab_GotFocus()
    Dim ilNext As Integer
    Dim ilIndex As Integer
    Dim llEnableRow As Long
    Dim llEnableCol As Long

    If GetFocus() <> pbcTab.hwnd Then
        Exit Sub
    End If
    If imCtrlVisible Then
        llEnableRow = lmEnableRow
        llEnableCol = lmEnableCol
        mSetShow
        grdContact.Row = llEnableRow
        grdContact.Col = llEnableCol
        Do
            ilNext = False
            Select Case grdContact.Col
                Case PBCINDEX
                    cmcExit.SetFocus
                    Exit Sub
                Case Else
                    grdContact.Col = grdContact.Col + 1
            End Select
            If mColOk() Then
                Exit Do
            Else
                ilNext = True
            End If
        Loop While ilNext
    Else
        cmcExit.SetFocus
        Exit Sub
    End If
    mEnableBox
End Sub


Private Sub pbcTitle_Paint()
    pbcTitle.CurrentX = 0
    pbcTitle.CurrentY = 0
    pbcTitle.Print "Send E-Mail To:"
End Sub

Private Sub tbcEMail_Click()
    grdContact.Visible = False
    lbcStations.Visible = False
    lbcTitle.Visible = False
    pbcTitle.Visible = False
    'dan M 10/13/11
    ckcEMail.Visible = False
    'dan M 10/14/11  clear fields? no.
   ' mClearEmailFields
    Select Case tbcEMail.SelectedItem.Index
        Case TABMASS  'Main
            lbcTitle.Visible = True
            pbcTitle.Visible = True
        Case TABINDIVIDUAL  'Personnel
            grdContact.Visible = True
            lbcStations.Visible = True
    End Select
    'frcTab(imTabIndex - 1).Visible = False
    imTabIndex = tbcEMail.SelectedItem.Index

End Sub

Private Sub tmcStart_Timer()
    tmcStart.Enabled = False
    mMousePointer vbHourglass
    mSetGridColumns
    mSetGridTitles
    'coming into email, set the station to what was selected on previous screen
    imShttCode = igContactEmailShttCode
    'mPopContactGrid
    mPopStations
    mSetCommands
    imFirstTime = False
    mMousePointer vbDefault
End Sub

Private Sub mMousePointer(ilMousepointer As Integer)
    Screen.MousePointer = ilMousepointer
    gSetMousePointer grdContact, grdContact, ilMousepointer
End Sub

Private Function mPopContactGrid() As Integer
    Dim llRow As Long
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilCol As Integer
    Dim ilIncludeUser As Integer
    
    mPopContactGrid = False
    On Error GoTo ErrHand:
    grdContact.Redraw = False
    imMktRepUstCode = 0
    imServRepUstCode = 0
    For llRow = 0 To UBound(tgStationInfo) - 1 Step 1
        If tgStationInfo(llRow).iCode = imShttCode Then
            imMktRepUstCode = tgStationInfo(llRow).iMktRepUstCode
            imServRepUstCode = tgStationInfo(llRow).iServRepUstCode
            Exit For
        End If
    Next llRow
    llRow = grdContact.FixedRows
    SQLQuery = "SELECT * FROM artt"
    SQLQuery = SQLQuery + " WHERE ("
    SQLQuery = SQLQuery & " arttShttCode = " & imShttCode & ")"
    SQLQuery = SQLQuery & " ORDER BY arttFirstName, arttLastName"
    Set rst_artt = gSQLSelectCall(SQLQuery)
    Do While Not rst_artt.EOF
        If llRow >= grdContact.Rows Then
            grdContact.AddItem ""
        End If
        If Trim$(rst_artt!arttEmail) <> "" Then
            grdContact.Row = llRow
            For ilCol = PCNAMEINDEX To PCEMAILINDEX Step 1
                grdContact.Col = ilCol
                grdContact.CellBackColor = LIGHTYELLOW
            Next ilCol
            grdContact.TextMatrix(llRow, PCNAMEINDEX) = Trim$(rst_artt!arttFirstName) & " " & Trim$(rst_artt!arttLastName)
            grdContact.TextMatrix(llRow, PCTITLEINDEX) = gGetTitleByTntCode(rst_artt!arttTntCode)
            grdContact.TextMatrix(llRow, PCEMAILINDEX) = Trim$(rst_artt!arttEmail)
            grdContact.TextMatrix(llRow, PCARTTCODEINDEX) = rst_artt!arttCode
            grdContact.TextMatrix(llRow, PUSTCODEINDEX) = "0"
            
            llRow = llRow + 1
        End If
        rst_artt.MoveNext
    Loop
    '5729
    SQLQuery = "SELECT ustname, ustReportName, ustEMailCefCode, ustDntCode, ustCode FROM Ust where ustState = 0 Order By ustReportName, ustname"
    'SQLQuery = "SELECT ustname, ustReportName, ustEMailCefCode, ustDntCode, ustCode FROM Ust Order By ustReportName, ustname"
    Set rst_Ust = gSQLSelectCall(SQLQuery)
    Do While Not rst_Ust.EOF
        If rst_Ust!ustEmailcefcode > 0 Then
            ilRet = mGetCefComment(rst_Ust!ustEmailcefcode, slStr)
            If slStr <> "" Then
                ilIncludeUser = True
                If rst_Ust!ustDntCode > 0 Then
                    SQLQuery = "SELECT dntName, dntColor, dntType FROM Dnt Where dntCode = " & rst_Ust!ustDntCode
                    Set rst_dnt = gSQLSelectCall(SQLQuery)
                    If Not rst_dnt.EOF Then
                        'If UserType is not defined or O, then show all
                        If smUserType = "M" Then
                            If rst_dnt!dntType <> "S" Then
                                ilIncludeUser = False
                            End If
                        End If
                        If smUserType = "S" Then
                            If rst_dnt!dntType <> "M" Then
                                ilIncludeUser = False
                            End If
                        End If
                    End If
                End If
                If ilIncludeUser Then
                    If llRow >= grdContact.Rows Then
                        grdContact.AddItem ""
                    End If
                    grdContact.Row = llRow
                    For ilCol = PCNAMEINDEX To PCEMAILINDEX Step 1
                        grdContact.Col = ilCol
                        If ilCol = PCNAMEINDEX Then
                            grdContact.CellBackColor = LIGHTGREENCOLOR
                        Else
                            grdContact.CellBackColor = LIGHTYELLOW
                        End If
                    Next ilCol
                    If Trim$(rst_Ust!ustReportName) <> "" Then
                        grdContact.TextMatrix(llRow, PCNAMEINDEX) = Trim$(rst_Ust!ustReportName)
                    Else
                        grdContact.TextMatrix(llRow, PCNAMEINDEX) = Trim$(rst_Ust!ustname)
                    End If
                    grdContact.TextMatrix(llRow, PCTITLEINDEX) = ""
                    If rst_Ust!ustDntCode > 0 Then
                        If Not rst_dnt.EOF Then
                            If (imMktRepUstCode = rst_Ust!ustCode) Or (imServRepUstCode = rst_Ust!ustCode) Then
                                grdContact.Col = PCTITLEINDEX
                                grdContact.CellBackColor = rst_dnt!dntColor
                            End If
                            grdContact.TextMatrix(llRow, PCTITLEINDEX) = Trim$(rst_dnt!dntName)
                        End If
                    End If
                    grdContact.TextMatrix(llRow, PCEMAILINDEX) = slStr
                    grdContact.TextMatrix(llRow, PCARTTCODEINDEX) = "0"
                    grdContact.TextMatrix(llRow, PUSTCODEINDEX) = rst_Ust!ustCode
                    llRow = llRow + 1
                End If
            End If
        End If
        rst_Ust.MoveNext
    Loop
    grdContact.Row = 0
    grdContact.Col = PCARTTCODEINDEX
    grdContact.Redraw = True
    Exit Function
ErrHand:
   'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError LOGFILE, FORMNAME & "-mPopContactGrid"
    grdContact.Redraw = True
End Function

'Private Function mGetTitle(ilCode As Integer) As String
'
'    '***** replaced by global funtion gGetTitle with same parameters located in modGenSubs ****
'
'    mGetTitle = ""
'    SQLQuery = "Select tntCode, tntTitle From Tnt where tntCode = " & ilCode
'    Set rst_tnt = gSQLSelectCall(SQLQuery)
'    If Not rst_tnt.EOF Then
'        mGetTitle = Trim$(rst_tnt!tntTitle)
'    End If
'    Exit Function
'End Function

Private Sub mEnableBox()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilLang                        slNameCode                *
'*  slCode                        ilCode                        ilRet                     *
'*  ilLoop                                                                                *
'******************************************************************************************

'
'   mEnableBox ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'   Dan M 10/14/11 don't set if in mass email tab
    If tbcEMail.SelectedItem.Index = TABMASS Then
        Exit Sub
    End If
    If (grdContact.Row < grdContact.FixedRows) Or (grdContact.Row >= grdContact.Rows) Or (grdContact.Col < grdContact.FixedCols) Or (grdContact.Col >= grdContact.Cols - 1) Then
        Exit Sub
    End If
    lmEnableRow = grdContact.Row
    lmEnableCol = grdContact.Col
    imCtrlVisible = True

    Select Case grdContact.Col
        Case PTOINDEX
            ckcEMail.Move grdContact.Left + grdContact.ColPos(grdContact.Col) + 30, grdContact.Top + grdContact.RowPos(grdContact.Row) + 15, grdContact.ColWidth(grdContact.Col) - 30
            grdContact.Col = PTOINDEX
            grdContact.CellFontName = "Monotype Sorts"
            'If ckcAffContact.Height > grdContact.RowHeight(grdContact.Row) - 15 Then
                ckcEMail.FontName = "Arial"
                ckcEMail.Height = grdContact.RowHeight(grdContact.Row) - 15
            'End If
            If grdContact.TextMatrix(grdContact.Row, PTOINDEX) = "4" Then
                ckcEMail.Value = vbChecked
            Else
                ckcEMail.Value = vbUnchecked
            End If
            
            ckcEMail.Visible = True
            ckcEMail.SetFocus
        Case PCCINDEX
            ckcEMail.Move grdContact.Left + grdContact.ColPos(grdContact.Col) + 30, grdContact.Top + grdContact.RowPos(grdContact.Row) + 15, grdContact.ColWidth(grdContact.Col) - 30
            grdContact.Col = PCCINDEX
            grdContact.CellFontName = "Monotype Sorts"
            'If ckcISCI2Contact.Height > grdContact.RowHeight(grdContact.Row) - 15 Then
                ckcEMail.FontName = "Arial"
                ckcEMail.Height = grdContact.RowHeight(grdContact.Row) - 15
            'End If
            If grdContact.TextMatrix(grdContact.Row, PCCINDEX) = "4" Then
                ckcEMail.Value = vbChecked
            Else
                ckcEMail.Value = vbUnchecked
            End If
            
            ckcEMail.Visible = True
            ckcEMail.SetFocus
        Case PBCINDEX
            ckcEMail.Move grdContact.Left + grdContact.ColPos(grdContact.Col) + 30, grdContact.Top + grdContact.RowPos(grdContact.Row) + 15, grdContact.ColWidth(grdContact.Col) - 30
            grdContact.Col = PCCINDEX
            grdContact.CellFontName = "Monotype Sorts"
            'If ckcISCI2Contact.Height > grdContact.RowHeight(grdContact.Row) - 15 Then
                ckcEMail.FontName = "Arial"
                ckcEMail.Height = grdContact.RowHeight(grdContact.Row) - 15
            'End If
            If grdContact.TextMatrix(grdContact.Row, PBCINDEX) = "4" Then
                ckcEMail.Value = vbChecked
            Else
                ckcEMail.Value = vbUnchecked
            End If
            
            ckcEMail.Visible = True
            ckcEMail.SetFocus
    End Select
    mSetFocus
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
Private Sub mSetShow()
    Dim slStr As String
    Dim ilRet As Integer
    Dim llCode As Long

    If (lmEnableRow >= grdContact.FixedRows) And (lmEnableRow < grdContact.Rows) Then
        Select Case lmEnableCol
            Case PTOINDEX
            Case PCCINDEX
            Case PBCINDEX
        End Select
    End If
    lmEnableRow = -1
    lmEnableCol = -1
    imCtrlVisible = False
    ckcEMail.Visible = False
    mSetCommands
    Exit Sub
    
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetFocus                       *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set focus to specified control *
'*                                                     *
'*******************************************************
Private Sub mSetFocus()
'
'   mSetFocus ilBoxNo
'   Where:
'       ilBoxNo (I)- Number of the Control to be enabled
'
    Dim llColPos As Long
    Dim ilCol As Integer
    Dim llColWidth As Long

    If (grdContact.Row < grdContact.FixedRows) Or (grdContact.Row >= grdContact.Rows) Or (grdContact.Col < grdContact.FixedCols) Or (grdContact.Col >= grdContact.Cols - 1) Then
        Exit Sub
    End If
    imCtrlVisible = True
    Select Case grdContact.Col
        Case PTOINDEX
            ckcEMail.Visible = True
            ckcEMail.SetFocus
        Case PCCINDEX
            ckcEMail.Visible = True
            ckcEMail.SetFocus
        Case PBCINDEX
            ckcEMail.Visible = True
            ckcEMail.SetFocus
    End Select
End Sub

Private Function mColOk() As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slStr                         ilPos                         ilValue                   *
'*                                                                                        *
'******************************************************************************************


    mColOk = True
    If grdContact.CellBackColor = LIGHTYELLOW Then
        mColOk = False
        Exit Function
    End If
    If grdContact.CellForeColor = vbRed Then
        mColOk = False
        Exit Function
    End If

End Function


Private Sub mGetUserInfo()
    On Error GoTo ErrHand:
    smUserType = ""
    SQLQuery = "SELECT ustDntCode FROM Ust Where ustCode = " & igUstCode
    Set rst_Ust = gSQLSelectCall(SQLQuery)
    If Not rst_Ust.EOF Then
        If rst_Ust!ustDntCode > 0 Then
            SQLQuery = "SELECT dntType FROM Dnt Where dntCode = " & rst_Ust!ustDntCode
            Set rst_dnt = gSQLSelectCall(SQLQuery)
            If Not rst_dnt.EOF Then
                smUserType = rst_dnt!dntType
            End If
        End If
    End If
    Exit Sub
ErrHand:
   'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError LOGFILE, FORMNAME & "-mGetUserInfo"
End Sub


Private Function mAddComment(ilShttCode As Integer, ilToUstCode As Integer) As Integer
    Dim slDate As String
    Dim ilSource As Integer
    Dim slOK As String
    Dim slComment As String
    Dim ilVefCode As Integer
    Dim llRow As Long
    Dim slToName As String
    Dim slName As String
    
    On Error GoTo ErrHand:
    slDate = cbcFollowUp.Text
    If slDate = "" Then
        slDate = "12/31/2069"
    End If
 'Dan M 10/17/11 E-Mail: Outgoing  ilsource =3 Mass E-Mail: Outgoing ilsource = 4
   'ilSource = 3
    ilSource = mGetEmailSource()
    slOK = "N"
    slToName = ""
    slComment = gFixQuote(Trim$(edcMessage.Text)) '& Chr(0)
    'ind(3) or mass(4)
    If ilSource = 3 Then
        'get each name for to address
        For llRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
            If grdContact.TextMatrix(llRow, PCNAMEINDEX) <> "" Then
                If grdContact.TextMatrix(llRow, PTOINDEX) = "4" Then
                    slName = Trim$(grdContact.TextMatrix(llRow, PCNAMEINDEX))
                    If slName = "" Then
                        slName = Trim$(grdContact.TextMatrix(llRow, PCEMAILINDEX))
                    End If
                    If slToName = "" Then
                        slToName = "To: " & slName
                    Else
                        slToName = slToName & ", " & slName
                    End If
                End If
            End If
        Next llRow
    Else
        ' no 'to' names because a mass mailing.  Get title sending to instead.
        For llRow = 0 To lbcTitle.ListCount - 1 Step 1
            If lbcTitle.Selected(llRow) Then
                slToName = slToName & " " & lbcTitle.List(llRow)
            End If
        Next llRow
    End If

'    For llRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
'        If grdContact.TextMatrix(llRow, PCNAMEINDEX) <> "" Then
'            If grdContact.TextMatrix(llRow, PTOINDEX) = "4" Then
'                slName = Trim$(grdContact.TextMatrix(llRow, PCNAMEINDEX))
'                If slName = "" Then
'                    slName = Trim$(grdContact.TextMatrix(llRow, PCEMAILINDEX))
'                End If
'                If slToName = "" Then
'                    slToName = "To: " & slName
'                Else
'                    slToName = slToName & ", " & slName
'                End If
'            End If
'        End If
'    Next llRow
    slComment = slToName & ". " & slComment
    ilVefCode = 0
    If cbcConcern.ListIndex > 0 Then
        ilVefCode = cbcConcern.ItemData(cbcConcern.ListIndex)
    End If
    SQLQuery = "INSERT INTO cct (cctShfCode, cctVefCode, cctActionDate, cctCstCode, cctDone, cctDoneUstCode, cctDoneDate, cctDoneTime, cctChgdUstCode, cctChgdDate, cctChgdTime, cctComment, cctUstCode, cctToEMailUstCode, cctEnteredDate, cctEnteredTime)"
    SQLQuery = SQLQuery & " VALUES ("
    'SQLQuery = SQLQuery & igContactEmailShttCode & ", "
    SQLQuery = SQLQuery & ilShttCode & ", "
    SQLQuery = SQLQuery & ilVefCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(slDate, sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & ilSource & ", "
    SQLQuery = SQLQuery & "'" & slOK & "', "
    If slOK = "Y" Then
        SQLQuery = SQLQuery & igUstCode & ", "
        SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLTimeForm) & "', "
    Else
        SQLQuery = SQLQuery & 0 & ", "
        SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
        SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "
    End If
    SQLQuery = SQLQuery & 0 & ", "
    SQLQuery = SQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$("12AM", sgSQLTimeForm) & "', "
    SQLQuery = SQLQuery & "'" & gFixQuote(slComment) & "', "
    SQLQuery = SQLQuery & igUstCode & ", "
    SQLQuery = SQLQuery & ilToUstCode & ", "
    SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
    SQLQuery = SQLQuery & "'" & Format$(gNow(), sgSQLTimeForm) & "' " & ")"
    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '6/10/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError LOGFILE, "ContactEMail-mAddComment"
        mAddComment = False
        Exit Function
    End If
    mAddComment = True
    Exit Function
ErrHand:
   'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError LOGFILE, FORMNAME & "-mAddComment"
    mAddComment = False
    On Error GoTo 0
End Function
Private Function mGetEmailSource() As Integer
    ' mass or ind?
    Select Case tbcEMail.SelectedItem.Index
        Case TABMASS
            mGetEmailSource = 4
        Case Else
            mGetEmailSource = 3
    End Select
End Function
Private Sub mPopVehicle()
    Dim ilLoop As Integer
    cbcConcern.Clear
    For ilLoop = 0 To UBound(tgVehicleInfo) - 1 Step 1
        If tgVehicleInfo(ilLoop).sState = "A" Then
            cbcConcern.AddItem Trim$(tgVehicleInfo(ilLoop).sVehicle)
            cbcConcern.ItemData(cbcConcern.NewIndex) = tgVehicleInfo(ilLoop).iCode
        End If
    Next ilLoop
    cbcConcern.AddItem "[All Vehicles]", 0
    cbcConcern.ItemData(cbcConcern.NewIndex) = 0
    cbcConcern.ListIndex = 0
End Sub

Private Sub mPopTitles()
    Dim slSave As String
    
    On Error GoTo ErrHand
    lbcTitle.Clear
    SQLQuery = "Select tntCode, tntTitle From Tnt"
    Set rst_tnt = gSQLSelectCall(SQLQuery)
    While Not rst_tnt.EOF
        lbcTitle.AddItem (Trim(rst_tnt!tntTitle))
        lbcTitle.ItemData(lbcTitle.NewIndex) = rst_tnt!tntCode
        rst_tnt.MoveNext
    Wend
    lbcTitle.AddItem "[Affiliate Log E-Mail Recipient]", 0
    lbcTitle.ItemData(lbcTitle.NewIndex) = -3
    lbcTitle.AddItem "[Affiliate ISCI Export]", 0
    lbcTitle.ItemData(lbcTitle.NewIndex) = -2
    lbcTitle.AddItem "[Affiliate Labels]", 0
    lbcTitle.ItemData(lbcTitle.NewIndex) = -1
    Exit Sub
ErrHand:
   'ttp 5217
    Screen.MousePointer = vbDefault
    gHandleError LOGFILE, FORMNAME & "-mPopTitles"
End Sub


Private Sub mPopStations()
    Dim llRow As Long
    
    For llRow = frmStationSearch!grdStations.FixedRows To frmStationSearch!grdStations.Rows - 1 Step 1
        If Trim$(frmStationSearch!grdStations.TextMatrix(llRow, SCALLLETTERINDEX)) <> "" Then
            lbcStations.AddItem Trim$(frmStationSearch!grdStations.TextMatrix(llRow, SCALLLETTERINDEX))
            lbcStations.ItemData(lbcStations.NewIndex) = frmStationSearch!grdStations.TextMatrix(llRow, SSHTTCODEINDEX)
        End If
    Next llRow
    If lbcStations.ListCount <= 0 Then
        For llRow = 0 To UBound(tgStationInfo) - 1 Step 1
            lbcStations.AddItem (Trim(tgStationInfo(llRow).sCallLetters))
            lbcStations.ItemData(lbcStations.NewIndex) = tgStationInfo(llRow).iCode
        Next llRow
    End If
    If imShttCode > 0 Then
        For llRow = 0 To lbcStations.ListCount - 1 Step 1
            If lbcStations.ItemData(llRow) = imShttCode Then
                lbcStations.ListIndex = llRow
                Exit For
            End If
        Next llRow
    End If

End Sub

Private Sub mClearGrid()
    Dim llRow As Long
    Dim llCol As Long
    'dan m 10/13/11
     ckcEMail.Visible = False
    grdContact.Rows = grdContact.FixedRows + 1
    For llRow = grdContact.FixedRows To grdContact.Rows - 1 Step 1
        For llCol = 0 To grdContact.Cols - 1 Step 1
            grdContact.Row = llRow
            grdContact.Col = llCol
            grdContact.CellBackColor = vbWhite
            grdContact.Text = ""
        Next llCol
    Next llRow
End Sub
Private Function mLoseLastLetterIfComma(slInput As String) As String
    Dim llLength As Long
    Dim slNewString As String
    Dim llLastLetter As Long
    
    llLength = Len(slInput)
    llLastLetter = InStrRev(slInput, ",")
    If llLength > 0 And llLastLetter = llLength Then
        slNewString = Mid(slInput, 1, llLength - 1)
    Else
        slNewString = slInput
    End If
    mLoseLastLetterIfComma = slNewString
End Function
