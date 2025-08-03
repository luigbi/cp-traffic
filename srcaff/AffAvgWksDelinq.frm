VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmAvgWksDelinq 
   Caption         =   "Calculate Avg #  Weeks Delinquent in Posting"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   Icon            =   "AffAvgWksDelinq.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7350
   Begin V81Affiliate.CSI_Calendar CSI_CalDateThru 
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
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
      CSI_AllowBlankDate=   0   'False
      CSI_AllowTFN    =   0   'False
      CSI_DefaultDateType=   1
   End
   Begin VB.TextBox edcDateThru 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox edcTotalAvgWks 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox edcTotalWksDelinq 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox edcTotalStations 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
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
      Caption         =   "Calculate"
      Height          =   375
      Left            =   2025
      TabIndex        =   0
      Top             =   4065
      Width           =   1335
   End
   Begin VB.Label lacDateThru 
      Caption         =   "Calculating Average #  Weeks Delinquent In Posting Thru"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label labCount 
      Caption         =   "Average # Weeks Delinquent in Posting: "
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.Label labCount 
      Caption         =   "Total Agreements:"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Label labCount 
      Caption         =   "Total # Weeks Unposted/Partially Posted:"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "frmAvgWksDelinq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmAvgWksDelinq - calculate avg # of weeks behind in posting
'               Gather all unique stations from current date on back
'               Determine the # of unposted weeks for all stations
'               Divide the # of unposted weeks by the # of station
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Dim rst_ActiveStations As ADODB.Recordset    'all unique stations with active ageements
Dim rst_StationCPTT As ADODB.Recordset  'active station cptts
Dim rst_StationAtt As ADODB.Recordset   'agreements for a station

Private Sub cmdCancel_Click()
    Unload frmAvgWksDelinq
End Sub

Private Sub cmdOk_Click()
    Screen.MousePointer = vbHourglass
  
    mAvgWksDelinq
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
Dim slEndDate As String

    Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbDefault
    slEndDate = Format$(gNow(), sgShowDateForm)
    Do While Weekday(slEndDate) <> vbMonday
        slEndDate = DateAdd("d", -1, slEndDate)
    Loop
    slEndDate = DateAdd("d", -1, slEndDate)     'get to sunday of current week
    CSI_CalDateThru.Text = slEndDate
    If (StrComp(sgUserName, "Guide", 1) <> 0) Then            'not csi or guide, date is defaulted
        edcDateThru.Move lacDateThru.Left + lacDateThru.Width, lacDateThru.Top
        edcDateThru.Visible = True
        edcDateThru.Text = slEndDate
    Else                                'csi or guide, allow the user to change the date
        CSI_CalDateThru.Move lacDateThru.Left + lacDateThru.Width, lacDateThru.Top
        CSI_CalDateThru.Visible = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rst_ActiveStations.Close    'all unique station active ageements
    rst_StationCPTT.Close
    rst_StationAtt.Close
    Set frmAvgWksDelinq = Nothing
End Sub

Private Sub mAvgWksDelinq()

Dim llTotalWksDelinq As Long
Dim ilShttInx As Integer
Dim slSQLQuery As String
Dim ilHowManyStations As Integer
Dim slEndDate As String
Dim slDate As String
Dim llTotalAVg As Long
Dim ilRet As Integer
Dim llAgreeCount As Long

            On Error GoTo ErrHand

'            slEndDate = Format$(gNow(), sgShowDateForm)
'            Do While Weekday(slEndDate) <> vbMonday
'                slEndDate = DateAdd("d", -1, slEndDate)
'            Loop
'            slEndDate = DateAdd("d", -1, slEndDate)     'get to sunday of current week
            slEndDate = CSI_CalDateThru.Text
            ilRet = gPopShttInfo()
            If Not ilRet Then
                gMsgBox "gPopShttInfo failed, call Counterpoint"
                Exit Sub
            End If
            'Create table of all unique stations to process for output.  Initialze all the fields within the array.
            ilHowManyStations = 0               'init total unique stations
            llTotalWksDelinq = 0                     'init total weeks delinquent
            llAgreeCount = 0
            On Error GoTo ErrNone
'            slSQLQuery = "Select Distinct shttCode from att Left Outer Join shtt on attShfCode = shttCode Where  attOnAir <= '" & Format$(slEndDate, sgSQLDateForm) & "' Order by shttcode"
'            Set rst_ActiveStations = gSQLSelectCall(slSQLQuery)
'            While Not rst_ActiveStations.EOF
'                ilShttInx = -1
'                'stations are in sorted internal code order
'                If rst_ActiveStations!shttCode >= 0 Then
'                    ilShttInx = gBinarySearchShtt(rst_ActiveStations!shttCode)
'                End If
'                If ilShttInx = -1 Then
'                    'missing station
'                Else
'                    ilHowManyStations = ilHowManyStations + 1
'                    'read all the agreements for this station
'                    slSQLQuery = "Select attcode from att WITH(INDEX(KEY2)) Where  attshfcode = " & rst_ActiveStations!shttCode & " and attOnAir <= '" & Format$(slEndDate, sgSQLDateForm) & "' and (attoffair > attonair) "
'                    Set rst_StationAtt = gSQLSelectCall(slSQLQuery)
'                    While Not rst_StationAtt.EOF
'                        llAgreeCount = llAgreeCount + 1
'                        'obtain all the cppts unposted for this station
'                        slSQLQuery = "Select count(*) as Unposted from cptt Where  cpttatfcode = " & rst_StationAtt!attCode & " and cpttPostingStatus < 2 and cpttStartDate <= '" & Format$(slEndDate, sgSQLDateForm) & "'"
'                        Set rst_StationCPTT = gSQLSelectCall(slSQLQuery)
'                        While Not rst_StationCPTT.EOF
'                             'cpttpostingstatus <2  are those weeks unposted or partially posted
'                             llTotalWksDelinq = llTotalWksDelinq + rst_StationCPTT!Unposted
'                        rst_StationCPTT.MoveNext            'next cptt for station agreement
'                        Wend
'                    rst_StationAtt.MoveNext                     'next vehicle cptt for same staion
'                    Wend
'                End If                          'if shttinx = -1
'            rst_ActiveStations.MoveNext          'next unique station
'            Wend

            slSQLQuery = "Select  Count(If(cpttPostingStatus <= 1, 1, Null)) as Unposted, Count(distinct cpttAtfCode) as NumberAgreements From CPTT inner join att on cpttatfcode = attcode  where  cpttStartDate <= '" & Format$(slEndDate, sgSQLDateForm) & "' and attonair <= '" & Format$(slEndDate, sgSQLDateForm) & "' "
            Set rst_StationAtt = gSQLSelectCall(slSQLQuery)
            If Not rst_StationAtt.EOF Then
                llAgreeCount = rst_StationAtt!NumberAgreements
                llTotalWksDelinq = rst_StationAtt!Unposted
            End If
            lacDateThru = "Calculating Average # Weeks Delinquent in Posting thru "     '& slEndDate
            edcTotalStations(0).Text = Str$(llAgreeCount)       'Str$(ilHowManyStations)
            edcTotalWksDelinq(1).Text = Str$(llTotalWksDelinq)
'            If ilHowManyStations > 0 Then
'                llTotalAVg = llTotalWksDelinq * 10 / ilHowManyStations
'            End If
            If llAgreeCount > 0 Then            'calc the avg # weeks delinq in posting from the # of active agreements
                llTotalAVg = llTotalWksDelinq * 10 / llAgreeCount
            End If

            edcTotalAvgWks(2).Text = FormatNumber(llTotalAVg / 10, 1)
            
            If (StrComp(sgUserName, "Guide", 1) = 0) Then            ' csi or guide, let user see the # of agreements
                edcTotalStations(0).Visible = True
                labCount(0).Visible = True
            End If

            edcTotalWksDelinq(1).Visible = True
            edcTotalAvgWks(2).Visible = True
            labCount(1).Visible = True
            labCount(2).Visible = True
            lacDateThru.Visible = True
    On Error GoTo 0
    Exit Sub
ErrNone:
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "AvgWksDelinq-mAvgWksDelinq"
End Sub

