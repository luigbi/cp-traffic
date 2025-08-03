VERSION 5.00
Begin VB.Form ExpGPBarter 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3240
   ClientLeft      =   825
   ClientTop       =   2400
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3240
   ScaleWidth      =   7095
   Begin VB.PictureBox plcMethods 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   4290
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   4290
      Begin VB.CheckBox ckcMethod 
         Caption         =   "Year"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3000
         TabIndex        =   18
         Top             =   0
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CheckBox ckcMethod 
         Caption         =   "Month"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1920
         TabIndex        =   17
         Top             =   0
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CheckBox ckcMethod 
         Caption         =   "Week"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   960
         TabIndex        =   16
         Top             =   0
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.ListBox lbcSelection 
      Appearance      =   0  'Flat
      Height          =   810
      ItemData        =   "ExpGPBarter.frx":0000
      Left            =   4560
      List            =   "ExpGPBarter.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   14
      Top             =   780
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox ckcInclRecd 
      Caption         =   "Update Barter Paid"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   885
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.CheckBox ckcInclPreviousPaid 
      Caption         =   "Include previously paid stations"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1155
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Timer tmcCancel 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6570
      Top             =   195
   End
   Begin VB.TextBox edcSelCFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   2235
      MaxLength       =   3
      TabIndex        =   0
      Top             =   570
      Width           =   615
   End
   Begin VB.TextBox edcSelCFrom1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   300
      Left            =   3810
      MaxLength       =   4
      TabIndex        =   1
      Top             =   570
      Width           =   615
   End
   Begin VB.TextBox edcLinkDestDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6555
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1305
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkDestHelpMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6450
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox edcLinkSrceDoneMsg 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6465
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmcExport 
      Appearance      =   0  'Flat
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2415
      TabIndex        =   2
      Top             =   2820
      Width           =   1050
   End
   Begin VB.CommandButton cmcCancel 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   3705
      TabIndex        =   3
      Top             =   2820
      Width           =   1050
   End
   Begin VB.Label lacTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Great Plains Barter Export"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   11
      Top             =   75
      Width           =   2610
   End
   Begin VB.Label lacSelCFrom1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3180
      TabIndex        =   10
      Top             =   630
      Width           =   420
   End
   Begin VB.Label lacSelCFrom 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Bdcst Month"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   630
      Width           =   1905
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.Image imcHelp 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      Top             =   2775
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   3390
   End
End
Attribute VB_Name = "ExpGPBarter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of ExpGPBarter.frm on Fri 3/12/10 @ 11:00 AM *
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Dim imFirstActivate As Integer
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmTo As Integer   'From file hanle
Dim imTerminate As Integer
Dim imBypassFocus As Integer
Dim imExporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim smStartStd As String    'Starting date for standard billing
Dim smEndStd As String      'Ending date for standard billing
Dim smStartCal As String    'Starting date for standard billing
Dim smEndCal As String      'Ending date for standard billing
Dim lmStartStd As Long    'Starting date for standard billing
Dim lmEndStd As Long      'Ending date for standard billing
Dim lmStartCal As Long
Dim lmEndCal As Long
'*******************************************************
'*                                                     *
'*      Procedure Name:mRepVeh                         *
'*                                                     *
'*             Created:9-3-02       By:D. Hosaka       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Populate the selection combo   *
'*                      box with REP vehicles only     *
'*                                                     *
'*******************************************************
Private Sub mRepVeh()
    Dim ilRet As Integer
        ilRet = gPopUserVehicleBox(ExpGPBarter, ACTIVEVEH + VEHREP_WO_CLUSTER + VEHREP_W_CLUSTER, lbcSelection, tgCSVNameCode(), sgCSVNameCodeTag)     'lbcCSVNameCode)
    If ilRet <> CP_MSG_NOPOPREQ Then
        On Error GoTo mRepVehERr
        gCPErrorMsg ilRet, "mRepVeh (gPopUserVehicleBox: Vehicle)", ExpGPBarter
        On Error GoTo 0
    End If
    Exit Sub
mRepVehERr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub cmcCancel_Click()

       If imExporting Then
           imTerminate = True
           Exit Sub
       End If
       mTerminate

End Sub

Private Sub cmcExport_Click()
    Dim ilYear As Integer
       Dim slStr As String
       Dim slToFile As String
       Dim ilRet As Integer
       Dim slDateTime As String
       Dim slMonth As String
       Dim slStr1 As String
       Dim ilCurrentMonth As Integer

       lacInfo(0).Visible = False
       lacInfo(1).Visible = False
       If imExporting Then
           Exit Sub
       End If
       On Error GoTo ExportError
       
        If Not Len(edcSelCFrom.Text) > 0 Then
            ''MsgBox "Month Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Month"
            gAutomationAlertAndLogHandler "Month Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Month"
            edcSelCFrom.SetFocus
            Exit Sub
        End If
        
        gGetMonthNoFromString edcSelCFrom.Text, ilCurrentMonth         'getmonth #
        If ilCurrentMonth = 0 Then                                 'input isn't text month name, try month #
            ilCurrentMonth = Val(slStr1)
            ''MsgBox "Month Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Month"
            gAutomationAlertAndLogHandler "Month Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Month"
        End If
        
        If Not Len(edcSelCFrom1.Text) > 0 Then
            ''MsgBox "Year Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Year"
            gAutomationAlertAndLogHandler "Year Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Year"
            edcSelCFrom1.SetFocus
            Exit Sub
        End If
        
        ilYear = Val(edcSelCFrom1.Text)
        If ilYear < 1970 Or ilYear > 2026 Then
            ''MsgBox "Year Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Year"
            gAutomationAlertAndLogHandler "Year Entered is Not Valid", vbOkOnly + vbApplicationModal, "Enter Year"
            edcSelCFrom1.SetFocus
            Exit Sub
        End If

        slStr1 = Trim$(str(ilCurrentMonth)) & "/15/" & Trim$(ExpGP!edcSelCFrom1.Text)     'form mm/dd/yy
        smStartStd = gObtainStartStd(slStr1)               'obtain std start date for month
        lmStartStd = gDateValue(smStartStd)
        smEndStd = gObtainEndStd(slStr1)                 'obtain std end date for month
        lmEndStd = gDateValue(smEndStd)

        slMonth = Trim$(str(ilCurrentMonth))
        If Len(Trim$(str(ilCurrentMonth))) = 1 Then
            slMonth = "0" & slMonth
        End If
        slToFile = sgExportPath & "Barter_" & slMonth & Trim$(edcSelCFrom1.Text) & ".csv"
        If DoesFileExist(slToFile) Then
            Kill slToFile
        End If
        If (InStr(slToFile, ":") = 0) And (Left$(slToFile, 2) <> "\\") Then
            slToFile = sgExportPath & slToFile
        End If
        ilRet = 0
        'On Error GoTo cmcExportErr:
        'slDateTime = FileDateTime(slToFile)
        ilRet = gFileExist(slToFile)
        If ilRet = 0 Then
            'hmTo = FreeFile
            'Open slToFile For Append As hmTo
            ilRet = gFileOpen(slToFile, "Append", hmTo)
            If ilRet <> 0 Then
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                Exit Sub
            End If
         Else
            ilRet = 0
            'hmTo = FreeFile
            'Open slToFile For Output As hmTo
            ilRet = gFileOpen(slToFile, "Output", hmTo)
            If ilRet <> 0 Then
                ''MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                gAutomationAlertAndLogHandler "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                Exit Sub
            End If
         End If

        

      Screen.MousePointer = vbHourglass
      imExporting = True
    sgMessageFile = sgDBPath & "Messages\" & "ExportBarter.txt"
    'gLogMsg "** Barter Export **", "ExportGreatPlains.txt", False
    
    gAutomationAlertAndLogHandler "** Export Barter **"
    gAutomationAlertAndLogHandler "* Export To = " & slToFile
    gAutomationAlertAndLogHandler "* Month = " & edcSelCFrom.Text
    gAutomationAlertAndLogHandler "* Year = " & edcSelCFrom1.Text
    

    'write out headers
    'slStr = """" & "Vendor ID" & """" & ","
    'slStr = slStr & """" & "Branch Code" & """" & ","
    'slStr = slStr & """" & "Date" & """" & ","
    'slStr = slStr & """" & "Contract" & """" & ","
    'slStr = slStr & """" & "Advertiser/Product" & """" & ","
    'slStr = slStr & """" & "Due Station" & """"
    
    slStr = "As of " & Format$(gNow(), "mm/dd/yy") & " "
    slStr = slStr & Format$(gNow(), "h:mm:ssAM/PM")
    slStr = slStr & " Barter Payment for " & edcSelCFrom.Text & " " & edcSelCFrom1.Text
    'On Error GoTo cmcExportErr
    Print #hmTo, slStr        'write header description
    On Error GoTo 0
    
    'slStr = "Vendor ID, Branch Code, Date, Contract, Advertiser/Product, Due Station"
    'write out headers
    slStr = """" & "Vendor ID" & """" & ","
    slStr = slStr & """" & "Branch Code" & """" & ","
    slStr = slStr & """" & "Date" & """" & ","
    slStr = slStr & """" & "Contract" & """" & ","
    slStr = slStr & """" & "Advertiser/Product" & """" & ","
    slStr = slStr & """" & "Due Station" & """"
    'On Error GoTo cmcExportErr
    Print #hmTo, slStr     'write header description
    On Error GoTo 0

      ilRet = gGenBarterPayment(ExpGPBarter, hmTo)
      If ilRet <> -1 Then       'error will be an error code
          lacInfo(0).Caption = "Export Failed"
          gLogMsg "Export failed: #" & Trim$(str$(ilRet)), "ExportBarter.txt", False
      Else
          lacInfo(0).Caption = "Export Successfully Completed"
          gLogMsg "Export Successfully Completed, Export Files: " & slToFile, "ExportBarter.txt", False
      End If
      lacInfo(1).Caption = "Export Files: " & slToFile

      lacInfo(0).Visible = True
      lacInfo(1).Visible = True
      Close hmTo
      cmcCancel.Caption = "&Done"
      cmcCancel.SetFocus
      cmcExport.Enabled = False
      Screen.MousePointer = vbDefault
      imExporting = False
       Exit Sub
'cmcExportErr:
'        ilRet = Err.Number
'        Resume Next

ExportError:
    gAutomationAlertAndLogHandler "Export Terminated, " & "Errors starting export..." & err & " - " & Error(err)
    
End Sub





Private Sub Form_Activate()

       If Not imFirstActivate Then
           DoEvents    'Process events so pending keys are not sent to this
           Me.KeyPreview = True
           Exit Sub
       End If
       imFirstActivate = False
       DoEvents    'Process events so pending keys are not sent to this
       Me.KeyPreview = True
       Me.Refresh
       edcSelCFrom.SetFocus

End Sub
Private Sub Form_Deactivate()

       Me.KeyPreview = False

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

       If (KeyCode = KEYF1) Or (KeyCode = KEYF4) Or (KeyCode = KEYF5) Then
           gFunctionKeyBranch KeyCode
       End If

End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)

       sgDoneMsg = CmdStr
       igChildDone = True
       Cancel = 0

End Sub
Private Sub Form_Load()

    mInit
    If imTerminate Then
        'cmcCancel_Click
        tmcCancel.Enabled = True
        Me.Left = 2 * Screen.Width      'move off the screen so screen won't flash
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ExpGPBarter = Nothing   'Remove data segment
End Sub

Private Sub tmcCancel_Timer()
    tmcCancel.Enabled = False       'screen has now been focused to show
    cmcCancel_Click         'simulate clicking of cancen button
End Sub
Private Sub mInit()

        Dim ilRet As Integer
        Dim slStdDate As String
        Dim llRg As Long
        Dim llRet As Long
        Dim ilValue As Integer
        
        imTerminate = False
        imFirstActivate = True
        Screen.MousePointer = vbHourglass
        imExporting = False
        imFirstFocus = True
        imBypassFocus = False
        lmTotalNoBytes = 0
        lmProcessedNoBytes = 0
        ilValue = True
       
        mRepVeh         'load all valid vehicles in list box
        llRg = CLng(lbcSelection.ListCount - 1) * &H10000 Or 0
        llRet = SendMessageByNum(lbcSelection.HWnd, LB_SELITEMRANGE, ilValue, llRg)

        gCenterStdAlone ExpGPBarter
        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStdDate
        edcSelCFrom.Text = Format$(slStdDate, "MMM")
        edcSelCFrom1.Text = Year(slStdDate)
        Screen.MousePointer = vbDefault
        
        gAutomationAlertAndLogHandler ""
        gAutomationAlertAndLogHandler "Selected Export=" & ExportList.lbcExport.List(ExportList.lbcExport.ListIndex)
        Exit Sub

mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

Private Sub mTerminate()

        Dim ilRet As Integer
        Screen.MousePointer = vbDefault
        igManUnload = YES
        Unload ExpGPBarter
        igManUnload = NO

End Sub


