VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form EngrMessages 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9510
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5790
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox PbcInch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      ScaleHeight     =   0.476
      ScaleMode       =   7  'Centimeter
      ScaleWidth      =   0.344
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8880
      Top             =   5160
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5790
      FormDesignWidth =   9510
   End
   Begin VB.FileListBox lbcFileName 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   7560
      Pattern         =   "*.txt"
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Timer tmcUsers 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   1260
      Top             =   5295
   End
   Begin VB.PictureBox pbcArial 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   615
      ScaleHeight     =   270
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   5325
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ckcAll 
      Caption         =   "All Users"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4920
      TabIndex        =   6
      Top             =   1350
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.ListBox lbcTextFiles 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      ItemData        =   "Engrmessages.frx":0000
      Left            =   4920
      List            =   "Engrmessages.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   405
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.ListBox lbcShowFile 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      ItemData        =   "Engrmessages.frx":000E
      Left            =   240
      List            =   "Engrmessages.frx":0010
      TabIndex        =   3
      Top             =   1680
      Width           =   9015
   End
   Begin VB.ListBox lbcFileSelect 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      ItemData        =   "Engrmessages.frx":0012
      Left            =   240
      List            =   "Engrmessages.frx":0059
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   405
      Width           =   4300
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4965
      TabIndex        =   1
      Top             =   5175
      Width           =   2010
   End
   Begin VB.CommandButton cmdEmail 
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2520
      TabIndex        =   0
      Top             =   5175
      Width           =   2010
   End
   Begin VB.Image imcPrint 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   7800
      Picture         =   "Engrmessages.frx":01A3
      Top             =   5160
      Width           =   480
   End
   Begin VB.Label lblHeader 
      Caption         =   "Message Viewer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   315
      TabIndex        =   4
      Top             =   60
      Width           =   4200
   End
End
Attribute VB_Name = "EngrMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private imAllClicked As Integer
Private imSetAll As Integer
Private smChoice As String
Private imNewDisplay As Integer
Private bmPrinting As Boolean
Private smSelectedFileName As String

'not used at this time
Private Sub ckcAll_Click()

    Dim Value As Integer
    Dim llRet As Long
    Dim llRg As Long
    Dim ilValue As Integer

    lbcShowFile.Clear
    'tmcUsers.Enabled = True

    Value = False
    If ckcAll.Value = vbChecked Then
        Value = True
    Else
        lbcShowFile.Clear
    End If

    ilValue = Value
    If imSetAll Then
        imAllClicked = True
        If lbcTextFiles.ListCount > 0 Then
            llRg = CLng(lbcTextFiles.ListCount - 1) * &H10000 Or 0
            llRet = SendMessageByNum(lbcTextFiles.hwnd, LB_SELITEMRANGE, ilValue, llRg)
        End If
        imAllClicked = False
    End If

End Sub

Private Sub cmdEmail_Click()

'    Dim ilRet As Integer
'    Dim ilLoop As Integer
'    Dim slFileName As String
'    Dim fs As New FileSystemObject
'    Dim tlTxtStream As TextStream
'
'    If sgFileAttachment <> "" Then
'        ilRet = MsgBox("Would you like to attach ** " & sgFileAttachmentName & " ** to your email?", vbYesNo)
'        If ilRet = vbNo Then
'            sgFileAttachment = ""
'        End If
'    End If
'    EngrEmail.Show vbModal
'    Exit Sub
    Dim ilRet As Integer
    Dim fs As New FileSystemObject
    Dim slFileAttachment As String
   
    slFileAttachment = sgFileAttachment 'sgMsgDirectory & smSelectedFileName
    If fs.FileExists(slFileAttachment) Then
        ilRet = MsgBox("Would you like to attach ** " & smSelectedFileName & " ** to your email?", vbYesNo)
        If ilRet = vbNo Then
            slFileAttachment = ""
        End If
    Else
        slFileAttachment = ""
    End If
    Set ogEmailer = New CEmail
    ogEmailer.Attachment = slFileAttachment
    EmailGeneric.isCounterpointService = True
    EmailGeneric.isZipAttachment = True
    EmailGeneric.Show vbModal
    Set ogEmailer = Nothing

End Sub

Private Sub cmdExit_Click()
    Unload EngrMessages
End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / (800 * 15 / Me.Width)
    Me.Height = (Screen.Height) / (600 * 15 / Me.Height)
    gSetFonts EngrMessages
    gCenterForm EngrMessages
End Sub

Private Sub Form_Load()
     mInit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set EngrMessages = Nothing
End Sub

Private Sub imcPrint_Click()
Dim ilRptDest As Integer            'disply, print, save as file
Dim slRptName As String
Dim slExportName As String
Dim slRptType As String
Dim llResult As Long
Dim ilExportType As Integer
Dim llRow As Long
Dim slStr As String
Dim llTime As Long
Dim llAirDate As Long
Dim slFilter As String              'filters selected by user
Dim llLoop As Long
Dim llInch As Long
Dim slOverflow As String
Dim slLine As String
Dim llTemp As Long
Dim ilLen As Integer
Dim ilSeq As Integer

    If bmPrinting Then
        Exit Sub
    End If
    bmPrinting = True
    igRptIndex = TEXT_RPT
    'ilRptDest = 0                   'force to display for debugging, else force to Print
    ilRptDest = 1                    'force to print
    slExportName = ""               'no export for now
    slRptType = ""
    
    Set rstTextrpt = New Recordset
    gGenerateRstText     'generate the ddfs for report
    
    rstTextrpt.Open
    'build the data definition (.ttx) file in the database path for crystal to access
    llResult = CreateFieldDefFile(rstTextrpt, sgDBPath & "\TextRpt.ttx", True)
    
   
    For llRow = 0 To lbcShowFile.ListCount - 1
        ilSeq = 0
        slStr = lbcShowFile.List(llRow)
        llInch = PbcInch.TextWidth(slStr)
        slOverflow = ""
        'the width of the landscape is 10" printable space, which is approx 25 centimeters.
        'If the width of the string is greater than 25 centimeters, make as many records
        'as necessary to that it won't truncate any data.  Crystal text string is 255 max.
        'also, if formatting crystal with wrap-around, it always skips a blank line if the
        'wrap-around doesnt occur.
        Do While llInch >= 26
            For llLoop = Len(slStr) To 1 Step -1         'loop from the end of the string to find the end of a word
                
                If Mid(slStr, llLoop, 1) = " " Then      'found a blank
                    slOverflow = Mid$(slStr, 1, llLoop)
                         
                    llTemp = PbcInch.TextWidth(slOverflow)
                    If llTemp < 26 Then
                        rstTextrpt.AddNew
                        slOverflow = Mid$(slStr, 1, llLoop)     'form line to write
                        If ilSeq = 0 Then                       'if coninuation line, need to indent it (by inserting blanks)
                            rstTextrpt.Fields("Text") = Trim$(slOverflow)
                        Else
                            rstTextrpt.Fields("Text") = "    " & Trim$(slOverflow)
                        End If
                        ilLen = Len(slStr) - llLoop             'determine string that is beyond the 25 centimeters (10")
                        'start over with the remainder of the string
                        slStr = Mid$(slStr, llLoop, ilLen + 1)    'start over to get next line less than 10"
                        llInch = PbcInch.TextWidth(slStr)       'determine if greater than 10"
                        ilSeq = ilSeq + 1                       'need to know if there are continuation lines for indention
                        Exit For
                    'continue finding the blanks to strip off another word
                    End If
                End If
            Next llLoop
        Loop
        rstTextrpt.AddNew
        If ilSeq = 0 Then           'is this a continuation line, if so, ident it by inserting blanks
            rstTextrpt.Fields("Text") = Trim$(slStr)
        Else        'indent the continuation lines
            rstTextrpt.Fields("Text") = "     " & Trim$(slStr)
        End If
    Next llRow
   

    'gObtainReportforCrystal slRptName, slExportName     'determine which .rpt to call and setup an export name is user selected output to export
    igRptSource = vbModal
    slRptName = "Text.rpt"      'concatenate the crystal report name plus extension
    sgCrystlFormula2 = "'" & Trim$(sgFileAttachmentName) & " (" & smSelectedFileName & ")'"
    EngrCrystal.gActiveCrystalReports ilExportType, ilRptDest, Trim$(slRptName) & Trim$(slRptType), slExportName, rstTextrpt
    
    Screen.MousePointer = vbDefault
    
    Set rstTextrpt = Nothing
    bmPrinting = False
    Exit Sub
End Sub

Private Sub lbcFileSelect_Click()

    Dim ilLoop As Integer
    Dim slLocation As String
    Dim llRet As Long
    Dim ilPos As Integer

    'Init
    lbcTextFiles.Visible = False
    'ckcAll.Visible = False
    'ckcAll.Value = vbUnchecked
    lbcShowFile.Clear
    slLocation = ""

    'clear the horz. scroll bar if its there
    llRet = SendMessageByNum(lbcShowFile.hwnd, LB_SETHORIZONTALEXTENT, 0, 0)

    'Find out which group was selected; adv, logs, inv. etc.
    For ilLoop = 0 To lbcFileSelect.ListCount - 1 Step 1
        If lbcFileSelect.Selected(ilLoop) Then
            smChoice = Trim$(lbcFileSelect.text)
            Exit For
        End If
    Next ilLoop

    lbcShowFile.Clear
    Select Case smChoice
        Case "Auto Export"
            sgFileAttachmentName = "Auto Export"
            mPopExports
            slLocation = sgMsgDirectory & smSelectedFileName
            lbcTextFiles.Visible = True
        Case "CSIStart"
            slLocation = "C:\CSI\CSIStart.txt"
            sgFileAttachmentName = "CSIStart"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "CSIStart.txt"     'send to crystl for heading
        Case "CSISetup"
            slLocation = "c:\csi\CSISetup.txt"
            sgFileAttachmentName = "CSISetup"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "CSISetup.txt"     'send to crystl for heading
        Case "CSIUnzip"
            slLocation = sgMsgDirectory & "CSIUnzip.txt"
            sgFileAttachmentName = "CSIUnzip"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "CSIUnzip.txt"     'send to crystl for heading
        Case "DBUnzip"
            slLocation = sgMsgDirectory & "DBUnzip.txt"
            sgFileAttachmentName = "DBUnzip"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "CSIDBUnzip.txt"     'send to crystl for heading
        Case "DDF Reorg"
            slLocation = sgMsgDirectory & "DDFREorg.txt"
            sgFileAttachmentName = "DDFReorg"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "DDEReorg.txt"     'send to crystl for heading
        Case "Engineering Errors"
            sgFileAttachmentName = "Engineer Messages"
            lbcTextFiles.Visible = False
            slLocation = sgMsgDirectory & "engrerrors.txt"
            mDisplayFile slLocation
            smSelectedFileName = "EngrErrors.txt"     'send to crystl for heading
        Case "Merge Spots"
            sgFileAttachmentName = "Merge Spots"
            mPopExports
            slLocation = sgMsgDirectory & smSelectedFileName
            lbcTextFiles.Visible = True
        Case "Conflict After Merge"
            sgFileAttachmentName = "Conflict After Merge"
            mPopExports
            slLocation = sgMsgDirectory & smSelectedFileName
            lbcTextFiles.Visible = True
        Case "Conflict From Schedule"
            sgFileAttachmentName = "Conflict From Schedule"
            mPopExports
            slLocation = sgMsgDirectory & smSelectedFileName
            lbcTextFiles.Visible = True
        Case "DCart-Client"
            slLocation = sgMsgDirectory & "commport_client.txt"
            sgFileAttachmentName = "DCart-Client"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "CommPort_Client.txt"     'send to crystl for heading
        Case "DCart-Server"
            slLocation = sgMsgDirectory & "servercommport.txt"
            sgFileAttachmentName = "DCart-Server"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "ServerCommPort.txt"     'send to crystl for heading
        Case "Engineering Service Errors"
            slLocation = sgMsgDirectory & "EngrServiceErrors.txt"
            sgFileAttachmentName = "Engineer Service Error"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "EngrServiceErrors.txt"     'send to crystl for heading
       Case "Conflict After Schedule"
            sgFileAttachmentName = "Conflict After Schedule"
            mPopExports
            slLocation = sgMsgDirectory & smSelectedFileName
            lbcTextFiles.Visible = True
            smSelectedFileName = "conflictAfterSchedule*.txt"     'send to crystl for heading
        Case "Extract Library"
            slLocation = sgMsgDirectory & "ExtractLibrary.txt"
            sgFileAttachmentName = "Extract Library"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "ExtractLibrary.txt"     'send to crystl for heading
        Case "Extract Template"
            slLocation = sgMsgDirectory & "ExtractTemplate.txt"
            sgFileAttachmentName = "Extract Template"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "ExtractTemplate.txt"     'send to crystl for heading
        Case "Import Audios"
            slLocation = sgMsgDirectory & "ImportAudio.txt"
            sgFileAttachmentName = "Import Audios"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "ImportAudio.txt"     'send to crystl for heading
        Case "Import Buses"
            slLocation = sgMsgDirectory & "ImportBus.txt"
            sgFileAttachmentName = "Extract Buses"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "ImportBus.txt"     'send to crystl for heading
        Case "Import Netcues"
            slLocation = sgMsgDirectory & "ImportNetcue.txt"
            sgFileAttachmentName = "Import Netcues"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "ImportNetcue.txt"     'send to crystl for heading
        Case "Import Relays"
            slLocation = sgMsgDirectory & "ImportRelay.txt"
            sgFileAttachmentName = "Import Relays"
            lbcTextFiles.Visible = False
            mDisplayFile slLocation
            smSelectedFileName = "ImportRelay.txt"     'send to crystl for heading
   End Select

    
    'mDisplayFile slLocation
    sgFileAttachment = slLocation

End Sub
Private Sub mInit()
    sgFileAttachment = ""
    imAllClicked = False
    imSetAll = True
    bmPrinting = False
End Sub

Private Sub lbcTextFiles_Click()

    Dim ilLoop As Integer

    lbcShowFile.Clear
    'If Not imAllClicked Then
    '    imSetAll = False
    '    ckcAll.Value = vbUnchecked
    '    imSetAll = True
    'End If

    For ilLoop = 0 To lbcTextFiles.ListCount - 1 Step 1
        If lbcTextFiles.Selected(ilLoop) Then
            smSelectedFileName = Trim$(lbcTextFiles.List(ilLoop))
            'tmcUsers.Enabled = False
            'tmcUsers.Enabled = True
            Exit For
        End If
    Next ilLoop
    mDisplayFile sgMsgDirectory & smSelectedFileName

End Sub
Private Sub mDisplayFile(sLocation As String)

    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim llRet As Long
    Dim slTemp As String
    Dim slRetString As String
    Dim llMaxWidth As Long
    Dim llValue As Long
    Dim llRg As Long

    EngrMessages.pbcArial.Width = 8925
    'Make Sure we start out each time without a horizontal scroll bar
    llValue = 0
    If imNewDisplay Then
        llRet = SendMessageByNum(lbcShowFile.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    End If
    llMaxWidth = 0
    If fs.FileExists(sLocation) Then
        Set tlTxtStream = fs.OpenTextFile(sLocation, ForReading, False)
    Else
        lbcShowFile.Clear
        lbcShowFile.AddItem "** No Data Available **"
        sgFileAttachment = ""
        Exit Sub
    End If
    slTemp = ""

    Do While tlTxtStream.AtEndOfStream <> True
        slRetString = tlTxtStream.ReadLine
        lbcShowFile.AddItem slRetString
        If (EngrMessages.pbcArial.TextWidth(slRetString)) > llMaxWidth Then
            llMaxWidth = (EngrMessages.pbcArial.TextWidth(slRetString))
        End If
    Loop

    'Show a horzontal scroll bar if needed
    If llMaxWidth > lbcShowFile.Width Then
        llValue = llMaxWidth / 15 + 120
        llRg = 0
        llRet = SendMessageByNum(lbcShowFile.hwnd, LB_SETHORIZONTALEXTENT, llValue, llRg)
    End If
    imNewDisplay = False
    tlTxtStream.Close
End Sub

Private Sub lbcTextFiles_GotFocus()
    'tmcUsers.Enabled = False
End Sub

Private Sub lbcTextFiles_Scroll()
    'tmcUsers.Enabled = False
    'tmcUsers.Enabled = True
End Sub

Private Sub tmcUsers_Timer()
    'tmcUsers.Enabled = False
    imNewDisplay = True

End Sub

Public Sub mPopExports()
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slWildCard As String
    
    
    lbcFileName.Path = Left$(sgMsgDirectory, Len(sgMsgDirectory) - 1)
    
    slWildCard = ""
    Select Case smChoice
        Case "Auto Export"
            slWildCard = "AutoExport*.txt"
        Case "Merge Spots"
            slWildCard = "MergeSpots*.txt"
        Case "Conflict After Merge"
            slWildCard = "ConflictAfterMerge*.txt"
        Case "Conflict After Schedule"
            slWildCard = "ConflictAfterSchedule*.txt"
        Case "Conflict From Schedule"
            slWildCard = "ConflictFromSchedule*.txt"
        Case "Load"
            slWildCard = "Load*.txt"
        Case "Test Auto"
            slWildCard = "TestAuto*.txt"
    End Select
    lbcFileName.Pattern = slWildCard
    'Move File names to list box for user to select
    lbcTextFiles.Clear
    For ilLoop = 0 To lbcFileName.ListCount - 1 Step 1
        'All names moved into lbcSelection(0)
        'Only names of log images moved into lbcSelection(1)
        slStr = lbcFileName.List(ilLoop)
        lbcTextFiles.AddItem slStr
       
    Next ilLoop
End Sub
