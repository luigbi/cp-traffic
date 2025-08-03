VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PodcastCntrImpt 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2625
   ClientLeft      =   360
   ClientTop       =   3600
   ClientWidth     =   8700
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
   Icon            =   "PodcastCntrImpt.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   8700
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   6420
      Top             =   1155
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.PictureBox plcScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   45
      ScaleHeight     =   270
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   0
      Width           =   2505
   End
   Begin VB.CommandButton cmcBrowser 
      Appearance      =   0  'Flat
      Caption         =   "&Browse..."
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
      Left            =   6855
      TabIndex        =   4
      Top             =   780
      Width           =   1485
   End
   Begin VB.PictureBox plcFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1035
      ScaleHeight     =   345
      ScaleWidth      =   5640
      TabIndex        =   2
      Top             =   750
      Width           =   5670
      Begin VB.TextBox edcFrom 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   45
         TabIndex        =   3
         Top             =   105
         Width           =   5580
      End
   End
   Begin VB.CommandButton cmcImport 
      Appearance      =   0  'Flat
      Caption         =   "&Import"
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
      Left            =   3225
      TabIndex        =   5
      Top             =   1800
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
      Left            =   4530
      TabIndex        =   6
      Top             =   1800
      Width           =   1050
   End
   Begin VB.Label lacToDB 
      Appearance      =   0  'Flat
      Caption         =   "Data To:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   135
      TabIndex        =   8
      Top             =   2265
      Width           =   7635
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2760
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.Label lacFrom 
      Appearance      =   0  'Flat
      Caption         =   "From File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   105
      TabIndex        =   1
      Top             =   825
      Width           =   810
   End
End
Attribute VB_Name = "PodcastCntrImpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: PodcastCntrImpt.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Import Rate Card input screen code
Option Explicit
Option Compare Text
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim hmFrom As Integer   'From file hanle
Dim smDefaultInitDir As String
Dim imTerminate As Integer
Dim imUserInput  As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim imImporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control
Dim lmNowDate As Long
Dim imNowYear As Integer
Dim imShowHelpMsg As Integer    'True=Show help message; False=Ignore help message system

Dim smAdvtName As String
Dim smProduct As String
Dim hmAdds As Integer
Dim bmItemsAdded As Boolean


'' MsgBox parameters
'Const vbOkOnly = 0                 ' OK button only
'Const vbCritical = 16          ' Critical message
'Const vbApplicationModal = 0
'Const INDEXKEY0 = 0
Private Sub cmcBrowser_Click()
    Dim slCurDir As String
    CMDialogBox.DialogTitle = "From File"
    CMDialogBox.Filter = "Import|PE*.csi"
    CMDialogBox.InitDir = smDefaultInitDir 'Left$(sgImportPath, Len(sgImportPath) - 1)
    CMDialogBox.DefaultExt = ".csi"
    CMDialogBox.Action = 1 'Open dialog
    edcFrom.Text = CMDialogBox.fileName
    slCurDir = CurDir$
    If Left$(slCurDir, 2) <> Left$(sgCurDir, 2) Then
        ChDrive Left$(sgCurDir, 2)    'windows 95 requires drive to be changed, then directory
    End If
    ChDir sgCurDir
    DoEvents
End Sub
Private Sub cmcBrowser_GotFocus()
    lacInfo.Visible = False
End Sub
Private Sub cmcCancel_Click()
    If imImporting Then
        imTerminate = True
        Exit Sub
    End If
    mTerminate
End Sub
Private Sub cmcImport_Click()
    Dim slFromFile As String
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slNameCode As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim slStr As String
    Dim ilDays As Integer
    Dim slName As String
    lacInfo.Visible = False
    If imImporting Then
        Exit Sub
    End If
    slFromFile = Trim$(edcFrom.Text)
    If Len(slFromFile) = 0 Then
        Beep
        edcFrom.SetFocus
        Exit Sub
    End If
    If InStr(slFromFile, ":") = 0 Then
        slFromFile = sgImportPath & slFromFile
    End If
    Screen.MousePointer = vbHourglass
    imImporting = True
    
    gLogMsgWODT "O", hmAdds, sgDBPath & "Messages\" & "PodcastImportAdds_" & Format(Now, "mmddyyyy") & ".csv"
    gLogMsgWODT "W", hmAdds, "Podcast Import Items added as of " & Format(Now, "m/d/yy") & " " & Format(Now, "h:mm:ssAM/PM")
    bmItemsAdded = False
    
    If ilRet = 0 Then
        lacInfo.Caption = "Import Successful"
    Else
        lacInfo.Caption = "Import Failed"
    End If
    
    If Not bmItemsAdded Then
        gLogMsgWODT "W", hmAdds, "No New Items Added"
    End If
    gLogMsgWODT "C", hmAdds, ""
        
    lacInfo.Visible = True
    cmcCancel.Caption = "Done"
    cmcCancel.SetFocus
    cmcImport.Enabled = False
    Screen.MousePointer = vbDefault
    imImporting = False
    If Not imUserInput Then
        imTerminate = True
        cmcCancel_Click
    End If
    Exit Sub
cmcImportErr:
    ilRet = err
    Resume Next
End Sub
Private Sub edcFrom_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    lacInfo.Visible = False
    gCtrlGotFocus ActiveControl
End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    sgDoneMsg = CmdStr
    igChildDone = True
    Cancel = 0
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
    On Error Resume Next
    btrStopAppl
    End
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInit                           *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
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
    Dim llNoRec As Long 'Number of records
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slDate As String
    imTerminate = False
    mParseCmmdLine
    Screen.MousePointer = vbHourglass
    imImporting = False
    imFirstFocus = True
    imBypassFocus = False
    lmTotalNoBytes = 0
    lmProcessedNoBytes = 0
    slDate = Format$(gNow(), "m/d/yy")   'Get year
    lmNowDate = gDateValue(slDate)
    lacToDB.Caption = "Data To: " & Left$(sgDBPath, Len(sgDBPath) - 1)
    If imTerminate Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    gCenterStdAlone PodcastCntrImpt
    Screen.MousePointer = vbDefault
    Exit Sub
mInitErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mInitBox                        *
'*                                                     *
'*             Created:6/30/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Set mouse and control locations*
'*                                                     *
'*******************************************************
Private Sub mInitBox()
'
'   mInitBox
'   Where:
'
    Dim flTextHeight As Single  'Standard text height
    Dim ilLoop As Integer
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mParseCmmdLine                  *
'*                                                     *
'*             Created:5/17/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Parse command line             *
'*                                                     *
'*******************************************************
Private Sub mParseCmmdLine()
    Dim slCommand As String
    Dim slStr As String
    Dim ilRet As Integer
    Dim slTestSystem As String
    Dim ilTestSystem As Integer
    Dim slHelpSystem As String
    Dim ilPos As Integer
    Dim slChar As String
    slCommand = Command$
    smDefaultInitDir = ""
    igStdAloneMode = True 'False  'Switch from/to stand alone mode
    sgCallAppName = ""
    slStr = "Guide"
    ilTestSystem = False
    ilPos = InStr(1, slCommand, "/S:Test", 1)
    If ilPos > 0 Then
        ilTestSystem = True
    End If
    ilPos = InStr(1, slCommand, "/S:Prod", 1)
    If ilPos > 0 Then
        ilTestSystem = False
    End If


    gInitGlobalVar

    ilPos = InStr(1, slCommand, "/UserInput", 1)
    If ilPos > 0 Then
        imUserInput = True
        igBkgdProg = 0
    Else
        imUserInput = False
        igBkgdProg = 22
        Me.Visible = False
    End If

    If imUserInput Then
        Signon.Show vbModal
        If igExitTraffic Then
            imTerminate = True
            Exit Sub
        End If

    End If
    gInitStdAlone PodcastCntrImpt, slStr, ilTestSystem
    If Len(smDefaultInitDir) = 0 Then
        smDefaultInitDir = Left$(sgImportPath, Len(sgImportPath) - 1)
    End If
End Sub
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
    igParentRestarted = False
    'If Not igStdAloneMode Then
    '    If StrComp(sgCallAppName, "Traffic", 1) = 0 Then
    '        edcLinkDestHelpMsg.LinkExecute "@" & "Done"
    '    Else
    '        edcLinkDestHelpMsg.LinkMode = vbLinkNone    'None
    '        edcLinkDestHelpMsg.LinkTopic = sgCallAppName & "|DoneMsg"
    '        edcLinkDestHelpMsg.LinkItem = "edcLinkSrceDoneMsg"
    '        edcLinkDestHelpMsg.LinkMode = vbLinkAutomatic    'Automatic
    '        edcLinkDestHelpMsg.LinkExecute "Done"
    '    End If
    '    Do While Not igParentRestarted
    '        DoEvents
    '    Loop
    'End If
    Screen.MousePointer = vbDefault
    igManUnload = YES
    Unload PodcastCntrImpt
    Set PodcastCntrImpt = Nothing   'Remove data segment
    igManUnload = NO
End Sub
Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Podcast Contract Import"
End Sub

Private Function mAddAdv() As Boolean

    Dim slSQLQuery As String
    Dim llPrfCode As Long
    
    mAddAdv = False
    sgAdfname = Trim(smAdvtName)
    sgAdfAbbrv = Mid$(Trim(smAdvtName), 1, 7)
    sgProdName = Trim(smProduct)
    slSQLQuery = gDefaultADF("R")
    lgAdfCode = gInsertAndReturnCode(slSQLQuery, "ADF_Advertisers", "AdfCode", "Replace")
    gLogMsgWODT "W", hmAdds, "Advertiser Added: " & smAdvtName

    mAddAdv = True
End Function
Private Function mAddProd() As Boolean

    Dim slSQLQuery As String
    Dim llPrfCode As Long
    
    mAddProd = False
    If Len(smProduct) > 0 Then
        sgProdName = smProduct
        slSQLQuery = gDefaultPRF()
        llPrfCode = gInsertAndReturnCode(slSQLQuery, "PRF_Product_Names", "prfCode", "Replace")
        slSQLQuery = "UPDATE PRF_Product_Names SET "
        slSQLQuery = slSQLQuery + "prfAutoCode = " & llPrfCode
        slSQLQuery = slSQLQuery + " WHERE (prfCode = " & llPrfCode & ")"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            gHandleError "PodcastCntrImpt.Txt", "PodcastCntrImpt: mAddProd"
        Else
            gLogMsgWODT "W", hmAdds, "Product Added: " & smProduct
        End If
    End If
    mAddProd = True
    
End Function

Private Function mAddAgy() As Boolean

    Dim slSQLQuery As String
    Dim ilAgfCode As Long
    
    mAddAgy = False
    slSQLQuery = gDefaultAgf()
    ilAgfCode = gInsertAndReturnCode(slSQLQuery, "Agf_Agenies", "agfCode", "Replace")
    slSQLQuery = "UPDATE Agf_Agenies SET "
    slSQLQuery = slSQLQuery + "agfAutoCode = " & ilAgfCode
    slSQLQuery = slSQLQuery + " WHERE (agfCode = " & ilAgfCode & ")"
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        gHandleError "PodcastCntrImpt.Txt", "PodcastCntrImpt: mAddAgy"
    Else
        gLogMsgWODT "W", hmAdds, "Agency Added: " & sgAgfname & "," & sgAgfCity
    End If
    mAddAgy = True
    
End Function
Private Function mAddSlf() As Boolean

    Dim slSQLQuery As String
    Dim ilSlfCode As Long
    
    mAddSlf = False
    slSQLQuery = gDefaultSlf()
    ilSlfCode = gInsertAndReturnCode(slSQLQuery, "SLF_Salespeople", "slfCode", "Replace")
    slSQLQuery = "UPDATE SLF_Salespeople SET "
    slSQLQuery = slSQLQuery + "slfAutoCode = " & ilSlfCode
    slSQLQuery = slSQLQuery + " WHERE (slfCode = " & ilSlfCode & ")"
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        gHandleError "PodcastCntrImpt.Txt", "PodcastCntrImpt: mAddSlf"
    Else
        gLogMsgWODT "W", hmAdds, "Salesperson Added: " & sgSlfFirstName & " " & sgSlfLastName
    End If
    mAddSlf = True
    
End Function
