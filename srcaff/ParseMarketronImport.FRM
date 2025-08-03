VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ParseMarketronImport 
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
      ScaleWidth      =   4530
      TabIndex        =   0
      Top             =   0
      Width           =   4530
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
      Caption         =   "&Parse"
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
      Caption         =   "Parse File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   105
      TabIndex        =   1
      Top             =   825
      Width           =   1005
   End
End
Attribute VB_Name = "ParseMarketronImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: ParseMarketronImport.Frm
'
' Release: 1.0
'
' Description:
'   This file contains the Import Rate Card input screen code
Option Explicit
Option Compare Text
Dim oMyFileObj As FileSystemObject
Dim olFileObj As FileSystemObject
Dim myFile As TextStream
Dim olFile As TextStream
Dim smLines As String
Dim lmTotalNoBytes As Long
Dim lmProcessedNoBytes As Long
Dim smDefaultInitDir As String
Dim imTerminate As Integer
Dim imBSMode As Integer     'Backspace flag
Dim imBypassFocus As Integer
Dim lmCurrentRow As Long
Dim imImporting As Integer
Dim imFirstFocus As Integer 'True = cbcSelect has not had focus yet, used to branch to another control

Private Const ROWNAME As String = "<AffidavitSpot>"

Private Sub cmcBrowser_Click()
    Dim slCurDir As String
    CMDialogBox.DialogTitle = "Parse File"
    CMDialogBox.Filter = "*.txt"
    CMDialogBox.InitDir = smDefaultInitDir 'Left$(sgImportPath, Len(sgImportPath) - 1)
    CMDialogBox.DefaultExt = ".txt"
    CMDialogBox.Action = 1 'Open dialog
    edcFrom.Text = CMDialogBox.FileName
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
    Screen.MousePointer = vbHourglass
    imImporting = True
    'Dan 2/8/17 returning false incorrectly
'    ilRet = mProcessXml(slFromFile)
'    If ilRet = 0 Then
'        lacInfo.Caption = "Import Successful"
'    Else
'        lacInfo.Caption = "Import Failed"
'    End If
    If mProcessXml(slFromFile) Then
        lacInfo.Caption = "Import Successful"
    Else
        lacInfo.Caption = "Import Failed"
    End If
    lacInfo.Visible = True
    cmcCancel.Caption = "Done"
    cmcCancel.SetFocus
    cmcImport.Enabled = False
    Screen.MousePointer = vbDefault
    imImporting = False
    Exit Sub
cmcImportErr:
    ilRet = Err
    Resume Next
End Sub
Private Sub edcFrom_GotFocus()
    If imFirstFocus Then 'Test if coming from sales source- if so, branch to first control
        imFirstFocus = False
        'Show branner
    End If
    lacInfo.Visible = False
End Sub
Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  'dan 2/8/17 removed ilret and added setting to nothing
  '  Dim ilRet As Integer
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
    Set oMyFileObj = Nothing
    Set olFileObj = Nothing
    Set ParseMarketronImport = Nothing   'Remove data segment

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
    Set oMyFileObj = New FileSystemObject
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
    ilPos = InStr(1, slCommand, "/I:", 1)
    If ilPos > 0 Then
        smDefaultInitDir = ""
        ilPos = ilPos + 3
        slChar = Trim$(Mid$(slCommand, ilPos, 1))
        Do While slChar = ""
            ilPos = ilPos + 1
            slChar = Mid$(slCommand, ilPos, 1)
        Loop
        Do While slChar <> ""
            smDefaultInitDir = smDefaultInitDir & slChar
            ilPos = ilPos + 1
            If ilPos > Len(slCommand) Then
                Exit Do
            End If
            slChar = Mid$(slCommand, ilPos, 1)
        Loop
        If Right$(smDefaultInitDir, 1) = "\" Then
            smDefaultInitDir = Left$(smDefaultInitDir, Len(smDefaultInitDir) - 1)
        End If
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
    Screen.MousePointer = vbDefault
    'Dan 2/8/17 moved to unload per Dick
  '  Set oMyFileObj = Nothing
    Unload ParseMarketronImport
   ' Set ParseMarketronImport = Nothing   'Remove data segment
End Sub

Private Sub plcScreen_Paint()
    plcScreen.CurrentX = 0
    plcScreen.CurrentY = 0
    plcScreen.Print "Parse Marketron Import"
End Sub
Private Function mProcessXml(slPath As String) As Boolean
'O- slStartDate
    Dim c As Integer
    Dim ilChunk As Integer
    Dim ilMultiplier As Integer
    Dim slStatus As String
    Dim slOutPath As String
    Dim ilPos As Integer
    
    ilMultiplier = 1
    ilChunk = 20
    smLines = vbNullString
    mProcessXml = False
    On Error GoTo ERRNOREAD
    Set myFile = oMyFileObj.OpenTextFile(slPath)
    smLines = myFile.ReadAll
    If mTestFirstLine(smLines) Then
        ilPos = InStr(1, UCase(slPath), ".TXT", vbTextCompare)
        If ilPos > 0 Then
            slOutPath = Left$(slPath, ilPos - 1) & ".Xml"
        Else
            slOutPath = slPath & ".Xml"
        End If
        Set olFileObj = CreateObject("Scripting.FileSystemObject")
        Set olFile = olFileObj.CreateTextFile(slOutPath, True, False)
        Call mParseXml(smLines, "OrderStartDate")
        Call mParseXml(smLines, "Signature")
        Call mParseXml(smLines, "ProgramName")
        Call mParseXml(smLines, "CallLetters")
        Call mParseXml(smLines, "Band")
        Do While bmMoreRows
            Call mParseXml(smLines, "SpotID")
            Call mParseXml(smLines, "DateAired")
            Call mParseXml(smLines, "TimeAired")
            Call mParseXml(smLines, "Length")
            Call mParseXml(smLines, "ISCICode")
            Call mParseXml(smLines, "AffidavitSpotStatus")
        Loop
        mProcessXml = True
    End If
Cleanup:
    On Error Resume Next
    myFile.Close
    olFile.Close
    Set myFile = Nothing
    Exit Function
ERRNOREAD:
    GoTo Cleanup
End Function

Private Sub mParseXml(sllines As String, slName As String)
    Dim slStartElement As String
    Dim slEndElement As String
    Dim slvalue As String
    Dim ilPos As Long
    Dim ilEndPos As Long
    Dim ilLength As Integer
    Dim ilStart As Long
    Dim slParse As String
    
    If lmCurrentRow = 0 Then
        lmCurrentRow = 1
    End If
    slStartElement = "<" & slName & ">"
    slEndElement = "</" & slName & ">"
    ilPos = InStr(lmCurrentRow, sllines, slStartElement)
    If ilPos > 0 Then
        ilStart = ilPos + Len(slStartElement)
        ilEndPos = InStr(lmCurrentRow, sllines, slEndElement)
        ilLength = ilEndPos - ilStart
        slvalue = Mid(sllines, ilStart, ilLength)
        slParse = mUnencodeXmlData(slvalue)
        olFile.WriteLine slStartElement & slParse & slEndElement
   End If
End Sub

Private Function mTestFirstLine(slLine As String) As Boolean
    If InStr(1, slLine, "<Affidavit>") = 1 Then
        mTestFirstLine = True
    Else
        mTestFirstLine = False
    End If
End Function
Private Function mUnencodeXmlData(slData As String) As String
    Dim slRet As String
    If InStr(1, slData, "&") > 0 Then
        slRet = Replace(slData, "&lt;", "<")
        slRet = Replace(slRet, "&gt;", ">")
        slRet = Replace(slRet, "&amp;", "&")
        slRet = Replace(slRet, "&apos;", "`")
        slRet = Replace(slRet, "&quot;", """")
        mUnencodeXmlData = slRet
    Else
        mUnencodeXmlData = slData
    End If
End Function
Private Property Get bmMoreRows() As Boolean
    'find next row name
    Dim llPos As Long
    
    If lmCurrentRow = 0 Then
        lmCurrentRow = 1
    End If
    llPos = InStr(lmCurrentRow + 1, smLines, ROWNAME)
    If llPos > 0 Then
        lmCurrentRow = llPos
        bmMoreRows = True
    Else
        lmCurrentRow = 0
        bmMoreRows = False
    End If
End Property

