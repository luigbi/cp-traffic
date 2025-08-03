VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DecodeXMLLine 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3180
   ClientLeft      =   360
   ClientTop       =   3600
   ClientWidth     =   9945
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
   Icon            =   "DecodeXMLLine.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "DoneMsg"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "DecodeXMLLine.frx":08CA
   ScaleHeight     =   3180
   ScaleWidth      =   9945
   Begin MSComDlg.CommonDialog CMDialogBox 
      Left            =   9285
      Top             =   255
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   4100
      FontSize        =   0
      MaxFileSize     =   256
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
      Left            =   8100
      TabIndex        =   3
      Top             =   1275
      Width           =   1485
   End
   Begin VB.PictureBox plcFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1035
      ScaleHeight     =   345
      ScaleWidth      =   6810
      TabIndex        =   1
      Top             =   1245
      Width           =   6840
      Begin VB.TextBox edcFrom 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   45
         TabIndex        =   2
         Top             =   105
         Width           =   6720
      End
   End
   Begin VB.CommandButton cmcImport 
      Appearance      =   0  'Flat
      Caption         =   "&Decode"
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
      TabIndex        =   4
      Top             =   2490
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
      Left            =   5775
      TabIndex        =   5
      Top             =   2490
      Width           =   1050
   End
   Begin VB.Label lacPurpose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "This applicate will convert a single line of XML Tags into separate Tag lines"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   60
      TabIndex        =   7
      Top             =   885
      Visible         =   0   'False
      Width           =   8250
   End
   Begin VB.Label lacInfo 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2760
      TabIndex        =   6
      Top             =   1935
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.Label lacFrom 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Decode File"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1005
   End
End
Attribute VB_Name = "DecodeXMLLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 1993 Counterpoint Software®, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: DecodeXMLLine.Frm
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
    cmcCancel.Caption = "&Cancel"
    cmcImport.Enabled = True
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
    ilRet = mProcessXml(slFromFile)
    If ilRet Then
        lacInfo.Caption = "Decode Successful"
    Else
        lacInfo.Caption = "Decode Failed"
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
    cmcCancel.Caption = "&Cancel"
    cmcImport.Enabled = True
End Sub

Private Sub Form_Click()
    lacInfo.Visible = False
    cmcCancel.Caption = "&Cancel"
    cmcImport.Enabled = True
End Sub

Private Sub Form_Load()
    mInit
    If imTerminate Then
        cmcCancel_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim ilRet As Integer
'Rm**    ilRet = btrReset(hgHlf)
'Rm**    btrDestroy hgHlf
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
    Set oMyFileObj = Nothing
    Unload DecodeXMLLine
    Set DecodeXMLLine = Nothing   'Remove data segment
End Sub

Private Function mProcessXml(slPath As String) As Boolean
    Dim slOutPath As String
    Dim ilPos As Integer
    Dim llChar As Long
    Dim slTag As String
    Dim llStartPos As Long
    Dim slTest As String
    
    smLines = vbNullString
    mProcessXml = False
    On Error GoTo ERRNOREAD
    Set myFile = oMyFileObj.OpenTextFile(slPath)
    smLines = myFile.ReadAll
    ilPos = InStr(1, UCase(slPath), ".TXT", vbTextCompare)
    If ilPos > 0 Then
        slOutPath = Left$(slPath, ilPos - 1) & ".Xml"
    Else
        slOutPath = slPath & ".Xml"
    End If
    Set olFileObj = CreateObject("Scripting.FileSystemObject")
    Set olFile = olFileObj.CreateTextFile(slOutPath, True, False)
    
    llChar = 1
    slTag = ""
    Do While llChar <= Len(smLines)
        llStartPos = llChar
        Do While llChar <= Len(smLines)
            slTest = Mid(smLines, llChar, 2)
            If slTest = "/>" Then
                slTag = Mid(smLines, llStartPos, llChar + 2 - llStartPos)
                olFile.WriteLine slTag
                llChar = llChar + 2
                Exit Do
            
            ElseIf slTest = "><" Then
                slTag = Mid(smLines, llStartPos, llChar + 1 - llStartPos)
                olFile.WriteLine slTag
                llChar = llChar + 1
                Exit Do
            Else
                llChar = llChar + 1
            End If
        Loop
    Loop
    If llStartPos < Len(smLines) Then
        slTag = Mid(smLines, llStartPos)
        olFile.WriteLine slTag
    End If
    mProcessXml = True
Cleanup:
    On Error Resume Next
    myFile.Close
    olFile.Close
    Set myFile = Nothing
    Exit Function
ERRNOREAD:
    GoTo Cleanup
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

