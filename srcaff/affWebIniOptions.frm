VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmWebIniOptions 
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6360
   Icon            =   "affWebIniOptions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox edcFTPImportDir 
      Height          =   360
      Left            =   1680
      TabIndex        =   21
      Top             =   4320
      Width           =   4095
   End
   Begin VB.TextBox edcFTPExportDir 
      Height          =   360
      Left            =   1680
      TabIndex        =   20
      Top             =   4920
      Width           =   4095
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CheckBox chkFtpIsOn 
      Caption         =   "FTP Is On"
      Height          =   360
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox edcFTPPWD 
      Height          =   360
      Left            =   1680
      TabIndex        =   9
      Top             =   6120
      Width           =   4095
   End
   Begin VB.TextBox edcFTPUID 
      Height          =   360
      Left            =   1680
      TabIndex        =   8
      Top             =   5520
      Width           =   4095
   End
   Begin VB.TextBox edcFTPPort 
      Height          =   360
      Left            =   1680
      TabIndex        =   7
      Top             =   3720
      Width           =   4095
   End
   Begin VB.TextBox edcWebImports 
      Height          =   360
      Left            =   1680
      TabIndex        =   3
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox edcWebExports 
      Height          =   360
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox edcRegSection 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox edcRootURL 
      Height          =   360
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   6720
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7380
      FormDesignWidth =   6360
   End
   Begin VB.TextBox edcFTPAddress 
      Height          =   360
      Left            =   1680
      TabIndex        =   6
      Top             =   3120
      Width           =   4095
   End
   Begin VB.Label lbcWebType 
      Caption         =   "Production Website"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "FTPImportDir"
      Height          =   315
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "FTPExportDir"
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "FTPPWD"
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "FTPUID"
      Height          =   300
      Left            =   120
      TabIndex        =   15
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "FTPPort"
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "WebImports"
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "WebExports"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "RegSection"
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "RootURL"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "FTPAddress"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "frmWebIniOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private imChangesOccured As Boolean



Private Sub Form_Load()
    
    frmWebIniOptions.Caption = "Affiliate Website INI Settings - " & sgClientName
    Me.Width = Screen.Width / 1.5
    Me.Height = Screen.Height / 1.6
    Me.Top = (Screen.Height - Me.Height) / 2.5
    Me.Left = (Screen.Width - Me.Width) / 2

    igWebIniOptionsOK = True
    imChangesOccured = False
    mInit
End Sub

Private Sub cmdDone_Click()
    If mSave(True) Then
        Unload Me
    End If
End Sub

Private Sub cmdSave_Click()
    mSave (False)
End Sub

Private Sub cmdCancel_Click()
    igWebIniOptionsOK = False
    Unload Me
End Sub

Private Sub mInit()
    Dim blRet As Boolean
    Dim sBuffer As String
    '10000
    lbcWebType.FontSize = 6
    If igDemoMode Then
        lbcWebType.Caption = "Demo Mode"
    ElseIf gIsTestWebServer() Then
        lbcWebType.Caption = "Test Website"
    End If
    
    mLoadINIValue sgWebServerSection, "RootURL", edcRootURL
    mLoadINIValue sgWebServerSection, "RegSection", edcRegSection
    mLoadINIValue sgWebServerSection, "WebExports", edcWebExports
    mLoadINIValue sgWebServerSection, "WebImports", edcWebImports
    mLoadINIValue sgWebServerSection, "FTPAddress", edcFTPAddress
    mLoadINIValue sgWebServerSection, "FTPPort", edcFTPPort
    mLoadINIValue sgWebServerSection, "FTPImportDir", edcFTPImportDir
    mLoadINIValue sgWebServerSection, "FTPExportDir", edcFTPExportDir
    mLoadINIValue sgWebServerSection, "FTPUID", edcFTPUID
    mLoadINIValue sgWebServerSection, "FTPPWD", edcFTPPWD
    mLoadINIValue sgWebServerSection, "FTPIsOn", chkFtpIsOn
End Sub

Private Sub mLoadINIValue(Section As String, Key As String, Ctrl As control)
    Dim sBuffer As String
    Dim ilRet As Boolean

    ilRet = gLoadOption(Section, Key, sBuffer)
    If TypeOf Ctrl Is TextBox Then
        Ctrl.Text = sBuffer
    ElseIf TypeOf Ctrl Is CheckBox Then
        Ctrl.Value = Val(sBuffer)
    End If
    If Len(sBuffer) < 1 Then
        ' Ctrl.SetFocus
    End If
End Sub

Private Function mSave(iAsk As Boolean) As Integer

    mSave = True
    If Not imChangesOccured Then
        Exit Function
    End If
    If iAsk Then
        If gMsgBox("Save all changes?", vbYesNo) <> vbYes Then
            Exit Function
        End If
    End If
    If Not gSaveOption(sgWebServerSection, "RootURL", edcRootURL.Text) Then
        gMsgBox "Unable to write RootURL value.", vbCritical
        Exit Function
    End If
    If Not gSaveOption(sgWebServerSection, "RegSection", edcRegSection.Text) Then
        gMsgBox "Unable to write RegSection value.", vbCritical
        Exit Function
    End If
    If Not gSaveOption(sgWebServerSection, "WebExports", edcWebExports.Text) Then
        gMsgBox "Unable to write WebExports value.", vbCritical
        Exit Function
    End If
    If Not gSaveOption(sgWebServerSection, "WebImports", edcWebImports.Text) Then
        gMsgBox "Unable to write WebImports value.", vbCritical
        Exit Function
    End If
    If Not gSaveOption(sgWebServerSection, "FTPAddress", edcFTPAddress.Text) Then
        gMsgBox "Unable to write FTPAddress value.", vbCritical
        Exit Function
    End If
    If Not gSaveOption(sgWebServerSection, "FTPPort", edcFTPPort.Text) Then
        gMsgBox "Unable to write FTPPort value.", vbCritical
        Exit Function
    End If
    If Not gSaveOption(sgWebServerSection, "FTPImportDir", edcFTPImportDir.Text) Then
        gMsgBox "Unable to write FTPImportDir value.", vbCritical
        Exit Function
    End If
    If Not gSaveOption(sgWebServerSection, "FTPExportDir", edcFTPExportDir.Text) Then
        gMsgBox "Unable to write FTPExportDir value.", vbCritical
        Exit Function
    End If
    If Not gSaveOption(sgWebServerSection, "FTPUID", edcFTPUID.Text) Then
        gMsgBox "Unable to write FTPUID value.", vbCritical
        Exit Function
    End If
    If Not gSaveOption(sgWebServerSection, "FTPPWD", edcFTPPWD.Text) Then
        gMsgBox "Unable to write FTPPWD value.", vbCritical
        Exit Function
    End If
    If Not gSaveOption(sgWebServerSection, "FTPIsOn", chkFtpIsOn.Value) Then
        gMsgBox "Unable to write FTPIsOn value.", vbCritical
        Exit Function
    End If
    imChangesOccured = False
    mSave = True
End Function

Private Sub edcFTPAddress_Change()
    imChangesOccured = True
End Sub

Private Sub edcFTPPort_Change()
    imChangesOccured = True
End Sub

Private Sub edcFTPPWD_Change()
    imChangesOccured = True
End Sub

Private Sub edcFTPUID_Change()
    imChangesOccured = True
End Sub

Private Sub edcRegSection_Change()
    imChangesOccured = True
End Sub

Private Sub edcRootURL_Change()
    imChangesOccured = True
End Sub

Private Sub edcWebExports_Change()
    imChangesOccured = True
End Sub

Private Sub edcWebImports_Change()
    imChangesOccured = True
End Sub

Private Sub chkFtpIsOn_Click()
    imChangesOccured = True
End Sub

Private Sub edcFTPExportDir_Change()
    imChangesOccured = True
End Sub

Private Sub edcFTPImportDir_Change()
    imChangesOccured = True
End Sub

