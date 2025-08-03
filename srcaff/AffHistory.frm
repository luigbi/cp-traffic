VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmHistory 
   Caption         =   "Station History"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   Icon            =   "AffHistory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   4785
   Begin VB.OptionButton OptHistory 
      Caption         =   "No"
      Height          =   195
      Index           =   1
      Left            =   4020
      TabIndex        =   8
      Top             =   885
      Width           =   690
   End
   Begin VB.OptionButton OptHistory 
      Caption         =   "Yes"
      Height          =   195
      Index           =   0
      Left            =   3255
      TabIndex        =   6
      Top             =   885
      Value           =   -1  'True
      Width           =   690
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3180
      Top             =   1260
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   2580
      FormDesignWidth =   4785
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2460
      TabIndex        =   4
      Top             =   1980
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Continue Save"
      Height          =   375
      Left            =   885
      TabIndex        =   3
      Top             =   1980
      Width           =   1335
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1365
      TabIndex        =   1
      Top             =   1305
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "Retain Change as part of Station History:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   885
      Width           =   2925
   End
   Begin VB.Label labCallLetters 
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   375
      Width           =   4725
   End
   Begin VB.Label labCallLetters 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   105
      Width           =   4755
   End
   Begin VB.Label Label2 
      Caption         =   "Last Air Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1335
      Width           =   1185
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmHistory - enters History Station information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer



Private Sub cmdCancel_Click()
    igHistoryStatus = 2
    Unload frmHistory
End Sub

Private Sub cmdOk_Click()
    Dim sDate As String
    
    If OptHistory(0).Value Then
        sDate = Trim$(txtDate.Text)
        If sDate = "" Then
            Beep
            gMsgBox "Dates must be specified.", vbOKOnly
            txtDate.SetFocus
            Exit Sub
        End If
        If gIsDate(sDate) = False Then
            Beep
            gMsgBox "Please enter a valid date (m/d/yy).", vbCritical
            txtDate.SetFocus
            Exit Sub
        End If
        igHistoryStatus = 1
        sgLastAirDate = sDate
    Else
        igHistoryStatus = 0
    End If
    Unload frmHistory

End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / 2
    Me.Height = (Screen.Height) / 3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    frmHistory.Caption = "Station History - " & sgClientName
    'Me.Width = (Screen.Width) / 1.55
    'Me.Height = (Screen.Height) / 2.4
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    labCallLetters(0).Caption = sgOrigCallLetters & " changed to"
    labCallLetters(1).Caption = sgNewCallLetters
    Screen.MousePointer = vbDefault
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmHistory = Nothing
End Sub







