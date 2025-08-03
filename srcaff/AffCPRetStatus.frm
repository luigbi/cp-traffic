VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmCPRetStatus 
   Caption         =   "CP Return Status"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   ControlBox      =   0   'False
   Icon            =   "AffCPRetStatus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCPRetStatus 
      Caption         =   "Un-post Date/Time of all Affiliate Spots"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   300
      TabIndex        =   6
      Top             =   2715
      Width           =   4020
   End
   Begin VB.OptionButton rbcPost 
      Caption         =   "Post Date/Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1395
      TabIndex        =   1
      Top             =   225
      Width           =   1680
   End
   Begin VB.OptionButton rbcPost 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   330
      TabIndex        =   0
      Top             =   225
      Value           =   -1  'True
      Width           =   870
   End
   Begin VB.CommandButton cmdCPRetStatus 
      Caption         =   "View Date/Time Only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   300
      TabIndex        =   2
      Top             =   675
      Width           =   4020
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4320
      Top             =   3630
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3855
      FormDesignWidth =   4755
   End
   Begin VB.CommandButton cmdCPRetStatus 
      Caption         =   "Return to Date selection screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   300
      TabIndex        =   7
      Top             =   3225
      Width           =   4020
   End
   Begin VB.CommandButton cmdCPRetStatus 
      Caption         =   "Post all Spots as ""None Aired"" w/o viewing the Spots"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   300
      TabIndex        =   5
      Top             =   2205
      Width           =   4020
   End
   Begin VB.CommandButton cmdCPRetStatus 
      Caption         =   "Post Spots Date/Time"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   300
      TabIndex        =   3
      Top             =   1185
      Width           =   4020
   End
   Begin VB.CommandButton cmdCPRetStatus 
      Caption         =   "Post all Spots as ""Pledged"" w/o viewing the Spots"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   315
      TabIndex        =   4
      Top             =   1695
      Width           =   4020
   End
End
Attribute VB_Name = "frmCPRetStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCPRetStatus_Click(Index As Integer)

    If cmdCPRetStatus(0).Value = True Then
        sgCPRetStatus = "Yes"
    ElseIf cmdCPRetStatus(1).Value = True Then
        sgCPRetStatus = "No"
    ElseIf cmdCPRetStatus(2).Value = True Then
        sgCPRetStatus = "None"
    ElseIf cmdCPRetStatus(3).Value = True Then
        sgCPRetStatus = "Cancel"
    ElseIf cmdCPRetStatus(4).Value = True Then
        sgCPRetStatus = "View"
    ElseIf cmdCPRetStatus(5).Value = True Then
        sgCPRetStatus = "Unpost"
    Else
        sgCPRetStatus = "UnKnown"
    End If
    
    Unload frmCPRetStatus
    
End Sub

Private Sub Form_Load()

    frmCPRetStatus.Caption = sgCPRetStatus  '"CP Return Status - " & sgClientName
    sgCPRetStatus = ""
    Me.Width = (Screen.Width) / 3
    Me.Height = (Screen.Height) / 3 '4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    gCenterForm frmCPRetStatus

End Sub

Private Sub rbcPost_Click(Index As Integer)
    If rbcPost(Index).Value Then
        If Index = 0 Then
            cmdCPRetStatus(0).Enabled = False
            cmdCPRetStatus(1).Enabled = False
            cmdCPRetStatus(2).Enabled = False
            cmdCPRetStatus(5).Enabled = False
        Else
            cmdCPRetStatus(0).Enabled = True
            cmdCPRetStatus(1).Enabled = True
            cmdCPRetStatus(2).Enabled = True
            cmdCPRetStatus(5).Enabled = True
        End If
    End If
End Sub
