VERSION 5.00
Begin VB.Form AffiliateTrafficCom 
   BorderStyle     =   0  'None
   Caption         =   "AffiliateTrafficCom"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox edcTrafficMsg 
      Height          =   555
      Left            =   780
      TabIndex        =   0
      Top             =   660
      Width           =   1575
   End
End
Attribute VB_Name = "AffiliateTrafficCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub edcTrafficMsg_Change()
    On Error Resume Next
    If edcTrafficMsg.Text = "" Then Exit Sub
    If lgShellAndWaitID <> 0 Then
        AppActivate lgShellAndWaitID
    Else
        If frmMain.WindowState = vbMinimized Then
            frmMain.WindowState = vbMaximized
        End If
    End If
    edcTrafficMsg.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set AffiliateTrafficCom = Nothing
End Sub
