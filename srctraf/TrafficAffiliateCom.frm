VERSION 5.00
Begin VB.Form TrafficAffiliateCom 
   BorderStyle     =   0  'None
   Caption         =   "TrafficAffiliateCom"
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
   Begin VB.TextBox edcAffiliateMsg 
      Height          =   555
      Left            =   780
      TabIndex        =   0
      Top             =   660
      Width           =   1575
   End
End
Attribute VB_Name = "TrafficAffiliateCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub edcAffiliateMsg_Change()
    On Error Resume Next
    If edcAffiliateMsg.Text = "" Then Exit Sub
    If lgShellAndWaitID <> 0 Then
        AppActivate lgShellAndWaitID
    Else
        If Traffic.WindowState = vbMinimized Then
            Traffic.WindowState = vbMaximized
        End If
    End If
    edcAffiliateMsg.Text = ""
End Sub
