VERSION 5.00
Begin VB.Form ProgressMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standby..."
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   ControlBox      =   0   'False
   Icon            =   "ProgressMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BTN_OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox edcMsg 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "ProgressMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub SetMessage(MsgType As Integer, Msg As String)

    Dim ilRet As Integer

    edcMsg.ForeColor = RGB(0, 0, 0)
    BTN_OK.Visible = False
    BTN_OK.Caption = "OK"
    
        
    'Show Black Text & Button
    If MsgType = 1 Then
        edcMsg.ForeColor = RGB(0, 0, 0)
        BTN_OK.Visible = True
    End If
    
    'Show Red Text & Button
    If MsgType = 2 Then
        edcMsg.ForeColor = RGB(255, 0, 0)
        BTN_OK.Visible = True
    End If
    
    edcMsg = Msg
End Sub

Private Sub BTN_OK_Click()
    Unload Me
End Sub


'Private Sub Form_Initialize()
'
'    gSetFonts Me
'    gCenterForm Me
'
'End Sub

Private Sub Form_Load()

    'gCenterForm ProgressMsg
    gCenterStdAlone ProgressMsg

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmProgressMst = Nothing
End Sub
