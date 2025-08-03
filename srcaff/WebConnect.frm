VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form WebConnect 
   Caption         =   "Counterpoint Documents"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15735
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   15735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDocuments 
      Caption         =   "Documents"
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   75
      Width           =   1290
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      CausesValidation=   0   'False
      Height          =   8910
      Left            =   135
      TabIndex        =   0
      Top             =   465
      Width           =   15555
      ExtentX         =   27437
      ExtentY         =   15716
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "WebConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bmHomePage As Boolean

Private Sub cmdDocuments_Click()
 WebBrowser.Navigate "www.counterpoint.net/clients/rsvpPost.php"
End Sub

Private Sub Form_Load()
    mInit
End Sub

Private Sub mInit()
    Dim slUrl As String
    
    If bmHomePage Then
        slUrl = "www.counterpoint.net"
        Me.Caption = "Counterpoint"
        cmdDocuments.Visible = False
    Else
        slUrl = "www.counterpoint.net/clients/rsvpPost.php"
        Me.Caption = "Counterpoint Documents"
        cmdDocuments.Enabled = False
   End If
    With WebBrowser
        .Navigate slUrl
    End With
    bmHomePage = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set WebConnect = Nothing
End Sub

Private Sub WebBrowser_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If InStr(1, URL, "rsvpPost") > 0 Then
        With WebBrowser
            .Document.Forms.Item(, 0).elements("CsiNo").Value = "Cs!3x42S!c"
            .Document.Forms.Item(, 0).elements("CsiCo").Value = mAdjustName(sgClientName)
            .Document.Forms.Item(, 0).submit
        End With
            'no need for button since on this page.
            cmdDocuments.Enabled = False
    'allow 'back to documents' if going to a document
    ElseIf Not bmHomePage And InStr(1, URL, "clientDocuments") > 0 Then
        cmdDocuments.Enabled = True
    Else
        'looking at csi home pages, or on our documentation.php
        cmdDocuments.Enabled = False
    End If
End Sub

Private Function mAdjustName(slClientName As String) As String
    Dim ilPos As Integer
    Dim slNewName As String
    
    slNewName = Trim$(slClientName)
    ilPos = InStr(1, slClientName, "\")
    If ilPos > 0 Then
        slNewName = Mid(slClientName, 1, ilPos - 1)
    End If
    ilPos = InStr(1, slNewName, "-")
    If ilPos > 0 Then
        slNewName = Mid(slNewName, 1, ilPos - 1)
    End If
    mAdjustName = slNewName
End Function

