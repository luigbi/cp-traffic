VERSION 5.00
Begin VB.Form ExportInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Info"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1545
      TabIndex        =   8
      Top             =   2970
      Width           =   900
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   435
      Left            =   345
      TabIndex        =   7
      Top             =   2985
      Width           =   900
   End
   Begin VB.Frame plcPages 
      Height          =   2520
      Left            =   240
      TabIndex        =   0
      Top             =   225
      Width           =   4155
      Begin VB.Frame plcLimitPages 
         Height          =   1200
         Left            =   405
         TabIndex        =   2
         Top             =   780
         Width           =   3480
         Begin VB.TextBox txtEndPage 
            Height          =   285
            Left            =   2220
            TabIndex        =   6
            Text            =   "0"
            Top             =   585
            Width           =   735
         End
         Begin VB.TextBox txtStartPage 
            Height          =   285
            Left            =   660
            TabIndex        =   5
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "To"
            Height          =   300
            Left            =   1650
            TabIndex        =   4
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label1 
            Caption         =   "From"
            Height          =   300
            Left            =   135
            TabIndex        =   3
            Top             =   615
            Width           =   645
         End
      End
      Begin VB.CheckBox clbPages 
         Caption         =   "Export all pages"
         Height          =   495
         Left            =   255
         TabIndex        =   1
         Top             =   210
         Value           =   1  'Checked
         Width           =   2175
      End
   End
End
Attribute VB_Name = "ExportInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function mValidate() As Boolean
'rules:  start can always be zero, but end cannot be zero if start not zero.
' end can be greater than what is in report.  end cannot be less than start.
'if start is greater than last page, only get last page.
    Dim blRet As Boolean
    Dim ilFirst As Integer
    Dim ilLast As Integer
    
    blRet = True
    If clbPages.Value <> vbChecked Then
        If IsNumeric(txtStartPage.Text) And IsNumeric(txtEndPage.Text) Then
            ilFirst = txtStartPage.Text
            ilLast = txtEndPage.Text
            If ilLast < ilFirst Then
                MsgBox "End page must be equal to or greater than starting page.", vbOKOnly, "Warning!"
                txtEndPage.SetFocus
                blRet = False
            End If
        Else
            blRet = False
        End If
    End If
    mValidate = blRet
End Function
Private Sub mInit()
    plcPages.Visible = True
    clbPages.Value = vbChecked
    txtStartPage.Text = "0"
    txtEndPage.Text = "0"
    txtStartPage.Enabled = False
    txtEndPage.Enabled = False
End Sub
Private Sub clbPages_Click()
    If clbPages.Value = vbChecked Then
        plcLimitPages.Enabled = False
        txtStartPage.Enabled = False
        txtEndPage.Enabled = False
    Else
        plcLimitPages.Enabled = True
        txtStartPage.Enabled = True
        txtEndPage.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    fCrViewerExport.bmContinue = False
    Unload ExportInfo
End Sub

Private Sub cmdOk_Click()
    fCrViewerExport.bmContinue = True
    If Not ogReport Is Nothing Then
        ogReport.PdfPageFirst = 0
        ogReport.PdfPageLast = 0
        If clbPages.Value <> vbChecked Then
            If mValidate() Then
                ogReport.PdfPageFirst = txtStartPage.Text
                ogReport.PdfPageLast = txtEndPage.Text
                Unload ExportInfo
            End If
        Else
            Unload ExportInfo
        End If
    End If
End Sub

Private Sub Form_Load()
    mInit
End Sub
