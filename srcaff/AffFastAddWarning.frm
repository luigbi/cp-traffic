VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmFastAddWarning 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9795
   Icon            =   "AffFastAddWarning.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   8640
      Top             =   3960
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4455
      FormDesignWidth =   9795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5175
      TabIndex        =   3
      Top             =   3855
      Width           =   1935
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   3855
      Width           =   1935
   End
   Begin VB.ListBox lbcFastAddWarning 
      Height          =   2400
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   9255
   End
   Begin VB.Label lblAdvise 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   9375
   End
   Begin VB.Label lblFastAddWarning 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9375
   End
End
Attribute VB_Name = "frmFastAddWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Sub cmdCancel_Click()
    igFastAddContinue = False
    gLogMsg "User chose to Cancel Process", "FastAddSummary.txt", False
    gLogMsg "User chose to Cancel Process", "FastAddVerbose.Txt", False
    Unload frmFastAddWarning
End Sub

Private Sub cmdContinue_Click()
    igFastAddContinue = True
    gLogMsg "User chose to Continue Process", "FastAddSummary.txt", False
    gLogMsg "User chose to Continue Process", "FastAddVerbose.Txt", False
    Unload frmFastAddWarning
End Sub

Private Sub Form_Initialize()
    'D.S. 5/22/18 moved from form load
    Me.Width = (Screen.Width) / 1.55
    Me.Height = (Screen.Height) / 2.4
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    'D.S. 5/22/18 added code below
    gSetFonts frmFastAddWarning
    gCenterForm frmFastAddWarning

End Sub

Private Sub Form_Load()
    
    'D.S. 5/22/18 added line below
    cmdCancel.Visible = bgFastAddCancelButton
    lblFastAddWarning.Caption = "The following problems were found prior to proccessing the station(s).  Each of the station(s) will be moved to the excluded list."
    frmFastAddWarning.Caption = "Affiliate Fast Add Warnings - " & sgClientName
    lblAdvise.Caption = "Note: To view or print this listing please go to: " & sgMsgDirectory & " and double-click on " & "FastAddSummary.txt"
    mProcessWarnings

End Sub

Public Function mProcessWarnings()

    Dim ilLoop As Integer
    Dim slWarning As String
    
    gLogMsg "", "FastAddSummary.txt", True
    gLogMsg "This is a listing of Stations followed by their associated error.", "FastAddSummary.txt", False
    gLogMsg "This is a listing of Stations followed by their associated error.", "FastAddVerbose.txt", False
    gLogMsg "", "FastAddSummary.txt", False
    
    For ilLoop = 0 To UBound(tgStaNameAndCode) - 1 Step 1
        slWarning = Trim$(tgStaNameAndCode(ilLoop).sStationName) & " - " & Trim$(tgStaNameAndCode(ilLoop).sInfo)
        gLogMsg slWarning, "FastAddSummary.txt", False
        gLogMsg slWarning, "FastAddVerbose.txt", False
        lbcFastAddWarning.AddItem slWarning
    Next ilLoop

End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmFastAddWarning = Nothing
End Sub
