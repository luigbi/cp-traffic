VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmCDStartTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CD/Tape Start Time"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   Icon            =   "AffCDStartTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frcStatus 
      Caption         =   "Pledge Status"
      Enabled         =   0   'False
      Height          =   1425
      Left            =   435
      TabIndex        =   4
      Top             =   1110
      Width           =   3390
      Begin VB.OptionButton rbcStatus 
         Caption         =   "Air in Daypart"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Top             =   315
         Value           =   -1  'True
         Width           =   2145
      End
      Begin VB.OptionButton rbcStatus 
         Caption         =   "Delay Cmml/Prg"
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   6
         Top             =   675
         Width           =   2295
      End
      Begin VB.OptionButton rbcStatus 
         Caption         =   "Air Cmml Only"
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   5
         Top             =   1035
         Width           =   2400
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   30
      Top             =   630
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3525
      FormDesignWidth =   4365
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2340
      TabIndex        =   3
      Top             =   2775
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   780
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtCDStartTime 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblCDStartTime 
      Caption         =   "Please Enter the Start Time for the CD or Tape"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmCDStartTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmAffDP - Affiliate Daypart Information
'*
'*  Created August,2001 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Dim smRetTime As String



Private Sub cmdOk_Click()

    sgCDStartTime = txtCDStartTime.Text
    igCDStartTimeOK = gIsTime(sgCDStartTime)
    If Not igCDStartTimeOK Then
        txtCDStartTime.Text = ""
        gMsgBox "Please Enter a Valid Time"
        Unload frmCDStartTime
        Exit Sub
    End If
    'sgCDStartTimeOK = True
    igReturnPledgeStatus = 0
    If frcStatus.Enabled Then
        If rbcStatus(1).Value Then
            igReturnPledgeStatus = 1
        ElseIf rbcStatus(2).Value Then
            igReturnPledgeStatus = 2
        End If
    End If
    Unload frmCDStartTime

End Sub

Private Sub cmdCancel_Click()

    'If the user cancels out of the daypart screen reset the radio buttons
    'back to all false
    frmAgmnt!optTimeType(0).Value = False
    frmAgmnt!optTimeType(1).Value = False
    frmAgmnt!optTimeType(2).Value = False
    igCDStartTimeOK = False
    igReload = False
    igReturnPledgeStatus = 0

    Unload frmCDStartTime


End Sub

Private Sub Form_Load()

    Dim slTime As String
    Dim slRetTime As String
    
    
    If Hour(Trim$(tgDat(0).sFdSTime)) <> "" Then
        slTime = Hour(Trim$(tgDat(0).sFdSTime))
        slRetTime = gConvertMilitaryHourToRegTime(slTime)
    End If
    
    smRetTime = slRetTime
    txtCDStartTime.Text = slRetTime
    Screen.MousePointer = vbHourglass
    Me.Width = (Screen.Width) / 2.75
    Me.Height = (Screen.Height) / 3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Screen.MousePointer = vbDefault
    

End Sub

Private Sub txtCDStartTime_Change()
    If smRetTime <> Trim$(txtCDStartTime.Text) Then
        frcStatus.Enabled = True
    Else
        frcStatus.Enabled = False
    End If
End Sub
