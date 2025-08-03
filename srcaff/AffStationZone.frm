VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#2.5#0"; "Resize32.ocx"
Begin VB.Form frmStationZone 
   Caption         =   "Station Zone"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   Icon            =   "AffStationZone.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   4785
   Begin VB.OptionButton optAuto 
      Caption         =   "No"
      Height          =   255
      Index           =   1
      Left            =   3810
      TabIndex        =   2
      Top             =   825
      Width           =   780
   End
   Begin VB.OptionButton optAuto 
      Caption         =   "Yes"
      Height          =   255
      Index           =   0
      Left            =   2925
      TabIndex        =   1
      Top             =   825
      Width           =   780
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4170
      Top             =   1995
      _Version        =   131077
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      AutoCenterFormOnLoad=   -1  'True
      FormDesignHeight=   2190
      FormDesignWidth =   4785
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1695
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Continue Save"
      Height          =   375
      Left            =   945
      TabIndex        =   5
      Top             =   1695
      Width           =   1335
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   2685
      TabIndex        =   4
      Top             =   1110
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "Automatically Update Agreements"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   825
      Width           =   2625
   End
   Begin VB.Label labTimeZone 
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   390
      Width           =   4725
   End
   Begin VB.Label labTimeZone 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   105
      Width           =   4755
   End
   Begin VB.Label Label2 
      Caption         =   "Start Date of Zone Change"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   1155
      Width           =   2100
   End
End
Attribute VB_Name = "frmStationZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmStationZone - enters Station Zone information
'*
'*  Created January,1998 by Wade Bjerke
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit
Private imFieldChgd As Integer



Private Sub cmdCancel_Click()
    igTimeZoneStatus = 2
    Unload frmStationZone
End Sub

Private Sub cmdOk_Click()
    Dim sDate As String
    

    If optAuto(1).Value = True Then
        igTimeZoneStatus = 0
        Unload frmStationZone
        Exit Sub
    End If
    
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
    If DateValue(gAdjYear(sDate)) < DateValue(gAdjYear(Format$(gNow(), "m/d/yy"))) Then
        gMsgBox "Date Cannot be Prior To: " & Format$(gNow(), sgShowDateForm), vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If
    If Weekday(sDate, vbSunday) <> vbMonday Then
        gMsgBox "Date Must be a Monday", vbOKOnly
        txtDate.SetFocus
        Exit Sub
    End If

    igTimeZoneStatus = 1
    sgTimeZoneChangeDate = sDate
    Unload frmStationZone

End Sub

Private Sub Form_Initialize()
    Me.Width = (Screen.Width) / 2
    Me.Height = (Screen.Height) / 3
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    frmStationZone.Caption = "Station Zone - " & sgClientName
    'Me.Width = (Screen.Width) / 1.55
    'Me.Height = (Screen.Height) / 2.4
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2
    If Trim$(sgOrigTimeZone) = "" Then
        labTimeZone(0).Caption = "Changing Zone To"
    Else
        labTimeZone(0).Caption = sgOrigTimeZone & " Changing To"
    End If
    labTimeZone(1).Caption = sgNewTimeZone
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmStationZone = Nothing
End Sub







