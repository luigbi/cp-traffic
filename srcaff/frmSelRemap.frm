VERSION 5.00
Begin VB.Form frmSelRemap 
   Caption         =   "Affiliate Listing"
   ClientHeight    =   6840
   ClientLeft      =   2460
   ClientTop       =   2745
   ClientWidth     =   9885
   Icon            =   "frmSelRemap.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9885
   Begin VB.Frame Frame1 
      Caption         =   "Selection Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   315
      TabIndex        =   6
      Top             =   105
      Width           =   9225
      Begin VB.OptionButton rbcWhichToRemap 
         Caption         =   "Begin With All Affiliates Re-Mapped As NA-Off Air"
         Height          =   285
         Index           =   1
         Left            =   315
         TabIndex        =   8
         Top             =   615
         Width           =   6300
      End
      Begin VB.OptionButton rbcWhichToRemap 
         Caption         =   "Begin With All Affiliates Re-Mapped As Live"
         Height          =   285
         Index           =   0
         Left            =   315
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   5850
      End
   End
   Begin VB.ListBox lbcSelRemap 
      Height          =   4350
      Index           =   1
      Left            =   5115
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1650
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5175
      TabIndex        =   2
      Top             =   6270
      Width           =   1635
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   3165
      TabIndex        =   1
      Top             =   6270
      Width           =   1635
   End
   Begin VB.ListBox lbcSelRemap 
      Height          =   4350
      Index           =   0
      Left            =   315
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1665
      Width           =   4455
   End
   Begin VB.Label LblGetsRemapped 
      Caption         =   "Affiliates Below Will Be Re-Mapped As NA-Off Air"
      Height          =   435
      Index           =   1
      Left            =   5130
      TabIndex        =   5
      Top             =   1305
      Width           =   4395
   End
   Begin VB.Label LblGetsRemapped 
      Caption         =   "Affiliates Below Will Be Re-Mapped As  Live"
      Height          =   435
      Index           =   0
      Left            =   330
      TabIndex        =   4
      Top             =   1305
      Width           =   3555
   End
End
Attribute VB_Name = "frmSelRemap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'*  frmSelRemap - Avail time re-mapping of Selected Affiliates.
'*                Used as an enhancement to frmAVRemap.
'*
'*  Created December, 2001 by Doug Smith
'*
'*  Copyright Counterpoint Software, Inc. 2001
'******************************************************

Option Explicit

Private Sub cmdSelRemap_Click(Index As Integer)
End Sub

Private Sub cmdCancel_Click()
    igOkToRemap = False
    Unload frmSelRemap
    Set frmSelRemap = Nothing
End Sub

Private Sub cmdOk_Click()

    Dim ilIdx As Integer
    Dim llAttIdx As Long
        
    On Error GoTo ErrHand
    
    If lbcSelRemap(0).ListCount = 0 Then
        gMsgBox "No affiliates were selected to be re-mapped.", vbOKOnly
        Exit Sub
    End If
    
    'Set all items in the to be Re-mapped list box as selected
    For ilIdx = 0 To lbcSelRemap(0).ListCount - 1 Step 1
        For llAttIdx = 0 To UBound(tgAttInfo) - 1 Step 1
            If lbcSelRemap(0).ItemData(ilIdx) = tgAttInfo(llAttIdx).lAttCode Then
                tgAttInfo(llAttIdx).iSelected = True
                Exit For
            End If
        Next llAttIdx
    Next ilIdx
    
    'Set all items in the Not to be Re-mapped list box as Not selected
    For ilIdx = 0 To lbcSelRemap(1).ListCount - 1 Step 1
        For llAttIdx = 0 To UBound(tgAttInfo) - 1 Step 1
            If lbcSelRemap(1).ItemData(ilIdx) = tgAttInfo(llAttIdx).lAttCode Then
                tgAttInfo(llAttIdx).iSelected = False
                Exit For
            End If
        Next llAttIdx
    Next ilIdx
                    
    'Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbDefault
    Unload frmSelRemap
    Exit Sub

ErrHand:
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSelRemap-cmdOK-Click: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
End Sub


Private Sub Form_Load()

    frmSelRemap.Caption = "Affiliate Listing - " & sgClientName
    ScaleMode = 3       'set screen mode to pixels
    Me.Top = (Screen.Height - Me.Height) / 2.2
    Me.Left = (Screen.Width - Me.Width) / 2
    
End Sub
Private Sub Label1_Click()
End Sub

Private Sub lbcSelRemap_Click(Index As Integer)

    On Error GoTo ErrHand

    'Move item from the remapped side to the not remapped side
    If lbcSelRemap(0).ListIndex >= 0 Then
        lbcSelRemap(1).AddItem lbcSelRemap(0).Text
        lbcSelRemap(1).ItemData(lbcSelRemap(1).NewIndex) = lbcSelRemap(0).ItemData(lbcSelRemap(0).ListIndex)
        lbcSelRemap(0).RemoveItem (lbcSelRemap(0).ListIndex)
        lbcSelRemap(1).ListIndex = -1
        lbcSelRemap(0).ListIndex = -1
    End If
    
    'Move item from the not remapped side to the remapped side
    If lbcSelRemap(1).ListIndex >= 0 Then
        lbcSelRemap(0).AddItem lbcSelRemap(1).Text
        lbcSelRemap(0).ItemData(lbcSelRemap(0).NewIndex) = lbcSelRemap(1).ItemData(lbcSelRemap(1).ListIndex)
        lbcSelRemap(1).RemoveItem (lbcSelRemap(1).ListIndex)
        lbcSelRemap(0).ListIndex = -1
        lbcSelRemap(1).ListIndex = -1
    End If
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSelRemap-lbcSelRemap-Click: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    
End Sub

Private Sub rbcWhichToRemap_Click(Index As Integer)

    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    If rbcWhichToRemap(0).Value Then
        ilRet = gPopSelRemap(frmSelRemap, frmAvRemap!lbcAvail(1).ListCount, 0)
    Else
        ilRet = gPopSelRemap(frmSelRemap, frmAvRemap!lbcAvail(1).ListCount, 1)
    End If
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in frmSelRemap-rbcWhichToRemap-Click: "
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number, vbCritical
    End If
    

End Sub
