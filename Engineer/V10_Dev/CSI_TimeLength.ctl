VERSION 5.00
Begin VB.UserControl CSI_TimeLength 
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   ScaleHeight     =   5430
   ScaleWidth      =   6795
   ToolboxBitmap   =   "CSI_TimeLength.ctx":0000
   Begin VB.PictureBox pc_Frame 
      Height          =   495
      Left            =   555
      ScaleHeight     =   435
      ScaleWidth      =   1245
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1005
      Width           =   1305
      Begin VB.TextBox ec_Decimal_1 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   810
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "."
         Top             =   0
         Width           =   75
      End
      Begin VB.TextBox ec_Colon_2 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   510
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   ":"
         Top             =   0
         Width           =   75
      End
      Begin VB.TextBox ec_Colon_1 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   225
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   ":"
         Top             =   0
         Width           =   75
      End
      Begin VB.TextBox ec_Tenths 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   885
         TabIndex        =   4
         Text            =   "1"
         Top             =   0
         Width           =   135
      End
      Begin VB.TextBox ec_Seconds 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   585
         TabIndex        =   3
         Text            =   "01"
         Top             =   0
         Width           =   225
      End
      Begin VB.TextBox ec_Minutes 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   285
         TabIndex        =   2
         Text            =   "01"
         Top             =   0
         Width           =   225
      End
      Begin VB.TextBox ec_Hours 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   0
         TabIndex        =   1
         Text            =   "01"
         Top             =   0
         Width           =   225
      End
   End
End
Attribute VB_Name = "CSI_TimeLength"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Public MaxLength As Integer

Private smText As String
Private imUseHours As Boolean
Private imUseTenths As Boolean
Private bmMainControlHadFocus As Boolean
Private bmLostFocusLeft As Boolean
Private bmLostFocusRight As Boolean
Private bmControlIsReady As Boolean
Private bmOkToTab As Boolean
Private imLastPos As Integer
Private m_Enabled As Boolean
Private bmIgnoreResize As Boolean
Private smStartText As String
Private bmSendChangeEvent As Boolean
Private bmHourHasFocus As Boolean
Private bmMinuteHasFocus As Boolean
Private bmSecondHasFocus As Boolean
Private bmTenthHasFocus As Boolean
Private imHourGotFocus As Boolean
Private imMinuteGotFocus As Boolean
Private imSecondGotFocus As Boolean
Private imTenthGotFocus As Boolean
Private bmIgnoreNextUpKey As Boolean
Private bmUserInputDetected As Boolean ' Is false until the user types something.

'If only a single digit is entered, provide 0 in front of it with a change event.
'Update the field during the lost focus but do not send a change event.
'When the control loses focus, update the HH:MM:SS so it is all zero's. This will
'help so it still works.

' Format when they are getting the text.
' Provide property to allow 60 seconds.
' I can change all fields from HH:MM:SS when any of them change.

Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

Enum CSI_TimeLength_BorderStyle
    No_Border = 0
    Single_Border = 1
End Enum

Event OnChange()
Event OnLostFocus()

Private Sub ec_Colon_1_GotFocus()
    If imUseHours Then
        ec_Hours.SetFocus
    Else
        ec_Minutes.SetFocus
    End If
End Sub

Private Sub ec_Colon_2_GotFocus()
    ec_Minutes.SetFocus
End Sub

Private Sub ec_Decimal_1_GotFocus()
    ec_Seconds.SetFocus
End Sub

'*****************************************************************
'
'*****************************************************************
Private Sub ec_Hours_Change()
    If Not bmControlIsReady Then Exit Sub
    If Not bmOkToTab Then Exit Sub
    
    If Len(ec_Hours.text) >= 2 Then
        If ec_Hours.text > 24 Then
            bmControlIsReady = False
            ec_Hours.text = "24"
            ec_Hours.SelStart = 0
            ec_Hours.SelLength = 2
            Call SetTextValue(True, ec_Hours.hwnd)
            bmControlIsReady = True
            Beep
            Exit Sub
        End If
        SendKeys "{TAB}", True
    End If

    Call SetTextValue(True, ec_Hours.hwnd)
End Sub

Private Sub ec_Hours_KeyPress(KeyAscii As Integer)
    Dim slOneChar As String

    If KeyAscii = vbKeyBack Then
        Exit Sub    ' Allow back space.
    End If
    slOneChar = Chr(KeyAscii)
'    If slOneChar = ":" Then
'        KeyAscii = 0    ' Throw this key away since there is already a colon there.
'        If Len(ec_Hours.text) < 2 Then
'            ec_Hours.text = "0" + ec_Hours.text
'        End If
'        Call SetTextValue(True, ec_Hours.hwnd)
'        ' ec_Minutes.SetFocus
'        Exit Sub
'    End If
    If InStr(1, "0123456789", slOneChar) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ec_Hours_LostFocus()
    'Exit Sub
    imHourGotFocus = False
    bmControlIsReady = False
    If ec_Hours.text = "HH" Or ec_Hours.text = "H" Then
        ec_Hours.text = "00"
        'RaiseEvent OnChange
        Call SetTextValue(True, 0)
    End If
    If Len(ec_Hours.text) < 2 Then
        ec_Hours.text = "0" + ec_Hours.text
        'RaiseEvent OnChange
        'Call SetTextValue(True, 0)
    End If
    bmControlIsReady = True
End Sub

Private Sub ec_Hours_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyRight
            ' If the user only has the first digit selected then highlight the next digit.
            If ec_Hours.SelStart = 0 And ec_Hours.SelLength = 1 Then
                ec_Hours.SetFocus
                ec_Hours.SelStart = 1
                ec_Hours.SelLength = 1
                bmIgnoreNextUpKey = True
                KeyCode = -1
            End If
        Case vbKeyLeft
            If ec_Hours.SelStart = 1 And ec_Hours.SelLength = 1 Then
                ec_Hours.SetFocus
                ec_Hours.SelStart = 0
                ec_Hours.SelLength = 1
                bmIgnoreNextUpKey = True
                KeyCode = -1
            End If
    End Select
End Sub

Private Sub ec_Hours_KeyUp(KeyCode As Integer, Shift As Integer)
    If bmIgnoreNextUpKey Then
        bmIgnoreNextUpKey = False
        Exit Sub
    End If
    Select Case KeyCode
        Case vbKeyUp
            bmOkToTab = False
            If ec_Hours.text = "HH" Or ec_Hours.text = "H" Then
                ec_Hours.text = "00"
            End If
            If ec_Hours.text < 24 Then
                ec_Hours.text = ec_Hours.text + 1
            Else
                ec_Hours.text = 0
            End If
            If Len(ec_Hours.text) < 2 Then
                'ec_Hours.Text = "0" + ec_Hours.Text
            End If
            ec_Hours.SelStart = 0
            ec_Hours.SelLength = 2
            Call SetTextValue(True, 0)
            bmOkToTab = True
        Case vbKeyDown
            bmOkToTab = False
            If ec_Hours.text = "HH" Or ec_Hours.text = "H" Then
                ec_Hours.text = "00"
            End If
            If ec_Hours.text > 0 Then
                ec_Hours.text = ec_Hours.text - 1
            Else
                ec_Hours.text = 24
            End If
            If Len(ec_Hours.text) < 2 Then
                'ec_Hours.Text = "0" + ec_Hours.Text
            End If
            ec_Hours.SelStart = 0
            ec_Hours.SelLength = 2
            Call SetTextValue(True, 0)
            bmOkToTab = True
        Case vbKeyRight
            ' Tab to the Minutes field.
            SendKeys "{TAB}", True
        Case vbKeyLeft
            ' Tab to the Minutes or Tenths field.
            SendKeys "{TAB}", True
            SendKeys "{TAB}", True
            If imUseTenths Then
                SendKeys "{TAB}", True
            End If
    End Select
End Sub

Private Sub ec_Hours_GotFocus()
    If ec_Hours.text = "HH" Or ec_Hours.text = "H" Then
        bmControlIsReady = False
        bmControlIsReady = True
    End If
    ec_Hours.SelStart = 0
    ec_Hours.SelLength = 2
    imHourGotFocus = True
End Sub

Private Sub ec_Hours_Click()
    If imUseHours Then
        'ec_Hours.SelStart = 0
        'ec_Hours.SelLength = 2
    End If
End Sub

Private Sub ec_Hours_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imHourGotFocus = False
End Sub

Private Sub ec_Hours_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilCurStartPos As Integer

    ilCurStartPos = ec_Hours.SelStart
    If imHourGotFocus Then
        ec_Hours.SelStart = 0
        ec_Hours.SelLength = 2
        imHourGotFocus = False
        Exit Sub
    End If
    If ec_Hours.text = "HH" Then
        ec_Hours.SelStart = 0
        ec_Hours.SelLength = 2
    Else
        If ilCurStartPos > 1 Then
            ilCurStartPos = 1
        End If
        ec_Hours.SelStart = ilCurStartPos
        ec_Hours.SelLength = 1
    End If
End Sub

'*****************************************************************
'
'*****************************************************************
Private Sub ec_Minutes_Change()
    If Not bmControlIsReady Then Exit Sub
    If Not bmOkToTab Then Exit Sub
    If Len(ec_Minutes.text) >= 2 Then
        If ec_Minutes.text > 59 Then
            bmControlIsReady = False
            ec_Minutes.text = "59"
            ec_Minutes.SelStart = 0
            ec_Minutes.SelLength = 2
            Call SetTextValue(True, ec_Minutes.hwnd)
            bmControlIsReady = True
            Beep
            Exit Sub
        End If
        SendKeys "{TAB}", True
    End If
    bmControlIsReady = False
    Call SetTextValue(True, ec_Minutes.hwnd)
    bmControlIsReady = True
End Sub

Private Sub ec_Minutes_Click()
    'ec_Minutes.SelStart = 0
    'ec_Minutes.SelLength = 2
End Sub

Private Sub ec_Minutes_KeyPress(KeyAscii As Integer)
    Dim slOneChar As String

    If KeyAscii = vbKeyBack Then
        Exit Sub    ' Allow back space.
    End If
    slOneChar = Chr(KeyAscii)
'    If slOneChar = ":" Then
'        KeyAscii = 0    ' Throw this key away since there is already a colon there.
'        If Len(ec_Minutes.text) < 2 Then
'            ec_Minutes.text = "0" + ec_Minutes.text
'        End If
'        Call SetTextValue(True, ec_Minutes.hwnd)
'        'ec_Seconds.SetFocus
'        Exit Sub
'    End If
    If InStr(1, "0123456789", slOneChar) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ec_Minutes_LostFocus()
    'Exit Sub
    
    bmControlIsReady = False
    If ec_Minutes.text = "MM" Or ec_Minutes.text = "M" Then
        ec_Minutes.text = "00"
        Call SetTextValue(True, 0)
    End If
    If Len(ec_Minutes.text) < 2 Then
        ec_Minutes.text = "0" + ec_Minutes.text
        'Call SetTextValue(True, 0)
    End If
    bmControlIsReady = True
End Sub

Private Sub ec_Minutes_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyRight
            ' If the user only has the first digit selected then highlight the next digit.
            If ec_Minutes.SelStart = 0 And ec_Minutes.SelLength = 1 Then
                ec_Minutes.SetFocus
                ec_Minutes.SelStart = 1
                ec_Minutes.SelLength = 1
                bmIgnoreNextUpKey = True
                KeyCode = -1
            End If
        Case vbKeyLeft
            If ec_Minutes.SelStart = 1 And ec_Minutes.SelLength = 1 Then
                ec_Minutes.SetFocus
                ec_Minutes.SelStart = 0
                ec_Minutes.SelLength = 1
                bmIgnoreNextUpKey = True
                KeyCode = -1
            End If
    End Select
End Sub

Private Sub ec_Minutes_KeyUp(KeyCode As Integer, Shift As Integer)
    If bmIgnoreNextUpKey Then
        bmIgnoreNextUpKey = False
        Exit Sub
    End If
    Select Case KeyCode
        Case vbKeyUp
            bmOkToTab = False
            If ec_Minutes.text = "MM" Or ec_Minutes.text = "M" Then
                ec_Minutes.text = "00"
            End If
            If ec_Minutes.text < 59 Then
                ec_Minutes.text = ec_Minutes.text + 1
            Else
                ec_Minutes.text = 0
            End If
            If Len(ec_Minutes.text) < 2 Then
                'ec_Minutes.Text = "0" + ec_Minutes.Text
            End If
            ec_Minutes.SelStart = 0
            ec_Minutes.SelLength = 2
            Call SetTextValue(True, 0)
            bmOkToTab = True
        Case vbKeyDown
            bmOkToTab = False
            If ec_Minutes.text = "MM" Or ec_Minutes.text = "M" Then
                ec_Minutes.text = "00"
            End If
            If ec_Minutes.text > 0 Then
                ec_Minutes.text = ec_Minutes.text - 1
            Else
                ec_Minutes.text = 59
            End If
            If Len(ec_Minutes.text) < 2 Then
                'ec_Minutes.Text = "0" + ec_Minutes.Text
            End If
            ec_Minutes.SelStart = 0
            ec_Minutes.SelLength = 2
            Call SetTextValue(True, 0)
            bmOkToTab = True
        Case vbKeyRight
            ' Tab to the Seconds field.
            SendKeys "{TAB}", True
        Case vbKeyLeft
            ' Tab to the Hours field.
            If imUseHours Then
                SendKeys "+" + "{TAB}", True
            Else
                If imUseTenths Then
                    SendKeys "{TAB}", True
                    SendKeys "{TAB}", True
                Else
                    SendKeys "{TAB}", True
                End If
            End If
    End Select
End Sub

Private Sub ec_Minutes_GotFocus()
    If ec_Minutes.text = "MM" Or ec_Minutes.text = "M" Then
        bmControlIsReady = False
        'ec_Minutes.Text = "01"
        bmControlIsReady = True
    End If
    ec_Minutes.SelStart = 0
    ec_Minutes.SelLength = 2
    imMinuteGotFocus = True
End Sub

Private Sub ec_Minutes_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imMinuteGotFocus = False
End Sub

Private Sub ec_Minutes_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilCurStartPos As Integer
    
    ilCurStartPos = ec_Minutes.SelStart
    If imMinuteGotFocus Then
        ec_Minutes.SelStart = 0
        ec_Minutes.SelLength = 2
        imMinuteGotFocus = False
        Exit Sub
    End If
    If ec_Minutes.text = "MM" Then
        ec_Minutes.SelStart = 0
        ec_Minutes.SelLength = 2
    Else
        If ilCurStartPos > 1 Then
            ilCurStartPos = 1
        End If
        ec_Minutes.SelStart = ilCurStartPos
        ec_Minutes.SelLength = 1
    End If
End Sub

'*****************************************************************
'
'*****************************************************************
Private Sub ec_Seconds_Change()
    If Not bmControlIsReady Then Exit Sub
    If Not bmOkToTab Then Exit Sub
    If Len(ec_Seconds.text) >= 2 Then
        If ec_Seconds.text > 59 Then
            bmControlIsReady = False
            ec_Seconds.text = "59"
            ec_Seconds.SelStart = 0
            ec_Seconds.SelLength = 2
            Call SetTextValue(True, ec_Seconds.hwnd)
            bmControlIsReady = True
            Beep
            Exit Sub
        End If
        If CSI_UseTenths Then
            SendKeys "{TAB}", True
        End If
    End If
    Call SetTextValue(True, ec_Seconds.hwnd)
End Sub

Private Sub ec_Seconds_Click()
    'ec_Seconds.SelStart = 0
    'ec_Seconds.SelLength = 2
End Sub

Private Sub ec_Seconds_KeyPress(KeyAscii As Integer)
    Dim slOneChar As String

    If KeyAscii = vbKeyBack Then
        Exit Sub    ' Allow back space.
    End If
    slOneChar = Chr(KeyAscii)
    If slOneChar = "." And imUseTenths Then
        KeyAscii = 0    ' Throw this key away since there is already a colon there.
        If Len(ec_Seconds.text) < 2 Then
            ec_Seconds.text = "0" + ec_Seconds.text
        End If
        Call SetTextValue(True, ec_Seconds.hwnd)
        ec_Tenths.SetFocus
        Exit Sub
    End If
    If InStr(1, "0123456789", slOneChar) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ec_Seconds_LostFocus()
    'Exit Sub
    
    bmControlIsReady = False
    If ec_Seconds.text = "SS" Or ec_Seconds.text = "S" Then
        ec_Seconds.text = "00"
        Call SetTextValue(True, 0)
    End If
    If Len(ec_Seconds.text) < 2 Then
        ec_Seconds.text = "0" + ec_Seconds.text
        'Call SetTextValue(True, 0)
    End If
    bmControlIsReady = True
End Sub

Private Sub ec_Seconds_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyRight
            ' If the user only has the first digit selected then highlight the next digit.
            If ec_Seconds.SelStart = 0 And ec_Seconds.SelLength = 1 Then
                ec_Seconds.SetFocus
                ec_Seconds.SelStart = 1
                ec_Seconds.SelLength = 1
                bmIgnoreNextUpKey = True
                KeyCode = -1
            End If
        Case vbKeyLeft
            If ec_Seconds.SelStart = 1 And ec_Seconds.SelLength = 1 Then
                ec_Seconds.SetFocus
                ec_Seconds.SelStart = 0
                ec_Seconds.SelLength = 1
                bmIgnoreNextUpKey = True
                KeyCode = -1
            End If
    End Select
End Sub

Private Sub ec_Seconds_KeyUp(KeyCode As Integer, Shift As Integer)
    If bmIgnoreNextUpKey Then
        bmIgnoreNextUpKey = False
        Exit Sub
    End If
    Select Case KeyCode
        Case vbKeyUp
            bmOkToTab = False
            If ec_Seconds.text = "SS" Or ec_Seconds.text = "S" Then
                ec_Seconds.text = "00"
            End If
            If ec_Seconds.text < 59 Then
                ec_Seconds.text = ec_Seconds.text + 1
            Else
                ec_Seconds.text = 0
            End If
            If Len(ec_Seconds.text) < 2 Then
                'ec_Seconds.Text = "0" + ec_Seconds.Text
            End If
            ec_Seconds.SelStart = 0
            ec_Seconds.SelLength = 2
            Call SetTextValue(True, 0)
            bmOkToTab = True
        Case vbKeyDown
            bmOkToTab = False
            If ec_Seconds.text = "SS" Or ec_Seconds.text = "S" Then
                ec_Seconds.text = "00"
            End If
            If ec_Seconds.text > 0 Then
                ec_Seconds.text = ec_Seconds.text - 1
            Else
                ec_Seconds.text = 59
            End If
            If Len(ec_Seconds.text) < 2 Then
                'ec_Seconds.Text = "0" + ec_Seconds.Text
            End If
            ec_Seconds.SelStart = 0
            ec_Seconds.SelLength = 2
            Call SetTextValue(True, 0)
            bmOkToTab = True
        Case vbKeyRight
            ' Tab to the Tenths or Hours field.
            If imUseTenths Then
                SendKeys "{TAB}", True
            Else
                ' Tab to hours or minutes.
                SendKeys "+" + "{TAB}", True
                If imUseHours Then
                    SendKeys "+" + "{TAB}", True
                End If
            End If
        Case vbKeyLeft
            ' Tab to the Minutes field.
            SendKeys "+" + "{TAB}", True
    End Select
End Sub

Private Sub ec_Seconds_GotFocus()
    If ec_Seconds.text = "SS" Or ec_Seconds.text = "S" Then
        bmControlIsReady = False
        'ec_Seconds.Text = "01"
        bmControlIsReady = True
    End If
    ec_Seconds.SelStart = 0
    ec_Seconds.SelLength = 2
    imSecondGotFocus = True
End Sub

Private Sub ec_Seconds_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imSecondGotFocus = False
End Sub

Private Sub ec_Seconds_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ilCurStartPos As Integer
    
    ilCurStartPos = ec_Seconds.SelStart
    If imSecondGotFocus Then
        ec_Seconds.SelStart = 0
        ec_Seconds.SelLength = 2
        imSecondGotFocus = False
        Exit Sub
    End If
    If ec_Seconds.text = "SS" Then
        ec_Seconds.SelStart = 0
        ec_Seconds.SelLength = 2
    Else
        If ilCurStartPos > 1 Then
            ilCurStartPos = 1
        End If
        ec_Seconds.SelStart = ilCurStartPos
        ec_Seconds.SelLength = 1
    End If
End Sub

'*****************************************************************
'
'*****************************************************************
Private Sub ec_Tenths_Change()
    If Not bmControlIsReady Then Exit Sub
    If Not bmOkToTab Then Exit Sub
    If Len(ec_Tenths.text) >= 2 Then
        ec_Tenths.text = Left(ec_Tenths.text, 1)
    End If
    If ec_Tenths.text > "9" Then
        bmControlIsReady = False
        ec_Tenths.text = "9"
        ec_Tenths.SelStart = 0
        ec_Tenths.SelLength = 1
        Call SetTextValue(True, ec_Tenths.hwnd)
        bmControlIsReady = True
        Beep
        Exit Sub
    End If
    ec_Tenths.SelStart = 0
    ec_Tenths.SelLength = 1
    'SendKeys "{TAB}", True
    Call SetTextValue(True, ec_Tenths.hwnd)
End Sub

Private Sub ec_Tenths_Click()
    ec_Tenths.SelStart = 0
    ec_Tenths.SelLength = 1
End Sub

Private Sub ec_Tenths_KeyPress(KeyAscii As Integer)
    Dim slOneChar As String

    If KeyAscii = vbKeyBack Then
        Exit Sub    ' Allow back space.
    End If
    slOneChar = Chr(KeyAscii)
    If InStr(1, "0123456789", slOneChar) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ec_Tenths_LostFocus()
    'Exit Sub
    
    bmControlIsReady = False
    If ec_Tenths.text = "T" Then
        ec_Tenths.text = "0"
        Call SetTextValue(True, 0)
    End If
    bmControlIsReady = True
End Sub

Private Sub ec_Tenths_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            bmOkToTab = False
            If ec_Tenths.text = "T" Then
                ec_Tenths.text = "0"
            End If
            If ec_Tenths.text < 9 Then
                ec_Tenths.text = ec_Tenths.text + 1
            Else
                ec_Tenths.text = 0
            End If
            ec_Tenths.SelStart = 0
            ec_Tenths.SelLength = 1
            Call SetTextValue(True, 0)
            bmOkToTab = True
        Case vbKeyDown
            bmOkToTab = False
            If ec_Tenths.text = "T" Then
                ec_Tenths.text = "0"
            End If
            If ec_Tenths.text > 0 Then
                ec_Tenths.text = ec_Tenths.text - 1
            Else
                ec_Tenths.text = 9
            End If
            ec_Tenths.SelStart = 0
            ec_Tenths.SelLength = 1
            Call SetTextValue(True, 0)
            bmOkToTab = True
        Case vbKeyRight
            ' Tab back to the hour or minute field.
            SendKeys "+" + "{TAB}", True
            SendKeys "+" + "{TAB}", True
            If imUseHours Then
                SendKeys "+" + "{TAB}", True
            End If

            ' Tab to the next field on the form.
            'SendKeys "{TAB}", True
        Case vbKeyLeft
            ' Tab to the Seconds field.
            SendKeys "+" + "{TAB}", True
    End Select
End Sub

Private Sub ec_Tenths_GotFocus()
    If ec_Tenths.text = "T" Then
        bmControlIsReady = False
        'ec_Tenths.Text = "0"
        bmControlIsReady = True
    End If
    ec_Tenths.SelStart = 0
    ec_Tenths.SelLength = 1
    imTenthGotFocus = True
End Sub

Private Sub ec_Tenths_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imTenthGotFocus = False
End Sub

Private Sub UserControl_EnterFocus()
    Dim blChangeEventNeeded As Boolean
    
    Exit Sub
    
    blChangeEventNeeded = False
    bmSendChangeEvent = False
    bmControlIsReady = False
    If ec_Hours.text = "HH" Or ec_Hours.text = "H" Then
        blChangeEventNeeded = True
        ec_Hours.text = "00"
    End If
    If Len(ec_Hours.text) < 2 Then
        'blChangeEventNeeded = True
        'ec_Hours.Text = "0" + ec_Hours.Text
    End If
    
    If ec_Minutes.text = "MM" Or ec_Minutes.text = "M" Then
        blChangeEventNeeded = True
        ec_Minutes.text = "00"
    End If
    If Len(ec_Minutes.text) < 2 Then
        'blChangeEventNeeded = True
        'ec_Minutes.Text = "0" + ec_Minutes.Text
    End If
    
    If ec_Seconds.text = "SS" Or ec_Seconds.text = "S" Then
        blChangeEventNeeded = True
        ec_Seconds.text = "00"
    End If
    If Len(ec_Seconds.text) < 2 Then
        'blChangeEventNeeded = True
        'ec_Seconds.Text = "0" + ec_Seconds.Text
    End If
    
    If ec_Tenths.text = "T" Then
        blChangeEventNeeded = True
        ec_Tenths.text = "0"
    End If
    
    bmControlIsReady = True
    bmSendChangeEvent = True
    smStartText = smText
    
    If blChangeEventNeeded Then
        Call SetTextValue(False, 0)
    End If
End Sub

Private Sub UserControl_Initialize()
    bmControlIsReady = False
    bmSendChangeEvent = True
    bmOkToTab = True
    Set mFont = New StdFont
    Set UserControl.Font = mFont
    pc_Frame.Top = 0
    pc_Frame.Left = 0
    ec_Hours.text = "HH"
    ec_Minutes.text = "MM"
    ec_Seconds.text = "SS"
    ec_Tenths.text = "T"
    bmMainControlHadFocus = False
    bmLostFocusRight = False
    bmLostFocusLeft = False
    bmControlIsReady = True
    imUseHours = True
    imUseTenths = True
    imHourGotFocus = False
    imMinuteGotFocus = False
    imSecondGotFocus = False
    imTenthGotFocus = False
    bmIgnoreNextUpKey = False
    ' MaxLength = 200
End Sub

Private Sub UserControl_ExitFocus()
    bmSendChangeEvent = False
    bmControlIsReady = False
    If ec_Hours.text = "HH" Or ec_Hours.text = "H" Then
        ec_Hours.text = "00"
    End If
    If Len(ec_Hours.text) < 2 Then
        'ec_Hours.Text = "0" + ec_Hours.Text
    End If
    
    If ec_Minutes.text = "MM" Or ec_Minutes.text = "M" Then
        ec_Minutes.text = "00"
    End If
    If Len(ec_Minutes.text) < 2 Then
        'ec_Minutes.Text = "0" + ec_Minutes.Text
    End If
    
    If ec_Seconds.text = "SS" Or ec_Seconds.text = "S" Then
        ec_Seconds.text = "00"
    End If
    If Len(ec_Seconds.text) < 2 Then
        'ec_Seconds.Text = "0" + ec_Seconds.Text
    End If
    
    If ec_Tenths.text = "T" Then
        ec_Tenths.text = "0"
    End If
    
    bmControlIsReady = True
    bmSendChangeEvent = True
End Sub

Private Sub UserControl_Resize()
    If bmIgnoreResize Then Exit Sub
    Call PositionAllControls
End Sub

Private Sub SetTextValue(bUpdateScreenText As Boolean, lhWnd As Long)
    bmControlIsReady = False
    Dim slValue As String

    If ec_Hours.text = "HH" And ec_Minutes.text = "MM" And ec_Seconds.text = "SS" Then
        Exit Sub
    End If

    smText = ""
    slValue = Trim(ec_Hours.text)
    If slValue = "HH" Then
        slValue = "00"
    End If
    If imUseHours Then
        If Len(slValue) < 1 Then
            smText = "00:"
        Else
            If Len(slValue) < 2 Then
                slValue = "0" + slValue
            End If
            smText = slValue + ":"
        End If
    End If
    
    If bUpdateScreenText Then
        If lhWnd <> ec_Hours.hwnd Then
            ' Fill this in only if the field is not selected
            ec_Hours.text = slValue
        End If
    End If

    slValue = Trim(ec_Minutes.text)
    If slValue = "MM" Then
        slValue = "00"
    End If
    If Len(slValue) < 1 Then
        smText = smText + "00:"
    Else
        If Len(slValue) < 2 Then
            slValue = "0" + slValue
        End If
        smText = smText + slValue + ":"
    End If

    If bUpdateScreenText Then
        If lhWnd <> ec_Minutes.hwnd Then
            ' Fill this in only if the field is not selected
            ec_Minutes.text = slValue
        End If
    End If

    slValue = Trim(ec_Seconds.text)
    If slValue = "SS" Then
        slValue = "00"
    End If
    If Len(slValue) < 1 Then
        smText = smText + "00"
    Else
        If Len(slValue) < 2 Then
            slValue = "0" + slValue
        End If
        smText = smText + slValue
    End If
    
    If bUpdateScreenText Then
        If lhWnd <> ec_Seconds.hwnd Then
            ' Fill this in only if the field is not selected
            ec_Seconds.text = slValue
        End If
    End If
    
    If imUseTenths Then
        slValue = Trim(ec_Tenths.text)
        If slValue = "T" Then
            slValue = "0"
        End If
        If Len(slValue) < 1 Then
            smText = smText + ".0"
        Else
            smText = smText + "." + slValue
        End If
    End If
    
    If bUpdateScreenText Then
        If lhWnd <> ec_Tenths.hwnd Then
            ' Fill this in only if the field is not selected
            ec_Tenths.text = slValue
        End If
    End If
    
    If bmSendChangeEvent Then
        RaiseEvent OnChange
    End If
    bmControlIsReady = True
End Sub

Private Sub PositionAllControls()
    Dim ilTotalWidth As Integer
    Dim iPixelWidth As Integer
    Dim ilWidthAdjustment As Integer
    
    bmIgnoreResize = True
    pc_Frame.Height = Height
    ec_Hours.Height = pc_Frame.Height
    ec_Colon_1.Height = pc_Frame.Height
    ec_Minutes.Height = pc_Frame.Height
    ec_Colon_2.Height = pc_Frame.Height
    ec_Seconds.Height = pc_Frame.Height
    ec_Decimal_1.Height = pc_Frame.Height
    ec_Tenths.Height = pc_Frame.Height

    ilWidthAdjustment = 0
    ec_Hours.Top = 0
    ec_Hours.Left = 0
    ' gLogMsg "PAC - UseHours:" & Str(imUseHours) & ", UseTenths:" & Str(imUseTenths), "CSI_Controls.txt", False
    If imUseHours Then    ' imUseHours
        iPixelWidth = TextWidth("HH")
        ec_Hours.Visible = True
        ec_Hours.Width = iPixelWidth + (Screen.TwipsPerPixelX * ilWidthAdjustment)

        iPixelWidth = TextWidth(":")
        ec_Colon_1.Visible = True
        ec_Colon_1.Left = ec_Hours.Width
        ec_Colon_1.Width = iPixelWidth + (Screen.TwipsPerPixelX * ilWidthAdjustment)
        
        iPixelWidth = TextWidth("MM")
        ec_Minutes.Left = ec_Hours.Width + ec_Colon_1.Width
        ec_Minutes.Width = iPixelWidth + (Screen.TwipsPerPixelX * ilWidthAdjustment)
        
        iPixelWidth = TextWidth(":")
        ec_Colon_2.Left = ec_Hours.Width + ec_Colon_1.Width + ec_Minutes.Width
        ec_Colon_2.Width = iPixelWidth + (Screen.TwipsPerPixelX * ilWidthAdjustment)
    
        iPixelWidth = TextWidth("SS")
        ec_Seconds.Left = ec_Hours.Width + ec_Colon_1.Width + ec_Minutes.Width + ec_Colon_2.Width
        ec_Seconds.Width = iPixelWidth + (Screen.TwipsPerPixelX * ilWidthAdjustment)
    
        iPixelWidth = TextWidth(".")
        ec_Decimal_1.Left = ec_Hours.Width + ec_Colon_1.Width + ec_Minutes.Width + ec_Colon_2.Width + ec_Seconds.Width
        ec_Decimal_1.Width = iPixelWidth + (Screen.TwipsPerPixelX * ilWidthAdjustment)
    
        iPixelWidth = TextWidth("T")
        ec_Tenths.Left = ec_Hours.Width + ec_Colon_1.Width + ec_Minutes.Width + ec_Colon_2.Width + ec_Seconds.Width + ec_Decimal_1.Width
        ec_Tenths.Width = iPixelWidth + (Screen.TwipsPerPixelX * ilWidthAdjustment)
        
        ilTotalWidth = ec_Hours.Width + _
                         ec_Colon_1.Width + _
                         ec_Minutes.Width + _
                         ec_Colon_2.Width + _
                         ec_Seconds.Width + _
                         ec_Decimal_1.Width + _
                         ec_Tenths.Width + _
                         (Screen.TwipsPerPixelX * 6)
    Else
        ec_Hours.Width = 0
        ec_Hours.Visible = False
        ec_Colon_1.Width = 0
        ec_Colon_1.Visible = False
    
        iPixelWidth = TextWidth("MM")
        ec_Minutes.Left = 0
        ec_Minutes.Width = iPixelWidth + (Screen.TwipsPerPixelX * ilWidthAdjustment)
        
        iPixelWidth = TextWidth(":")
        ec_Colon_2.Left = ec_Minutes.Width
        ec_Colon_2.Width = iPixelWidth + (Screen.TwipsPerPixelX * ilWidthAdjustment)
    
        iPixelWidth = TextWidth("SS")
        ec_Seconds.Left = ec_Minutes.Width + ec_Colon_2.Width
        ec_Seconds.Width = iPixelWidth + (Screen.TwipsPerPixelX * ilWidthAdjustment)
    
        iPixelWidth = TextWidth(".")
        ec_Decimal_1.Left = ec_Minutes.Width + ec_Colon_2.Width + ec_Seconds.Width
        ec_Decimal_1.Width = iPixelWidth + (Screen.TwipsPerPixelX * ilWidthAdjustment)
    
        iPixelWidth = TextWidth("T")
        ec_Tenths.Left = ec_Minutes.Width + ec_Colon_2.Width + ec_Seconds.Width + ec_Decimal_1.Width
        ec_Tenths.Width = iPixelWidth + (Screen.TwipsPerPixelX * ilWidthAdjustment)
    
        ilTotalWidth = ec_Minutes.Width + _
                         ec_Colon_2.Width + _
                         ec_Seconds.Width + _
                         ec_Decimal_1.Width + _
                         ec_Tenths.Width + _
                         (Screen.TwipsPerPixelX * 6)
    End If

    If imUseHours Then
        ec_Hours.Enabled = True
    Else
        ec_Hours.Enabled = False
    End If
    
    Width = ilTotalWidth
    'Height = ec_Hours.Height + (Screen.TwipsPerPixelX * 1)
    pc_Frame.Width = ilTotalWidth

    If Not imUseTenths Then
        ec_Decimal_1.Width = 0
        ec_Decimal_1.Visible = False
        ec_Tenths.Width = 0
        ec_Tenths.Visible = False
        ' Stretch the width of the last control to fill in the rest of the area.
        ec_Seconds.Width = (pc_Frame.Width - (ec_Minutes.Left + ec_Minutes.Width)) * Screen.TwipsPerPixelX * 2
    Else
        ' Stretch the width of the last control to fill in the rest of the area.
		ec_Tenths.Visible = True
        If (ec_Tenths.Left + ec_Tenths.Width) <= pc_Frame.Width Then
            ec_Tenths.Width = (pc_Frame.Width - (ec_Tenths.Left + ec_Tenths.Width)) * Screen.TwipsPerPixelX * 2
        Else
            ec_Tenths.Width = pc_Frame.Width * Screen.TwipsPerPixelX * 2
        End If
    End If
    bmIgnoreResize = False
End Sub

'****************************************************************************
' Control Properties from within the designer.
'
'
'****************************************************************************

'****************************************************************************
' Load property values from storage
'****************************************************************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    smText = PropBag.ReadProperty("Text", "")
    BackColor = PropBag.ReadProperty("BackColor", RGB(255, 255, 255))
    ForeColor = PropBag.ReadProperty("ForeColor", RGB(255, 255, 255))
    imUseHours = PropBag.ReadProperty("CSI_UseHours", True)
    imUseTenths = PropBag.ReadProperty("CSI_UseTenths", True)

    ec_Hours.BackColor = BackColor
    ec_Colon_1.BackColor = BackColor
    ec_Minutes.BackColor = BackColor
    ec_Colon_2.BackColor = BackColor
    ec_Seconds.BackColor = BackColor
    ec_Decimal_1.BackColor = BackColor
    ec_Tenths.BackColor = BackColor

    ec_Hours.ForeColor = ForeColor
    ec_Colon_1.ForeColor = ForeColor
    ec_Minutes.ForeColor = ForeColor
    ec_Colon_2.ForeColor = ForeColor
    ec_Seconds.ForeColor = ForeColor
    ec_Decimal_1.ForeColor = ForeColor
    ec_Tenths.ForeColor = ForeColor
    
    pc_Frame.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)

    Exit Sub
    
    ec_Hours.Font = Font
    ec_Hours.FontSize = FontSize
    ec_Hours.FontBold = FontBold
    ec_Hours.FontItalic = FontItalic
    
    ec_Minutes.Font = Font
    ec_Minutes.FontSize = FontSize
    ec_Minutes.FontBold = FontBold
    ec_Minutes.FontItalic = FontItalic
    
    ec_Seconds.Font = Font
    ec_Seconds.FontSize = FontSize
    ec_Seconds.FontBold = FontBold
    ec_Seconds.FontItalic = FontItalic
    
    ec_Tenths.Font = Font
    ec_Tenths.FontSize = FontSize
    ec_Tenths.FontBold = FontBold
    ec_Tenths.FontItalic = FontItalic
    
    ec_Colon_1.Font = Font
    ec_Colon_1.FontSize = FontSize
    ec_Colon_1.FontBold = FontBold
    ec_Colon_1.FontItalic = FontItalic
    
    ec_Colon_2.Font = Font
    ec_Colon_2.FontSize = FontSize
    ec_Colon_2.FontBold = FontBold
    ec_Colon_2.FontItalic = FontItalic
    
    ec_Decimal_1.Font = Font
    ec_Decimal_1.FontSize = FontSize
    ec_Decimal_1.FontBold = FontBold
    ec_Decimal_1.FontItalic = FontItalic
    
    ' When the font changes, all controls have to be re-positioned.
    Call PositionAllControls
End Sub

'****************************************************************************
' Write property values to storage
'****************************************************************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", smText, "")
    Call PropBag.WriteProperty("BackColor", BackColor, RGB(255, 255, 255))
    Call PropBag.WriteProperty("ForeColor", ForeColor, RGB(255, 255, 255))
    Call PropBag.WriteProperty("BorderStyle", pc_Frame.BorderStyle)
    Call PropBag.WriteProperty("CSI_UseHours", imUseHours)
    Call PropBag.WriteProperty("CSI_UseTenths", imUseTenths)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
End Sub

Private Function GetString(iPos As Integer, sText As String) As String
    Dim ilLoop As Integer
    Dim ilLen As Integer
    Dim slOneChar As String
    
    GetString = ""
    ilLen = Len(sText)
    For ilLoop = iPos To ilLen
        slOneChar = Mid(sText, ilLoop, 1)
        If slOneChar = ":" Or slOneChar = "." Then
            imLastPos = ilLoop + 1
            Exit Function
        End If
        GetString = GetString + slOneChar
        imLastPos = ilLoop
    Next
End Function
'****************************************************************************
'
'****************************************************************************
Private Function SetupTextForBlankTime() As String
    If imUseHours Then
        SetupTextForBlankTime = "HH:MM:SS"
        ec_Hours.text = "HH"
        ec_Minutes.text = "MM"
        ec_Seconds = "SS"
    Else
        SetupTextForBlankTime = "MM:SS"
        ec_Hours.text = ""
        ec_Minutes.text = "MM"
        ec_Seconds = "SS"
    End If
    If imUseTenths Then
        SetupTextForBlankTime = SetupTextForBlankTime + ".T"
        ec_Tenths.text = "T"
    End If
End Function
Public Property Get text() As String
    bmSendChangeEvent = False
    Call SetTextValue(False, 0)
    bmSendChangeEvent = True
    text = smText
    'MsgBox "Giving it back as " & smText
End Property
Public Property Let text(sText As String)
    On Error GoTo Err_Text
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    Dim ilLoop As Integer
    Dim ilLen As Integer
    Dim slFields() As String
    Dim slNewText As String

    smStartText = smText

    'MsgBox "Setting time length to " & sText
    bmControlIsReady = False
    ec_Hours.text = ""
    ec_Minutes.text = ""
    ec_Seconds = ""
    ec_Tenths.text = ""
    
    ' The format expected is as follows.
    ' HH:MM:SS.T    ' Hours and Tenths turned on.
    ' MM:SS.T       ' Hours off and Tenths on.
    ' MM:SS         ' Hours and Tenths both off.
    ' HH:MM:SS      ' Hours on and and Tenths off.
    slNewText = Trim(sText)
    If Len(sText) < 1 Then
        ' User has passed in an empty string.
        slNewText = SetupTextForBlankTime()
        smText = ""
    Else
        mCSIParseCDFields slNewText, 0, slFields()
        If imUseHours Then
            If UBound(slFields) < 3 Then
                ' Time is not formated correctly.
                slNewText = SetupTextForBlankTime()
            Else
                ec_Hours.text = slFields(1)
                ec_Minutes.text = slFields(2)
                ec_Seconds = slFields(3)
                If imUseTenths And UBound(slFields) > 3 Then
                    ec_Tenths = slFields(4)
                End If
            End If
        Else
            ' Hours is not turned on
            If UBound(slFields) < 1 Then
                ' Time is not formated correctly. It needed to be at least MM:SS
                slNewText = SetupTextForBlankTime()
            Else
                ec_Minutes.text = slFields(1)
                ec_Seconds = slFields(2)
                If imUseTenths And UBound(slFields) > 2 Then
                    ec_Tenths = slFields(3)
                Else
                    ec_Tenths = "0"
                End If
            End If
        End If
    End If
    sText = slNewText
    bmControlIsReady = True
    PropertyChanged "Text"
    Exit Property
    
Err_Text:
    Exit Property
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get BackColor() As OLE_COLOR
   BackColor = ec_Hours.BackColor
End Property
Public Property Let BackColor(BKColor As OLE_COLOR)
    ec_Hours.BackColor = BKColor
    PropertyChanged "BackColor"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get ForeColor() As OLE_COLOR
   ForeColor = ec_Hours.ForeColor  ' lmForeColor
End Property
Public Property Let ForeColor(FGColor As OLE_COLOR)
    ec_Hours.ForeColor = FGColor
    PropertyChanged "ForeColor"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get BorderStyle() As CSI_TimeLength_BorderStyle
   BorderStyle = pc_Frame.BorderStyle
End Property
Public Property Let BorderStyle(BorderStyle As CSI_TimeLength_BorderStyle)
    pc_Frame.BorderStyle = BorderStyle
    PropertyChanged "BorderStyle"
    pc_Frame.BorderStyle = BorderStyle
    'Call PositionAllControls
End Property

''****************************************************************************
''
''****************************************************************************
'Public Property Get Font() As StdFont
'   Set Font = mFont
'End Property
'Public Property Set Font(ByVal New_Font As Font)
'   'Set UserControl.Font = New_Font
'   With mFont
'      .Bold = New_Font.Bold
'      .Italic = New_Font.Italic
'      .Name = New_Font.Name
'      .Size = New_Font.Size
'   End With
'   PropertyChanged "Font"
'   ec_Hours.Font = New_Font
'   ec_Colon_1.Font = New_Font
'   ec_Minutes.Font = New_Font
'   ec_Colon_2.Font = New_Font
'   ec_Seconds.Font = New_Font
'   ec_Decimal_1.Font = New_Font
'   ec_Tenths.Font = New_Font
'   Call PositionAllControls
'End Property
Private Sub mFont_FontChanged(ByVal PropertyName As String)
   Set UserControl.Font = mFont
   Refresh
End Sub

Public Property Get FontName() As String
    FontName = mFont.Name
End Property
Public Property Let FontName(sInFontName As String)
    mFont.Name = sInFontName
    ec_Hours.FontName = mFont.Name
    ec_Colon_1.FontName = mFont.Name
    ec_Minutes.FontName = mFont.Name
    ec_Colon_2.FontName = mFont.Name
    ec_Seconds.FontName = mFont.Name
    ec_Decimal_1.FontName = mFont.Name
    ec_Tenths.FontName = mFont.Name
    Call PositionAllControls
End Property
Public Property Get FontSize() As Double
    FontSize = mFont.Size
End Property
Public Property Let FontSize(dInFontSize As Double)
    mFont.Size = dInFontSize
    ec_Hours.FontSize = mFont.Size
    ec_Colon_1.FontSize = mFont.Size
    ec_Minutes.FontSize = mFont.Size
    ec_Colon_2.FontSize = mFont.Size
    ec_Seconds.FontSize = mFont.Size
    ec_Decimal_1.FontSize = mFont.Size
    ec_Tenths.FontSize = mFont.Size
    Call PositionAllControls
End Property
Public Property Get FontBold() As Integer
    FontBold = mFont.Bold
End Property
Public Property Let FontBold(dInFontBold As Integer)
    mFont.Bold = dInFontBold
    ec_Hours.FontBold = mFont.Bold
    ec_Colon_1.FontBold = mFont.Bold
    ec_Minutes.FontBold = mFont.Bold
    ec_Colon_2.FontBold = mFont.Bold
    ec_Seconds.FontBold = mFont.Bold
    ec_Decimal_1.FontBold = mFont.Bold
    ec_Tenths.FontBold = mFont.Bold
    Call PositionAllControls
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_UseHours() As Boolean
   CSI_UseHours = imUseHours
End Property
Public Property Let CSI_UseHours(Setting As Boolean)
    imUseHours = Setting
    Call PositionAllControls
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_UseTenths() As Boolean
   CSI_UseTenths = imUseTenths
End Property
Public Property Let CSI_UseTenths(Setting As Boolean)
    imUseTenths = Setting
    Call PositionAllControls
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
End Property


'*******************************************************
'*                                                     *
'*      Procedure Name:mCSIParseCDFields               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Parse colon delimited fields    *
'*                     Note:including quotes that are  *
'*                     enclosed within quotes          *
'*                     ""xxxxxxxx"":"xxxxx",           *
'*                                                     *
'*******************************************************
Sub mCSIParseCDFields(slCDStr As String, ilLower As Integer, slFields() As String)
'
'   gParseCDFields slCDStr, ilLower, slFields()
'   Where:
'       slCDStr(I)- Comma delinited string
'       ilLower(I)- True=Convert string fields characters to lower case (preceding character is A-Z)
'       slFields() (O)- fields parsed from comma delimited string
'
    Dim ilFieldNo As Integer
    Dim ilFieldType As Integer  '0=String, 1=Number
    Dim slChar As String
    Dim ilIndex As Integer
    Dim ilAscChar As Integer
    Dim ilAddToStr As Integer
    Dim slNextChar As String

'    For ilIndex = LBound(slFields) To UBound(slFields) Step 1
'        slFields(ilIndex) = ""
'    Next ilIndex
    ReDim slFields(1 To 1) As String
    slFields(UBound(slFields)) = ""
    ilFieldNo = 1
    ilIndex = 1
    ilFieldType = -1
    Do While ilIndex <= Len(Trim$(slCDStr))
        slChar = Mid$(slCDStr, ilIndex, 1)
        If ilFieldType = -1 Then
            If slChar = ":" Or slChar = "." Then    'Comma was followed by a comma-blank field
                ilFieldType = -1
                ilFieldNo = ilFieldNo + 1
                If ilFieldNo > UBound(slFields) Then
                    ReDim Preserve slFields(LBound(slFields) To UBound(slFields) + 1) As String
                    slFields(UBound(slFields)) = ""
                End If
            ElseIf slChar <> """" Then
                ilFieldType = 1
                slFields(ilFieldNo) = slChar
            Else
                ilFieldType = 0 'Quote field
            End If
        Else
            If ilFieldType = 0 Then 'Started with a Quote
                'Add to string unless "
                ilAddToStr = True
                If slChar = """" Then
                    If ilIndex = Len(Trim$(slCDStr)) Then
                        ilAddToStr = False
                    Else
                        slNextChar = Mid$(slCDStr, ilIndex + 1, 1)
                        If slNextChar = ":" Or slChar = "." Then
                            ilAddToStr = False
                        End If
                    End If
                End If
                If ilAddToStr Then
                    If (slFields(ilFieldNo) <> "") And ilLower Then
                        ilAscChar = Asc(UCase(Right$(slFields(ilFieldNo), 1)))
                        If ((ilAscChar >= Asc("A")) And (ilAscChar <= Asc("Z"))) Then
                            slChar = LCase$(slChar)
                        End If
                    End If
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        ReDim Preserve slFields(LBound(slFields) To UBound(slFields) + 1) As String
                        slFields(UBound(slFields)) = ""
                    End If
                    ilIndex = ilIndex + 1   'bypass comma
                End If
            Else
                'Add to string unless :
                If slChar <> ":" And slChar <> "." Then
                    slFields(ilFieldNo) = slFields(ilFieldNo) & slChar
                Else
                    ilFieldType = -1
                    ilFieldNo = ilFieldNo + 1
                    If ilFieldNo > UBound(slFields) Then
                        ReDim Preserve slFields(LBound(slFields) To UBound(slFields) + 1) As String
                        slFields(UBound(slFields)) = ""
                    End If
                End If
            End If
        End If
        ilIndex = ilIndex + 1
    Loop
End Sub

