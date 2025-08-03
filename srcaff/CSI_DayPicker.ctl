VERSION 5.00
Begin VB.UserControl CSI_DayPicker 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   ClipBehavior    =   0  'None
   KeyPreview      =   -1  'True
   ScaleHeight     =   3315
   ScaleMode       =   0  'User
   ScaleWidth      =   4455
   ToolboxBitmap   =   "CSI_DayPicker.ctx":0000
   Begin VB.CommandButton btn_DownArrow 
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Monotype Sorts"
         Size            =   6
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      UseMaskColor    =   -1  'True
      Width           =   300
   End
   Begin VB.PictureBox pbcDropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   255
      ScaleHeight     =   975
      ScaleWidth      =   3405
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1215
      Width           =   3405
      Begin VB.Label btn_BreakoutDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Compact Days"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   30
         TabIndex        =   14
         Top             =   705
         Width           =   3330
      End
      Begin VB.Label btn_Clear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clr"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2550
         TabIndex        =   13
         Top             =   450
         Width           =   690
      End
      Begin VB.Label btn_All 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "All"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   30
         TabIndex        =   12
         Top             =   450
         Width           =   720
      End
      Begin VB.Label btn_SaSu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sa-Su"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1695
         TabIndex        =   11
         Top             =   435
         Width           =   690
      End
      Begin VB.Label btn_MoFr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M-F"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   885
         TabIndex        =   10
         Top             =   435
         Width           =   585
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   15
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   1
         Left            =   495
         TabIndex        =   3
         Top             =   15
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   2
         Left            =   975
         TabIndex        =   4
         Top             =   15
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Th"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   3
         Left            =   1455
         TabIndex        =   5
         Top             =   15
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   4
         Left            =   1935
         TabIndex        =   6
         Top             =   15
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   5
         Left            =   2415
         TabIndex        =   7
         Top             =   15
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Su"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   6
         Left            =   2895
         TabIndex        =   8
         Top             =   15
         Width           =   480
      End
   End
   Begin VB.TextBox ec_InputBox 
      Height          =   315
      Left            =   270
      TabIndex        =   0
      Top             =   540
      Width           =   1710
   End
End
Attribute VB_Name = "CSI_DayPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' This event is fired when the text box is changed.
Event OnChange()

Public MaxLength As Integer

Private smText As String
Private bmShowSelectRangeButtons As Boolean
Private bmAllowMultiSelection As Boolean
Private bmShowDropDownOnFocus As Boolean
Private bmAllowBreakoutDays As Boolean

Private bmIsDroppedDown As Boolean
Private imDropDownWidth As Integer
Private imDropDownHeight As Integer
Private imEditBoxAlignment As Integer        ' 0=Edit box is on the Left, 1=Edit box is on the right
Private imIgnoreChangeEvent As Boolean
Private imDayOnColor As OLE_COLOR
Private imDayOffColor As OLE_COLOR
Private imRangePickerFGColor As OLE_COLOR
Private imRangePickerBGColor As OLE_COLOR
Private imCurrentButtonAnchor As Integer
Private bmIgnoreResize As Boolean

' Highlight the entire field when focus
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

Private Const MON = 0
Private Const TUE = 1
Private Const WED = 2
Private Const THU = 3
Private Const FRI = 4
Private Const SAT = 5
Private Const SUN = 6

Private Type DayStruct
    IsSlected As Boolean
    Text As String
End Type
Private DaysArray(0 To 6) As DayStruct

Private Sub btn_BreakoutDays_Click()
    If btn_BreakoutDays.Caption = "Compact Days" Then
        btn_BreakoutDays.Caption = "Breakout Days"
    Else
        btn_BreakoutDays.Caption = "Compact Days"
    End If
    RaiseEvent OnChange
End Sub

Private Sub ec_InputBox_GotFocus()
    ec_InputBox.SelStart = 0
    ec_InputBox.SelLength = Len(ec_InputBox.Text)
End Sub

Private Sub UserControl_Initialize()
    Dim ilLoop As Integer
    
    Set mFont = New StdFont
    Set UserControl.Font = mFont
    bmIsDroppedDown = False
    ec_InputBox.Top = 0
    ec_InputBox.Left = 0
    
    imDropDownWidth = pbcDropDown.Width
    imDropDownHeight = ec_InputBox.Height + pbcDropDown.Height
    For ilLoop = 0 To 6
        DaysArray(ilLoop).IsSlected = False
    Next
'    DaysArray(0).Text = "Mo"
'    DaysArray(1).Text = "Tu"
'    DaysArray(2).Text = "We"
'    DaysArray(3).Text = "Th"
'    DaysArray(4).Text = "Fr"
'    DaysArray(5).Text = "Sa"
'    DaysArray(6).Text = "Su"
    DaysArray(0).Text = "M"
    DaysArray(1).Text = "Tu"
    DaysArray(2).Text = "W"
    DaysArray(3).Text = "Th"
    DaysArray(4).Text = "F"
    DaysArray(5).Text = "Sa"
    DaysArray(6).Text = "Su"

    KeyPreview = True
    If imDayOnColor < 1 Then
        imDayOnColor = RGB(70, 200, 70)
        ' Set all other pre-defined values here as well.
        imRangePickerFGColor = RGB(0, 0, 0)
        imRangePickerBGColor = vbButtonFace
        bmShowSelectRangeButtons = True
        bmShowDropDownOnFocus = True
    End If
    If imDayOffColor < 1 Then
        imDayOffColor = vbButtonFace
    End If
    bmAllowMultiSelection = True
    bmAllowBreakoutDays = False
    imCurrentButtonAnchor = -1
    bmIgnoreResize = False
End Sub

Private Sub UserControl_InitProperties()
    Set Font = Ambient.Font
End Sub


Private Sub btn_All_Click()
    Call SetButtonStates(0, 6, True)
End Sub

Private Sub btn_MoFr_Click()
    Call SetButtonStates(0, 4, True)
End Sub

Private Sub btn_SaSu_Click()
    Call SetButtonStates(5, 6, True)
End Sub

Private Sub btn_Clear_Click()
    Call SetButtonStates(0, 6, False)
End Sub

Private Sub ec_InputBox_Change()
    Call ShowDaysBasedOnDaySettings
    RaiseEvent OnChange
End Sub

Private Sub ec_InputBox_Click()
    If bmShowDropDownOnFocus Then
        ec_InputBox.SetFocus
        Exit Sub
    End If
    Call UserControl_ExitFocus
End Sub

Private Sub UserControl_EnterFocus()
    If bmShowDropDownOnFocus Then
        Call btn_DownArrow_Click
        ec_InputBox.SetFocus
    End If
End Sub

Private Sub PositionAllControls()
    Dim ilTotalWidth As Integer
    Dim ilTotalHeight As Integer
    Dim iPixelWidth As Integer
    Dim iPixelHeight As Integer
    Dim CurLeft As Integer
    Dim ilButtonWidths As Integer
    Dim ilButtonHeights As Integer
    Dim ilLoop As Integer

    ' Adjust the size of all the buttons according to the current font and size.
    bmIgnoreResize = True
    CurLeft = 0
    iPixelWidth = TextWidth("SS")
    iPixelHeight = TextHeight("SS")
    ilButtonWidths = iPixelWidth + (Screen.TwipsPerPixelX * 4)
    ilButtonHeights = iPixelHeight + (Screen.TwipsPerPixelY * 2)
    For ilLoop = 0 To 6
        KeyPad_Buttons(ilLoop).Top = 0
        KeyPad_Buttons(ilLoop).Left = CurLeft
        KeyPad_Buttons(ilLoop).Width = ilButtonWidths
        KeyPad_Buttons(ilLoop).Height = ilButtonHeights
        CurLeft = CurLeft + ilButtonWidths - (Screen.TwipsPerPixelX) ' - (Screen.TwipsPerPixelX) removes the double border effect.
    Next
    ilTotalWidth = CurLeft
    btn_All.Top = KeyPad_Buttons(0).Height - (Screen.TwipsPerPixelY)    ' - (Screen.TwipsPerPixelY) removes the double border effect.
    btn_MoFr.Top = KeyPad_Buttons(0).Height - (Screen.TwipsPerPixelY)   ' - (Screen.TwipsPerPixelY) removes the double border effect.
    btn_SaSu.Top = KeyPad_Buttons(0).Height - (Screen.TwipsPerPixelY)   ' - (Screen.TwipsPerPixelY) removes the double border effect.
    btn_Clear.Top = KeyPad_Buttons(0).Height - (Screen.TwipsPerPixelY)  ' - (Screen.TwipsPerPixelY) removes the double border effect.
    
    btn_All.Height = ilButtonHeights
    btn_MoFr.Height = ilButtonHeights
    btn_SaSu.Height = ilButtonHeights
    btn_Clear.Height = ilButtonHeights
    btn_BreakoutDays.Height = ilButtonHeights
    btn_BreakoutDays.Top = btn_All.Top + btn_All.Height - (Screen.TwipsPerPixelY)
    
    btn_All.Width = (ilTotalWidth / 4)
    btn_All.Width = btn_All.Width - (Screen.TwipsPerPixelX * 12)
    btn_MoFr.Width = (ilTotalWidth / 4) + (Screen.TwipsPerPixelX)
    btn_MoFr.Width = btn_MoFr.Width + (Screen.TwipsPerPixelX * 12)
    btn_SaSu.Width = (ilTotalWidth / 4)
    btn_SaSu.Width = btn_SaSu.Width + (Screen.TwipsPerPixelX * 12)
    btn_Clear.Width = (ilTotalWidth / 4) + (Screen.TwipsPerPixelY)
    
    btn_All.Left = 0
    btn_MoFr.Left = btn_All.Left + btn_All.Width - (Screen.TwipsPerPixelX)
    btn_SaSu.Left = btn_MoFr.Left + btn_MoFr.Width
    btn_Clear.Left = btn_SaSu.Left + btn_SaSu.Width - (Screen.TwipsPerPixelX)
    btn_Clear.Width = ilTotalWidth - btn_Clear.Left + (Screen.TwipsPerPixelX)
    btn_BreakoutDays.Left = 0

    If bmAllowBreakoutDays Then
        ilTotalHeight = KeyPad_Buttons(0).Height + btn_All.Height + btn_BreakoutDays.Height - (Screen.TwipsPerPixelY * 3)
    Else
        ilTotalHeight = KeyPad_Buttons(0).Height + btn_All.Height - (Screen.TwipsPerPixelY * 2)
    End If
    pbcDropDown.Width = ilTotalWidth + (Screen.TwipsPerPixelX * 2)
    pbcDropDown.Height = ilTotalHeight + (Screen.TwipsPerPixelY * 2)
    btn_BreakoutDays.Width = ilTotalWidth + Screen.TwipsPerPixelX
    
    imDropDownWidth = pbcDropDown.Width
    imDropDownHeight = ec_InputBox.Height + pbcDropDown.Height
    bmIgnoreResize = False
End Sub

Private Sub ShowDaysBasedOnDaySettings()
    Dim ilLoop As Integer
    Dim ilLen As Integer
    Dim slOneChar As String
    Dim blRangeIsSelected As Boolean

    If imIgnoreChangeEvent Then Exit Sub
    ' Prepare to rebuild the enabled list depending on whats selected.
    For ilLoop = 0 To 6
        DaysArray(ilLoop).IsSlected = False
    Next

    ' Parse the input string and set each days button values
    blRangeIsSelected = False
    ilLen = Len(ec_InputBox.Text)
    If ilLen = 1 Then
        If UCase(ec_InputBox.Text) = "A" Then
            ec_InputBox.Text = "M-Su"
        End If
    End If
    For ilLoop = 1 To ilLen
        slOneChar = UCase(Mid(ec_InputBox.Text, ilLoop, 1))
        If ilLen = 1 And slOneChar = "A" Then
            ' Special code if the user types A as the first character.
            DaysArray(MON).IsSlected = True
            DaysArray(TUE).IsSlected = True
            DaysArray(WED).IsSlected = True
            DaysArray(THU).IsSlected = True
            DaysArray(FRI).IsSlected = True
            DaysArray(SAT).IsSlected = True
            DaysArray(SUN).IsSlected = True
            Exit For
        End If
        Select Case slOneChar
            Case "M"
                DaysArray(MON).IsSlected = True
            Case "W"
                If blRangeIsSelected Then
                    Call SetAllDaysNeededBeforeThisDay(WED)
                Else
                    DaysArray(WED).IsSlected = True
                End If
                blRangeIsSelected = False
            Case "T"
                slOneChar = UCase(Mid(ec_InputBox.Text, ilLoop + 1, 1))
                Select Case slOneChar
                    Case "U"
                        DaysArray(TUE).IsSlected = True
                        blRangeIsSelected = False
                    Case "H"
                        DaysArray(THU).IsSlected = True
                        If blRangeIsSelected Then
                            Call SetAllDaysNeededBeforeThisDay(THU)
                        End If
                        blRangeIsSelected = False
                End Select
            Case "F"
                If blRangeIsSelected Then
                    Call SetAllDaysNeededBeforeThisDay(FRI)
                Else
                    DaysArray(FRI).IsSlected = True
                End If
                blRangeIsSelected = False
            Case "S"
                slOneChar = UCase(Mid(ec_InputBox.Text, ilLoop + 1, 1))
                Select Case slOneChar
                    Case "A"
                        DaysArray(SAT).IsSlected = True
                        If blRangeIsSelected Then
                            Call SetAllDaysNeededBeforeThisDay(SAT)
                        End If
                        blRangeIsSelected = False
                    Case "U"
                        DaysArray(SUN).IsSlected = True
                        If blRangeIsSelected Then
                            Call SetAllDaysNeededBeforeThisDay(SUN)
                        End If
                        blRangeIsSelected = False
                    Case "S"
                        ' (SS) Special code to select Sa,Su
                        DaysArray(SAT).IsSlected = True
                        DaysArray(SUN).IsSlected = True
                End Select
            Case "-"
                blRangeIsSelected = True
        End Select
    Next
    Call SetButtonColors
End Sub

Private Sub SetAllDaysNeededBeforeThisDay(ThisDay As Integer)
    Dim iLoop1 As Integer
    Dim iLoop2 As Integer
    
    ' Back up to the first day that is select, then go forward to here selecting all between.
    For iLoop1 = ThisDay - 1 To MON Step -1
        If DaysArray(iLoop1).IsSlected Then
            ' ok, we found a day selected before the passed in day.
            ' Now move forward and select all of them between these two.
            For iLoop2 = iLoop1 + 1 To ThisDay '- 1
                DaysArray(iLoop2).IsSlected = True
            Next
            Exit Sub
        End If
    Next
End Sub

Private Sub ec_InputBox_KeyPress(KeyAscii As Integer)
    Dim CurChar As String
    Dim AllowedCharacters As String
    Dim PrevCharEntered As String
    Dim CurLen As Integer
    Dim ilLoop As Integer
    Dim ilTotalSelected As Integer
    
    If KeyAscii < 32 Then
        ' Allow all control keys to be processed as normal.
        Exit Sub
    End If
    AllowedCharacters = ",-MmoOTtuUWweEThHFfrRSsaASu"
    CurChar = Chr(KeyAscii)
    ' Convert Begining day names to upper case
    Select Case CurChar
        Case "m"
            CurChar = "M"
        Case "t"
            CurChar = "T"
        Case "s"
            CurChar = "S"
        Case "f"
            CurChar = "F"
        Case "w"
            CurChar = "W"
    End Select
    If Not bmAllowMultiSelection Then
        If ec_InputBox.SelStart = 0 And ec_InputBox.SelLength = Len(ec_InputBox.Text) Then
            ' Replace whatever is selected now with the current character entered if
            ' the entire text is selected.
            Call UnselectAllButtons
        End If
        ilTotalSelected = GetTotalSelected()
        If ilTotalSelected > 0 Then
            Beep
            KeyAscii = 0 ' Instruct VB to ignore this key.
            Exit Sub
        End If
    End If
    KeyAscii = Asc(CurChar)
    If InStr(1, AllowedCharacters, CurChar) = 0 Then
        Beep
        KeyAscii = 0 ' Instruct VB to ignore this key.
    End If
End Sub

Private Sub ec_InputBox_LostFocus()
    Dim ilLoop As Integer
    Dim slvalue As String

    ' Verify the string the user has entered.

' Don't change the values when losing focus !
'    slValue = UCase(ec_InputBox.Text)
'    Select Case slValue
'        Case "MF"
'            ec_InputBox.Text = "M-F"
'        Case "SS"
'            ec_InputBox.Text = "Sa,Su"
'        Case "A"
'            ec_InputBox.Text = "M-Su"
'    End Select
    Call FillInTextWithArrayValues
    'ec_InputBox.SetFocus
End Sub

Private Sub SetButtonColors()
    Dim ilLoop As Integer
    
    For ilLoop = 0 To 6
        If DaysArray(ilLoop).IsSlected Then
            KeyPad_Buttons(ilLoop).BackColor = imDayOnColor
        Else
            KeyPad_Buttons(ilLoop).BackColor = imDayOffColor
        End If
    Next
End Sub

Private Sub SetButtonStates(iStart As Integer, iEnd As Integer, bStatus As Boolean)
    Dim ilLoop As Integer
    Dim blAllAreSelected As Boolean
    
    blAllAreSelected = True
    For ilLoop = iStart To iEnd
        If Not DaysArray(ilLoop).IsSlected Then
            blAllAreSelected = False
            Exit For
        End If
    Next
    If blAllAreSelected Then
        ' All states are set to on. Set the back off regardless of the status being passed in.
        bStatus = False
    End If
    For ilLoop = iStart To iEnd
        DaysArray(ilLoop).IsSlected = bStatus
    Next
    
' This code flip flops them all on and off depending on whats selected now.
'    For ilLoop = iStart To iEnd
'        If DaysArray(ilLoop).IsSlected = bStatus Then
'            DaysArray(ilLoop).IsSlected = (Not bStatus)
'        Else
'            DaysArray(ilLoop).IsSlected = bStatus
'        End If
'    Next
    Call FillInTextWithArrayValues
End Sub

Private Sub SelectAllButtonsBetweenCurrentAnchor(iClickedButtonIndex As Integer)
    Dim ilLoop As Integer

    If imCurrentButtonAnchor > iClickedButtonIndex Then
        For ilLoop = iClickedButtonIndex To imCurrentButtonAnchor
            DaysArray(ilLoop).IsSlected = True
            KeyPad_Buttons(ilLoop).BackColor = imDayOnColor
        Next
    Else
        For ilLoop = imCurrentButtonAnchor To iClickedButtonIndex
            DaysArray(ilLoop).IsSlected = True
            KeyPad_Buttons(ilLoop).BackColor = imDayOnColor
        Next
    End If
End Sub

Private Sub UnselectAllButtons()
    Dim ilLoop As Integer
    
    For ilLoop = 0 To 6
        DaysArray(ilLoop).IsSlected = False
    Next
End Sub
Private Sub KeyPad_Buttons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bmAllowMultiSelection Then
        ' Anytime a day is selected and MultiSelection is turned off, clear all others first.
        Call UnselectAllButtons
    End If
    
    If DaysArray(Index).IsSlected Then
        DaysArray(Index).IsSlected = False
        KeyPad_Buttons(Index).BackColor = imDayOffColor
    Else
        DaysArray(Index).IsSlected = True
        KeyPad_Buttons(Index).BackColor = imDayOnColor
        If Shift And bmAllowMultiSelection Then
            If imCurrentButtonAnchor <> -1 Then
                ' Select all buttons between this button and the current anchor button.
                Call SelectAllButtonsBetweenCurrentAnchor(Index)
            End If
        End If
        imCurrentButtonAnchor = Index
    End If
    Call FillInTextWithArrayValues
End Sub

Private Function GetCountSelectedFromHere(Index As Integer, MaxCount As Integer) As Integer
    Dim ilLoop As Integer
    
    GetCountSelectedFromHere = 0
    For ilLoop = Index To MaxCount
        If DaysArray(ilLoop).IsSlected Then
            GetCountSelectedFromHere = GetCountSelectedFromHere + 1
        Else
            Exit Function
        End If
    Next
End Function

Private Function GetTotalSelected() As Integer
    Dim ilLoop As Integer
    
    GetTotalSelected = 0
    For ilLoop = 0 To 6
        If DaysArray(ilLoop).IsSlected Then
            GetTotalSelected = GetTotalSelected + 1
        End If
    Next
End Function

Private Sub FillInTextWithArrayValues()
    Dim ilLoop As Integer
    Dim slvalue As String
    Dim slDelimiter As String
    Dim TotalSelectedFromHere As Integer
    
    slvalue = ""
    For ilLoop = 0 To 6
        If DaysArray(ilLoop).IsSlected Then
            TotalSelectedFromHere = GetCountSelectedFromHere(ilLoop, 6)
            If TotalSelectedFromHere >= 3 Then
                slvalue = slvalue + DaysArray(ilLoop).Text + "-"
                ilLoop = ilLoop + TotalSelectedFromHere - 2
            Else
                slvalue = slvalue + DaysArray(ilLoop).Text + ","
            End If
        End If
    Next
    ' Check for and remove the final comma if it exists.
    If Len(slvalue) > 0 Then
        If right(slvalue, 1) = "," Then
            slvalue = Left(slvalue, Len(slvalue) - 1)
        End If
    End If
    imIgnoreChangeEvent = True
    ec_InputBox.Text = slvalue
    imIgnoreChangeEvent = False
    For ilLoop = 0 To 6
        If DaysArray(ilLoop).IsSlected Then
            KeyPad_Buttons(ilLoop).BackColor = imDayOnColor
        Else
            KeyPad_Buttons(ilLoop).BackColor = imDayOffColor
        End If
    Next
End Sub

Private Sub UserControl_ExitFocus()
    On Error Resume Next
    If bmIsDroppedDown Then
        'Width = ec_InputBox.Width
        pbcDropDown.Visible = False
        Height = ec_InputBox.Height
        'ec_InputBox.SetFocus
        bmIsDroppedDown = False
    End If
End Sub

Private Sub UserControl_Resize()
    If bmIgnoreResize Then Exit Sub
    If bmIsDroppedDown Then
        Exit Sub
    End If
    
    ec_InputBox.Width = Width
    ec_InputBox.Height = Height
    
    If ec_InputBox.BorderStyle = 0 Then
        ' The edit control does not have a border. Make it as tall as the edit control.
        btn_DownArrow.Height = Height
        btn_DownArrow.Top = ec_InputBox.Top
        btn_DownArrow.Left = ec_InputBox.Width - btn_DownArrow.Width
    Else
        ' The edit area has a border. Adjust it so it looks correct.
        btn_DownArrow.Height = Height - (Screen.TwipsPerPixelX * 4)
        btn_DownArrow.Top = ec_InputBox.Top + Screen.TwipsPerPixelY * 2
        btn_DownArrow.Left = ec_InputBox.Width - btn_DownArrow.Width - (Screen.TwipsPerPixelX * 2)
    End If
    
    'Height = ec_InputBox.Height
    
End Sub

Private Sub btn_DownArrow_Click()
    If bmIsDroppedDown Then
        'Width = ec_InputBox.Width
        pbcDropDown.Visible = False
        Height = ec_InputBox.Height
        ec_InputBox.SetFocus
        bmIsDroppedDown = False
        Exit Sub
    End If
    bmIsDroppedDown = True

    If Not bmShowSelectRangeButtons Or Not bmAllowMultiSelection Then
        ' If bmAllowMultiSelection is not enabled, then the bottom range selection is
        ' not turned on either.
        pbcDropDown.Height = KeyPad_Buttons(0).Height
    End If

    pbcDropDown.Top = (ec_InputBox.Top + ec_InputBox.Height)
    pbcDropDown.Left = ScaleLeft
    If imEditBoxAlignment = 1 Then  ' Are we aligning the edit box on the right ?
        If pbcDropDown.Left + pbcDropDown.Width < ec_InputBox.Left + ec_InputBox.Width Then
            ' Reposition the drop down so it always lines up on the right side.
            pbcDropDown.Left = (ec_InputBox.Left + ec_InputBox.Width) - pbcDropDown.Width
        End If
    End If
    pbcDropDown.Visible = True
    pbcDropDown.SetFocus

    If Width < pbcDropDown.Width Then
        Width = pbcDropDown.Width
    End If
    Height = imDropDownHeight + Screen.TwipsPerPixelY + 2000
End Sub

'****************************************************************************
' Control Properties from within the designer.
'
'
'****************************************************************************
Private Sub AssignControlProperties()
    ec_InputBox.Font = Font
    ec_InputBox.FontSize = FontSize
    ec_InputBox.FontBold = FontBold
    ec_InputBox.FontItalic = FontItalic

    If imEditBoxAlignment = 0 Then
        ec_InputBox.Left = 0
    Else
        ec_InputBox.Left = Width - ec_InputBox.Width
    End If
    btn_DownArrow.Top = ec_InputBox.Top + Screen.TwipsPerPixelY * 2
    btn_DownArrow.Left = ec_InputBox.Left + ec_InputBox.Width - btn_DownArrow.Width - (Screen.TwipsPerPixelX * 2)
    
    btn_All.ForeColor = imRangePickerFGColor
    btn_All.BackColor = imRangePickerBGColor
    btn_MoFr.ForeColor = imRangePickerFGColor
    btn_MoFr.BackColor = imRangePickerBGColor
    btn_SaSu.ForeColor = imRangePickerFGColor
    btn_SaSu.BackColor = imRangePickerBGColor
    btn_Clear.ForeColor = imRangePickerFGColor
    btn_Clear.BackColor = imRangePickerBGColor
    btn_BreakoutDays.Visible = bmAllowBreakoutDays
End Sub

'****************************************************************************
' Load property values from storage
'****************************************************************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    smText = PropBag.ReadProperty("Text", "")
    ec_InputBox.BackColor = PropBag.ReadProperty("BackColor", RGB(255, 255, 255))
    ec_InputBox.ForeColor = PropBag.ReadProperty("ForeColor", RGB(0, 0, 0))
    ec_InputBox.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    bmShowSelectRangeButtons = PropBag.ReadProperty("CSI_ShowSelectRangeButtons", True)
    bmAllowMultiSelection = PropBag.ReadProperty("CSI_AllowMultiSelection", True)
    bmAllowBreakoutDays = PropBag.ReadProperty("CSI_AllowBreakoutDays", False)
    bmShowDropDownOnFocus = PropBag.ReadProperty("CSI_ShowDropDownOnFocus", True)
    imEditBoxAlignment = PropBag.ReadProperty("CSI_InputBoxBoxAlignment", 0)
    imDayOnColor = PropBag.ReadProperty("CSI_DayOnColor", RGB(70, 200, 70))
    imDayOffColor = PropBag.ReadProperty("CSI_DayOffColor", vbButtonFace)
    imRangePickerFGColor = PropBag.ReadProperty("CSI_RangeFGColor", RGB(0, 0, 0))
    imRangePickerBGColor = PropBag.ReadProperty("CSI_RangeBGColor", vbButtonFace)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    Call AssignControlProperties
End Sub

'****************************************************************************
' Write property values to storage
'****************************************************************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", smText, "")
    Call PropBag.WriteProperty("BackColor", ec_InputBox.BackColor, RGB(255, 255, 255))
    Call PropBag.WriteProperty("ForeColor", ec_InputBox.ForeColor, RGB(0, 0, 0))
    Call PropBag.WriteProperty("BorderStyle", ec_InputBox.BorderStyle)
    Call PropBag.WriteProperty("CSI_ShowSelectRangeButtons", bmShowSelectRangeButtons)
    Call PropBag.WriteProperty("CSI_AllowMultiSelection", bmAllowMultiSelection)
    Call PropBag.WriteProperty("CSI_ShowDropDownOnFocus", bmShowDropDownOnFocus)
    Call PropBag.WriteProperty("CSI_InputBoxBoxAlignment", imEditBoxAlignment)
    Call PropBag.WriteProperty("CSI_DayOnColor", imDayOnColor)
    Call PropBag.WriteProperty("CSI_DayOffColor", imDayOffColor)
    Call PropBag.WriteProperty("CSI_RangeFGColor", imRangePickerFGColor)
    Call PropBag.WriteProperty("CSI_RangeBGColor", imRangePickerBGColor)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
End Sub

'****************************************************************************
'
'****************************************************************************
Public Property Get Text() As String
   Text = ec_InputBox.Text
End Property
Public Property Let Text(sText As String)
    smText = sText
    ec_InputBox.Text = smText
    PropertyChanged "Text"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get BackColor() As OLE_COLOR
   BackColor = ec_InputBox.BackColor
End Property
Public Property Let BackColor(BKColor As OLE_COLOR)
    ec_InputBox.BackColor = BKColor
    Call AssignControlProperties
    PropertyChanged "BackColor"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get ForeColor() As OLE_COLOR
   ForeColor = ec_InputBox.ForeColor  ' lmForeColor
End Property
Public Property Let ForeColor(FGColor As OLE_COLOR)
    ec_InputBox.ForeColor = FGColor
    Call AssignControlProperties
    PropertyChanged "ForeColor"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get BorderStyle() As BorderStyleConstants
   BorderStyle = ec_InputBox.BorderStyle
End Property
Public Property Let BorderStyle(BorderStyle As BorderStyleConstants)
    ec_InputBox.BorderStyle = BorderStyle
    Call AssignControlProperties
    PropertyChanged "BorderStyle"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_ShowSelectRangeButtons() As Boolean
   CSI_ShowSelectRangeButtons = bmShowSelectRangeButtons
End Property
Public Property Let CSI_ShowSelectRangeButtons(Setting As Boolean)
    bmShowSelectRangeButtons = Setting
    Call AssignControlProperties
    PropertyChanged "CSI_ShowSelectRangeButtons"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_AllowMultiSelection() As Boolean
   CSI_AllowMultiSelection = bmAllowMultiSelection
End Property
Public Property Let CSI_AllowMultiSelection(Setting As Boolean)
    bmAllowMultiSelection = Setting
    Call AssignControlProperties
    PropertyChanged "CSI_AllowMultiSelection"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_ShowDropDownOnFocus() As Boolean
   CSI_ShowDropDownOnFocus = bmShowDropDownOnFocus
End Property
Public Property Let CSI_ShowDropDownOnFocus(Setting As Boolean)
    bmShowDropDownOnFocus = Setting
    Call AssignControlProperties
    PropertyChanged "CSI_ShowDropDownOnFocus"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_DayOnColor() As OLE_COLOR
   CSI_DayOnColor = imDayOnColor
End Property
Public Property Let CSI_DayOnColor(Setting As OLE_COLOR)
    imDayOnColor = Setting
    Call AssignControlProperties
    PropertyChanged "CSI_DayOnColor"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_DayOffColor() As OLE_COLOR
   CSI_DayOffColor = imDayOffColor
End Property
Public Property Let CSI_DayOffColor(Setting As OLE_COLOR)
    imDayOffColor = Setting
    Call AssignControlProperties
    PropertyChanged "CSI_DayOffColor"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_RangeFGColor() As OLE_COLOR
   CSI_RangeFGColor = imRangePickerFGColor
End Property
Public Property Let CSI_RangeFGColor(Setting As OLE_COLOR)
    imRangePickerFGColor = Setting
    Call AssignControlProperties
    PropertyChanged "CSI_RangeFGColor"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_RangeBGColor() As OLE_COLOR
   CSI_RangeBGColor = imRangePickerBGColor
End Property
Public Property Let CSI_RangeBGColor(Setting As OLE_COLOR)
    imRangePickerBGColor = Setting
    Call AssignControlProperties
    PropertyChanged "CSI_RangeBGColor"
End Property
'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_AllowBreakoutDays() As Boolean
   CSI_AllowBreakoutDays = bmAllowBreakoutDays
End Property
Public Property Let CSI_AllowBreakoutDays(Setting As Boolean)
    bmAllowBreakoutDays = Setting
    Call AssignControlProperties
    Call PositionAllControls
    PropertyChanged "CSI_AllowBreakoutDays"
End Property
'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_BreakoutDays() As Boolean
    If btn_BreakoutDays.Caption = "Compact Days" Then
        CSI_BreakoutDays = False
    Else
        CSI_BreakoutDays = True
    End If
End Property
Public Property Let CSI_BreakoutDays(Setting As Boolean)
    If Setting Then
        btn_BreakoutDays.Caption = "Breakout Days"
    Else
        btn_BreakoutDays.Caption = "Compact Days"
    End If
    PropertyChanged "CSI_BreakoutDays"
End Property

''****************************************************************************
''
''****************************************************************************
'Public Property Get Font() As StdFont
'   Set Font = mFont
'End Property
'
'Public Property Set Font(ByVal New_Font As Font)
'   Dim ilLoop As Integer
'
'   With mFont
'      .Bold = New_Font.Bold
'      .Italic = New_Font.Italic
'      .Name = New_Font.Name
'      .Size = New_Font.Size
'   End With
'   ec_InputBox.Font = New_Font
'   For ilLoop = 0 To 6
'      KeyPad_Buttons(ilLoop).Font = New_Font
'   Next
'   btn_All.Font = New_Font
'   btn_MoFr.Font = New_Font
'   btn_SaSu.Font = New_Font
'   btn_Clear.Font = New_Font
'
'   PropertyChanged "Font"
'   Call AssignControlProperties
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
    Dim ilLoop As Integer
    
    mFont.Name = sInFontName
    UserControl.FontName = sInFontName
    ec_InputBox.FontName = sInFontName
    For ilLoop = 0 To 6
       KeyPad_Buttons(ilLoop).FontName = mFont.Name
    Next
    btn_All.FontName = mFont.Name
    btn_MoFr.FontName = mFont.Name
    btn_SaSu.FontName = mFont.Name
    btn_Clear.FontName = mFont.Name
    
    Call AssignControlProperties
    Call PositionAllControls
End Property
Public Property Get FontSize() As Double
    FontSize = mFont.Size
End Property
Public Property Let FontSize(dInFontSize As Double)
    Dim ilLoop As Integer

    mFont.Size = dInFontSize
    UserControl.FontSize = mFont.Size
    ec_InputBox.FontSize = mFont.Size
    For ilLoop = 0 To 6
       KeyPad_Buttons(ilLoop).FontSize = mFont.Size
    Next
    'btn_DownArrow.FontSize = mFont.Size
    btn_All.FontSize = mFont.Size
    btn_MoFr.FontSize = mFont.Size
    btn_SaSu.FontSize = mFont.Size
    btn_Clear.FontSize = mFont.Size
    
    Call AssignControlProperties
    Call PositionAllControls
End Property
Public Property Get FontBold() As Integer
    FontBold = mFont.Bold
End Property
Public Property Let FontBold(dInFontBold As Integer)
    Dim ilLoop As Integer

    mFont.Bold = dInFontBold
    UserControl.FontBold = mFont.Bold
    ec_InputBox.FontBold = mFont.Bold
    For ilLoop = 0 To 6
       KeyPad_Buttons(ilLoop).FontBold = mFont.Bold
    Next
    'btn_DownArrow.FontBold = mFont.Bold
    btn_All.FontBold = mFont.Bold
    btn_MoFr.FontBold = mFont.Bold
    btn_SaSu.FontBold = mFont.Bold
    btn_Clear.FontBold = mFont.Bold
    btn_BreakoutDays.FontBold = mFont.Bold
    Call AssignControlProperties
    ' Call PositionAllControls
End Property


