VERSION 5.00
Begin VB.UserControl CSI_Calendar 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   ScaleHeight     =   4830
   ScaleWidth      =   5400
   ToolboxBitmap   =   "CSI_Calendar.ctx":0000
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
      Left            =   2325
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   390
      UseMaskColor    =   -1  'True
      Width           =   300
   End
   Begin VB.PictureBox pbcDayNameFontHolder 
      Height          =   330
      Left            =   330
      ScaleHeight     =   270
      ScaleWidth      =   930
      TabIndex        =   6
      Top             =   3705
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox pbcCalFrame 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   300
      ScaleHeight     =   2715
      ScaleWidth      =   3705
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   870
      Width           =   3705
      Begin VB.VScrollBar sb_Year 
         Height          =   270
         Left            =   2730
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton btn_MonthNext 
         Caption         =   "t"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   6
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3060
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   345
      End
      Begin VB.CommandButton btn_MonthPrev 
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Monotype Sorts"
            Size            =   6
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   330
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   225
         Width           =   345
      End
      Begin VB.Label ec_Year 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   2040
         TabIndex        =   9
         Top             =   255
         Width           =   525
      End
      Begin VB.Label lc_WeekNames 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mo"
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   0
         Left            =   630
         TabIndex        =   8
         Top             =   840
         Width           =   450
      End
      Begin VB.Label ec_MonthName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Month Name"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   780
         TabIndex        =   7
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   615
         TabIndex        =   5
         Top             =   1380
         Width           =   480
      End
   End
   Begin VB.TextBox ec_InputBox 
      Height          =   315
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "CSI_Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum CalDateTypeList
   LongMonth
   ShortMonth
End Enum

Public Enum CalDefaultDateList
   csiNoDate
   csiCurrentDate
   csiThisMonday
   csiNextMonday
   csiStandardBroadcastMonday
End Enum

' This event is fired when the text box is changed.
Event CalendarChanged()
Event Change()
Event DateClicked()

Private Const NDAYS = 34

' Create the day boxes array and place in the proper places.
' Assign day values to each according to the date setting.
' Assign color values based on settings and also based on whether the day can be picked.
' Allow user to pick a day if the day is available.
' Process up, down, right, left arrows to adjust the date properly

Private smCalendarDate As String
Private dmCurDate As Date
Private bmShowDropDownOnFocus As Boolean
Private bmIgnoreResize As Boolean
Private bmIsDroppedDown As Boolean
Private imDropDownWidth As Integer
Private imDropDownHeight As Integer
Private imEditBoxAlignment As Integer        ' 0=Edit box is on the Left, 1=Edit box is on the right
Private imIgnoreChangeEvent As Boolean
Private cmCurrentDayBGColor As OLE_COLOR
Private cmCurrentDayFGColor As OLE_COLOR
Private cmCalendarBGColor As OLE_COLOR
Private bmCloseCalAfterSelection As Boolean
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
Private fmDayNameFont As StdFont
Private fmMonthNameFont As StdFont
Private imCalDateFormat As CalDateTypeList
Private cmDaySelectedColor As OLE_COLOR
Private cmDayUnSelectedColor As OLE_COLOR
Private imLastSelectedDay As Integer
Private bmFreezeCalendar As Boolean
Private bmIgnoreYearChange As Boolean
Private bmForceMondaySelectionOnly As Boolean
Private dmFirstDate As Date
Private dmLastDate As Date
Private bmAllowBlankDate As Boolean
Private bmAllowTFN As Boolean
Private bmDefaultDateType As CalDefaultDateList
Private bmUserIsTyping As Boolean

' This structure is used to define the properties for each day shown on the calendar.
Private Type OneDayProperty
    iDay As Integer
    lFGColor As Long
    lBGColor As Long
    lSelectedFGColor As Long
    lSelectedBGColor As Long
    bCanSelect As Boolean
    dDate As Date
End Type
Private tmDayProperties(0 To NDAYS) As OneDayProperty

'****************************************************************************
'
'****************************************************************************
Private Sub ec_Year_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilDay As Integer
    
    If Not gIsDate(smCalendarDate) Then
        smCalendarDate = DateSerial(year(gNow()), month(gNow()), Day(gNow()))
    End If
    If Button = 1 Then
        ec_Year.Caption = ec_Year.Caption + 1
    Else
        ec_Year.Caption = ec_Year.Caption - 1
    End If
    ilMonth = month(smCalendarDate)
    ilDay = Day(smCalendarDate)
    ilYear = ec_Year.Caption
    smCalendarDate = DateSerial(ilYear, ilMonth, 1)
    Call mSetCalendarDays(smCalendarDate)
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub UserControl_Initialize()
    Dim ilLoop As Integer
    
    bmIgnoreResize = True
    bmIgnoreYearChange = True
    ' Create the calendar array of controls from the single control on the form.
    For ilLoop = 1 To NDAYS
        Load KeyPad_Buttons(ilLoop)
        KeyPad_Buttons(ilLoop).Visible = True
    Next
    For ilLoop = 1 To 6
        Load lc_WeekNames(ilLoop)
        lc_WeekNames(ilLoop).Visible = True
    Next
    lc_WeekNames(0).Caption = "M"
    lc_WeekNames(1).Caption = "Tu"
    lc_WeekNames(2).Caption = "W"
    lc_WeekNames(3).Caption = "Th"
    lc_WeekNames(4).Caption = "F"
    lc_WeekNames(5).Caption = "Sa"
    lc_WeekNames(6).Caption = "Su"

    Set mFont = New StdFont
    Set UserControl.Font = mFont
    Set fmDayNameFont = mFont
    Set fmMonthNameFont = mFont

    imCalDateFormat = ShortMonth
    bmIsDroppedDown = False
    ec_InputBox.Top = 0
    ec_InputBox.Left = 0

    KeyPreview = True
    Call ResetDayProperties
    imLastSelectedDay = 0
    cmCalendarBGColor = RGB(170, 255, 255)
    cmDaySelectedColor = RGB(0, 0, 170)
    cmDayUnSelectedColor = RGB(255, 255, 255)
    cmCurrentDayBGColor = RGB(255, 255, 255)
    cmCurrentDayFGColor = RGB(0, 200, 0)
    
    sb_Year.Min = 1800
    sb_Year.Max = 4000
    sb_Year.LargeChange = 1
    sb_Year.SmallChange = 1
    bmFreezeCalendar = False
    smCalendarDate = DateSerial(year(gNow()), month(gNow()), Day(gNow()))
    sb_Year.Value = year(smCalendarDate)
    Call mSetCalendarDays(smCalendarDate)
    bmCloseCalAfterSelection = True
    bmIgnoreYearChange = False
    bmIgnoreResize = False
    bmUserIsTyping = False
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub PositionAllControls()
    Dim ilTotalWidth As Integer
    Dim ilTotalHeight As Integer
    Dim iPixelWidth As Integer
    Dim iPixelHeight As Integer
    Dim ilCurTop As Integer
    Dim ilCurLeft As Integer
    Dim ilButtonWidths As Integer
    Dim ilButtonHeights As Integer
    Dim ilLoop As Integer
    Dim ilTopBorder As Integer
    Dim ilBottomBorder As Integer
    Dim ilLeftBorder As Integer
    Dim ilRightBorder As Integer
    
    ilTopBorder = 50
    ilBottomBorder = 50
    ilLeftBorder = 50
    ilRightBorder = 50
    
    ec_InputBox.Font = Font
    ec_InputBox.FontSize = FontSize
    ec_InputBox.FontBold = FontBold
    ec_InputBox.FontItalic = FontItalic

    ec_MonthName.FontName = fmMonthNameFont.Name
    ec_MonthName.FontSize = fmMonthNameFont.Size
    ec_MonthName.FontBold = fmMonthNameFont.Bold
    ec_MonthName.FontItalic = fmMonthNameFont.Italic
    
    ec_Year.FontName = fmMonthNameFont.Name
    ec_Year.FontSize = fmMonthNameFont.Size
    ec_Year.FontBold = fmMonthNameFont.Bold
    ec_Year.FontItalic = fmMonthNameFont.Italic

    'btn_MonthPrev.FontSize = fmMonthNameFont.Size
    'btn_MonthNext.FontSize = fmMonthNameFont.Size
    'sb_Year.Width = fmMonthNameFont.Size * 16

    If imEditBoxAlignment = 0 Then
        ec_InputBox.Left = 0
    Else
        ec_InputBox.Left = Width - ec_InputBox.Width
    End If
    btn_DownArrow.Top = ec_InputBox.Top + Screen.TwipsPerPixelY * 2
    btn_DownArrow.Left = ec_InputBox.Left + ec_InputBox.Width - btn_DownArrow.Width - (Screen.TwipsPerPixelX * 2)

    For ilLoop = 0 To 6
        lc_WeekNames(ilLoop).Font = pbcDayNameFontHolder.Font
        lc_WeekNames(ilLoop).FontBold = pbcDayNameFontHolder.FontBold
        lc_WeekNames(ilLoop).FontSize = pbcDayNameFontHolder.FontSize
        lc_WeekNames(ilLoop).FontItalic = pbcDayNameFontHolder.FontItalic
    Next

    For ilLoop = 0 To NDAYS
        KeyPad_Buttons(ilLoop).Font = pbcDayNameFontHolder.Font
        KeyPad_Buttons(ilLoop).FontBold = pbcDayNameFontHolder.FontBold
        KeyPad_Buttons(ilLoop).FontSize = pbcDayNameFontHolder.FontSize
        KeyPad_Buttons(ilLoop).FontItalic = pbcDayNameFontHolder.FontItalic
    Next

    ' Adjust the size of all the buttons according to the current font and size.
    bmIgnoreResize = True
    iPixelWidth = pbcDayNameFontHolder.TextWidth("00") + (Screen.TwipsPerPixelX * 1)
    iPixelHeight = pbcDayNameFontHolder.TextHeight("00")
    
    ilButtonWidths = iPixelWidth + (Screen.TwipsPerPixelX * 4)
    ilButtonHeights = iPixelHeight + (Screen.TwipsPerPixelY * 2)

    btn_MonthPrev.height = ilButtonHeights - Screen.TwipsPerPixelY
    btn_MonthNext.height = ilButtonHeights - Screen.TwipsPerPixelY
    ec_MonthName.height = ilButtonHeights
    ec_Year.height = ilButtonHeights - Screen.TwipsPerPixelY
    'sb_Year.Height = ilButtonHeights - Screen.TwipsPerPixelY

    ilCurTop = ec_MonthName.Top + ec_MonthName.height
    ilCurLeft = ilLeftBorder
    For ilLoop = 0 To 6
        lc_WeekNames(ilLoop).Top = ilCurTop
        lc_WeekNames(ilLoop).height = ilButtonHeights + (Screen.TwipsPerPixelY)
        lc_WeekNames(ilLoop).Left = ilCurLeft
        lc_WeekNames(ilLoop).Width = ilButtonWidths + (Screen.TwipsPerPixelX)
        ilCurLeft = ilCurLeft + ilButtonWidths - (Screen.TwipsPerPixelX)
    Next

    ilCurTop = ilCurTop + ilButtonHeights
    ilCurLeft = ilLeftBorder
    For ilLoop = 0 To NDAYS
        If ilLoop And ilLoop Mod 7 = 0 Then
            ilCurTop = ilCurTop + ilButtonHeights
            ilCurLeft = ilLeftBorder
        End If
        KeyPad_Buttons(ilLoop).Top = ilCurTop
        KeyPad_Buttons(ilLoop).height = ilButtonHeights + (Screen.TwipsPerPixelY)
        KeyPad_Buttons(ilLoop).Left = ilCurLeft
        KeyPad_Buttons(ilLoop).Width = ilButtonWidths + (Screen.TwipsPerPixelX)
        ilCurLeft = ilCurLeft + ilButtonWidths - (Screen.TwipsPerPixelX)
    Next

    ilTotalWidth = ilCurLeft
    ilTotalHeight = ilTopBorder + KeyPad_Buttons(NDAYS).Top + KeyPad_Buttons(NDAYS).height - (Screen.TwipsPerPixelY * 4) ' - (Screen.TwipsPerPixelX) removes the double border effect.

    ec_MonthName.Top = ilTopBorder
    ec_Year.Top = ilTopBorder
    ec_Year.Width = pbcDayNameFontHolder.TextWidth("888888")
    'sb_Year.Top = ilTopBorder
    btn_MonthPrev.Top = ilTopBorder
    btn_MonthPrev.Left = ilLeftBorder
    btn_MonthNext.Top = ilTopBorder

    btn_MonthNext.Left = KeyPad_Buttons(6).Left + KeyPad_Buttons(6).Width - btn_MonthNext.Width
    'sb_Year.Left = btn_MonthNext.Left - sb_Year.Width
    'ec_Year.Left = sb_Year.Left - ec_Year.Width
    ec_Year.Left = btn_MonthNext.Left - ec_Year.Width
    ec_MonthName.Left = btn_MonthPrev.Left + btn_MonthPrev.Width
    ec_MonthName.Width = ec_Year.Left - ec_MonthName.Left
    
    pbcCalFrame.Top = ec_InputBox.Top + ec_InputBox.height

    pbcCalFrame.Width = KeyPad_Buttons(6).Left + KeyPad_Buttons(6).Width + ilRightBorder - Screen.TwipsPerPixelX
    pbcCalFrame.height = ilTotalHeight + (Screen.TwipsPerPixelX * 4)

    imDropDownWidth = pbcCalFrame.Width
    imDropDownHeight = ec_InputBox.height + pbcCalFrame.height
    bmIgnoreResize = False
End Sub

Private Sub SelectAndHighlightMainDate()
    ec_InputBox.SetFocus
    ec_InputBox.SelStart = 0
    ec_InputBox.SelLength = Len(ec_InputBox.Text)
End Sub

Private Sub pbcCalFrame_Paint()
    pbcCalFrame.ForeColor = RGB(0, 0, 0)
    pbcCalFrame.Line (0, 0)-Step(pbcCalFrame.Width - Screen.TwipsPerPixelX, pbcCalFrame.height - Screen.TwipsPerPixelY), , B
End Sub

Private Sub sb_Year_Change()
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilDay As Integer

    If bmIgnoreYearChange Then
        Exit Sub
    End If
    If Not gIsDate(smCalendarDate) Then
        smCalendarDate = DateSerial(year(gNow()), month(gNow()), Day(gNow()))
    End If

    ilMonth = month(smCalendarDate)
    ilDay = Day(smCalendarDate)
    ilYear = sb_Year.Value
    smCalendarDate = DateSerial(ilYear, ilMonth, 1)
    Call mSetCalendarDays(smCalendarDate)
    'RaiseEvent CalendarChanged
End Sub

Private Sub SelectCalendarDay(nDay As Integer)
    ' Unselect any previous day selection.
    KeyPad_Buttons(imLastSelectedDay).BackColor = tmDayProperties(imLastSelectedDay).lBGColor
    KeyPad_Buttons(imLastSelectedDay).ForeColor = tmDayProperties(imLastSelectedDay).lFGColor
    'KeyPad_Buttons(imLastSelectedDay).BackColor = cmDayUnSelectedColor
    imLastSelectedDay = nDay

    If bmForceMondaySelectionOnly Then
        While Weekday(tmDayProperties(imLastSelectedDay).dDate) <> vbMonday
            imLastSelectedDay = imLastSelectedDay - 1
        Wend
    End If
    KeyPad_Buttons(imLastSelectedDay).BackColor = tmDayProperties(imLastSelectedDay).lSelectedBGColor
    KeyPad_Buttons(imLastSelectedDay).ForeColor = tmDayProperties(imLastSelectedDay).lSelectedFGColor
End Sub

Private Sub KeyPad_Buttons_Click(Index As Integer)
    If Not tmDayProperties(Index).bCanSelect Then
        ' Exit Sub
    End If
    Call SelectCalendarDay(Index)
    bmFreezeCalendar = True
    ec_InputBox.Text = tmDayProperties(imLastSelectedDay).dDate
    bmFreezeCalendar = False
    Call SelectAndHighlightMainDate
    If bmCloseCalAfterSelection Then
        Call btn_DownArrow_Click
        'SendKeys "{TAB}", False
    End If
    RaiseEvent DateClicked
End Sub

Private Sub MoveUp()
    Dim slNewDate As Date
    
    If imLastSelectedDay >= 7 Then
        imLastSelectedDay = imLastSelectedDay - 7
        smCalendarDate = tmDayProperties(imLastSelectedDay).dDate
        bmFreezeCalendar = True
        ec_InputBox.Text = smCalendarDate
        bmFreezeCalendar = False
        Call mSelectDay
        Exit Sub
    End If
    ' Back up to the prior month
    slNewDate = DateAdd("m", -1, tmDayProperties(14).dDate)
    smCalendarDate = slNewDate
    Call mSetCalendarDays(smCalendarDate)
    imLastSelectedDay = imLastSelectedDay + (NDAYS - 6)
    While tmDayProperties(imLastSelectedDay).dDate = "1/1/1800"
        imLastSelectedDay = imLastSelectedDay - 7
    Wend
    'imLastSelectedDay = imLastSelectedDay - 7
    smCalendarDate = tmDayProperties(imLastSelectedDay).dDate
    bmFreezeCalendar = True
    ec_InputBox.Text = smCalendarDate
    bmFreezeCalendar = False
    Call mSelectDay
End Sub
Private Sub MoveDown()
    Dim slNewDate As Date
    
    If imLastSelectedDay <= (NDAYS - 7) Then
        imLastSelectedDay = imLastSelectedDay + 7
        If tmDayProperties(imLastSelectedDay).dDate <> "1/1/1800" Then
            smCalendarDate = tmDayProperties(imLastSelectedDay).dDate
            bmFreezeCalendar = True
            ec_InputBox.Text = smCalendarDate
            bmFreezeCalendar = False
            Call mSelectDay
            Exit Sub
        End If
    End If
    ' Move to the future month
    slNewDate = DateAdd("m", 1, tmDayProperties(14).dDate)
    smCalendarDate = slNewDate
    Call mSetCalendarDays(smCalendarDate)
    'RaiseEvent CalendarChanged
    While imLastSelectedDay >= 7
        imLastSelectedDay = imLastSelectedDay - 7
    Wend
    smCalendarDate = tmDayProperties(imLastSelectedDay).dDate
    bmFreezeCalendar = True
    ec_InputBox.Text = smCalendarDate
    bmFreezeCalendar = False
    Call mSelectDay
End Sub
Private Sub MoveLeft()
    Dim slNewDate As Date
    
    If imLastSelectedDay > 0 Then
        imLastSelectedDay = imLastSelectedDay - 1
        smCalendarDate = tmDayProperties(imLastSelectedDay).dDate
        bmFreezeCalendar = True
        ec_InputBox.Text = smCalendarDate
        bmFreezeCalendar = False
        Call mSelectDay
        Exit Sub
    End If
    ' Back up to the prior month
    slNewDate = DateAdd("m", -1, tmDayProperties(14).dDate)
    smCalendarDate = slNewDate
    Call mSetCalendarDays(smCalendarDate)
    imLastSelectedDay = NDAYS
    While imLastSelectedDay >= 0 And tmDayProperties(imLastSelectedDay).dDate = "1/1/1800"
        imLastSelectedDay = imLastSelectedDay - 1
    Wend
    smCalendarDate = tmDayProperties(imLastSelectedDay).dDate
    bmFreezeCalendar = True
    ec_InputBox.Text = smCalendarDate
    bmFreezeCalendar = False
    Call mSelectDay
End Sub
Private Sub MoveRight()
    Dim slNewDate As Date
    
    If imLastSelectedDay < NDAYS Then
        imLastSelectedDay = imLastSelectedDay + 1
        If tmDayProperties(imLastSelectedDay).dDate <> "1/1/1800" Then
            smCalendarDate = tmDayProperties(imLastSelectedDay).dDate
            bmFreezeCalendar = True
            ec_InputBox.Text = smCalendarDate
            bmFreezeCalendar = False
            Call mSelectDay
            Exit Sub
        End If
    End If
    ' Back up to the prior month
    slNewDate = DateAdd("m", 1, tmDayProperties(14).dDate)
    smCalendarDate = slNewDate
    Call mSetCalendarDays(smCalendarDate)
    imLastSelectedDay = 0
    smCalendarDate = tmDayProperties(imLastSelectedDay).dDate
    bmFreezeCalendar = True
    ec_InputBox.Text = smCalendarDate
    bmFreezeCalendar = False
    Call mSelectDay
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub ec_InputBox_KeyPress(KeyAscii As Integer)
    Dim slChar As String
    
    slChar = Chr(KeyAscii)
    Select Case slChar
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "/", "-"
            Exit Sub
        Case Else
            If KeyAscii = 8 Then    ' Back Space
                Exit Sub
            End If
            KeyAscii = 0    ' Ignore this key.
    End Select
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub ec_InputBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            bmUserIsTyping = False
            Call MoveUp
            Call SelectAndHighlightMainDate
            KeyCode = -1
        Case vbKeyDown
            bmUserIsTyping = False
            Call MoveDown
            Call SelectAndHighlightMainDate
            KeyCode = -1
        Case vbKeyLeft
            If bmUserIsTyping Then
                Exit Sub
            End If
            Call MoveLeft
            Call SelectAndHighlightMainDate
            KeyCode = -1
        Case vbKeyRight
            If bmUserIsTyping Then
                Exit Sub
            End If
            Call MoveRight
            Call SelectAndHighlightMainDate
            KeyCode = -1
        Case vbKeyReturn
            bmUserIsTyping = False
            Call SelectAndHighlightMainDate
            Call btn_DownArrow_Click
            'SendKeys "{TAB}", False
            KeyCode = -1
        Case vbKeyT, vbKeyF, vbKeyN
            If Not bmAllowTFN Then
                KeyCode = -1
            End If
        Case vbKeyA To vbKeyZ
            KeyCode = -1
        Case Else
            ' Once the user starts typing a date into the edit box, the arrow keys
            ' will then stop working until enter is pressed.
            bmUserIsTyping = True
    End Select
End Sub
Private Sub pbcCalFrame_KeyDown(KeyCode As Integer, Shift As Integer)
    Call ec_InputBox_KeyDown(KeyCode, Shift)
End Sub


'****************************************************************************
'
'****************************************************************************
Private Sub btn_MonthPrev_Click()
    Dim ilMonth As Integer
    Dim ilYear As Integer

    If Not gIsDate(smCalendarDate) Then
        smCalendarDate = DateSerial(year(gNow()), month(gNow()), Day(gNow()))
    End If
    dmCurDate = smCalendarDate
    ilMonth = month(dmCurDate)
    ilYear = year(dmCurDate)

    ilMonth = ilMonth - 1
    If ilMonth = 0 Then
        ilYear = ilYear - 1
        ilMonth = 12
    End If
    smCalendarDate = DateSerial(ilYear, ilMonth, 1)
    'ec_InputBox.text = smCalendarDate
    Call mSetCalendarDays(smCalendarDate)
    'RaiseEvent CalendarChanged
    ec_InputBox.SetFocus
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub btn_MonthNext_Click()
    Dim ilMonth As Integer
    Dim ilYear As Integer
    
    If Not gIsDate(smCalendarDate) Then
        smCalendarDate = DateSerial(year(gNow()), month(gNow()), Day(gNow()))
    End If
    dmCurDate = smCalendarDate
    ilMonth = month(dmCurDate)
    ilYear = year(dmCurDate)
    ilMonth = ilMonth + 1
    If ilMonth > 12 Then
        ilYear = ilYear + 1
        ilMonth = 1
    End If
    smCalendarDate = DateSerial(ilYear, ilMonth, 1)
    'ec_InputBox.text = smCalendarDate
    Call mSetCalendarDays(smCalendarDate)
    'RaiseEvent CalendarChanged
    ec_InputBox.SetFocus
End Sub

'****************************************************************************
'
'****************************************************************************
Public Sub Refresh()
    Call mSetCalendarDays(smCalendarDate)
End Sub

'****************************************************************************
'
'****************************************************************************
Public Sub ResetDayProperties()
    Dim ilLoop As Integer

    For ilLoop = 0 To NDAYS
        tmDayProperties(ilLoop).lBGColor = cmCalendarBGColor
        tmDayProperties(ilLoop).lFGColor = RGB(0, 0, 0)
        tmDayProperties(ilLoop).lSelectedBGColor = RGB(0, 0, 255)
        tmDayProperties(ilLoop).lSelectedFGColor = RGB(255, 255, 255)
        tmDayProperties(ilLoop).bCanSelect = True
    Next
End Sub

'****************************************************************************
'
'****************************************************************************
Public Sub SetDateProperties(dDate As Date, FGColor As Long, CanSelect As Boolean)
    Dim ilLoop As Integer
    Dim dlTodaysDate As Date
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilDay As Integer
    
    ilYear = year(gNow())
    ilMonth = month(gNow())
    ilDay = Day(gNow())
    dlTodaysDate = DateSerial(ilYear, ilMonth, ilDay)
    
    For ilLoop = 0 To NDAYS
        If tmDayProperties(ilLoop).dDate = dDate Then
            tmDayProperties(ilLoop).lFGColor = FGColor
            tmDayProperties(ilLoop).bCanSelect = CanSelect
        End If
        If tmDayProperties(ilLoop).dDate = dlTodaysDate Then
            tmDayProperties(ilLoop).lBGColor = cmCurrentDayBGColor
            tmDayProperties(ilLoop).lFGColor = cmCurrentDayFGColor
        End If
    Next
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub ec_InputBox_GotFocus()
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilDay As Integer

    If Trim(ec_InputBox.Text) = "" Then
        Select Case bmDefaultDateType
            Case csiNoDate
                ' Nothing to do here.
            Case csiCurrentDate
                ec_InputBox.Text = Format(gNow(), sgShowDateForm)
            Case csiThisMonday
                ec_InputBox.Text = mGetMondayForThisWeek(ec_InputBox.Text)
            Case csiNextMonday
                ec_InputBox.Text = mGetMondayForNextWeek(ec_InputBox.Text)
            Case csiStandardBroadcastMonday
        End Select
    End If

    Call SelectAndHighlightMainDate
End Sub

'****************************************************************************
'
'****************************************************************************
Private Function mDateIsOnCalendar(dDate As Date)
    Dim ilLoop As Integer

    mDateIsOnCalendar = False
    'Exit Function

    For ilLoop = 0 To NDAYS
        If tmDayProperties(ilLoop).dDate = dDate Then
            mDateIsOnCalendar = True
            Exit Function
        End If
    Next
End Function

'****************************************************************************
'
'****************************************************************************
Private Sub mSelectDay()
    Dim ilLoop As Integer
    Dim dlDate As Date
    
    If Trim(ec_InputBox.Text) = "" Then
        Exit Sub
    End If
    If Trim(smCalendarDate) = "" Then
        Exit Sub
    End If

    smCalendarDate = gAdjYear(smCalendarDate)
    If Not gIsDate(smCalendarDate) Then
        smCalendarDate = DateSerial(year(gNow()), month(gNow()), Day(gNow()))
    End If
    
    dlDate = gAdjYear(ec_InputBox.Text)
    'dlDate = smCalendarDate
    If Not mDateIsOnCalendar(dlDate) Then
        Exit Sub
    End If
    
    ' dlDate = smCalendarDate
    For ilLoop = 0 To NDAYS
        If tmDayProperties(ilLoop).dDate = dlDate Then
            Call SelectCalendarDay(ilLoop)
            'imLastSelectedDay = ilLoop
            'KeyPad_Buttons(ilLoop).BackColor = tmDayProperties(ilLoop).lSelectedBGColor
            'KeyPad_Buttons(ilLoop).ForeColor = tmDayProperties(ilLoop).lSelectedFGColor
        Else
            KeyPad_Buttons(ilLoop).BackColor = tmDayProperties(ilLoop).lBGColor
            KeyPad_Buttons(ilLoop).ForeColor = tmDayProperties(ilLoop).lFGColor
        End If
    Next
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub mUnSelectAllDays()
    Dim ilLoop As Integer

    For ilLoop = 0 To NDAYS
        KeyPad_Buttons(ilLoop).BackColor = tmDayProperties(ilLoop).lBGColor
        KeyPad_Buttons(ilLoop).ForeColor = tmDayProperties(ilLoop).lFGColor
    Next
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub mSetFirstAndLastDate(sDate As String)
    Dim dlFirstDay As Date
    Dim dlLastDay As Date
    Dim nDOW As VbDayOfWeek
    Dim slTempDate As String
    
    slTempDate = gAdjYear(sDate)
    dlFirstDay = DateSerial(year(slTempDate), month(slTempDate), 1)
    nDOW = Weekday(dlFirstDay)
    While nDOW <> vbMonday
        dlFirstDay = dlFirstDay - 1
        nDOW = Weekday(dlFirstDay)
    Wend
    dmFirstDate = dlFirstDay    ' This is the first monday to show on the calendar.
    ' Find the last sunday in the current month.
    dlLastDay = DateSerial(year(slTempDate), month(slTempDate) + 1, 1)
    dlLastDay = dlLastDay - 1
    nDOW = Weekday(dlLastDay)
    While nDOW <> vbSunday
        dlLastDay = dlLastDay - 1
        nDOW = Weekday(dlLastDay)
    Wend
    'LastDay = LastDay + 1
    dmLastDate = dlLastDay
End Sub

'****************************************************************************
'
'****************************************************************************
Function mGetMondayForThisWeek(sDate As String)
    Dim slThisDate As String
    
    slThisDate = Trim(sDate)
    If slThisDate = "" Then
        slThisDate = Format(gNow(), sgShowDateForm)
    End If
    While Weekday(slThisDate) <> vbMonday
        slThisDate = DateAdd("d", -1, slThisDate)
    Wend
    mGetMondayForThisWeek = slThisDate
End Function

'****************************************************************************
'
'****************************************************************************
Function mGetMondayForNextWeek(sDate As String)
    Dim slThisDate As String

    slThisDate = Trim(sDate)
    If slThisDate = "" Then
        slThisDate = Format(gNow(), sgShowDateForm)
    End If
    While Weekday(slThisDate) <> vbMonday
        slThisDate = DateAdd("d", 1, slThisDate)
    Wend
    mGetMondayForNextWeek = slThisDate
End Function

'****************************************************************************
'
' Broadcast Calendar
' Always starts on Monday and ends on Sunday.
' Example: 1-1-2004
' Backup to monday which may be in the previous month and year
' Show the last sunday of the current month and then quit.
' Provide a different color for days that are not on the current month.
' Must be a rule that only allows monday to be selected.
' Provide a function to select the Foreground, selectable
'
'****************************************************************************
Private Sub mSetCalendarDays(sDate As String)
    Dim ilLoop As Integer
    Dim ilDayNum As Integer
    Dim dlDayCounter As Date
    Dim blOkToAssignDay As Boolean

    If Trim(sDate) = "" Then
        Exit Sub
    End If

    On Error GoTo Err_mSetCalendarDays
    If bmFreezeCalendar Then
        Exit Sub
    End If

    ' JD 12-22-22 Force the date to be a US date.
    sDate = Format(sDate, "mm/dd/yyyy")

    smCalendarDate = gAdjYear(sDate)
    dmCurDate = smCalendarDate
    ' Back up to the first monday.
    If Not gIsDate(smCalendarDate) Then
        smCalendarDate = DateSerial(year(gNow()), month(gNow()), Day(gNow()))
    End If

    Call mSetFirstAndLastDate(smCalendarDate)
    'If bmIsDroppedDown Then
        If dmCurDate > dmLastDate Then
        ' This date is not on this calendar.
            dmCurDate = DateAdd("d", 15, dmCurDate)  ' Move forward 1 month.
            smCalendarDate = dmCurDate
            Call mSetFirstAndLastDate(smCalendarDate)
        End If
    'End If

    dmCurDate = smCalendarDate
    'sb_Year.Value = Year(dmCurDate)

    dlDayCounter = dmFirstDate
    blOkToAssignDay = True

    If Not mDateIsOnCalendar(dmCurDate) Then
        ' Do not draw the calendar if the date is already on it.
        ' This is support for the broadcast calendar.
        Call ResetDayProperties
        For ilLoop = 0 To NDAYS
            If blOkToAssignDay Then
                KeyPad_Buttons(ilLoop).Caption = Day(dlDayCounter)
                tmDayProperties(ilLoop).dDate = dlDayCounter
            Else
                KeyPad_Buttons(ilLoop).Caption = ""
                tmDayProperties(ilLoop).dDate = "1/1/1800"
            End If
            dlDayCounter = DateAdd("d", 1, dlDayCounter)
            If dlDayCounter > dmLastDate Then ' stop at last day
                blOkToAssignDay = False
                ' Exit For
            End If
        Next
        
        ' Tell the parent that the dates on the calendar have been updated and we
        ' need it to set the colors for this month.
        RaiseEvent CalendarChanged
        
        ' Now set the colors the parent has ordered.
        For ilLoop = 0 To NDAYS
             KeyPad_Buttons(ilLoop).BackColor = tmDayProperties(ilLoop).lBGColor
             KeyPad_Buttons(ilLoop).ForeColor = tmDayProperties(ilLoop).lFGColor
        Next
    End If
    Call SetMonthName
    Call mSelectDay
    Exit Sub
    
Err_mSetCalendarDays:
    gMsgBox "ERROR: mSetCalendarDays, Date = " & sDate
    Resume Next
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub SetMonthName()
    Dim ilMonth As Integer
    
    If Not gIsDate(smCalendarDate) Then
        smCalendarDate = DateSerial(year(gNow()), month(gNow()), Day(gNow()))
    End If
    ilMonth = GetMonth(smCalendarDate)
    ' dlDate = smCalendarDate
    
    Select Case imCalDateFormat
        Case LongMonth
            ec_MonthName.Caption = MonthName(ilMonth)
        Case ShortMonth
            ec_MonthName.Caption = Left(MonthName(ilMonth), 3)
    End Select
    ec_Year.Caption = str(GetYear(gAdjYear(smCalendarDate)))
    ' For some reason the very first time the user clicks on the year and executes this
    ' code, the year shifts to the left a pixel or so. So here we do it now so it will
    ' not have this effect when first clicked on.
    ec_Year.Caption = ec_Year.Caption + 1
    ec_Year.Caption = ec_Year.Caption - 1
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub ec_InputBox_Change()
    Dim ilLen As Integer

    ilLen = Len(ec_InputBox.Text)
    ' Minimum length is 6 for a date like 1/1/20
    'If ilLen < 6 Then
    '    Exit Sub
    'End If
    If Not gIsDate(ec_InputBox.Text) Then
        Call mUnSelectAllDays
        'gMsgBox "Change"
        RaiseEvent Change
        Exit Sub
    End If
'    If Mid(ec_InputBox.text, ilLen, 1) = "/" Then
'        ec_InputBox.text = ec_InputBox.text + ec_Year.Caption
'        ec_InputBox.SelStart = ilLen ' + 2
'        ec_InputBox.SelLength = 4
'        'ec_InputBox.Text = ec_InputBox.Text + Mid(ec_Year.Caption, 2, 2)
'        'ec_InputBox.SelStart = ilLen ' + 2
'        'ec_InputBox.SelLength = 2
'    End If
    smCalendarDate = ec_InputBox.Text
    Call mSetCalendarDays(smCalendarDate)
    'gMsgBox "Change"
    RaiseEvent Change
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub ec_InputBox_Click()
    If ec_InputBox.SelLength < 1 Then
        ' The user has clicked in the edit box again and does not have any
        ' digits selected.
        bmUserIsTyping = True
        Exit Sub
    End If
    If bmShowDropDownOnFocus Then
        ec_InputBox.SetFocus
        Exit Sub
    End If
    Call UserControl_ExitFocus
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub UserControl_EnterFocus()
    If Not ec_InputBox.Enabled Then
        Exit Sub
    End If
    If bmShowDropDownOnFocus Then
        Call btn_DownArrow_Click
        ec_InputBox.SetFocus
    End If
    bmUserIsTyping = False
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub ec_InputBox_LostFocus()
    If bmForceMondaySelectionOnly Then
        ec_InputBox.Text = tmDayProperties(imLastSelectedDay).dDate
    End If
    Call ValidateEnteredDate
End Sub

Private Function GetMonth(sDate As String) As Integer
    Dim ilPos As Integer
    Dim slTempDate As String
    Dim slMonth As String
    
    GetMonth = 1
    slTempDate = sDate
    slTempDate = Replace(slTempDate, "/", "-")
    ilPos = InStr(slTempDate, "-")
    If ilPos > 0 Then
        slMonth = Mid(slTempDate, 1, ilPos - 1)
        GetMonth = Val(slMonth)
    End If
End Function
Private Function GetDay(sDate As String) As Integer
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    Dim slTempDate As String
    Dim slDay As String
    
    GetDay = 1
    slTempDate = sDate
    slTempDate = Replace(slTempDate, "/", "-")
    ilPos1 = InStr(slTempDate, "-")
    If ilPos1 > 0 Then
        ' Find the next one
        ilPos2 = InStr(ilPos1 + 1, slTempDate, "-")
        If ilPos2 > 0 Then
            slDay = Mid(slTempDate, ilPos1 + 1, (ilPos2 - ilPos1) - 1)
            GetDay = Val(slDay)
        Else
            ' the year may not have been entered. If the length of the string is greater than
            ' what ilPos1 is then take that portion.
            If ilPos1 < Len(slTempDate) Then
                slDay = Mid(slTempDate, ilPos1 + 1, Len(slTempDate))
                GetDay = Val(slDay)
        End If
    End If
    End If
End Function
Private Function GetYear(sDate As String) As Integer
    Dim ilPos1 As Integer
    Dim ilPos2 As Integer
    Dim slTempDate As String
    Dim slYear As String

    GetYear = year(gNow())
    slTempDate = gAdjYear(sDate)
    slTempDate = Replace(slTempDate, "/", "-")
    ilPos1 = InStr(slTempDate, "-")
    If ilPos1 > 0 Then
        ' Find the next one
        ilPos2 = InStr(ilPos1 + 1, slTempDate, "-")
        If ilPos2 > 0 Then
            If ilPos2 < Len(slTempDate) Then
                slYear = Mid(slTempDate, ilPos2 + 1, Len(slTempDate))
                GetYear = Val(slYear)
            End If
        End If
    End If
End Function

Private Sub ValidateEnteredDate()
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilDay As Integer

    If gIsDate(ec_InputBox.Text) Then
        ilMonth = GetMonth(ec_InputBox.Text)
        ilDay = GetDay(ec_InputBox.Text)
        ilYear = GetYear(gAdjYear(ec_InputBox.Text))
        'smCalendarDate = DateSerial(ec_Year.Caption, ilMonth, ilDay)
        'ec_InputBox.text = DateSerial(ec_Year.Caption, ilMonth, ilDay)
        ec_InputBox.Text = DateSerial(ilYear, ilMonth, ilDay)
    End If
    If UCase(ec_InputBox.Text) = "TFN" Then
        If bmAllowTFN Then
            ec_InputBox.Text = "TFN"
            Exit Sub
        Else
            ec_InputBox.Text = Format(gNow(), sgShowDateForm)
        End If
    End If
    If ec_InputBox.Text = "" Then
        If Not bmAllowBlankDate Then
            ec_InputBox.Text = Format(gNow(), sgShowDateForm)
        End If
    End If
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub UserControl_ExitFocus()
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilDay As Integer
    
    On Error Resume Next
    If bmIsDroppedDown Then
        'Width = ec_InputBox.Width
        pbcCalFrame.Visible = False
        height = ec_InputBox.height
        'ec_InputBox.SetFocus
        bmIsDroppedDown = False
    End If

    Call ValidateEnteredDate
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub UserControl_Resize()
    If bmIgnoreResize Or bmIsDroppedDown Then Exit Sub
    
    ec_InputBox.Width = Width
    ec_InputBox.height = height

    If ec_InputBox.BorderStyle = 0 Then
        ' The edit control does not have a border. Make it as tall as the edit control.
        btn_DownArrow.height = height
        btn_DownArrow.Top = ec_InputBox.Top
        btn_DownArrow.Left = ec_InputBox.Width - btn_DownArrow.Width
    Else
        ' The edit area has a border. Adjust it so it looks correct.
        btn_DownArrow.height = ec_InputBox.height - (Screen.TwipsPerPixelX * 4)
        btn_DownArrow.Top = ec_InputBox.Top + (Screen.TwipsPerPixelY * 2)
        btn_DownArrow.Left = ec_InputBox.Width - btn_DownArrow.Width - (Screen.TwipsPerPixelX * 2)
    End If
    
    'Height = ec_InputBox.Height
    
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub btn_DownArrow_Click()
    If Not btn_DownArrow.Enabled Then
        Exit Sub
    End If
    If bmIsDroppedDown Then
        'Width = ec_InputBox.Width
        pbcCalFrame.Visible = False
        height = ec_InputBox.height
        ec_InputBox.SetFocus
        bmIsDroppedDown = False
        Exit Sub
    End If
    bmIsDroppedDown = True
    'ec_InputBox.Height = Height

    pbcCalFrame.Top = (ec_InputBox.Top + ec_InputBox.height)
    pbcCalFrame.Left = ScaleLeft

    If imEditBoxAlignment = 1 Then  ' Are we aligning the edit box on the right ?
        If pbcCalFrame.Left + pbcCalFrame.Width < ec_InputBox.Left + ec_InputBox.Width Then
            ' Reposition the drop down so it always lines up on the right side.
            pbcCalFrame.Left = (ec_InputBox.Left + ec_InputBox.Width) - pbcCalFrame.Width
        End If
    End If
    pbcCalFrame.Visible = True
    pbcCalFrame.SetFocus

    If Width < pbcCalFrame.Width Then
        Width = pbcCalFrame.Width
    End If
    height = imDropDownHeight
    Call SelectAndHighlightMainDate
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
    '4/8/20: If date blank, random date appears. TTP 9800
    If ec_InputBox.Text <> "" Then
        smCalendarDate = PropBag.ReadProperty("Text", "")
    End If
    ec_InputBox.BackColor = PropBag.ReadProperty("BackColor", RGB(255, 255, 255))
    ec_InputBox.ForeColor = PropBag.ReadProperty("ForeColor", RGB(0, 0, 0))
    ec_InputBox.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    bmShowDropDownOnFocus = PropBag.ReadProperty("CSI_ShowDropDownOnFocus", True)
    imEditBoxAlignment = PropBag.ReadProperty("CSI_InputBoxBoxAlignment", 0)
    cmCalendarBGColor = PropBag.ReadProperty("CSI_CalBackColor", RGB(170, 255, 255))
    cmCurrentDayBGColor = PropBag.ReadProperty("CSI_CurDayBackColor", RGB(255, 255, 255))
    cmCurrentDayFGColor = PropBag.ReadProperty("CSI_CurDayForeColor", RGB(0, 0, 0))
    bmForceMondaySelectionOnly = PropBag.ReadProperty("CSI_ForceMondaySelectionOnly", False)
    bmAllowBlankDate = PropBag.ReadProperty("CSI_AllowBlankDate", True)
    bmAllowTFN = PropBag.ReadProperty("CSI_AllowTFN", True)
    bmDefaultDateType = PropBag.ReadProperty("CSI_DefaultDateType", csiCurrentDate)
    
    bmCloseCalAfterSelection = PropBag.ReadProperty("CSI_CloseCalAfterSelection", True)
    imCalDateFormat = PropBag.ReadProperty("CSI_CalDateFormat", LongMonth)
    
    Set Font = PropBag.ReadProperty("Font", mFont)
    Set fmDayNameFont = PropBag.ReadProperty("CSI_DayNameFont", mFont)
    Set fmMonthNameFont = PropBag.ReadProperty("CSI_MonthNameFont", mFont)

    Set pbcDayNameFontHolder.Font = fmDayNameFont
    Call PositionAllControls
End Sub

'****************************************************************************
' Write property values to storage
'****************************************************************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", smCalendarDate, "")
    Call PropBag.WriteProperty("BackColor", ec_InputBox.BackColor, RGB(255, 255, 255))
    Call PropBag.WriteProperty("ForeColor", ec_InputBox.ForeColor, RGB(0, 0, 0))
    Call PropBag.WriteProperty("BorderStyle", ec_InputBox.BorderStyle)
    Call PropBag.WriteProperty("CSI_ShowDropDownOnFocus", bmShowDropDownOnFocus)
    Call PropBag.WriteProperty("CSI_InputBoxBoxAlignment", imEditBoxAlignment)
    Call PropBag.WriteProperty("CSI_CalBackColor", cmCalendarBGColor)
    Call PropBag.WriteProperty("CSI_CloseCalAfterSelection", bmCloseCalAfterSelection, True)
    Call PropBag.WriteProperty("CSI_CalDateFormat", imCalDateFormat, LongMonth)
    Call PropBag.WriteProperty("Font", mFont, Ambient.Font)
    Call PropBag.WriteProperty("CSI_DayNameFont", fmDayNameFont)
    Call PropBag.WriteProperty("CSI_MonthNameFont", fmMonthNameFont)
    Call PropBag.WriteProperty("CSI_CurDayBackColor", cmCurrentDayBGColor)
    Call PropBag.WriteProperty("CSI_CurDayForeColor", cmCurrentDayFGColor)
    Call PropBag.WriteProperty("CSI_ForceMondaySelectionOnly", bmForceMondaySelectionOnly)
    Call PropBag.WriteProperty("CSI_AllowBlankDate", bmAllowBlankDate)
    Call PropBag.WriteProperty("CSI_AllowTFN", bmAllowTFN)
    Call PropBag.WriteProperty("CSI_DefaultDateType", bmDefaultDateType)
    
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub UserControl_Terminate()
    Dim ilLoop As Integer
    For ilLoop = 1 To NDAYS
        Unload KeyPad_Buttons(ilLoop)
    Next
    For ilLoop = 1 To 6
        Unload lc_WeekNames(ilLoop)
    Next
End Sub

'****************************************************************************
'
'****************************************************************************
Public Property Get Text() As String
   Text = ec_InputBox.Text
End Property
Public Property Let Text(sText As String)
    Dim ilYear As Integer
    Dim ilMonth As Integer
    Dim ilDay As Integer
    Dim blSetText As Boolean
    
    On Error GoTo Err_Text
    ' gMsgBox "Setting Text"
    blSetText = True
    bmIgnoreYearChange = True
    smCalendarDate = Trim(sText)
    If UCase(smCalendarDate) = "TFN" And bmAllowTFN Then
        smCalendarDate = "TFN"
        ec_InputBox.Text = smCalendarDate
        Exit Property
    End If
    If smCalendarDate = "" And bmAllowBlankDate Then
        ec_InputBox.Text = smCalendarDate
        smCalendarDate = Format(gNow(), sgShowDateForm)
        blSetText = False
    End If
    
    If Not gIsDate(smCalendarDate) Then
        smCalendarDate = Format(gNow(), sgShowDateForm)
    End If
    ilYear = year(smCalendarDate)
    ilMonth = month(smCalendarDate)
    ilDay = Day(smCalendarDate)
    smCalendarDate = DateSerial(ilYear, ilMonth, ilDay)
    sb_Year.Value = ilYear
    If blSetText Then
        ec_InputBox.Text = smCalendarDate
    End If
    PropertyChanged "Text"
    Call mSetCalendarDays(smCalendarDate)
    bmIgnoreYearChange = False
    RaiseEvent Change
    Exit Property

Err_Text:
    smCalendarDate = "1/1/1990"
    sb_Year.Value = year(smCalendarDate)
    ec_InputBox.Text = smCalendarDate
    PropertyChanged "Text"
    Call mSetCalendarDays(smCalendarDate)
    bmIgnoreYearChange = False
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get BackColor() As OLE_COLOR
   BackColor = ec_InputBox.BackColor
End Property
Public Property Let BackColor(BKColor As OLE_COLOR)
    ec_InputBox.BackColor = BKColor
    Call PositionAllControls
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
    Call PositionAllControls
    PropertyChanged "ForeColor"
End Property
'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_InputBoxBoxAlignment() As Integer
   CSI_InputBoxBoxAlignment = imEditBoxAlignment
End Property
Public Property Let CSI_InputBoxBoxAlignment(Setting As Integer)
    imEditBoxAlignment = Setting
    PropertyChanged "CSI_InputBoxBoxAlignment"
End Property
'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_CloseCalAfterSelection() As Boolean
   CSI_CloseCalAfterSelection = bmCloseCalAfterSelection
End Property
Public Property Let CSI_CloseCalAfterSelection(Setting As Boolean)
    bmCloseCalAfterSelection = Setting
    PropertyChanged "CSI_CloseCalAfterSelection"
End Property

''****************************************************************************
''
''****************************************************************************
'Public Property Get CSI_RangeFGColor() As OLE_COLOR
'   CSI_RangeFGColor = imRangePickerFGColor
'End Property
'Public Property Let CSI_RangeFGColor(Setting As OLE_COLOR)
'    imRangePickerFGColor = Setting
'    Call PositionAllControls
'    PropertyChanged "CSI_RangeFGColor"
'End Property
'
''****************************************************************************
''
''****************************************************************************
'Public Property Get CSI_RangeBGColor() As OLE_COLOR
'   CSI_RangeBGColor = imRangePickerBGColor
'End Property
'Public Property Let CSI_RangeBGColor(Setting As OLE_COLOR)
'    imRangePickerBGColor = Setting
'    Call PositionAllControls
'    PropertyChanged "CSI_RangeBGColor"
'End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get BorderStyle() As BorderStyleConstants
   BorderStyle = ec_InputBox.BorderStyle
End Property
Public Property Let BorderStyle(BorderStyle As BorderStyleConstants)
    ec_InputBox.BorderStyle = BorderStyle
    Call PositionAllControls
    PropertyChanged "BorderStyle"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_ShowDropDownOnFocus() As Boolean
   CSI_ShowDropDownOnFocus = bmShowDropDownOnFocus
End Property
Public Property Let CSI_ShowDropDownOnFocus(Setting As Boolean)
    bmShowDropDownOnFocus = Setting
    Call PositionAllControls
    PropertyChanged "CSI_ShowDropDownOnFocus"
End Property

''****************************************************************************
''
''****************************************************************************
'Public Property Get CSI_HourOnColor() As OLE_COLOR
'   CSI_HourOnColor = imHourOnColor
'End Property
'Public Property Let CSI_HourOnColor(Setting As OLE_COLOR)
'    imHourOnColor = Setting
'    Call PositionAllControls
'    PropertyChanged "CSI_HourOnColor"
'End Property
'
''****************************************************************************
''
''****************************************************************************
'Public Property Get CSI_HourOffColor() As OLE_COLOR
'   CSI_HourOffColor = imHourOffColor
'End Property
'Public Property Let CSI_HourOffColor(Setting As OLE_COLOR)
'    imHourOffColor = Setting
'    Call PositionAllControls
'    PropertyChanged "CSI_HourOffColor"
'End Property

'****************************************************************************
' Day Name Fonts
'****************************************************************************
Public Property Get CSI_DayNameFont() As StdFont
    Set CSI_DayNameFont = fmDayNameFont
End Property
Public Property Set CSI_DayNameFont(ByVal New_Font As Font)
    Dim ilLoop As Integer

    fmDayNameFont = New_Font
    fmDayNameFont.Name = New_Font.Name
    fmDayNameFont.Size = New_Font.Size
    fmDayNameFont.Bold = New_Font.Bold
    fmDayNameFont.Italic = New_Font.Italic

    pbcDayNameFontHolder.Font = fmDayNameFont
    pbcDayNameFontHolder.FontName = fmDayNameFont.Name
    pbcDayNameFontHolder.FontSize = fmDayNameFont.Size
    pbcDayNameFontHolder.FontBold = fmDayNameFont.Bold
    pbcDayNameFontHolder.FontItalic = fmDayNameFont.Italic

    PropertyChanged "CSI_DayNameFont"
    Call PositionAllControls
End Property

'****************************************************************************
' Month Name Font
'****************************************************************************
Public Property Get CSI_MonthNameFont() As StdFont
    Set CSI_MonthNameFont = fmMonthNameFont
End Property
Public Property Set CSI_MonthNameFont(ByVal New_Font As Font)
    fmMonthNameFont = New_Font
    fmMonthNameFont.Name = New_Font.Name
    fmMonthNameFont.Size = New_Font.Size
    fmMonthNameFont.Bold = New_Font.Bold
    fmMonthNameFont.Italic = New_Font.Italic

    PropertyChanged "CSI_MonthNameFont"
    Call PositionAllControls
End Property

''****************************************************************************
''
''****************************************************************************
'Public Property Get Font() As StdFont
'   Set Font = mFont
'End Property
'Public Property Set Font(ByVal New_Font As Font)
'    Dim ilLoop As Integer
'
'    With mFont
'        .Bold = New_Font.Bold
'        .Italic = New_Font.Italic
'        .Name = New_Font.Name
'        .Size = New_Font.Size
'    End With
'    ec_InputBox.Font = New_Font
'    Set CSI_MonthNameFont = New_Font
'    Set CSI_DayNameFont = New_Font
'
'    PropertyChanged "Font"
'    Call PositionAllControls
'End Property
Private Sub mFont_FontChanged(ByVal PropertyName As String)
   ' Set UserControl.Font = mFont
   Set Font = mFont
End Sub

'****************************************************************************
'
'****************************************************************************
Public Property Get FontName() As String
    FontName = mFont.Name
End Property
Public Property Let FontName(sInFontName As String)
    Dim ilLoop As Integer

    mFont.Name = sInFontName
    UserControl.FontName = sInFontName
    ec_InputBox.FontName = sInFontName
    CSI_MonthNameFont.Name = sInFontName
    CSI_DayNameFont.Name = sInFontName
    mFont.Name = sInFontName
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
    'btn_DownArrow.FontSize = mFont.Size
    CSI_MonthNameFont.Size = mFont.Size
    CSI_DayNameFont.Size = mFont.Size
    mFont.Size = mFont.Size
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
    'btn_DownArrow.FontBold = mFont.Bold
    CSI_MonthNameFont.Bold = mFont.Bold
    CSI_DayNameFont.Bold = mFont.Bold
    mFont.Bold = mFont.Bold
    
    Call PositionAllControls
End Property

'****************************************************************************
'
'****************************************************************************
Public Function GetFirstDate()
    GetFirstDate = dmFirstDate
End Function
Public Function GetLastDate()
    GetLastDate = dmLastDate
End Function
Public Function GetCalendarDate() As Date
    ' Call mSetFirstAndLastDate(smCalendarDate)
    GetCalendarDate = smCalendarDate
End Function

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_CalDateFormat() As CalDateTypeList
   CSI_CalDateFormat = imCalDateFormat
End Property
Public Property Let CSI_CalDateFormat(Setting As CalDateTypeList)
    imCalDateFormat = Setting
    PropertyChanged "CSI_CalDateFormat"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_CalBackColor() As OLE_COLOR
   CSI_CalBackColor = cmCalendarBGColor
End Property
Public Property Let CSI_CalBackColor(Setting As OLE_COLOR)
    cmCalendarBGColor = Setting
    PropertyChanged "CSI_CalBackColor"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_CurDayBackColor() As OLE_COLOR
   CSI_CurDayBackColor = cmCurrentDayBGColor
End Property
Public Property Let CSI_CurDayBackColor(Setting As OLE_COLOR)
    cmCurrentDayBGColor = Setting
    PropertyChanged "CSI_CurDayBackColor"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_CurDayForeColor() As OLE_COLOR
   CSI_CurDayForeColor = cmCurrentDayFGColor
End Property
Public Property Let CSI_CurDayForeColor(Setting As OLE_COLOR)
    cmCurrentDayFGColor = Setting
    PropertyChanged "CSI_CurDayForeColor"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_ForceMondaySelectionOnly() As Boolean
   CSI_ForceMondaySelectionOnly = bmForceMondaySelectionOnly
End Property
Public Property Let CSI_ForceMondaySelectionOnly(Setting As Boolean)
    bmForceMondaySelectionOnly = Setting
    PropertyChanged "CSI_ForceMondaySelectionOnly"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_AllowBlankDate() As Boolean
   CSI_AllowBlankDate = bmAllowBlankDate
End Property
Public Property Let CSI_AllowBlankDate(Setting As Boolean)
    bmAllowBlankDate = Setting
    PropertyChanged "CSI_AllowBlankDate"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_AllowTFN() As Boolean
   CSI_AllowTFN = bmAllowTFN
End Property
Public Property Let CSI_AllowTFN(Setting As Boolean)
    bmAllowTFN = Setting
    PropertyChanged "CSI_AllowTFN"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_DefaultDateType() As CalDefaultDateList
   CSI_DefaultDateType = bmDefaultDateType
End Property
Public Property Let CSI_DefaultDateType(Setting As CalDefaultDateList)
    bmDefaultDateType = Setting
    PropertyChanged "CSI_DefaultDateType"
End Property

Public Sub SetEnabled(State As Boolean)
    btn_DownArrow.Enabled = State
    ec_InputBox.Enabled = State
End Sub

Public Property Get GetEnabled() As Boolean
   GetEnabled = btn_DownArrow.Enabled
End Property

