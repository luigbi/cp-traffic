VERSION 5.00
Begin VB.UserControl CSI_HourPicker 
   BackStyle       =   0  'Transparent
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12210
   ScaleHeight     =   8730
   ScaleWidth      =   12210
   ToolboxBitmap   =   "CSI_HourPicker.ctx":0000
   Begin VB.PictureBox pb_DropDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   615
      ScaleHeight     =   1125
      ScaleWidth      =   5775
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1245
      Width           =   5775
      Begin VB.Label btn_Clear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clr"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4980
         TabIndex        =   33
         Top             =   855
         Width           =   690
      End
      Begin VB.Label btn_7to23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7-12"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4170
         TabIndex        =   32
         Top             =   885
         Width           =   690
      End
      Begin VB.Label btn_3to6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3-7"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3300
         TabIndex        =   31
         Top             =   855
         Width           =   690
      End
      Begin VB.Label btn_10to2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10-3"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2505
         TabIndex        =   30
         Top             =   855
         Width           =   690
      End
      Begin VB.Label btn_0to5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12-6"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   825
         TabIndex        =   4
         Top             =   870
         Width           =   675
      End
      Begin VB.Label btn_6to9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6-10"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1635
         TabIndex        =   3
         Top             =   870
         Width           =   705
      End
      Begin VB.Label btn_All 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "All"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   0
         TabIndex        =   2
         Top             =   870
         Width           =   720
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
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
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "23"
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
         Index           =   23
         Left            =   5280
         TabIndex        =   28
         Top             =   435
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "22"
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
         Index           =   22
         Left            =   4800
         TabIndex        =   27
         Top             =   435
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "21"
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
         Index           =   21
         Left            =   4320
         TabIndex        =   26
         Top             =   435
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
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
         Index           =   20
         Left            =   3840
         TabIndex        =   25
         Top             =   435
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "19"
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
         Index           =   19
         Left            =   3360
         TabIndex        =   24
         Top             =   435
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "18"
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
         Index           =   18
         Left            =   2880
         TabIndex        =   23
         Top             =   435
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "17"
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
         Index           =   17
         Left            =   2400
         TabIndex        =   22
         Top             =   435
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "16"
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
         Index           =   16
         Left            =   1920
         TabIndex        =   21
         Top             =   435
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15"
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
         Index           =   15
         Left            =   1440
         TabIndex        =   20
         Top             =   435
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "14"
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
         Index           =   14
         Left            =   960
         TabIndex        =   19
         Top             =   435
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "13"
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
         Index           =   13
         Left            =   480
         TabIndex        =   18
         Top             =   435
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
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
         Index           =   12
         Left            =   0
         TabIndex        =   17
         Top             =   435
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11"
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
         Index           =   11
         Left            =   5280
         TabIndex        =   16
         Top             =   0
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
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
         Index           =   10
         Left            =   4800
         TabIndex        =   15
         Top             =   0
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "09"
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
         Index           =   9
         Left            =   4320
         TabIndex        =   14
         Top             =   0
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "08"
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
         Index           =   8
         Left            =   3840
         TabIndex        =   13
         Top             =   0
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "07"
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
         Index           =   7
         Left            =   3360
         TabIndex        =   12
         Top             =   0
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "06"
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
         Left            =   2880
         TabIndex        =   11
         Top             =   0
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "05"
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
         Left            =   2400
         TabIndex        =   10
         Top             =   0
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "04"
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
         Left            =   1920
         TabIndex        =   9
         Top             =   0
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "03"
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
         Left            =   1440
         TabIndex        =   8
         Top             =   0
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "02"
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
         Left            =   960
         TabIndex        =   7
         Top             =   0
         Width           =   480
      End
      Begin VB.Label KeyPad_Buttons 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
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
         Left            =   480
         TabIndex        =   6
         Top             =   0
         Width           =   480
      End
   End
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
      Left            =   2535
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   570
      UseMaskColor    =   -1  'True
      Width           =   300
   End
   Begin VB.TextBox ec_InputBox 
      Height          =   315
      Left            =   525
      TabIndex        =   5
      Top             =   540
      Width           =   1710
   End
End
Attribute VB_Name = "CSI_HourPicker"
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
Private bmShowDayPartButtons As Boolean
Private bmShowDropDownOnFocus As Boolean
Private bmIgnoreResize As Boolean

Private bmIsDroppedDown As Boolean
Private imDropDownWidth As Integer
Private imDropDownHeight As Integer
Private imEditBoxAlignment As Integer        ' 0=Edit box is on the Left, 1=Edit box is on the right
Private imIgnoreChangeEvent As Boolean
Private imHourOnColor As OLE_COLOR
Private imHourOffColor As OLE_COLOR
Private imRangePickerFGColor As OLE_COLOR
Private imRangePickerBGColor As OLE_COLOR
Private smAllowedCharacters As String
Private smLastCharEntered As String
Private imCurrentButtonAnchor As Integer

Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
' Highlight the entire field when focus

Private Type HourStruct
    IsSlected As Boolean
    text As String
End Type
Private HoursArray(0 To 23) As HourStruct


Private Sub ec_InputBox_GotFocus()
    ec_InputBox.SelStart = 0
    ec_InputBox.SelLength = Len(ec_InputBox.text)
End Sub

Private Sub UserControl_Initialize()
    Dim iLoop As Integer
    
    Set mFont = New StdFont
    Set UserControl.Font = mFont
    
    bmIsDroppedDown = False
    ec_InputBox.Top = 0
    ec_InputBox.Left = 0
    
    imDropDownWidth = pb_DropDown.Width
    imDropDownHeight = ec_InputBox.Height + pb_DropDown.Height
    For iLoop = 0 To 23
        HoursArray(iLoop).IsSlected = False
        HoursArray(iLoop).text = iLoop
        If Len(HoursArray(iLoop).text) < 2 Then
            HoursArray(iLoop).text = "0" + HoursArray(iLoop).text
        End If
    Next
    KeyPreview = True
    If imHourOnColor < 1 Then
        imHourOnColor = RGB(70, 200, 70)
        ' Set all other pre-defined values here as well.
        imRangePickerFGColor = RGB(0, 0, 0)
        imRangePickerBGColor = vbButtonFace
        bmShowSelectRangeButtons = True
        bmShowDayPartButtons = True
        bmShowDropDownOnFocus = True
    End If
    If imHourOffColor < 1 Then
        imHourOffColor = vbButtonFace
    End If
    smAllowedCharacters = "0123456789"
    smLastCharEntered = ""
    bmAllowMultiSelection = True
    bmIgnoreResize = False
    imCurrentButtonAnchor = -1
End Sub

Private Sub UserControl_InitProperties()
    Set Font = Ambient.Font
End Sub

Private Sub SetButtonStates(iStart As Integer, iEnd As Integer, bStatus As Boolean)
    Dim iLoop As Integer
    Dim blAllAreSelected As Boolean
    
    blAllAreSelected = True
    For iLoop = iStart To iEnd
        If Not HoursArray(iLoop).IsSlected Then
            blAllAreSelected = False
            Exit For
        End If
    Next
    If blAllAreSelected Then
        ' All states are set to on. Set the back off regardless of the status being passed in.
        bStatus = False
    End If
    For iLoop = iStart To iEnd
        HoursArray(iLoop).IsSlected = bStatus
    Next
    
    Call FillInTextWithArrayValues
End Sub

Private Sub btn_All_Click()
    Call SetButtonStates(0, 23, True)
End Sub
Private Sub btn_0to5_Click()
    Call SetButtonStates(0, 5, True)
End Sub
Private Sub btn_6to9_Click()
    Call SetButtonStates(6, 9, True)
End Sub
Private Sub btn_10to2_Click()
    Call SetButtonStates(10, 14, True)
End Sub
Private Sub btn_3to6_Click()
    Call SetButtonStates(15, 18, True)
End Sub
Private Sub btn_7to23_Click()
    Call SetButtonStates(19, 23, True)
End Sub
Private Sub btn_Clear_Click()
    Call SetButtonStates(0, 23, False)
End Sub

Private Sub ec_InputBox_Change()
    Call ShowHoursBasedOnHourSettings
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

Private Function GetDigitsValueFromHere(sText As String, ByRef iPos As Integer) As Integer
    Dim ilLen As Integer
    Dim slOneChar As String
    Dim slCurDigitText As String

    GetDigitsValueFromHere = 0
    slCurDigitText = ""
    ilLen = Len(sText)
    While iPos <= ilLen
        slOneChar = Mid(sText, iPos, 1)
        If slOneChar = "," Or slOneChar = "-" Then
            GetDigitsValueFromHere = Val(slCurDigitText)
            iPos = iPos - 1
            Exit Function
        End If
        iPos = iPos + 1
        slCurDigitText = slCurDigitText + slOneChar
    Wend
    GetDigitsValueFromHere = Val(slCurDigitText)
End Function

Private Sub SelectAllPreviousHoursFromHere(iPos As Integer)
    Dim iLoop As Integer

    For iLoop = iPos - 1 To 1 Step -1
        If HoursArray(iLoop).IsSlected Then
            Exit Sub
        End If
        HoursArray(iLoop).IsSlected = True
    Next
End Sub

Private Sub UnselectAllButtons()
    Dim ilLoop As Integer
    
    For ilLoop = 0 To 23
        HoursArray(ilLoop).IsSlected = False
    Next
End Sub

Private Function GetTotalSelected() As Integer
    Dim ilLoop As Integer
    
    GetTotalSelected = 0
    For ilLoop = 0 To 23
        If HoursArray(ilLoop).IsSlected Then
            GetTotalSelected = GetTotalSelected + 1
        End If
    Next
End Function

Private Sub ShowHoursBasedOnHourSettings()
    Dim iLoop As Integer
    Dim iLoop2 As Integer
    Dim ilLen As Integer
    Dim slOneChar As String
    Dim slCurValue As String
    Dim ilCurValue As Integer
    Dim blRangeIsSelected As Boolean

    If imIgnoreChangeEvent Then Exit Sub
    ' Prepare to rebuild the enabled list depending on whats selected.
    For iLoop = 0 To 23
        HoursArray(iLoop).IsSlected = False
    Next

    ' Parse the input string and set each days button values
    blRangeIsSelected = False
    ilLen = Len(ec_InputBox.text)
    If ilLen = 1 Then
        If UCase(ec_InputBox.text) = "A" Then
            ec_InputBox.text = "00-23"
        End If
    End If
    slCurValue = ""
    For iLoop = 1 To ilLen
        slOneChar = Mid(ec_InputBox.text, iLoop, 1)
        Select Case slOneChar
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                ilCurValue = GetDigitsValueFromHere(ec_InputBox.text, iLoop)
                If ilCurValue > 23 Then
                    Exit Sub
                End If
                HoursArray(ilCurValue).IsSlected = True
                If blRangeIsSelected Then
                    ' Set all hours from the last hour entered.
                    Call SelectAllPreviousHoursFromHere(ilCurValue)
                    blRangeIsSelected = False
                End If
            Case ","
                'ilCurValue = Val(slCurValue)
                'HoursArray(ilCurValue).IsSlected = True
                slCurValue = ""
            Case "-"
                blRangeIsSelected = True
                'ilCurValue = Val(slCurValue)
                'HoursArray(ilCurValue).IsSlected = True
                slCurValue = ""
        End Select
    Next
    If Len(slCurValue) > 0 Then
        ilCurValue = Val(slCurValue)
        If ilCurValue < 23 Then
            HoursArray(ilCurValue).IsSlected = True
        End If
    End If
    Call SetButtonColors
End Sub

' This function returns the number of digits entered counting back from Len(sText)
Private Function PreviousDigitCount(iCurPos As Integer, sText As String)
    Dim iLen As Integer
    Dim iLoop As Integer
    Dim slOneChar As String

    PreviousDigitCount = 0
    iLen = Len(sText)
    For iLoop = iCurPos To 1 Step -1
        slOneChar = Mid(sText, iLoop, 1)
        If slOneChar = "," Or slOneChar = "-" Then
            Exit Function
        End If
        If InStr(1, "0123456789", slOneChar) <> 0 Then
            PreviousDigitCount = PreviousDigitCount + 1
        End If
    Next
End Function

Private Function ValueOfCurrentDigits(iCurPos As Integer, sText As String)
    Dim iLen As Integer
    Dim iLoop As Integer
    Dim slOneChar As String
    Dim slCurDigitText As String
    
    ValueOfCurrentDigits = 0
    slCurDigitText = ""
    iLen = Len(sText)
    For iLoop = iCurPos To iLen
        slOneChar = Mid(sText, iLoop, 1)
        If slOneChar = "," Or slOneChar = "-" Then
            ValueOfCurrentDigits = Val(slCurDigitText)
            Exit Function
        End If
        slCurDigitText = slCurDigitText + slOneChar
    Next
    ValueOfCurrentDigits = Val(slCurDigitText)
End Function

Private Sub ec_InputBox_KeyPress(KeyAscii As Integer)
    Dim CurChar As String
    Dim iLoop As Integer
    Dim ilLen As Integer
    Dim ilCurPos As Integer
    Dim ilTotalDigits As Integer
    Dim slOneChar As String
    Dim ilTotalSelected As Integer

    If KeyAscii < 32 Then
        ' Allow all control keys to be processed as normal.
        Exit Sub
    End If

    CurChar = Chr(KeyAscii)

    ilLen = ec_InputBox.SelLength
    ilCurPos = ec_InputBox.SelStart
    If ilCurPos < 1 Then
        ilCurPos = 1
    End If
    If ilLen = Len(ec_InputBox.text) Then
        ilTotalDigits = 0
        ec_InputBox.text = ""
    Else
        ilTotalDigits = PreviousDigitCount(ilCurPos, ec_InputBox.text)
    End If

    Select Case ilTotalDigits
        Case 0
            smAllowedCharacters = "A0123456789"
        Case 1
            smAllowedCharacters = ",-0123456789"
        Case 2
            smAllowedCharacters = ",-"
    End Select

    CurChar = UCase(CurChar)
    If Not bmAllowMultiSelection Then
        If ec_InputBox.SelStart = 0 And ec_InputBox.SelLength = Len(ec_InputBox.text) Then
            ' Replace whatever is selected now with the current character entered if
            ' the entire text is selected.
            Call UnselectAllButtons
        End If
        ilTotalSelected = GetTotalSelected()
        If ilTotalSelected = 1 Then
            Select Case ec_InputBox.text
                Case "0"
                    Call UnselectAllButtons
                    ilTotalSelected = 0 ' Allow the next key to go through
                Case "1"
                    Call UnselectAllButtons
                    ilTotalSelected = 0 ' Allow the next key to go through
                Case "2"
                    Call UnselectAllButtons
                    If CurChar >= "0" And CurChar < "4" Then
                        ilTotalSelected = 0 ' Allow the next key to go through
                    End If
            End Select
        End If
        If ilTotalSelected > 0 Then
            Beep
            KeyAscii = 0 ' Instruct VB to ignore this key.
            Exit Sub
        End If
    End If

    If InStr(1, smAllowedCharacters, CurChar) = 0 Then
        Beep
        KeyAscii = 0 ' Ignore this key.
        Exit Sub
    End If

    If ValueOfCurrentDigits(ilCurPos, ec_InputBox.text + CurChar) > 23 Then
        Beep
        KeyAscii = 0 ' Ignore this key.
    End If
End Sub

Private Sub ec_InputBox_LostFocus()
    Dim iLoop As Integer

    ' Verify the string the user has entered.
    Select Case ec_InputBox.text
        Case "a", "A"
            ec_InputBox.text = "0-23"
    End Select
    Call FillInTextWithArrayValues
    'ec_InputBox.SetFocus
End Sub

Private Sub SetButtonColors()
    Dim iLoop As Integer
    
    For iLoop = 0 To 23
        If HoursArray(iLoop).IsSlected Then
            KeyPad_Buttons(iLoop).BackColor = imHourOnColor
        Else
            KeyPad_Buttons(iLoop).BackColor = imHourOffColor
        End If
    Next
End Sub

Private Sub SelectAllButtonsBetweenCurrentAnchor(iClickedButtonIndex As Integer)
    Dim ilLoop As Integer

    If imCurrentButtonAnchor > iClickedButtonIndex Then
        For ilLoop = iClickedButtonIndex To imCurrentButtonAnchor
            HoursArray(ilLoop).IsSlected = True
            KeyPad_Buttons(ilLoop).BackColor = imHourOnColor
        Next
    Else
        For ilLoop = imCurrentButtonAnchor To iClickedButtonIndex
            HoursArray(ilLoop).IsSlected = True
            KeyPad_Buttons(ilLoop).BackColor = imHourOnColor
        Next
    End If
End Sub

Private Sub KeyPad_Buttons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not bmAllowMultiSelection Then
        ' Anytime a hour is selected and MultiSelection is turned off, clear all others first.
        Call UnselectAllButtons
    End If
    If HoursArray(Index).IsSlected Then
        HoursArray(Index).IsSlected = False
        KeyPad_Buttons(Index).BackColor = imHourOffColor
    Else
        HoursArray(Index).IsSlected = True
        KeyPad_Buttons(Index).BackColor = imHourOnColor
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
    Dim iLoop As Integer
    
    GetCountSelectedFromHere = 0
    For iLoop = Index To MaxCount
        If HoursArray(iLoop).IsSlected Then
            GetCountSelectedFromHere = GetCountSelectedFromHere + 1
        Else
            Exit Function
        End If
    Next
End Function

Private Sub FillInTextWithArrayValues()
    Dim iLoop As Integer
    Dim slValue As String
    Dim sldelimiter As String
    Dim TotalSelectedFromHere As Integer
    
    slValue = ""
    For iLoop = 0 To 23
        If HoursArray(iLoop).IsSlected Then
            TotalSelectedFromHere = GetCountSelectedFromHere(iLoop, 23)
            If TotalSelectedFromHere >= 3 Then
                slValue = slValue + HoursArray(iLoop).text + "-"
                iLoop = iLoop + TotalSelectedFromHere - 2
            Else
                slValue = slValue + HoursArray(iLoop).text + ","
            End If
        End If
    Next
    ' Check for and remove the final comma if it exists.
    If Len(slValue) > 0 Then
        If Right(slValue, 1) = "," Then
            slValue = Left(slValue, Len(slValue) - 1)
        End If
    End If
    imIgnoreChangeEvent = True
    ec_InputBox.text = slValue
    imIgnoreChangeEvent = False
    For iLoop = 0 To 23
        If HoursArray(iLoop).IsSlected Then
            KeyPad_Buttons(iLoop).BackColor = imHourOnColor
        Else
            KeyPad_Buttons(iLoop).BackColor = imHourOffColor
        End If
    Next
End Sub

Private Sub UserControl_ExitFocus()
    On Error Resume Next
    If bmIsDroppedDown Then
        'Width = ec_InputBox.Width
        pb_DropDown.Visible = False
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
        btn_DownArrow.Height = ec_InputBox.Height - (Screen.TwipsPerPixelX * 3)
        btn_DownArrow.Top = ec_InputBox.Top + (Screen.TwipsPerPixelY * 2)
        btn_DownArrow.Left = ec_InputBox.Width - btn_DownArrow.Width - (Screen.TwipsPerPixelX * 2)
    End If
    
    'Height = ec_InputBox.Height
    
End Sub

Private Sub btn_DownArrow_Click()
    If bmIsDroppedDown Then
        'Width = ec_InputBox.Width
        pb_DropDown.Visible = False
        Height = ec_InputBox.Height
        ec_InputBox.SetFocus
        bmIsDroppedDown = False
        Exit Sub
    End If
    bmIsDroppedDown = True

    If Not bmShowSelectRangeButtons Or Not bmAllowMultiSelection Then
        pb_DropDown.Height = KeyPad_Buttons(0).Height + KeyPad_Buttons(0).Height ' + (Screen.TwipsPerPixelY * 2)
    End If

    pb_DropDown.Top = (ec_InputBox.Top + ec_InputBox.Height)
    pb_DropDown.Left = ScaleLeft
    If imEditBoxAlignment = 1 Then  ' Are we aligning the edit box on the right ?
        If pb_DropDown.Left + pb_DropDown.Width < ec_InputBox.Left + ec_InputBox.Width Then
            ' Reposition the drop down so it always lines up on the right side.
            pb_DropDown.Left = (ec_InputBox.Left + ec_InputBox.Width) - pb_DropDown.Width
        End If
    End If
    pb_DropDown.Visible = True
    pb_DropDown.SetFocus

    If Width < pb_DropDown.Width Then
        Width = pb_DropDown.Width
    End If
    Height = imDropDownHeight + Screen.TwipsPerPixelY + 2000
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
    iPixelWidth = TextWidth("00")
    iPixelHeight = TextHeight("00")
    ilButtonWidths = iPixelWidth + (Screen.TwipsPerPixelX * 4)
    ilButtonHeights = iPixelHeight + (Screen.TwipsPerPixelY * 2)
    For ilLoop = 0 To 11
        KeyPad_Buttons(ilLoop).Top = 0
        KeyPad_Buttons(ilLoop).Left = CurLeft
        KeyPad_Buttons(ilLoop).Width = ilButtonWidths
        KeyPad_Buttons(ilLoop).Height = ilButtonHeights
        CurLeft = CurLeft + ilButtonWidths - (Screen.TwipsPerPixelX) ' - (Screen.TwipsPerPixelX) removes the double border effect.
    Next
    CurLeft = 0
    For ilLoop = 12 To 23
        KeyPad_Buttons(ilLoop).Top = KeyPad_Buttons(0).Height - (Screen.TwipsPerPixelY) ' - (Screen.TwipsPerPixelX) removes the double border effect.
        KeyPad_Buttons(ilLoop).Left = CurLeft
        KeyPad_Buttons(ilLoop).Width = ilButtonWidths
        KeyPad_Buttons(ilLoop).Height = ilButtonHeights
        CurLeft = CurLeft + ilButtonWidths - (Screen.TwipsPerPixelX) ' - (Screen.TwipsPerPixelX) removes the double border effect.
    Next
    ilTotalWidth = CurLeft
    If bmShowDayPartButtons Then
        btn_All.Top = KeyPad_Buttons(0).Height + KeyPad_Buttons(13).Height - (Screen.TwipsPerPixelY * 2) ' - (Screen.TwipsPerPixelX) removes the double border effect.
        btn_0to5.Top = KeyPad_Buttons(0).Height + KeyPad_Buttons(13).Height - (Screen.TwipsPerPixelY * 2) ' - (Screen.TwipsPerPixelX) removes the double border effect.
        btn_6to9.Top = KeyPad_Buttons(0).Height + KeyPad_Buttons(13).Height - (Screen.TwipsPerPixelY * 2) ' - (Screen.TwipsPerPixelX) removes the double border effect.
        btn_10to2.Top = KeyPad_Buttons(0).Height + KeyPad_Buttons(13).Height - (Screen.TwipsPerPixelY * 2) ' - (Screen.TwipsPerPixelX) removes the double border effect.
        btn_3to6.Top = KeyPad_Buttons(0).Height + KeyPad_Buttons(13).Height - (Screen.TwipsPerPixelY * 2) ' - (Screen.TwipsPerPixelX) removes the double border effect.
        btn_7to23.Top = KeyPad_Buttons(0).Height + KeyPad_Buttons(13).Height - (Screen.TwipsPerPixelY * 2) ' - (Screen.TwipsPerPixelX) removes the double border effect.
        btn_Clear.Top = KeyPad_Buttons(0).Height + KeyPad_Buttons(13).Height - (Screen.TwipsPerPixelY * 2) ' - (Screen.TwipsPerPixelX) removes the double border effect.
        
        btn_All.Height = ilButtonHeights
        btn_0to5.Height = ilButtonHeights
        btn_6to9.Height = ilButtonHeights
        btn_10to2.Height = ilButtonHeights
        btn_3to6.Height = ilButtonHeights
        btn_7to23.Height = ilButtonHeights
        btn_Clear.Height = ilButtonHeights
        
        btn_All.Width = (ilTotalWidth / 7)
        btn_0to5.Width = (ilTotalWidth / 7) + (Screen.TwipsPerPixelX * 2)
        btn_6to9.Width = (ilTotalWidth / 7) + (Screen.TwipsPerPixelX * 2)
        btn_10to2.Width = (ilTotalWidth / 7)
        btn_3to6.Width = (ilTotalWidth / 7)
        btn_7to23.Width = (ilTotalWidth / 7)
        btn_Clear.Width = (ilTotalWidth / 7) - (Screen.TwipsPerPixelX * 2)
        
        btn_All.Left = 0
        btn_0to5.Left = btn_All.Width - Screen.TwipsPerPixelX
        btn_6to9.Left = (btn_0to5.Left + btn_0to5.Width) - (Screen.TwipsPerPixelX * 2)
        btn_10to2.Left = btn_All.Width + btn_0to5.Width + btn_6to9.Width - (Screen.TwipsPerPixelX * 4)
        btn_3to6.Left = btn_All.Width + btn_0to5.Width + btn_6to9.Width + btn_10to2.Width - (Screen.TwipsPerPixelX * 5)
        btn_7to23.Left = btn_All.Width + btn_0to5.Width + btn_6to9.Width + btn_10to2.Width + btn_3to6.Width - (Screen.TwipsPerPixelX * 6)
        btn_Clear.Left = btn_All.Width + btn_0to5.Width + btn_6to9.Width + btn_10to2.Width + btn_3to6.Width + btn_7to23.Width - (Screen.TwipsPerPixelX * 7)
        btn_Clear.Width = ilTotalWidth - btn_Clear.Left + (Screen.TwipsPerPixelX)
    Else
        ilTotalWidth = ilTotalWidth - (Screen.TwipsPerPixelX)
        btn_All.Top = KeyPad_Buttons(0).Height + KeyPad_Buttons(13).Height - (Screen.TwipsPerPixelY * 2) ' - (Screen.TwipsPerPixelX) removes the double border effect.
        btn_Clear.Top = KeyPad_Buttons(0).Height + KeyPad_Buttons(13).Height - (Screen.TwipsPerPixelY * 2) ' - (Screen.TwipsPerPixelX) removes the double border effect.
        
        btn_All.Height = ilButtonHeights
        btn_Clear.Height = ilButtonHeights
        
        btn_All.Width = (ilTotalWidth / 2) + (Screen.TwipsPerPixelX)
        btn_Clear.Width = (ilTotalWidth / 2)
        
        btn_All.Left = 0
        btn_Clear.Left = btn_All.Width - (Screen.TwipsPerPixelX)
        btn_Clear.Width = (ilTotalWidth - btn_Clear.Left) + (Screen.TwipsPerPixelX * 2)
    End If
    
    ilTotalHeight = KeyPad_Buttons(0).Height + KeyPad_Buttons(13).Height + btn_All.Height - (Screen.TwipsPerPixelY * 4) ' - (Screen.TwipsPerPixelX) removes the double border effect.
    pb_DropDown.Width = ilTotalWidth + (Screen.TwipsPerPixelX * 2)
    pb_DropDown.Height = ilTotalHeight + (Screen.TwipsPerPixelY * 2)
    
    imDropDownWidth = pb_DropDown.Width
    imDropDownHeight = ec_InputBox.Height + pb_DropDown.Height
    bmIgnoreResize = False
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
    btn_0to5.ForeColor = imRangePickerFGColor
    btn_0to5.BackColor = imRangePickerBGColor
    btn_6to9.ForeColor = imRangePickerFGColor
    btn_6to9.BackColor = imRangePickerBGColor
    btn_10to2.ForeColor = imRangePickerFGColor
    btn_10to2.BackColor = imRangePickerBGColor
    btn_3to6.ForeColor = imRangePickerFGColor
    btn_3to6.BackColor = imRangePickerBGColor
    btn_7to23.ForeColor = imRangePickerFGColor
    btn_7to23.BackColor = imRangePickerBGColor
    btn_Clear.ForeColor = imRangePickerFGColor
    btn_Clear.BackColor = imRangePickerBGColor
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
    bmShowDayPartButtons = PropBag.ReadProperty("CSI_ShowDayPartButtons", True)
    bmShowDropDownOnFocus = PropBag.ReadProperty("CSI_ShowDropDownOnFocus", True)
    imEditBoxAlignment = PropBag.ReadProperty("CSI_InputBoxBoxAlignment", 0)
    imHourOnColor = PropBag.ReadProperty("CSI_HourOnColor", RGB(70, 200, 70))
    imHourOffColor = PropBag.ReadProperty("CSI_HourOffColor", vbButtonFace) ' &H8000000F)
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
    Call PropBag.WriteProperty("CSI_ShowDayPartButtons", bmShowDayPartButtons)
    Call PropBag.WriteProperty("CSI_ShowDropDownOnFocus", bmShowDropDownOnFocus)
    Call PropBag.WriteProperty("CSI_InputBoxBoxAlignment", imEditBoxAlignment)
    Call PropBag.WriteProperty("CSI_HourOnColor", imHourOnColor)
    Call PropBag.WriteProperty("CSI_HourOffColor", imHourOffColor)
    Call PropBag.WriteProperty("CSI_RangeFGColor", imRangePickerFGColor)
    Call PropBag.WriteProperty("CSI_RangeBGColor", imRangePickerBGColor)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
End Sub

'****************************************************************************
'
'****************************************************************************
Public Property Get text() As String
   text = ec_InputBox.text
End Property
Public Property Let text(sText As String)
    smText = sText
    ec_InputBox.text = smText
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
Public Property Get CSI_ShowDayPartButtons() As Boolean
   CSI_ShowDayPartButtons = bmShowDayPartButtons
End Property
Public Property Let CSI_ShowDayPartButtons(Setting As Boolean)
    bmShowDayPartButtons = Setting
    Call AssignControlProperties
    PropertyChanged "CSI_ShowDayPartButtons"
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
Public Property Get CSI_HourOnColor() As OLE_COLOR
   CSI_HourOnColor = imHourOnColor
End Property
Public Property Let CSI_HourOnColor(Setting As OLE_COLOR)
    imHourOnColor = Setting
    Call AssignControlProperties
    PropertyChanged "CSI_HourOnColor"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get CSI_HourOffColor() As OLE_COLOR
   CSI_HourOffColor = imHourOffColor
End Property
Public Property Let CSI_HourOffColor(Setting As OLE_COLOR)
    imHourOffColor = Setting
    Call AssignControlProperties
    PropertyChanged "CSI_HourOffColor"
End Property

''****************************************************************************
''
''****************************************************************************
'Public Property Get Font() As StdFont
'   Set Font = mFont
'End Property
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
'   For ilLoop = 0 To 23
'      KeyPad_Buttons(ilLoop).Font = New_Font
'   Next
'   btn_All.Font = New_Font
'   btn_0to5.Font = New_Font
'   btn_6to9.Font = New_Font
'   btn_10to2.Font = New_Font
'   btn_3to6.Font = New_Font
'   btn_7to23.Font = New_Font
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
    For ilLoop = 0 To 23
       KeyPad_Buttons(ilLoop).FontName = mFont.Name
    Next
    btn_All.FontName = mFont.Name
    btn_0to5.FontName = mFont.Name
    btn_6to9.FontName = mFont.Name
    btn_10to2.FontName = mFont.Name
    btn_3to6.FontName = mFont.Name
    btn_7to23.FontName = mFont.Name
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
    For ilLoop = 0 To 23
       KeyPad_Buttons(ilLoop).FontSize = mFont.Size
    Next
    'btn_DownArrow.FontSize = mFont.Size
    btn_All.FontSize = mFont.Size
    btn_0to5.FontSize = mFont.Size
    btn_6to9.FontSize = mFont.Size
    btn_10to2.FontSize = mFont.Size
    btn_3to6.FontSize = mFont.Size
    btn_7to23.FontSize = mFont.Size
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
    For ilLoop = 0 To 23
       KeyPad_Buttons(ilLoop).FontBold = mFont.Bold
    Next
    'btn_DownArrow.FontBold = mFont.Bold
    btn_All.FontBold = mFont.Bold
    btn_0to5.FontBold = mFont.Bold
    btn_6to9.FontBold = mFont.Bold
    btn_10to2.FontBold = mFont.Bold
    btn_3to6.FontBold = mFont.Bold
    btn_7to23.FontBold = mFont.Bold
    btn_Clear.FontBold = mFont.Bold
    Call AssignControlProperties
    ' Call PositionAllControls
End Property


