VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl CSI_RTFEdit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9165
   ForeColor       =   &H00FF0000&
   ScaleHeight     =   5325
   ScaleWidth      =   9165
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7800
      Top             =   1200
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   4680
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Color"
      Top             =   15
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
   Begin RichTextLib.RichTextBox rtfRichTextBox1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4895
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"CSI_RTFEdit.ctx":0000
   End
   Begin VB.PictureBox plcPanel 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   9135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9165
      Begin MSComctlLib.ImageList ilsImageList1 
         Left            =   7800
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   22
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   17
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":0082
               Key             =   "bold_u"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":0704
               Key             =   "spellcheck_u"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":0D86
               Key             =   "spellcheck_d"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":1408
               Key             =   "preview_d"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":1A8A
               Key             =   "preview_u"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":210C
               Key             =   "bold_m"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":278E
               Key             =   "bold_d"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":2E10
               Key             =   "font_u"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":3492
               Key             =   "font_d"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":3B14
               Key             =   "underline_u"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":4196
               Key             =   "underline_d"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":4818
               Key             =   "italic_u"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":4E9A
               Key             =   "italic_d"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":551C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":5B9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":6220
               Key             =   "upper_d"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":68A2
               Key             =   "upper_u"
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cbcFontCombo 
         Height          =   315
         ItemData        =   "CSI_RTFEdit.ctx":6F24
         Left            =   120
         List            =   "CSI_RTFEdit.ctx":6F26
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Font"
         Top             =   0
         Width           =   2475
      End
      Begin VB.ComboBox cbcSizeCombo 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Size"
         Top             =   0
         Width           =   735
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog1 
         Left            =   7200
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageListColors 
         Left            =   8400
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":6F28
               Key             =   ""
               Object.Tag             =   "0,0,0"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":7032
               Key             =   ""
               Object.Tag             =   "0,0,255"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":713C
               Key             =   ""
               Object.Tag             =   "0,0,128"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":7246
               Key             =   ""
               Object.Tag             =   "128,128,128"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":7350
               Key             =   ""
               Object.Tag             =   "0,128,0"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":745A
               Key             =   ""
               Object.Tag             =   "128,128,0"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":7564
               Key             =   ""
               Object.Tag             =   "0,255,255"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":766E
               Key             =   ""
               Object.Tag             =   "192,192,192"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":7778
               Key             =   ""
               Object.Tag             =   "0,255,0"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":7882
               Key             =   ""
               Object.Tag             =   "255,0,255"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":798C
               Key             =   ""
               Object.Tag             =   "128,0,128"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":7A96
               Key             =   ""
               Object.Tag             =   "255,0,0"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":7BA0
               Key             =   ""
               Object.Tag             =   "0,128,128"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":7CAA
               Key             =   ""
               Object.Tag             =   "255,255,255"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CSI_RTFEdit.ctx":7DB4
               Key             =   ""
               Object.Tag             =   "255,255,0"
            EndProperty
         EndProperty
      End
      Begin VB.Image imgFButtons 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   4320
         Picture         =   "CSI_RTFEdit.ctx":7EBE
         ToolTipText     =   "Upper"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgFButtons 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   5730
         Picture         =   "CSI_RTFEdit.ctx":8530
         ToolTipText     =   "Check Spelling"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgFButtons 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   5310
         Picture         =   "CSI_RTFEdit.ctx":8BA2
         ToolTipText     =   "Preview"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgFButtons 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   3600
         Picture         =   "CSI_RTFEdit.ctx":9214
         ToolTipText     =   "Bold"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgFButtons 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   3960
         Picture         =   "CSI_RTFEdit.ctx":9886
         ToolTipText     =   "Italic"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgFButtons 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   6720
         Picture         =   "CSI_RTFEdit.ctx":9EF8
         ToolTipText     =   "Underline"
         Top             =   15
         Visible         =   0   'False
         Width           =   360
      End
   End
End
Attribute VB_Name = "CSI_RTFEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of CSI_RTFEdit.ctl on Wed 6/17/09 @ 12:5
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  hmScr                         tmScr                         imScrRecLen               *
'*                                                                                        *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  InitFonts                     mIsAscii                      mOpenSCR                  *
'*                                                                                        *
'*                                                                                        *
'* Public Property Procedures (Marked)                                                    *
'*  MaxLength(Get)                FontName(Get)                 FontName(Let)             *
'*  FontSize(Get)                 FontSize(Let)                 BackColor(Get)            *
'*  ForeColor(Get)                ForeColor(Let)                                          *
'******************************************************************************************

Option Explicit

Private smText As String
Private imMaxLength As Integer
Private bmIgnoreChangeEvent As Boolean
Private bmControlIsReady As Boolean
Private smFontName As String
Private imFontSize As Integer
Private cmForeGroundColor As ColorConstants
Private cmBackGroundColor As ColorConstants
Private SpellCheck As Object

Event Change()
'Event SpellCheckerStarting()
'Event SpellCheckerCompleted()

Private bmFontsAreInitialized As Boolean

'****************************************************************************
'
'****************************************************************************
Private Sub rtfRichTextBox1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 2 Then
        Call mToggleBold
        KeyAscii = 0
    ElseIf KeyAscii = 9 Then
        Call mToggleItalic
        KeyAscii = 0
    ElseIf KeyAscii = 21 Then
        Call mToggleUnderline
        KeyAscii = 0
    End If
End Sub

Private Sub UserControl_EnterFocus()
    Dim slStr As String

    slStr = rtfRichTextBox1.Text
    If mIsAllUpperCase(slStr) Then
        imgFButtons(5).Picture = ilsImageList1.ListImages("upper_d").Picture
    Else
        imgFButtons(5).Picture = ilsImageList1.ListImages("upper_u").Picture
    End If
End Sub

Private Sub UserControl_ExitFocus()
    ' bmControlIsReady = False
End Sub
Private Sub UserControl_LostFocus()
    bmControlIsReady = False
End Sub


'***************************************************
'
'***************************************************
Private Sub UserControl_Initialize()
    bmControlIsReady = False
    bmIgnoreChangeEvent = False
    smText = ""
    imMaxLength = 5000
    If Not bmFontsAreInitialized Then
        ' Call InitFonts
        Call LoadComboFontList
        Call mLoadColorComboList
        'mOpenSCR
    End If
    rtfRichTextBox1.BackColor = cmBackGroundColor
    rtfRichTextBox1.SelColor = cmForeGroundColor

    bmControlIsReady = True
End Sub

'***************************************************
' Purpose:  Set the text font based on the selection.
'***************************************************
Private Sub cbcFontCombo_Click()
    On Error GoTo Err_cbcFontCombo_Click
    If Not bmControlIsReady Then
        Exit Sub
    End If
    If cbcFontCombo.Text = "Font..." Then
        Call mSetFontDialog
        Exit Sub
    End If
    rtfRichTextBox1.SelFontName = cbcFontCombo.Text
    rtfRichTextBox1.SetFocus
    Exit Sub
Err_cbcFontCombo_Click:
    MsgBox "Error: " & Err.Number & "  " & Err.Description, vbExclamation, "Critical Error"
    Resume Next
End Sub










'****************************************************************************
'
'****************************************************************************
Private Sub LoadComboFontList()
    Dim ilLoop As Integer

    For ilLoop = 0 To Screen.FontCount - 1
        cbcFontCombo.AddItem Screen.Fonts(ilLoop)
    Next
    For ilLoop = 0 To Screen.FontCount - 1
        cbcFontCombo.ListIndex = ilLoop
        If cbcFontCombo.Text = "Arial" Then
            Exit For
        End If
    Next

    rtfRichTextBox1.SelFontName = "Arial"
    rtfRichTextBox1.SelFontSize = 10
    For ilLoop = 6 To 70
        cbcSizeCombo.AddItem str(ilLoop)
    Next
    cbcFontCombo.AddItem "Font...", 0
    ' set default font size to 10
    cbcSizeCombo.ListIndex = 4
    bmFontsAreInitialized = True
    Exit Sub

'    cbcFontCombo.AddItem "Arial"
'    cbcFontCombo.AddItem "Arial Black"
'    cbcFontCombo.AddItem "Centry Gothic"
'    cbcFontCombo.AddItem "Courier New"
'    cbcFontCombo.AddItem "Franklin Gothic Book"
'    cbcFontCombo.AddItem "Franklin Gothic Demi"
'    cbcFontCombo.AddItem "Franklin Gothic Demi Cond"
'    cbcFontCombo.AddItem "Franklin Gothic Heavy"
'    cbcFontCombo.AddItem "Franklin Gothic Medium"
'    cbcFontCombo.AddItem "Franklin Gothic Heavy Cond"
'    cbcFontCombo.AddItem "Gill Sans MT"
'    cbcFontCombo.AddItem "Gill Sans MT Condensed"
'    cbcFontCombo.AddItem "Gill Sans MT Ext Condensed Bold"
'    cbcFontCombo.AddItem "Gill Sans Ultra Bold"
'    cbcFontCombo.AddItem "Gill Sans Ultra Bold Condensed"
'    cbcFontCombo.AddItem "Impact"
'    cbcFontCombo.AddItem "Lucida Sans Typewriter"
'    cbcFontCombo.AddItem "MS Sans Serif"
'    cbcFontCombo.AddItem "Palatino Linotype"
'    cbcFontCombo.AddItem "Rockwell"
'    cbcFontCombo.AddItem "Symbols"
'    cbcFontCombo.AddItem "Tohama"
'    cbcFontCombo.AddItem "Times New Roman"
'    cbcFontCombo.AddItem "Tw Cen MT"
'    cbcFontCombo.AddItem "Tw Cen MT Condensed"
'    cbcFontCombo.AddItem "Tw Cen MT Condensed Extra Bold"
'    cbcFontCombo.AddItem "Verdana"

End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub mLoadColorComboList()
    Dim ilLoop As Integer

    Set ImageCombo1.ImageList = ImageListColors
    ImageCombo1.ComboItems.Clear
    For ilLoop = 1 To ImageListColors.ListImages.Count
        ImageCombo1.ComboItems.Add ilLoop, ImageListColors.ListImages(ilLoop).Key, ImageListColors.ListImages(ilLoop).Tag, ilLoop, ilLoop
    Next
    ImageCombo1.Locked = True 'this way you can't mess with the box
    Set ImageCombo1.SelectedItem = ImageCombo1.GetFirstVisible
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub cbcSizeCombo_Click()
    On Error Resume Next
    ' change size
    rtfRichTextBox1.SelFontSize = cbcSizeCombo.Text
    ' set focus back to RTF
    rtfRichTextBox1.SetFocus

End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub ImageCombo1_Click()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIndex                                                                               *
'******************************************************************************************

    Dim slColorSetting As String
    Dim ilNextPos As Integer
    Dim llRed As Long
    Dim llGreen As Long
    Dim llBlue As Long

    slColorSetting = ImageCombo1.SelectedItem
    llRed = Val(slColorSetting)
    ilNextPos = InStr(1, slColorSetting, ",")
    llGreen = Val(Mid(slColorSetting, ilNextPos + 1))
    ilNextPos = InStr(ilNextPos + 1, slColorSetting, ",")
    llBlue = Val(Mid(slColorSetting, ilNextPos + 1))

    rtfRichTextBox1.SelColor = RGB(llRed, llGreen, llBlue)
    rtfRichTextBox1.SetFocus
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub imgFButtons_Click(Index As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRet                                                                                 *
'******************************************************************************************



    Select Case Index
        Case 0 'bold clicked
            Call mToggleBold

        Case 1 'italic clicked
            Call mToggleItalic

        Case 2 'underline clicked
            Call mToggleUnderline

        Case 3 ' Print Preview
            mPrintPreview

        Case 4 ' Spell Check
            Call mSpellCheckIt

        Case 5 ' Upper case
            Call mToggleCase
    End Select
End Sub

'****************************************************************************
'
'****************************************************************************
Public Sub mSpellCheckIt()
    Call mSpellCheckUsingMSWord
End Sub

'****************************************************************************
'
'****************************************************************************
Private Function mPrintPreview() As Integer
Dim ilRet As Integer
    ilRet = gGenRTF("RTFPreview.rpt", rtfRichTextBox1)

End Function

'****************************************************************************
'
'****************************************************************************
Private Sub imgFButtons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0 'bold clicked
        Case 1 'italic clicked
        Case 2 'underline clicked
        Case 3 'preview
            imgFButtons(3).Picture = ilsImageList1.ListImages("preview_d").Picture
        Case 4 'spell checker
            imgFButtons(4).Picture = ilsImageList1.ListImages("spellcheck_d").Picture
        Case 5 'Upper case
            imgFButtons(5).Picture = ilsImageList1.ListImages("upper_d").Picture
    End Select
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub imgFButtons_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0 'bold clicked
        Case 1 'italic clicked
        Case 2 'underline clicked
        Case 3 'preview
            imgFButtons(3).Picture = ilsImageList1.ListImages("preview_u").Picture
        Case 4 'spell checker
            imgFButtons(4).Picture = ilsImageList1.ListImages("spellcheck_u").Picture
        Case 5 'Upper case
            imgFButtons(5).Picture = ilsImageList1.ListImages("upper_u").Picture
    End Select
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub rtfRichTextBox1_Change()
    Dim slStr As String

    If bmIgnoreChangeEvent Then
        Exit Sub
    End If
    RaiseEvent Change

    slStr = rtfRichTextBox1.Text
    If mIsAllUpperCase(slStr) Then
        imgFButtons(5).Picture = ilsImageList1.ListImages("upper_d").Picture
    Else
        imgFButtons(5).Picture = ilsImageList1.ListImages("upper_u").Picture
    End If
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub rtfRichTextBox1_SelChange()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        ilStartListIndex                                        *
'******************************************************************************************

    Dim slColor As String
    Dim ilRow As Long

    On Error GoTo ERR_SelChange
    With rtfRichTextBox1
        ' Cause the color selector to track with whatever color the user is over.
        If Not IsNull(.SelColor) Then
            slColor = str(rtfRichTextBox1.SelColor)
            Call mSetColor(slColor)
        End If
        ' Same for bold, italic and underline.
        If (IsNull(.SelBold) = True) Or (.SelBold = False) Then
            imgFButtons(0).Picture = ilsImageList1.ListImages("bold_u").Picture
        ElseIf .SelBold = True Then
            imgFButtons(0).Picture = ilsImageList1.ListImages("bold_d").Picture
        End If
        If (IsNull(.SelItalic) = True) Or (.SelItalic = False) Then
            imgFButtons(1).Picture = ilsImageList1.ListImages("italic_u").Picture
        ElseIf .SelItalic = True Then
            imgFButtons(1).Picture = ilsImageList1.ListImages("italic_d").Picture
        End If
        If (IsNull(.SelUnderline) = True) Or (.SelUnderline = False) Then
            imgFButtons(2).Picture = ilsImageList1.ListImages("underline_u").Picture
        ElseIf .SelUnderline = True Then
            imgFButtons(2).Picture = ilsImageList1.ListImages("underline_d").Picture
        End If
        bmControlIsReady = False
        ilRow = SendMessageByString(cbcFontCombo.hwnd, CB_FINDSTRING, -1, .SelFontName)
        If ilRow >= 0 Then
            cbcFontCombo.ListIndex = ilRow
        End If
        ilRow = SendMessageByString(cbcSizeCombo.hwnd, CB_FINDSTRING, -1, str(.SelFontSize))
        If ilRow >= 0 Then
            cbcSizeCombo.ListIndex = ilRow
        End If
        bmControlIsReady = True
    End With
    Exit Sub

ERR_SelChange:
    bmControlIsReady = True
    Exit Sub
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub UserControl_Resize()
    rtfRichTextBox1.Top = plcPanel.Top + plcPanel.Height
    rtfRichTextBox1.Left = 0 ' Screen.TwipsPerPixelX
    rtfRichTextBox1.Width = Width ' - Screen.TwipsPerPixelX
    rtfRichTextBox1.Height = Height - plcPanel.Height ' + Screen.TwipsPerPixelX
    ImageCombo1.Height = imgFButtons(0).Height
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub mToggleBold()
    With rtfRichTextBox1
        If (IsNull(.SelBold) = True) Or (.SelBold = False) Then
            ' selection is mixed or not bold
            ' set it
            .SelBold = True
            imgFButtons(0).Picture = ilsImageList1.ListImages("bold_d").Picture
        ElseIf .SelBold = True Then
            ' selection is bold, toggle it
            .SelBold = False
            imgFButtons(0).Picture = ilsImageList1.ListImages("bold_u").Picture
        End If
    End With
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub mToggleCase()
    Dim slStr As String

    slStr = rtfRichTextBox1.Text
    If mIsAllUpperCase(slStr) Then
        slStr = LCase(slStr)
        imgFButtons(5).Picture = ilsImageList1.ListImages("upper_u").Picture
    Else
        slStr = UCase(slStr)
        imgFButtons(5).Picture = ilsImageList1.ListImages("upper_d").Picture
    End If
    rtfRichTextBox1.Text = slStr
End Sub

'****************************************************************************
' look at ascii characters only and determine if they are all upper case.
' If even one of them is not, return false.
'
'****************************************************************************
Private Function mIsAllUpperCase(sStr As String) As Boolean
    Dim ilLen As Integer
    Dim ilLoop As Integer
    Dim slOneChar As String

    mIsAllUpperCase = True
    ilLen = Len(sStr)
    For ilLoop = 1 To ilLen
        slOneChar = Mid(sStr, ilLoop, 1)
        If slOneChar >= "a" And slOneChar <= "z" Then
            mIsAllUpperCase = False
            Exit Function
        End If
    Next
End Function


'****************************************************************************
'
'****************************************************************************
Private Sub mToggleItalic()
    With rtfRichTextBox1
        If (IsNull(.SelItalic) = True) Or (.SelItalic = False) Then
            ' selection is italic or mixed, so set italic
            .SelItalic = True
            imgFButtons(1).Picture = ilsImageList1.ListImages("italic_d").Picture
        ElseIf .SelItalic = True Then
            'selection is italic, so toggle it
            .SelItalic = False
            imgFButtons(1).Picture = ilsImageList1.ListImages("italic_u").Picture
        End If
    End With
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub mToggleUnderline()
    With rtfRichTextBox1
        If (IsNull(.SelUnderline) = True) Or _
            (.SelUnderline = False) Then
            ' selection is underlined or mixed,
            ' so set to underlined
            .SelUnderline = True
            imgFButtons(2).Picture = ilsImageList1.ListImages("underline_d").Picture
        ElseIf .SelUnderline = True Then
            'selection is not underlined,
            ' so toggle it.
            .SelUnderline = False
            imgFButtons(2).Picture = ilsImageList1.ListImages("underline_u").Picture
        End If
    End With
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub mSetFontDialog()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim sCurrentFontName As String
    Dim ilRow As Long

    On Error GoTo CxlError5 ' set error trap
    sCurrentFontName = ""
    If Not IsNull(rtfRichTextBox1.SelFontName) Then
        sCurrentFontName = rtfRichTextBox1.SelFontName
    End If
    With dlgCommonDialog1
        ' set printer & screen fonts flag
        .flags = cdlCFBoth
        ' force the selection of a real font
        .flags = .flags + cdlCFForceFontExist
        ' set the special effects flag
        .flags = .flags + cdlCFEffects
        .CancelError = True ' set error trigger
        ' set common dialog info to current values
        .FontName = cbcFontCombo.Text
        If Len(cbcSizeCombo.Text) > 0 Then
            .FontSize = cbcSizeCombo.Text
        End If
        If rtfRichTextBox1.SelBold = True Then
            .FontBold = True
        End If
        If rtfRichTextBox1.SelItalic = True Then
            .FontItalic = True
        End If
        If rtfRichTextBox1.SelUnderline = True Then
            .FontUnderline = True
        End If
        If rtfRichTextBox1.SelStrikeThru = True Then
            .FontStrikethru = True
        End If
        ' show font selection common dialog
        .ShowFont
        ' Get the new values from the dialog
        If .FontBold Then
            Call mToggleBold
        End If
        If .FontItalic Then
            Call mToggleItalic
        End If
        If .FontUnderline Then
            Call mToggleUnderline
        End If
        If .FontStrikethru Then
            'rtfRichTextBox1.SelStrikeThru = True
        End If

        rtfRichTextBox1.SelColor = .Color
        ' set the font based on the selection
        Call mSetFont(.FontName, .FontSize)
        Call mSetColor(.Color)
        rtfRichTextBox1.SetFocus
        Exit Sub
    End With
CxlError5: ' cancel selected
    ' Select the font that was selected before this operation occurred.
    ilRow = SendMessageByString(cbcFontCombo.hwnd, CB_FINDSTRING, -1, sCurrentFontName)
    If ilRow >= 0 Then
        cbcFontCombo.ListIndex = ilRow
    End If
    Exit Sub
End Sub

'****************************************************************************
'
'****************************************************************************
Sub mSetFont(strFontName As String, strFontSize As String)
    Dim intI As Integer
    For intI = 0 To cbcFontCombo.ListCount - 1
        If cbcFontCombo.List(intI) = strFontName Then
            cbcFontCombo.ListIndex = intI
            intI = cbcFontCombo.ListCount
        End If
    Next intI
    If Val(strFontSize) <= 70 And Val(strFontSize) >= 6 Then
        cbcSizeCombo.ListIndex = Val(strFontSize) - 6
    Else
        cbcSizeCombo.ListIndex = -1
    End If
End Sub

'****************************************************************************
'
'****************************************************************************
Sub mSetColor(strColor As String)
    Dim slColorSetting As String
    Dim ilLoop As Integer
    Dim ilNextPos As Integer
    Dim llRed As Long
    Dim llGreen As Long
    Dim llBlue As Long

    For ilLoop = 1 To ImageCombo1.ComboItems.Count
        slColorSetting = ImageCombo1.ComboItems.Item(ilLoop).Text
        llRed = Val(slColorSetting)
        ilNextPos = InStr(1, slColorSetting, ",")
        llGreen = Val(Mid(slColorSetting, ilNextPos + 1))
        ilNextPos = InStr(ilNextPos + 1, slColorSetting, ",")
        llBlue = Val(Mid(slColorSetting, ilNextPos + 1))

        If RGB(llRed, llGreen, llBlue) = Val(strColor) Then
            ImageCombo1.ComboItems.Item(ilLoop).Selected = ilLoop
            Exit For
        End If
    Next
End Sub

'****************************************************************************
' Load property values from storage
'****************************************************************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    smText = PropBag.ReadProperty("Text", "")
    imMaxLength = PropBag.ReadProperty("MaxLength", 5000)
    smFontName = PropBag.ReadProperty("FontName", "Arial")
    imFontSize = PropBag.ReadProperty("FontSize", "10")
    cmBackGroundColor = PropBag.ReadProperty("BackColor", RGB(255, 255, 255))
    cmForeGroundColor = PropBag.ReadProperty("ForeColor", RGB(0, 0, 0))

    rtfRichTextBox1.SelFontName = smFontName
    rtfRichTextBox1.SelFontSize = imFontSize
    rtfRichTextBox1.BackColor = cmBackGroundColor
    rtfRichTextBox1.SelColor = cmForeGroundColor
End Sub

'****************************************************************************
' Write property values to storage
'****************************************************************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", smText, "")
    Call PropBag.WriteProperty("MaxLength", imMaxLength, 5000)
    Call PropBag.WriteProperty("FontName", smFontName, "Arial")
    Call PropBag.WriteProperty("FontSize", imFontSize, 10)
    Call PropBag.WriteProperty("BackColor", cmBackGroundColor, RGB(255, 255, 255))
    Call PropBag.WriteProperty("ForeColor", cmForeGroundColor, RGB(0, 0, 0))

    rtfRichTextBox1.SelFontName = smFontName
    rtfRichTextBox1.SelFontSize = imFontSize
    rtfRichTextBox1.BackColor = cmBackGroundColor
    rtfRichTextBox1.SelColor = cmForeGroundColor
End Sub

'****************************************************************************
'
'****************************************************************************
Public Sub SetText(sText As String)
    bmIgnoreChangeEvent = True
    smText = sText
    rtfRichTextBox1.TextRTF = smText
    bmIgnoreChangeEvent = False
End Sub

'****************************************************************************
'
'****************************************************************************
Public Property Get Text() As String
    Dim slTempText As String

    slTempText = rtfRichTextBox1.TextRTF
    If Len(slTempText) > imMaxLength Then
        slTempText = Left(slTempText, imMaxLength)
    End If
    smText = slTempText
    Text = slTempText
End Property
Public Property Let Text(sText As String)
    Dim s As String

    On Error GoTo Error_Text
    If IsNull(rtfRichTextBox1.SelFontName) Then
        Exit Property
    End If

    smText = sText
    s = rtfRichTextBox1.SelFontName

    rtfRichTextBox1.TextRTF = smText
    PropertyChanged "Text"
    Exit Property

Error_Text:
    MsgBox "An error occured"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get TextOnly() As String
    Dim slTempText As String

    slTempText = rtfRichTextBox1.Text
    If Len(slTempText) > imMaxLength Then
        slTempText = Left(slTempText, imMaxLength)
    End If
    TextOnly = slTempText
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get MaxLength() As Integer 'VBC NR
    MaxLength = imMaxLength 'VBC NR
End Property 'VBC NR
Public Property Let MaxLength(iMaxLength As Integer)
    imMaxLength = iMaxLength
    PropertyChanged "MaxLength"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get FontName() As String 'VBC NR
    FontName = smFontName 'VBC NR
End Property 'VBC NR
Public Property Let FontName(sFontName As String) 'VBC NR
    Dim ilRow As Long 'VBC NR

    smFontName = sFontName 'VBC NR
    ilRow = SendMessageByString(cbcFontCombo.hwnd, CB_FINDSTRING, -1, smFontName) 'VBC NR
    If ilRow >= 0 Then 'VBC NR
        bmControlIsReady = False 'VBC NR
        cbcFontCombo.ListIndex = ilRow 'VBC NR
        rtfRichTextBox1.SelFontName = sFontName 'VBC NR
        bmControlIsReady = True 'VBC NR
    End If 'VBC NR
    PropertyChanged "FontName" 'VBC NR
End Property 'VBC NR

'****************************************************************************
'
'****************************************************************************
Public Property Get FontSize() As Integer 'VBC NR
    FontSize = imFontSize 'VBC NR
End Property 'VBC NR
Public Property Let FontSize(iFontSize As Integer) 'VBC NR
    Dim ilRow As Long 'VBC NR

    imFontSize = iFontSize 'VBC NR
    ilRow = SendMessageByString(cbcSizeCombo.hwnd, CB_FINDSTRING, -1, str(imFontSize)) 'VBC NR
    If ilRow >= 0 Then 'VBC NR
        bmControlIsReady = False 'VBC NR
        cbcSizeCombo.ListIndex = ilRow 'VBC NR
        bmControlIsReady = True 'VBC NR
    End If 'VBC NR
    PropertyChanged "FontSize" 'VBC NR
End Property 'VBC NR

'****************************************************************************
'
'****************************************************************************
Public Property Get BackColor() As ColorConstants 'VBC NR
   BackColor = rtfRichTextBox1.BackColor 'VBC NR
End Property 'VBC NR
Public Property Let BackColor(BKColor As ColorConstants)
    cmBackGroundColor = BKColor
    rtfRichTextBox1.BackColor = BKColor
    PropertyChanged "BackColor"
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Get ForeColor() As ColorConstants 'VBC NR
   ForeColor = rtfRichTextBox1.SelColor 'VBC NR
End Property 'VBC NR
Public Property Let ForeColor(FGColor As ColorConstants) 'VBC NR
    cmForeGroundColor = FGColor 'VBC NR
    rtfRichTextBox1.SelColor = FGColor 'VBC NR
    Call mSetColor(rtfRichTextBox1.SelColor) 'VBC NR
    PropertyChanged "ForeColor" 'VBC NR
End Property 'VBC NR




'****************************************************************************
'
'****************************************************************************
Public Sub mSpellCheckUsingMSWord()
    On Error GoTo Err_SpellCheckUsingMSWord
    Dim sText As String

    Screen.MousePointer = vbHourglass
    'RaiseEvent SpellCheckerStarting
    sText = rtfRichTextBox1.TextRTF

    App.OleRequestPendingTimeout = 999999   ' Prevent the "Switch To" dialog from appearing.
    'App.OleServerBusyMsgText = "Press Alt-Esc to see the spell checking results"
    'App.OleRequestPendingMsgText = "Press Alt-Esc to see the spell checking results"
'    DoEvents
'    App.OleServerBusyTimeout = 1000
'    App.OleServerBusyRaiseError = True
    Set SpellCheck = CreateObject("Word.Application")
    SpellCheck.Visible = False
    Call MinimizeWordIfOpen
    SpellCheck.Documents.Add                              'Open New Document (Hidden)
    Clipboard.Clear
    Clipboard.SetText sText, vbCFRTF                      'Copy Text To Clipboard
    SpellCheck.Selection.Paste                            'Paste Text Into WORD
    Call BringWindowToTopMost
    SpellCheck.Visible = False
    SpellCheck.ActiveDocument.CheckSpelling               'Activate The Spell Checker
    'SpellCheck.ActiveDocument.CheckGrammar                ' Does both spelling and grammer.
    SpellCheck.Visible = False                            'Hide WORD From User
    SpellCheck.ActiveDocument.Select                      'Select The Corrected Text
    SpellCheck.Selection.Cut                              'Cut The Text To Clipboard
    rtfRichTextBox1.TextRTF = Clipboard.GetText(vbCFRTF)  'Assign Text To SPELLCHECKER Function
    SpellCheck.ActiveDocument.Close False
    SpellCheck.Quit
    Set SpellCheck = Nothing
    'RaiseEvent SpellCheckerCompleted
    Screen.MousePointer = vbNormal
    MsgBox "Spell Checking is Complete"
    rtfRichTextBox1.SetFocus
    Exit Sub

Err_SpellCheckUsingMSWord:
    Screen.MousePointer = vbNormal
    SpellCheck.ActiveDocument.Close False
    SpellCheck.Quit
    Set SpellCheck = Nothing
    MsgBox "Error: " & Err.Number & ", " & Err.Description & vbCrLf & vbCrLf & "Please note you must have Microsoft Word installed to utilize the spell check feature.", vbExclamation, "Spell Check Problem"
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub BringWindowToTopMost()
    Dim hwnd As Long
    Dim ilresult As Long

    'hWnd = FindWindow(vbNullString, "Spelling: English (U.S.)")
    hwnd = FindWindow(vbNullString, "Document1 - Microsoft Word")

    If hwnd <> 0 Then
        ilresult = SetWindowPos(hwnd, WNDNOTOPMOST, 0, 0, 0, 0, FRMNOMOVE Or FRMNOSIZE)
    End If
End Sub

'****************************************************************************
' This function will look for a word doc that is currently open with the
' title of "Document1 - Microsoft Word", indicating a new blank word doc.
' If this is found, we need to minimize it to avoid having it become the
' top most visible window.
'
'****************************************************************************
Private Sub MinimizeWordIfOpen()
    Dim hwnd As Long
    Dim wp As WINDOWPLACEMENT

    ' Const WM_COMMAND = &H111
    hwnd = FindWindow(vbNullString, "Document1 - Microsoft Word")

    If hwnd <> 0 Then
        If GetWindowPlacement(hwnd, wp) > 0 Then
            wp.Length = Len(wp)
            wp.flags = 0&
            wp.showCmd = SW_SHOWMINIMIZED
            SetWindowPlacement hwnd, wp
        End If
        ' SendMessage hWnd, &H111, MIN_ALL, ByVal 0&
    End If
End Sub




