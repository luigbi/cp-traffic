VERSION 5.00
Begin VB.UserControl CSI_ComboBoxMS 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3555
   ClipBehavior    =   0  'None
   FontTransparent =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3555
   ToolboxBitmap   =   "CSI_ComboBoxMS.ctx":0000
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
      Left            =   3090
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   300
   End
   Begin VB.ListBox lbcMasterList 
      Height          =   840
      ItemData        =   "CSI_ComboBoxMS.ctx":0312
      Left            =   1320
      List            =   "CSI_ComboBoxMS.ctx":0319
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox ec_InputBox 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1710
   End
   Begin VB.ListBox lbcList 
      Height          =   840
      ItemData        =   "CSI_ComboBoxMS.ctx":032C
      Left            =   0
      List            =   "CSI_ComboBoxMS.ctx":0333
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "CSI_ComboBoxMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private imSelectedIndex As Long
Private bmIsDroppedDown As Boolean
Private imBSMode As Boolean
Private imDropDownWidth As Integer
Private imDropDownHeight As Integer
Private imDefaultListNoRows As Integer
Private imUserListNoRows As Integer
Private imToolTipArray() As String
Private cgList As String
Private igFoundRow As Integer
Private llForeColor As Long
Private llBackColor As Long
Private bmControlIsReady As Boolean
Private smFontName As String
Private imFontName As Long
Private imFontSize As Long
Private imFontBold As Integer
Private bmShowDropDownOnFocus As Boolean
Private smPopupListDirection As String

Event OnChange()
Event DblClick()
Event ReSetLoc()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event LostInputFocus()
Event GotInputFocus()

Const m_def_Text = ""
Dim m_Text As String

Private Sub btn_DownArrow_GotFocus()
    RaiseEvent GotInputFocus
End Sub

Private Sub ec_InputBox_LostFocus()
    m_Text = lbcList.Text
    ec_InputBox.Text = lbcList.Text
    imSelectedIndex = lbcList.ListIndex
    
    
    mSetListResult
    
End Sub

Private Sub lbcList_GotFocus()
    RaiseEvent GotInputFocus
End Sub

Private Sub lbcList_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub lbcList_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim A As Integer
End Sub

Private Sub lbcList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mHideDropDown
    ec_InputBox.Text = lbcList.Text
    If ec_InputBox.Visible Then
        ec_InputBox.SetFocus
        If lbcList.ListIndex + 1 <= UBound(imToolTipArray) Then
            ec_InputBox.ToolTipText = imToolTipArray(lbcList.ListIndex + 1)
        End If
    End If
    
    If lbcList.ListIndex = -1 Then
        'ec_InputBox.Text = lbcList.List(lbcList.ListIndex)
        gMatchLookAhead ec_InputBox, lbcList, imBSMode, 0
    End If
    m_Text = lbcList.Text
    imSelectedIndex = lbcList.ListIndex
    Button = 0
    
    RaiseEvent OnChange
    
    mSetListResult
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub lbcList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim llRow As Long
    
    llRow = gGetListBoxRow(lbcList, Y)
    If llRow < 0 Then Exit Sub
    If llRow + 1 <= UBound(imToolTipArray) Then
        lbcList.ToolTipText = imToolTipArray(llRow + 1)
    End If
    'lbcList.ListIndex = llRow
End Sub

Private Sub ec_InputBox_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub ec_InputBox_GotFocus()
    ec_InputBox.SelStart = 0
    ec_InputBox.SelLength = Len(ec_InputBox.Text)
    'RaiseEvent LostInputFocus
End Sub

Private Sub UserControl_EnterFocus()
'    If bmShowDropDownOnFocus Then
'        bmIsDroppedDown = False
'        Call btn_DownArrow_Click
'    End If
RaiseEvent GotInputFocus
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub UserControl_Initialize()
    ec_InputBox.Top = 0
    ec_InputBox.Left = 0

    bmControlIsReady = False
    imBSMode = False
    bmIsDroppedDown = False
    imSelectedIndex = 0
    imDropDownWidth = 0
    imDropDownHeight = 0
    smPopupListDirection = "B"
    imDefaultListNoRows = 8
    imUserListNoRows = -1
    ReDim imToolTipArray(0 To 0) As String
    SetDropDownNumRows 8, False
    SetDropDownCharWidth 20
    smFontName = "Arial"
    bmShowDropDownOnFocus = True
    bmControlIsReady = True
End Sub

Private Sub UserControl_LostFocus()
    RaiseEvent LostInputFocus
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub UserControl_Resize()
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
        btn_DownArrow.Height = ec_InputBox.Height - (Screen.TwipsPerPixelY * 6)
        btn_DownArrow.Top = ec_InputBox.Top + (Screen.TwipsPerPixelY * 2)
        btn_DownArrow.Left = ec_InputBox.Width - btn_DownArrow.Width - (Screen.TwipsPerPixelX * 2)
    End If
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub btn_DownArrow_Click()
    If bmIsDroppedDown Then
        mHideDropDown
        ec_InputBox.SetFocus
        Exit Sub
    End If
    Dim ilHightErr As Integer
        
    If lbcList.ListCount = 0 Then
        FilterList (ec_InputBox.Text)
    End If
    If ec_InputBox.Text <> "" Then
        'only item is selected, perhaps show the entire list, with this item selected
        FilterList ("")
        ilHightErr = 0
        gMatchLookAhead ec_InputBox, lbcList, imBSMode, ilHightErr
        m_Text = lbcList.Text
        imSelectedIndex = lbcList.ListIndex
    End If
    If ec_InputBox.Text = "" Then
        imSelectedIndex = -1
        m_Text = ""
    End If
    
    
    mShowDropDown
    
End Sub

Private Sub ec_InputBox_KeyDown(KeyCode As Integer, Shift As Integer)
    imBSMode = False
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        KeyCode = 0
        If bmIsDroppedDown = False Then
        'mShowDropDown
            btn_DownArrow_Click
        End If
    End If
    
    If (KeyCode = vbKeyTab Or KeyCode = vbKeyReturn) Then
        m_Text = lbcList.Text
        ec_InputBox.Text = lbcList.Text
        imSelectedIndex = lbcList.ListIndex
        
        RaiseEvent OnChange
        mHideDropDown
        mSetListResult
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub ec_InputBox_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lbcList.ListIndex < lbcList.ListCount - 1 Then
            lbcList.ListIndex = lbcList.ListIndex + 1
            'ec_InputBox.Text = lbcList.Text
            'ec_InputBox.SelStart = 0
            'ec_InputBox.SelLength = Len(ec_InputBox.Text)
           KeyCode = 0
        End If
        If lbcList.ListIndex = -1 Then
            If lbcList.ListCount > 0 Then
                lbcList.ListIndex = 0
            End If
        End If
    End If
    If KeyCode = vbKeyUp Then
        If lbcList.ListIndex > 0 Then
            lbcList.ListIndex = lbcList.ListIndex - 1
        End If
        If lbcList.ListIndex = -1 Then
            lbcList.ListIndex = lbcList.ListCount - 1
        End If
        If bmIsDroppedDown = False Then
            mShowDropDown
        End If
        KeyCode = 0
    End If
    If lbcList.ListIndex > -1 Then
        m_Text = lbcList.Text
        imSelectedIndex = lbcList.ListIndex
        RaiseEvent OnChange
        'mSetListResult
    End If
End Sub

Private Sub ec_InputBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If ec_InputBox.SelLength <> 0 Then
            imBSMode = True
        End If
    End If
    
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub lbcList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    mHideDropDown
'    ec_InputBox.Text = lbcList.Text
'    If ec_InputBox.Visible Then
'        ec_InputBox.SetFocus
'        If lbcList.ListIndex + 1 <= UBound(imToolTipArray) Then
'            ec_InputBox.ToolTipText = imToolTipArray(lbcList.ListIndex + 1)
'        End If
'    End If
'
'    If lbcList.ListIndex = -1 Then
'        'ec_InputBox.Text = lbcList.List(lbcList.ListIndex)
'        gMatchLookAhead ec_InputBox, lbcList, imBSMode, 0
'    End If
'    m_Text = lbcList.Text
'    imSelectedIndex = lbcList.ListIndex
'    RaiseEvent OnChange
'
'
'    mSetListResult
'
End Sub

'****************************************************************************
'
'****************************************************************************
Private Sub lbcList_LostFocus()
    'ec_InputBox.Text = lbcList.Text
    mHideDropDown
    
    If ec_InputBox.Visible Then
        ec_InputBox.SetFocus
        If lbcList.ListIndex + 1 <= UBound(imToolTipArray) Then
            ec_InputBox.ToolTipText = imToolTipArray(lbcList.ListIndex + 1)
        End If
    End If
    If lbcList.ListIndex > -1 Then
        m_Text = lbcList.Text
        imSelectedIndex = lbcList.ListIndex
        'RaiseEvent OnChange
        mSetListResult
    End If
    RaiseEvent KeyDown(9, 0)
End Sub


'****************************************************************************
'
'****************************************************************************
Private Sub ec_InputBox_Change()
    Dim sName As String
    Dim lRow As Long
    Dim iLen As Integer
    Dim ilHightErr As Integer
    If Not bmControlIsReady Then
        Exit Sub
    End If
    
    FilterList (ec_InputBox.Text)
    If ec_InputBox.Text = "" Then
        imSelectedIndex = -1
        m_Text = ""
        RaiseEvent OnChange
    End If
    'ilHightErr = 0
    'gMatchLookAhead ec_InputBox, lbcList, imBSMode, ilHightErr
    'm_Text = lbcList.Text
    'imSelectedIndex = lbcList.ListIndex
    'RaiseEvent OnChange
    Exit Sub

    sName = ec_InputBox.Text
    iLen = Len(sName)
    If imBSMode Then
        iLen = iLen - 1
        If iLen > 0 Then
            sName = Left$(sName, iLen)
        End If
        imBSMode = False
    End If
    ' Look up this value in the list control
    lRow = SendMessageByString(lbcList.HWnd, LB_FINDSTRING, -1, sName)
    If lRow >= 0 Then
        imSelectedIndex = lRow
        lbcList.ListIndex = lRow
        'ec_InputBox.Text = lbcList.Text
        'ec_InputBox.SelStart = iLen
        'ec_InputBox.SelLength = Len(lbcList.Text)
        If lRow + 1 <= UBound(imToolTipArray) Then
            ec_InputBox.ToolTipText = imToolTipArray(lRow + 1)
        End If
    End If
    m_Text = lbcList.Text
    RaiseEvent OnChange
End Sub

'****************************************************************************
'
'****************************************************************************
Public Sub AddItem(sItem As String)
    lbcList.AddItem (sItem)
    lbcMasterList.AddItem (sItem)
End Sub

'****************************************************************************
'
'****************************************************************************
Public Sub SetToolTip(iIndex As Integer, sTipMsg As String)
    Dim ilCount As Integer

    ilCount = lbcList.ListCount
    If iIndex > UBound(imToolTipArray) Then
        ReDim Preserve imToolTipArray(0 To ilCount) As String
    End If
    imToolTipArray(iIndex) = sTipMsg
End Sub

'****************************************************************************
'
'****************************************************************************
Public Sub SetDropDownNumRows(iNumRows As Integer, Optional blSaveRows As Boolean = True)
    Dim slStr As String
    Dim ilPixelsPerLine As Integer
    
    If blSaveRows Then
        imUserListNoRows = iNumRows
    End If
    ilPixelsPerLine = SendMessageByString(lbcList.HWnd, LB_GETITEMHEIGHT, 0, slStr)
    imDropDownHeight = Screen.TwipsPerPixelY * ilPixelsPerLine * (iNumRows + 1)
End Sub
Public Sub SetDropDownHeight(iHeight As Integer)
    imDropDownHeight = iHeight
End Sub
Public Sub SetDropDownCharWidth(iNumChars As Integer)
    imDropDownWidth = iNumChars * (lbcList.FontSize * Screen.TwipsPerPixelX)
End Sub
Public Sub SetDropDownWidth(iWidth As Integer)
    imDropDownWidth = iWidth
End Sub
Public Sub Clear()
    lbcList.Clear
    lbcMasterList.Clear
    ec_InputBox.Text = ""
    imSelectedIndex = 0
    RaiseEvent OnChange
End Sub
Public Sub Enabled(blEnabled As Boolean)
    lbcList.Enabled = blEnabled
    ec_InputBox.Enabled = blEnabled
    btn_DownArrow.Enabled = blEnabled
End Sub

Public Sub SetFont(sFontName As String, iFontSize As Double)
    lbcList.FontName = sFontName
    lbcList.FontSize = iFontSize
    ec_InputBox.FontName = sFontName
    ec_InputBox.FontSize = iFontSize
End Sub
Public Sub PopUpListDirection(slDirection As String)
    'A=Above, B=Below. Test for A
    smPopupListDirection = slDirection
    If slDirection = "A" Then
    
    End If
End Sub
Public Sub SetFocus()
    Me.SetFocus
End Sub


'*                                                     *
'*      Procedure Name:gManLookAhead                   *
'*                                                     *
'*             Created:5/3/93        By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Validates combo box input and,  *
'*                     if invalid, resets the input to *
'*                     the specified last valid input. *
'*                     It validates input by testing   *
'*                     to see if the list box contains *
'*                     any matching string to the input*
'*                     string.  This is used for combo *
'*                     boxes made up of text control/  *
'*                     button control/list box control.*
'*                                                     *
'*******************************************************
Private Sub gMatchLookAhead(edcTextBox As TextBox, lbcListBox As ListBox, ByVal ilBSMode As Integer, ilErrHighLightIndex As Integer)
'
'   gMatchLookAhead edcText, lbcCtrl, ilBSMode, ilHighlightIndex
'   Where:
'       edcText (I)- Text box control (containing input to be validated)
'       lbcCtrl (I)- List box control containing values to be matched
'       ilBSMode (I/O)- Backspace flag(True = backspace key was pressed, False =                        '       backspace key was not pressed)
'       ilHighlightIndex (I)- Selection to be highlighted if input is invalid
'

    Dim ilLen As Integer    'Length of current enter text
    Dim slStr As String     'Text entered
    Dim ilIndex As Integer  'Current index selected from combo box
    Dim ilSelStart As Integer
    Dim ilBracket As Integer
    Dim ilSearch As Integer
    Dim ilSvLastFound As Integer
    Dim illoop As Integer
    Dim ilPos As Integer

    ilSelStart = edcTextBox.SelStart
    slStr = LTrim$(edcTextBox.Text)    'Remove leading blanks only
    ilLen = Len(edcTextBox.Text)
    ilIndex = lbcListBox.ListIndex
    If slStr = "" Then  'If space bar selected, text will be blank- ListIndex will contain a value
        lbcListBox.ListIndex = -1
        Exit Sub
        If lbcListBox.ListIndex >= 0 Then
            slStr = lbcListBox.List(lbcListBox.ListIndex)
            ilLen = 0
            ilIndex = -1    'Force dispaly of selected item by space bar
        Else
            Beep
            If ilErrHighLightIndex >= 0 Then
                lbcListBox.ListIndex = ilErrHighLightIndex
            End If
            Exit Sub
        End If
    End If
    If ilBSMode Then    'If backspace mode- reduce string by one character
        ilLen = ilLen - 1
        If ilLen > 0 Then
            slStr = Left$(slStr, ilLen)
        End If
        ilBSMode = False
        If ilSelStart > 0 Then
            ilSelStart = ilSelStart - 1
        End If
    End If
    If Left$(slStr, 1) = "[" Then   'Search does not work when starting with [
        ilSvLastFound = -1
        For illoop = 0 To lbcListBox.ListCount - 1 Step 1
            ilPos = InStr(1, lbcListBox.List(illoop), slStr, 1)
            If ilPos = 1 Then
                ilSvLastFound = illoop
                Exit For
            Else
                If Left$(lbcListBox.List(illoop), 1) <> "[" Then
                    Exit For
                End If
            End If
        Next illoop
    Else
        'Test if matching string found in the combo box- if so display it (look ahead typing)
        'lbcListBox.ListIndex = 0
        gFndFirst lbcListBox, slStr
        ilBracket = False
        Do
            If gLastFound(lbcListBox) >= 0 Then
                If (Left$(lbcListBox.List(gLastFound(lbcListBox)), 1) = "[") And (Left$(slStr, 1) <> "[") Then
                    gFndNext lbcListBox, slStr
                    ilBracket = True
                Else
                    ilBracket = False
                End If
            Else
                ilBracket = False
            End If
        Loop While ilBracket
        'Test if another name matches encase names are not in sorted order
        ilSearch = True
        ilSvLastFound = gLastFound(lbcListBox)
        Do
            If gLastFound(lbcListBox) >= 0 Then
                If StrComp(slStr, lbcListBox.List(gLastFound(lbcListBox)), 1) = 0 Then
                    ilSvLastFound = gLastFound(lbcListBox)   'lbcListBox.LastFound
                    Exit Do
                End If
                gFndNext lbcListBox, slStr
            Else
                Exit Do
            End If
        Loop While ilSearch
    End If
    If ilSvLastFound >= 0 Then
        'If item found not same as current selected- change current
        If (ilIndex <> ilSvLastFound) Or ((ilIndex = ilSvLastFound) And Not ilBSMode) Then
            lbcListBox.ListIndex = ilSvLastFound 'This will cause a change event (reason for imChgMode)
        End If
'        If (ilIndex <> lbcListBox.LastFound) Or ((ilIndex = lbcListBox.LastFound) And Not ilBSMode) Then 'If indices not equal- highlight look ahead text
'            lbcListBox.SelStart = ilLen
'            lbcListBox.SelLength = Len(lbcListBox.Text)
'        End If
        ilErrHighLightIndex = ilSvLastFound
    Else
        Beep
        If ilErrHighLightIndex >= 0 And lbcListBox.ListCount > 0 Then
            lbcListBox.ListIndex = ilErrHighLightIndex
            ilSelStart = 0
        End If
    End If
    If lbcListBox.ListIndex >= 0 Then
        edcTextBox.Text = lbcListBox.List(lbcListBox.ListIndex)
    Else
        edcTextBox.Text = lbcListBox.Text
    End If
    If ilSelStart <= Len(edcTextBox.Text) Then
        edcTextBox.SelStart = ilSelStart
        edcTextBox.SelLength = Len(edcTextBox.Text)
    Else
        edcTextBox.SelStart = 0
        edcTextBox.SelLength = Len(edcTextBox.Text)
    End If
End Sub

Private Sub gFndFirst(lbcList As Control, slInMatch As String)
    cgList = lbcList
    If TypeOf lbcList Is ComboBox Then
        igFoundRow = SendMessageByString(lbcList.HWnd, CB_FINDSTRING, -1, slInMatch)
    Else
        igFoundRow = SendMessageByString(lbcList.HWnd, LB_FINDSTRING, -1, slInMatch)
    End If
End Sub

Private Sub gFndNext(lbcList As Control, slInMatch As String)
    Dim slNext As String
    Dim ilTestRow As Integer

    If cgList <> lbcList Then
        igFoundRow = -1
        Exit Sub
    End If
    ilTestRow = igFoundRow + 1
    Do While ilTestRow < lbcList.ListCount
        slNext = lbcList.List(ilTestRow)
        If InStr(1, slNext, slInMatch, vbTextCompare) = 1 Then
            igFoundRow = ilTestRow
            Exit Sub
        End If
        ilTestRow = ilTestRow + 1
    Loop
    igFoundRow = -1

End Sub

Private Function gLastFound(lbcList As Control) As Integer
    If cgList <> lbcList Then
        gLastFound = -1
    Else
        gLastFound = igFoundRow
    End If
End Function

'****************************************************************************
' Methods
'
'
'
'****************************************************************************
Public Property Get Text() As String
    Text = m_Text
End Property

'****************************************************************************
'
'****************************************************************************
Public Property Let Text(ByVal New_Text As String)
    Dim llRow As Long
    Dim ilLen As Integer
    
    m_Text = Trim(New_Text)
    ec_InputBox.Text = m_Text
    ' Look up this value in the list control so it will also be the selected text.
    ilLen = Len(m_Text)
    llRow = SendMessageByString(lbcList.HWnd, LB_FINDSTRING, -1, m_Text)
    If llRow >= 0 Then
        imSelectedIndex = llRow
        lbcList.ListIndex = llRow
        ec_InputBox.Text = lbcList.Text
        ec_InputBox.SelStart = ilLen
        ec_InputBox.SelLength = Len(lbcList.Text)
        If llRow + 1 <= UBound(imToolTipArray) Then
            ec_InputBox.ToolTipText = imToolTipArray(llRow + 1)
        End If
    End If
    PropertyChanged "Text"
End Property

Public Sub SelText(sText As String)
    Dim llRow As Long
    Dim ilLen As Integer
    
    bmControlIsReady = False
    ilLen = Len(sText)
    ' Look up this value in the list control
    llRow = SendMessageByString(lbcList.HWnd, LB_FINDSTRING, -1, sText)
    If llRow >= 0 Then
        imSelectedIndex = llRow
        lbcList.ListIndex = llRow
        ec_InputBox.Text = lbcList.Text
        ec_InputBox.SetFocus
        ec_InputBox.SelStart = ilLen
        ec_InputBox.SelLength = Len(lbcList.Text)
        If llRow + 1 <= UBound(imToolTipArray) Then
            ec_InputBox.ToolTipText = imToolTipArray(llRow + 1)
        End If
    End If
    m_Text = lbcList.Text
    bmControlIsReady = True
End Sub

'****************************************************************************
' Properties
'
'
'
'****************************************************************************
Private Sub UserControl_InitProperties()
    bmControlIsReady = False
    m_Text = UserControl.Name
    ec_InputBox.Text = m_Text
    ec_InputBox.BackColor = llBackColor
    ec_InputBox.ForeColor = llForeColor
    lbcList.BackColor = llBackColor
    lbcList.ForeColor = llForeColor
    imBSMode = False
    bmControlIsReady = True
End Sub

'****************************************************************************
' Load property values from storage
'****************************************************************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    llBackColor = PropBag.ReadProperty("BackColor", RGB(0, 0, 0))
    llForeColor = PropBag.ReadProperty("ForeColor", RGB(0, 0, 0))
    ec_InputBox.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    ec_InputBox.Text = m_Text
End Sub

'****************************************************************************
' Write property values to storage
'****************************************************************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("BackColor", llBackColor, RGB(0, 0, 0))
    Call PropBag.WriteProperty("ForeColor", llBackColor, RGB(0, 0, 0))
    Call PropBag.WriteProperty("BorderStyle", ec_InputBox.BorderStyle)
End Sub

'****************************************************************************
'
'****************************************************************************
Public Property Get BorderStyle() As BorderStyleConstants
   BorderStyle = ec_InputBox.BorderStyle
End Property
Public Property Let BorderStyle(BorderStyle As BorderStyleConstants)
    ec_InputBox.BorderStyle = BorderStyle
    PropertyChanged "BorderStyle"
End Property
Public Property Get BackColor() As SystemColorConstants
   BackColor = llBackColor
End Property
Public Property Let BackColor(BKColor As SystemColorConstants)
    llBackColor = BKColor
    ec_InputBox.BackColor = llBackColor
    lbcList.BackColor = llBackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As SystemColorConstants
   ForeColor = llForeColor
End Property
Public Property Let ForeColor(FGColor As SystemColorConstants)
    llForeColor = FGColor
    ec_InputBox.ForeColor = llForeColor
    lbcList.BackColor = llForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get FontBold() As Integer
    FontBold = imFontBold
End Property
Public Property Let FontBold(dInFontBold As Integer)
    Dim illoop As Integer

    imFontBold = dInFontBold
    ec_InputBox.FontBold = imFontBold
    lbcList.FontBold = imFontBold
    
    PropertyChanged "FontBold"
End Property
Public Sub SetEditBoxBorderStyle(BorderStyle As Integer)
    ec_InputBox.BorderStyle = BorderStyle
End Sub

Public Property Get FontName() As String
   FontName = smFontName
End Property
Public Property Let FontName(sFontName As String)
    smFontName = FontName
    ec_InputBox.FontName = smFontName
    lbcList.FontName = smFontName
    PropertyChanged "FontName"
End Property

Public Property Get FontSize()
   FontSize = imFontSize
End Property
Public Property Let FontSize(ByRef iFontSize)
    imFontSize = iFontSize
    ec_InputBox.FontSize = imFontSize
    lbcList.FontSize = imFontSize
    btn_DownArrow.FontSize = imFontSize
    PropertyChanged "FontSize"
End Property

Public Property Get ListCount()
   ListCount = lbcMasterList.ListCount
End Property
Public Property Get VisibleListCount()
   ListCount = lbcList.ListCount
End Property
Public Property Get GetItemData(idx As Integer)
   GetItemData = lbcList.ItemData(idx)
End Property
Public Property Let SetItemData(ItemData As Long)
   lbcList.ItemData(lbcList.NewIndex) = ItemData
   lbcMasterList.ItemData(lbcMasterList.NewIndex) = ItemData
End Property
Public Property Get ListIndex()
   ListIndex = lbcList.ListIndex
End Property
Public Property Get GetName(idx As Integer)
    If idx >= 0 And idx <= lbcList.ListCount - 1 Then
        GetName = lbcList.List(idx)
    Else
        GetName = ""
    End If
End Property
Public Property Get GetVisibleName(idx As Integer)
    If idx >= 0 And idx <= lbcList.ListCount - 1 Then
        GetVisibleName = lbcList.List(idx)
    Else
        GetVisibleName = ""
    End If
End Property
Public Property Let SetListIndex(ListIndex As Long)
    If ListIndex < 0 Or ListIndex > lbcList.ListCount - 1 Then
        lbcList.ListIndex = -1
        ec_InputBox.Text = ""
    Else
        lbcList.ListIndex = ListIndex
        'm_Text = lbcList.Text
        ec_InputBox.Text = lbcList.Text
    End If
End Property
Public Property Let RemoveListIndex(idx As Integer)
    If idx >= 0 And idx <= lbcList.ListCount - 1 Then
        If lbcList.List(idx) = ec_InputBox.Text Then
            lbcList.RemoveItem idx
            lbcList.ListIndex = -1
            ec_InputBox.Text = ""
        Else
            lbcList.RemoveItem idx
            lbcList.ListIndex = -1
        End If
    End If
End Property

Public Property Let ReSizeFont(slSource As String)
    If slSource = "A" Then
        mSetFonts
        btn_DownArrow.FontSize = 6
    End If
End Property

Private Sub mSetFonts()
    Dim Ctrl As Control
    Dim ilFontSize As Integer
    Dim ilColorFontSize As Integer
    Dim ilBold As Integer
    Dim ilChg As Integer
    Dim slStr As String
    Dim slFontName As String
    
    
    On Error Resume Next
    ilFontSize = 14
    ilBold = True
    ilColorFontSize = 10
    slFontName = "Arial"
    If Screen.Height / Screen.TwipsPerPixelY <= 480 Then
        ilFontSize = 8
        ilBold = False
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 600 Then
        ilFontSize = 8
        ilBold = False
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 768 Then
        ilFontSize = 10
        ilBold = False
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 800 Then
        ilFontSize = 10
        ilBold = True
        ilColorFontSize = 8
    ElseIf Screen.Height / Screen.TwipsPerPixelY <= 1024 Then
        ilFontSize = 12
        ilBold = True
    End If
    For Each Ctrl In UserControl.Controls
        'If TypeOf Ctrl Is MSHFlexGrid Then
        '    Ctrl.Font.Name = slFontName
        '    Ctrl.FontFixed.Name = slFontName
        '    Ctrl.Font.Size = ilFontSize
        '    Ctrl.FontFixed.Size = ilFontSize
        '    Ctrl.Font.Bold = ilBold
        '    Ctrl.FontFixed.Bold = ilBold
        'ElseIf TypeOf Ctrl Is TabStrip Then
        '    Ctrl.Font.Name = slFontName
        '    Ctrl.Font.Size = ilFontSize
        '    Ctrl.Font.Bold = ilBold
        'Else
            ilChg = 0
            If (Ctrl.ForeColor = vbBlack) Or (Ctrl.ForeColor = &H80000008) Or (Ctrl.ForeColor = &H80000012) Or (Ctrl.ForeColor = &H8000000F) Then
                ilChg = 1
            Else
                ilChg = 2
            End If
            slStr = Ctrl.Name
            If (InStr(1, slStr, "Arrow", vbTextCompare) > 0) Or ((InStr(1, slStr, "Dropdown", vbTextCompare) > 0) And (TypeOf Ctrl Is CommandButton)) Then
                ilChg = 0
            End If
            If ilChg = 1 Then
                Ctrl.FontName = slFontName
                Ctrl.FontSize = ilFontSize
                Ctrl.FontBold = ilBold
            ElseIf ilChg = 2 Then
                Ctrl.FontName = slFontName
                Ctrl.FontSize = ilColorFontSize
                Ctrl.FontBold = False
            End If
        'End If
    Next Ctrl
End Sub

Private Sub ShowMasterList()
    Dim illoop As Integer
    mHideDropDown
    lbcList.Clear
    For illoop = 0 To lbcMasterList.ListCount - 1
        lbcList.AddItem lbcMasterList.List(illoop)
        lbcList.ItemData(lbcList.NewIndex) = lbcMasterList.ItemData(illoop)
    Next illoop
    'show list
    If bmIsDroppedDown = True Then
        mHideDropDown
    End If
End Sub

Private Sub FilterList(lsFilterText As String)
    'mHideDropDown
    lbcList.Visible = False 'for Performance
    lbcList.Clear
    Dim illoop As Integer
    Dim ilTermsLoop As Integer
    Dim alTerms
    Dim blFound As Boolean
    blFound = False
    If Trim(lsFilterText) = "" Then
        ShowMasterList
        'If lbcList.Visible = True Then bmIsDroppedDown = True
        Exit Sub
    End If
    For illoop = 0 To lbcMasterList.ListCount - 1
        If InStr(1, lsFilterText, " ") = 0 Then
            'Single Term Search
            If InStr(1, LCase(lbcMasterList.List(illoop)), LCase(lsFilterText)) > 0 Then
                lbcList.AddItem lbcMasterList.List(illoop)
                lbcList.ItemData(lbcList.NewIndex) = lbcMasterList.ItemData(illoop)
            End If
        Else
            'Multi Term Search
            alTerms = Split(lsFilterText, " ")
            For ilTermsLoop = 0 To UBound(alTerms)
                blFound = False
                If InStr(1, LCase(lbcMasterList.List(illoop)), LCase(alTerms(ilTermsLoop))) > 0 Then blFound = True
                If blFound = False Then Exit For
            Next ilTermsLoop
            If blFound = True Then
                lbcList.AddItem lbcMasterList.List(illoop)
                lbcList.ItemData(lbcList.NewIndex) = lbcMasterList.ItemData(illoop)
            End If
        End If
    Next illoop
    'show list
    If bmIsDroppedDown = False Then
        mShowDropDown
    End If
    lbcList.Visible = True
    m_Text = ""
    imSelectedIndex = -1
    RaiseEvent OnChange
End Sub

Private Sub mSetListResult()
    Dim ilHightErr As Integer
    ilHightErr = 0
        
    If Trim(ec_InputBox.Text) = "" Then
        imSelectedIndex = -1
        lbcList.ListIndex = -1
        m_Text = ""
    Else
        If lbcList.ListIndex = -1 Then
            'ec_InputBox.Text = lbcList.List(lbcList.ListIndex)
            gMatchLookAhead ec_InputBox, lbcList, imBSMode, ilHightErr
        End If
        m_Text = lbcList.Text
        imSelectedIndex = lbcList.ListIndex
        ec_InputBox.Text = lbcList.Text
        
        If ec_InputBox.Visible Then
            'ec_InputBox.SetFocus
            If lbcList.ListIndex + 1 <= UBound(imToolTipArray) Then
                ec_InputBox.ToolTipText = imToolTipArray(lbcList.ListIndex + 1)
            End If

        End If
        'hide DDL
        mHideDropDown
    End If
    
    RaiseEvent OnChange
End Sub

Sub mShowDropDown()
    lbcList.Visible = True
    bmIsDroppedDown = True
    If imUserListNoRows <> -1 Then
        If lbcList.ListCount < imUserListNoRows Then
            SetDropDownNumRows lbcList.ListCount, False
        Else
            SetDropDownNumRows imUserListNoRows, False
        End If
    Else
        If lbcList.ListCount < imDefaultListNoRows Then
            SetDropDownNumRows lbcList.ListCount, False
        Else
            SetDropDownNumRows imDefaultListNoRows, False
        End If
    End If
    lbcList.Left = ec_InputBox.Left
    lbcList.Width = imDropDownWidth
    lbcList.Height = imDropDownHeight
    If imDropDownWidth > ec_InputBox.Width Then
        Width = imDropDownWidth
    End If
    'Height = imDropDownHeight + (Screen.TwipsPerPixelY * lbcList.FontSize) * 2
    Height = imDropDownHeight + ec_InputBox.Height
    If smPopupListDirection <> "A" Then
        lbcList.Top = (ec_InputBox.Top + ec_InputBox.Height)
    Else
        ec_InputBox.Top = Height - ec_InputBox.Height ' + (Screen.TwipsPerPixelY * lbcList.FontSize) * 2 - ec_InputBox.Height
        If ec_InputBox.BorderStyle = 0 Then
            btn_DownArrow.Top = ec_InputBox.Top
        Else
            btn_DownArrow.Top = ec_InputBox.Top + (Screen.TwipsPerPixelY * 2)
        End If
        'lbcList.Top = 0    'ec_InputBox.Top - lbcList.Height
        lbcList.Top = ec_InputBox.Top - lbcList.Height
        RaiseEvent ReSetLoc
    End If
    If (lbcList.ListCount > 0) And (imSelectedIndex >= 0) Then
        If imSelectedIndex <= lbcList.ListCount - 1 Then
            lbcList.Selected(imSelectedIndex) = True
            m_Text = lbcList.Text
            imSelectedIndex = lbcList.ListIndex
            
            RaiseEvent OnChange
        End If
        'ec_InputBox.Text = lbcList.Text
    End If
End Sub

Sub mHideDropDown()
    bmIsDroppedDown = False
    lbcList.Visible = False
End Sub
