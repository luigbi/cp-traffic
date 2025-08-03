VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents omPictureBox As PictureBox
Attribute omPictureBox.VB_VarHelpID = -1
Public Event Paint()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event GotFocus()

Public Property Set NewBox(myPict As PictureBox)
    Set omPictureBox = myPict
End Property
Public Property Let ForeColor(llColor As Long)
    omPictureBox.ForeColor = llColor
End Property
Public Property Get ForeColor() As Long
    ForeColor = omPictureBox.ForeColor
End Property
Public Property Let FontName(slFont As String)
    omPictureBox.FontName = slFont
End Property
Public Property Get FontName() As String
    FontName = omPictureBox.FontName
End Property
Public Property Let FontSize(ilSize As Integer)
    omPictureBox.FontSize = ilSize
End Property
Public Property Get FontSize() As Integer
    FontSize = omPictureBox.FontSize
End Property
Public Property Let FontBold(blValue As Boolean)
    omPictureBox.FontBold = blValue
End Property
Public Property Get FontBold() As Boolean
    FontBold = omPictureBox.FontBold
End Property
Public Property Let FontItalic(blValue As Boolean)
    omPictureBox.FontItalic = blValue
End Property
Public Property Get FontItalic() As Boolean
    FontItalic = omPictureBox.FontItalic
End Property
Public Property Let FontUnderline(blValue As Boolean)
    omPictureBox.FontUnderline = blValue
End Property
Public Property Get FontUnderline() As Boolean
    FontUnderline = omPictureBox.FontUnderline
End Property
Public Property Let Fontstrikethru(blValue As Boolean)
    omPictureBox.Fontstrikethru = blValue
End Property
Public Property Get Fontstrikethru() As Boolean
    Fontstrikethru = omPictureBox.Fontstrikethru
End Property
Public Property Let CurrentX(llX As Long)
    omPictureBox.CurrentX = llX
End Property
Public Property Get CurrentX() As Long
    CurrentX = omPictureBox.CurrentX
End Property
Public Property Let CurrentY(llY As Long)
    omPictureBox.CurrentY = llY
End Property
Public Property Get CurrentY() As Long
    CurrentY = omPictureBox.CurrentY
End Property
Public Property Let FillStyle(ilFillStyle As Integer)
    omPictureBox.FillStyle = ilFillStyle
End Property
Public Property Get FillStyle() As Integer
    FillStyle = omPictureBox.FillStyle
End Property
Public Property Let FillColor(llColor As Long)
    omPictureBox.FillColor = llColor
End Property
Public Property Get FillColor() As Long
    FillColor = omPictureBox.FillColor
End Property
Public Sub mPrint(slText As String)
    omPictureBox.Print slText
End Sub
Public Sub Cls()
    omPictureBox.Cls
End Sub
Public Sub mLine(blStep1 As Boolean, x1 As Single, y1 As Single, blStep2 As Boolean, x2 As Single, y2 As Single, slColor As String, slBox As String)
    Dim slCase As String
    Dim llColor As Long
    
    If (blStep1 = True) Then   'vbTrue) Then
        slCase = "T1"
    Else
        slCase = "F1"
    End If
    If (blStep2 = True) Then   'vbTrue) Then
        slCase = slCase & ":T2"
    Else
        slCase = slCase & ":F2"
    End If
    If slColor <> "" Then
        llColor = Val(slColor)
        slCase = slCase & "-C"
    End If
    If slBox = "B" Then
        slCase = slCase & "#B"
    ElseIf slBox = "BF" Then
        slCase = slCase & "#BF"
    End If
    Select Case slCase
        Case "T1:T2"
            omPictureBox.Line Step(x1, y1)-Step(x2, y2)
        Case "F1:T2"
            omPictureBox.Line (x1, y1)-Step(x2, y2)
        Case "F1:F2"
            omPictureBox.Line (x1, y1)-(x2, y2)
        Case "T1:F2"
            omPictureBox.Line Step(x1, y1)-(x2, y2)
    
        Case "T1:T2-C"
            omPictureBox.Line Step(x1, y1)-Step(x2, y2), llColor
        Case "F1:T2-C"
            omPictureBox.Line (x1, y1)-Step(x2, y2), llColor
        Case "F1:F2-C"
            omPictureBox.Line (x1, y1)-(x2, y2), llColor
        Case "T1:F2-C"
            omPictureBox.Line Step(x1, y1)-(x2, y2), llColor
    
        Case "T1:T2-C#B"
            omPictureBox.Line Step(x1, y1)-Step(x2, y2), llColor, B
        Case "F1:T2-C#B"
            omPictureBox.Line (x1, y1)-Step(x2, y2), llColor, B
        Case "F1:F2-C#B"
            omPictureBox.Line (x1, y1)-(x2, y2), llColor, B
        Case "T1:F2-C#B"
            omPictureBox.Line Step(x1, y1)-(x2, y2), llColor, B
    
        Case "T1:T2-C#BF"
            omPictureBox.Line Step(x1, y1)-Step(x2, y2), llColor, BF
        Case "F1:T2-C#BF"
            omPictureBox.Line (x1, y1)-Step(x2, y2), llColor, BF
        Case "F1:F2-C#BF"
            omPictureBox.Line (x1, y1)-(x2, y2), llColor, BF
        Case "T1:F2-C#BF"
            omPictureBox.Line Step(x1, y1)-(x2, y2), llColor, BF
    
    
        Case "T1:T2#B"
            omPictureBox.Line Step(x1, y1)-Step(x2, y2), , B
        Case "F1:T2#B"
            omPictureBox.Line (x1, y1)-Step(x2, y2), , B
        Case "F1:F2#B"
            omPictureBox.Line (x1, y1)-(x2, y2), , B
        Case "T1:F2#B"
            omPictureBox.Line Step(x1, y1)-(x2, y2), , B
    
    End Select
End Sub
Public Sub Refresh()
    omPictureBox.Refresh
End Sub
Public Function TextHeight(slText As String) As Single
    TextHeight = omPictureBox.TextHeight(slText)
End Function
Public Function TextWidth(slText As String) As Single
    TextWidth = omPictureBox.TextWidth(slText)
End Function
Private Sub omPictureBox_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub omPictureBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub omPictureBox_Paint()
    RaiseEvent Paint
End Sub
Private Sub omPictureBox_Gotfocus()
    RaiseEvent GotFocus
End Sub

