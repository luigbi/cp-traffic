Attribute VB_Name = "modCRC"
Option Explicit
Option Compare Text

Private Crc32Table(255) As Long
Private lCrc32Value As Long


Public Function InitCrc32(Optional ByVal Seed As Long = &HEDB88320, Optional ByVal Precondition As Long = &HFFFFFFFF) As Long

   Dim iBytes As Integer
   Dim iBits As Integer
   Dim lCrc32 As Long
   Dim lTempCrc32 As Long

   On Error Resume Next

   ' Iterate 256 times
   For iBytes = 0 To 255

      ' Initiate lCrc32 to counter variable
      lCrc32 = iBytes

      ' Now iterate through each bit in counter byte
      For iBits = 0 To 7
         ' Right shift unsigned long 1 bit
         lTempCrc32 = lCrc32 And &HFFFFFFFE
         lTempCrc32 = lTempCrc32 \ &H2
         lTempCrc32 = lTempCrc32 And &H7FFFFFFF

         ' Now check if temporary is less than zero and then mix Crc32 checksum with Seed value
         If (lCrc32 And &H1) <> 0 Then
            lCrc32 = lTempCrc32 Xor Seed
         Else
            lCrc32 = lTempCrc32
         End If
      Next

      ' Put Crc32 checksum value in the holding array
      Crc32Table(iBytes) = lCrc32
   Next

   ' After this is done, set function value to the precondition value
   InitCrc32 = Precondition

End Function

Public Function AddCrc32(ByVal Item As String, ByVal Crc32 As Long) As Long


   Dim bCharValue As Byte
   Dim iCounter As Integer
   Dim lIndex As Long
   Dim lAccValue As Long
   Dim lTableValue As Long

   On Error Resume Next

   ' Iterate through the string that is to be checksum-computed
   For iCounter = 1 To Len(Item)

      ' Get ASCII value for the current character
      bCharValue = Asc(Mid$(Item, iCounter, 1))

      ' Right shift an Unsigned Long 8 bits
      lAccValue = Crc32 And &HFFFFFF00
      lAccValue = lAccValue \ &H100
      lAccValue = lAccValue And &HFFFFFF

      ' Now select the right adding value from the holding table
      lIndex = Crc32 And &HFF
      lIndex = lIndex Xor bCharValue
      lTableValue = Crc32Table(lIndex)

      ' Then mix new Crc32 value with previous accumulated Crc32 value
      Crc32 = lAccValue Xor lTableValue
   Next

   ' Set function value the the new Crc32 checksum
   AddCrc32 = Crc32

End Function

Public Function GetCrc32(ByVal Crc32 As Long) As Long
   
   On Error Resume Next

   ' Set function to the current Crc32 value
   GetCrc32 = Crc32 Xor &HFFFFFFFF

End Function

Public Sub Main()

   On Error Resume Next
   lCrc32Value = InitCrc32()
   mReadFile

End Sub


Private Sub mReadFile()

    Dim tlTxtStream As TextStream
    Dim fs As New FileSystemObject
    Dim slLocation As String
    Dim slStr As String
    Dim llCRC As Long
    Dim ilLen As Integer
    
    slLocation = "c:\test.txt"

    If fs.FileExists(slLocation) Then
        Set tlTxtStream = fs.OpenTextFile(slLocation, ForReading, False)
        Do While tlTxtStream.AtEndOfStream <> True
            slStr = tlTxtStream.ReadLine
            ilLen = Len(slStr)
            lCrc32Value = AddCrc32(slStr, lCrc32Value)
        Loop
        tlTxtStream.Close
    Else
        MsgBox "I can't find the stinking file!"
    End If
    llCRC = GetCrc32(lCrc32Value)
End Sub



