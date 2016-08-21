Attribute VB_Name = "modCRC"
Option Explicit

Private pInititialized As Boolean
Private pTable(0 To 255) As Long


Public Sub CRCInit(Optional ByVal Poly As Long = &HEDB88320)
 'Deklarationen:
 Dim crc As Long
  Dim i As Integer
  Dim j As Integer

  For i = 0 To 255

    crc = i
    For j = 0 To 7

      If crc And &H1 Then
        'CRC = (CRC >>> 1) ^ Poly
        crc = ((crc And &HFFFFFFFE) \ &H2 And &H7FFFFFFF) Xor Poly
      Else
        'CRC = (CRC >>> 1)
        crc = crc \ &H2 And &H7FFFFFFF
      End If

    Next j
    pTable(i) = crc

  Next i
  pInititialized = True
End Sub


Public Function CRC32File(Path As String) As Long
  AddDebug "Calculating CRC of file: " & Path
  Dim Buffer() As Byte
  Dim BufferSize As Long
  Dim crc As Long
  Dim FileNr As Integer
  Dim Length As Long
  Dim i As Long

  If Not pInititialized Then CRCInit


  BufferSize = &H1000 '4 KB
  ReDim Buffer(1 To BufferSize)

  FileNr = FreeFile
  Open Path For Binary As #FileNr

    Length = LOF(FileNr)

    crc = &HFFFFFFFF

    Do While Length

      If Length < BufferSize Then
        BufferSize = Length
        ReDim Buffer(1 To Length)
      End If
      Get #FileNr, , Buffer

      For i = 1 To BufferSize
        crc = ((crc And &HFFFFFF00) \ &H100) And &HFFFFFF Xor pTable(Buffer(i) Xor crc And &HFF&)
      Next i

      Length = Length - BufferSize

    Loop
    CRC32File = Not crc

  Close #FileNr
End Function


