VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdler32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Adler-32 Checksum (1995-1998 Mark Adler)
' ---------------------------------------------------------------
' Faithfully ported to Visual Basic by,
' 2002 Ed Preston - epreston@selectedsystems.com
'
' Interface
'
'   Adler32(ByVal lngAdler32 As Long, ByRef bArrayIn() As Byte, _
'   ByVal dblLength As Double) As Long
'
' Notes
'
'   The Adler-32 algorithm is much faster than the CRC32
'   algorithm yet still provides an extremely low probability of
'   undetected errors. Used in the ZLIB/Deflate Compressed Data
'   Format Specification version 3.3. More info can be found at
'   http://www.info-zip.org/pub/infozip/zlib/rfc-zlib.htm
'
'   Mark Adler states that the Adler32 Checksum should be initialized
'   with 1 but he does not do this in his zLib implementation. Humm..
'
'   When passing byte arrays be carefull of the file size.  It is best
'   to break the file into chunks and call Adler32 multiple times.
'   Byte array is limited to the amount of addressable memory in the
'   process space.  The outcome is, pass a 712 meg array and watch your
'   machine grind to a halt while the hard disk tries to page things
'   in and out of memory.
'
' Limitations
'
'   Files size assumed to be less than 2 gig.
' ---------------------------------------------------------------


' Constants
Const BASE = 65521  ' largest prime smaller than 65536

' Largest n such that 255n(n+1)/2 + (n+1)(BASE-1) <= 2^32-1
Const NMAX = 5552

' What this means is that in the worst case sinario, each byte
' could have a value of 255.  As we are adding numbers together
' we could go beyond the value of a unsigned long or 2^32.
' Thats a value of 4294967296.  To be sure we are safe, we take
' things in handfulls of 5552 bytes at a time.


' ----------------------------
' Public Methods
' ----------------------------

' Return value needs to be a Long to work with our Hex$ method.
Public Function Adler32(ByVal lngAdler32 As Long, ByRef bArrayIn() As Byte, _
    ByVal dblLength As Double)
    
    
    Dim intPos As Integer
    Dim lngPosInArray As Long
    Dim lngLengthRemaining As Long
    Dim dblLow As Double
    Dim dblHigh As Double
    
    If lngAdler32 <> 0 Then
        dblLow = lngAdler32 And 65535
        dblHigh = RShiftNoRound(lngAdler32, 16) And 65535
    End If
    
    ' Array could be empty
    If UBound(bArrayIn) < LBound(bArrayIn) Then
        Adler32 = 1
    Else
        ' Get the initial length of the data
        lngLengthRemaining = dblLength
        
        Do While (lngLengthRemaining > 0)
            ' Are we at the end of the file?
            If lngLengthRemaining < NMAX Then
                ' Process the remaining data
                intPos = lngLengthRemaining
                lngLengthRemaining = 0
            Else
                ' No, process a chunk of data the size of NMAX
                intPos = NMAX
                ' We start counting at zero so add one to the NMAX
                lngLengthRemaining = lngLengthRemaining - (NMAX + 1)
            End If
            
            Do
                dblLow = dblLow + bArrayIn(lngPosInArray)
                dblHigh = dblHigh + dblLow
                
                lngPosInArray = lngPosInArray + 1
                intPos = intPos - 1
            Loop While intPos >= 0
            
            ' Use the overflow safe modulus function instead of the
            ' operator.
            dblLow = Modulus(dblLow, BASE)
            dblHigh = Modulus(dblHigh, BASE)
            
        Loop
        
        ' Done, return the result (rightmost 4 bytes).
        Adler32 = LShift4Byte(dblHigh, 16) Or dblLow
    End If
End Function

' ----------------------------
' Support Routines
' ----------------------------

' We can not use the Mod operator because floating point number could
' be greater than the maximum value of a Long (2,147,483,647).  The
' Mod operator in VB converts numerator and denominator to Longs just to
' mess with your head. ;)
Private Function Modulus(ByVal dblValue As Double, ByVal dblModValue As Double) As Double
    Modulus = dblValue - (dblModValue * Fix(dblValue / dblModValue))
End Function

' Thank you Redeye.  It seems VB does not automaticly trunicate leftmost bits when
' trying to type cast a Double into a Long.  Large numbers give overflow. Not good.
Private Function LShift4Byte(ByVal pnValue As Double, ByVal pnShift As Integer) As Long
    ' Mask For fixing msb
    Dim lngMask As Long
    
    ' Determine whether the sign bit should be set in the bit-shifted result
    If pnValue And (2 ^ (31 - pnShift)) Then
        lngMask = &H80000000
    End If
    
    ' Unset leftmost bits that will be discarded and do the bit shift.
    LShift4Byte = ((pnValue And ((2 ^ (31 - pnShift)) - 1)) * (2 ^ pnShift)) Or lngMask
End Function

Private Function RShiftNoRound(ByVal pnValue As Double, ByVal pnShift As Integer) As Double
    ' Equivilant to C's Bitwise >> operator
    RShiftNoRound = Int(pnValue / (2 ^ pnShift))
End Function







