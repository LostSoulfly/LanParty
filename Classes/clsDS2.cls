VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' DS2 Cipher (aka Digitally Secure Encryption)
' By: David Greenwood <dsguk@lycos.com>
' and David Midkiff <mdj2023@hotmail.com>
'
' Copyright © 2001-2002 David Greenwood and David Midkiff.
' All rights reserved.
'
' This algorithm is free for use in any non-commercial project but
' you must receive permission from both David Greenwood and David
' Midkiff to use this algorithm in commercial projects. Information
' on the algorithm can be found in the attached text file or by visiting
' our website at http://go.to/ds2cipher.

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGUID As GUID) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function StringFromGUID2 Lib "OLE32.DLL" (pGUID As GUID, ByVal PointerToString As Long, ByVal MaxLength As Long) As Long

Public Event Progress(Percent As Integer)

' Default rounds is 4. Lower rounds provide faster speed but
' less security. Higher rounds provide insane security but less speed.
' We would recommend simply leaving it at 4.
Const Rounds = 4

Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long

Private sbox_a(255, 255) As Byte
Private sbox_b(255, 255) As Byte

Private Type GUID
    Guid1 As Long
    Guid2 As Long
    Guid3 As Long
    Guid4(0 To 7) As Byte
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Function GSeed() As Double
    ' Retrieves the GUID value into a double variable
    Dim udtGUID As GUID, sGUID As String, lResult As Long

    lResult = CoCreateGuid(udtGUID)
    If lResult Then
        GSeed = 0
    Else
        With udtGUID
            GSeed = CInt(CLng((CDbl(.Guid1) + CDbl(.Guid2) + CDbl(.Guid3)) / CDbl(32766)) Mod 32766)
        End With
    End If
End Function
Private Function Hash(Data As String) As Double
    ' Simple, primitive yet fast 80-bit one-way hash function designed by
    ' David Midkiff for use only in Ethereal Chaos Generator. To use this
    ' elsewhere you must have David's permission.
    Dim i As Double, NBox As Long, OBox As Long, PBox As Long, QBox As Long, RBox As Long
    For i = 1 To Len(Data)
        NBox = NBox + Asc(Mid$(Data, i, 1))
        NBox = (NBox * 1765 + 1066) Mod 1048576
        OBox = OBox + Asc(Mid$(Data, i, 1))
        OBox = (OBox * 815 + 1801) Mod 1048576
        PBox = PBox + Asc(Mid$(Data, i, 1))
        PBox = (PBox * 1364 + 1083) Mod 1048576
        QBox = QBox + Asc(Mid$(Data, i, 1))
        QBox = (QBox * 1624 + 1564) Mod 1048576
        RBox = RBox + Asc(Mid$(Data, i, 1))
        RBox = (RBox * 1454 + 1905) Mod 1048576
    Next
    
    NBox = (NBox * 1765 + 1066) Mod 1048576: NBox = (NBox * 1765 + 1066) Mod 1048576
    NBox = (NBox * 1765 + 1066) Mod 1048576: NBox = (NBox * 1765 + 1066) Mod 1048576
    NBox = (NBox * 1765 + 1066) Mod 1048576: NBox = (NBox * 1765 + 1066) Mod 1048576
    NBox = (NBox * 1765 + 1066) Mod 1048576
    
    OBox = (OBox * 815 + 1801) Mod 1048576: OBox = (OBox * 815 + 1801) Mod 1048576
    OBox = (OBox * 815 + 1801) Mod 1048576: OBox = (OBox * 815 + 1801) Mod 1048576
    OBox = (OBox * 815 + 1801) Mod 1048576: OBox = (OBox * 815 + 1801) Mod 1048576
    OBox = (OBox * 815 + 1801) Mod 1048576
    
    PBox = (PBox * 1364 + 1083) Mod 1048576: PBox = (PBox * 1364 + 1083) Mod 1048576
    PBox = (PBox * 1364 + 1083) Mod 1048576: PBox = (PBox * 1364 + 1083) Mod 1048576
    PBox = (PBox * 1364 + 1083) Mod 1048576: PBox = (PBox * 1364 + 1083) Mod 1048576
    PBox = (PBox * 1364 + 1083) Mod 1048576
    
    QBox = (QBox * 1624 + 1564) Mod 1048576: QBox = (QBox * 1624 + 1564) Mod 1048576
    QBox = (QBox * 1624 + 1564) Mod 1048576: QBox = (QBox * 1624 + 1564) Mod 1048576
    QBox = (QBox * 1624 + 1564) Mod 1048576: QBox = (QBox * 1624 + 1564) Mod 1048576
    QBox = (QBox * 1624 + 1564) Mod 1048576
    
    RBox = (RBox * 1454 + 1905) Mod 1048576: RBox = (RBox * 1454 + 1905) Mod 1048576
    RBox = (RBox * 1454 + 1905) Mod 1048576: RBox = (RBox * 1454 + 1905) Mod 1048576
    RBox = (RBox * 1454 + 1905) Mod 1048576: RBox = (RBox * 1454 + 1905) Mod 1048576
    RBox = (RBox * 1454 + 1905) Mod 1048576

    Hash = NBox + OBox + PBox + QBox + RBox
End Function
Public Function PRNG(Lower As Double, Upper As Double, Optional CustomSeed As Double) As Double
    ' An implementation of ECG for the DS2 cipher. ECG is a secure
    ' pseudo-random number generator for Visual Basic developed by
    ' David Midkiff <mdj2023@hotmail.com>.
    Dim HSeed As Double, SSeed As Double, RSeed As Double
    Dim TSeed As Double, USeed As Double, VSeed As Double
    Dim WSeed As Double, XSeed As Double, FinalSeed As Double

    Dim iLoop As Long, R As String
    Dim Mouse As POINTAPI, Dimensions As RECT
    
    GetCursorPos Mouse
    Randomize (GetTickCount + Mouse.X + Mouse.Y)
    HSeed = Hash(CStr((Int(Rnd * 1000000) + Mouse.X - Mouse.Y)))

    For iLoop = 1 To 1000
        GetWindowRect iLoop, Dimensions
        RSeed = RSeed + GetWindowTextLength(iLoop)
        XSeed = XSeed + Dimensions.Bottom + Dimensions.Top + Dimensions.Left + Dimensions.Right
    Next
    
    GetCursorPos Mouse
    SSeed = Upper + Lower
    TSeed = Mouse.X
    USeed = Mouse.Y
    
    VSeed = GetTickCount
    Randomize (GetTickCount + RSeed + XSeed + TSeed + USeed)
    WSeed = Int(Rnd * 9999999999#) + 1
    
    FinalSeed = GSeed + HSeed + SSeed + RSeed + TSeed + USeed + VSeed + WSeed + XSeed + CustomSeed
    
    Randomize FinalSeed
    PRNG = Int(Rnd * (Upper - Lower + 1)) + Lower
End Function
Private Sub Append(ByRef StringData As String, Optional Length As Long)
    ' Appends data into memory block
    Dim DataLength As Long
    If Length > 0 Then DataLength = Length Else DataLength = Len(StringData)
    If DataLength + hiByte > hiBound Then
        hiBound = hiBound + 1024
        ReDim Preserve byteArray(hiBound)
    End If
    CopyMem ByVal VarPtr(byteArray(hiByte)), ByVal StringData, DataLength
    hiByte = hiByte + DataLength
End Sub
Private Function DeHex(Data As String) As String
    ' Decodes a Hexadecimal string into ASCII
    Dim iCount As Double
    Reset
    For iCount = 1 To Len(Data) Step 2
        Append Chr$(Val("&H" & Mid$(Data, iCount, 2)))
    Next
    DeHex = GData
    Reset
End Function
Public Function EnHex(Data As String) As String
    ' Encodes a string into Hexadecimal
    Dim iCount As Double, sTemp As String
    Reset
    For iCount = 1 To Len(Data)
        sTemp = Hex$(Asc(Mid$(Data, iCount, 1)))
        If Len(sTemp) < 2 Then sTemp = "0" & sTemp
        Append sTemp
    Next
    EnHex = GData
    Reset
End Function
Private Function FileExist(FileName As String) As Boolean
' Checks if a file exists
On Error GoTo errorhandler
GoSub begin
    
errorhandler:
    FileExist = False
    Exit Function

begin:
    Call FileLen(FileName)
    FileExist = True
End Function
Private Property Get GData() As String
    ' Retrieves block from memory
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
    GData = StringData
End Property
Public Function EncryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String) As Boolean
On Error GoTo errorhandler
GoSub begin
    
errorhandler:
    EncryptFile = False
    Exit Function
    
begin:
    If FileExist(InFile) = False Then
        EncryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True And Overwrite = False Then
        EncryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte, bKey() As Byte, bOut() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO))
        Buffer(LOF(1)) = 32
        Get #FileO, , Buffer()
    Close #FileO
    
    bKey() = StrConv(Key, vbFromUnicode)
    bOut() = EncryptByte(Buffer(), bKey())
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        Put #FileO, , bOut()
    Close #FileO
    EncryptFile = True
End Function
Public Function EncryptString(Text As String, Optional Key As String, Optional OutputInHex As Boolean) As String
    On Error Resume Next
    EncryptString = StrConv(EncryptByte(StrConv(Text, vbFromUnicode), StrConv(Key, vbFromUnicode)), vbUnicode)
    If OutputInHex = True Then EncryptString = EnHex(EncryptString)
End Function
Public Function DecryptString(Text As String, Optional Key As String, Optional IsTextInHex As Boolean) As String
    On Error Resume Next
    If IsTextInHex = True Then Text = DeHex(Text)
    DecryptString = StrConv(DecryptByte(StrConv(Text, vbFromUnicode), StrConv(Key, vbFromUnicode)), vbUnicode)
End Function
Public Function DecryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String) As Boolean
On Error GoTo errorhandler
GoSub begin
    
errorhandler:
    DecryptFile = False
    Exit Function
    
begin:
    If FileExist(InFile) = False Then
        DecryptFile = False
        Exit Function
    End If
    If FileExist(OutFile) = True Then
        DecryptFile = False
        Exit Function
    End If
    Dim FileO As Integer, Buffer() As Byte, bKey() As Byte, bOut() As Byte
    FileO = FreeFile
    Open InFile For Binary As #FileO
        ReDim Buffer(0 To LOF(FileO) - 1)
        Get #FileO, , Buffer()
    Close #FileO
    bKey() = StrConv(Key, vbFromUnicode)
    bOut() = DecryptByte(Buffer(), bKey())
    If FileExist(OutFile) = True Then Kill OutFile
    FileO = FreeFile
    Open OutFile For Binary As #FileO
        Put #FileO, , bOut()
    Close #FileO
    DecryptFile = True
End Function
Private Sub Reset()
    ' Clears block from memory
    hiByte = 0
    hiBound = 1024
    ReDim byteArray(hiBound)
End Sub
Public Function EncryptByte(DS() As Byte, Key() As Byte)
    Dim tmp2() As Byte, p As Double, i As Double, Bound As Double, R As Integer, Total As Double, Prog As Double
    
    ' DS2 requires a 16-bit key or greater. If the received key is smaller
    ' than 16-bits then simple key expansion is performed.
    If UBound(Key) < 0 Then
        ReDim Key(1) As Byte: Key(0) = 1: Key(1) = 1
    ElseIf UBound(Key) = 0 Then
        ReDim tmp2(0) As Byte: tmp2(0) = Key(0)
        ReDim Key(1) As Byte: Key(0) = tmp2(0): Key(1) = 1
        ReDim tmp2(0) As Byte
    End If
    
    ' Initialize sbox with Key
    InitBox Key()
    
    ' Generate and apply salt to array
    ReDim tmp2((UBound(DS)) + 6)
    tmp2(0) = PRNG(1, 255)
    tmp2(1) = PRNG(1, 255)
    tmp2(2) = PRNG(1, 255)
    tmp2(UBound(tmp2)) = PRNG(1, 255)
    tmp2(UBound(tmp2) - 1) = PRNG(1, 255)
    tmp2(UBound(tmp2) - 2) = PRNG(1, 255)
    
    Call CopyMem(tmp2(3), DS(0), UBound(DS) + 1)
    ReDim DS(UBound(tmp2)) As Byte
    DS() = tmp2()
    ReDim tmp2(0)
    
    ' Encrypt block x rounds
    Bound = UBound(Key)
    Total = UBound(DS) * Rounds
    For R = 1 To Rounds
        p = 0
        For i = 0 To UBound(DS) - 1
            If p = Bound Then p = 0
            DS(i) = sbox_b(DS(i), sbox_a(DS(i + 1), Key(p)))
            DS(i + 1) = sbox_b(DS(i), DS(i + 1))
            DS(i) = sbox_b(DS(i), sbox_a(DS(i + 1), Key(p + 1)))
            p = p + 1
            Prog = Prog + 1: RaiseEvent Progress((Prog / Total) * 100)
        Next
    Next
    
    EncryptByte = DS()
End Function

Public Function DecryptByte(ByRef DS() As Byte, Key() As Byte)
    On Error Resume Next
    Dim tmp2() As Byte, p As Double, i As Double, Bound As Integer, R As Integer, Total As Double, Prog As Double
    
    ' DS2 requires a 16-bit key or greater. If the received key is smaller
    ' than 16-bits then simple key expansion is performed.
    If UBound(Key) < 0 Then
        ReDim Key(1) As Byte: Key(0) = 1: Key(1) = 1
    ElseIf UBound(Key) = 0 Then
        ReDim tmp2(0) As Byte: tmp2(0) = Key(0)
        ReDim Key(1) As Byte: Key(0) = tmp2(0): Key(1) = 1
        ReDim tmp2(0) As Byte
    End If
    
    ' Initialize sbox with Key
    InitBox Key()
   
    ' Decrypt block
    Total = UBound(DS) * Rounds
    For R = 1 To Rounds
        Bound = (UBound(Key))
        p = (UBound(DS)) Mod (UBound(Key))
        For i = (UBound(DS)) To 1 Step -1
            If p = 0 Then p = Bound
            DS(i - 1) = sbox_b(DS(i - 1), sbox_a(DS(i), Key(p)))
            DS(i) = sbox_b(DS(i - 1), DS(i))
            DS(i - 1) = sbox_b(DS(i - 1), sbox_a(DS(i), Key(p - 1)))
            p = p - 1
            Prog = Prog + 1: RaiseEvent Progress((Prog / Total) * 100)
        Next
    Next
    
    ' Filter out salt
    tmp2() = DS()
    ReDim DS((UBound(tmp2)) - 3) As Byte
    Call CopyMem(DS(0), tmp2(3), UBound(DS))
    ReDim Preserve DS(UBound(DS) - 3) As Byte
    
    DecryptByte = DS()
End Function
Public Sub InitBox(Key() As Byte)
On Error GoTo oops
    Dim i As Integer, j As Integer, p As Double, Bound As Double
    Bound = UBound(Key())
    For i = 0 To 255
        For j = 0 To 255
            If p = Bound Then p = 0
            sbox_b(i, j) = CByte(i Xor j)
            sbox_a(i, j) = CByte((i + j * Key(p)) Mod 255)
            p = p + 1
        Next
    Next
Exit Sub

oops:

AddDebug "Error in DS2: " & err.Number & ": " & err.Description
End Sub



