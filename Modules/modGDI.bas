Attribute VB_Name = "modGDI"
' ----==== API Declarations ====----

Private Type GdiplusStartupInput
   GdiplusVersion As Long
   DebugEventCallback As Long
   SuppressBackgroundThread As Long
   SuppressExternalCodecs As Long
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" ( _
   token As Long, _
   inputbuf As GdiplusStartupInput, _
   Optional ByVal outputbuf As Long = 0) As Long

Private Declare Function GdiplusShutdown Lib "GDIPlus" ( _
   ByVal token As Long) As Long
   
Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" ( _
   ByVal FileName As Long, _
   bitmap As Long) As Long

Private Declare Function GdipDisposeImage Lib "GDIPlus" ( _
   ByVal image As Long) As Long

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" ( _
   ByVal bitmap As Long, _
   hbmReturn As Long, _
   ByVal background As Long) As Long

Private Declare Function GdipDeleteGraphics Lib "gdiplus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal GpImage As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus.dll" (ByVal image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus.dll" (ByVal image As Long, Height As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hBmp As Long, ByVal hPal As Long, GpBitmap As Long) As Long
Private Declare Function GdipCreateMetafileFromWmf Lib "gdiplus.dll" (ByVal hWmf As Long, ByVal deleteWmf As Long, WmfHeader As wmfPlaceableFileHeader, Metafile As Long) As Long
Private Declare Function GdipCreateMetafileFromEmf Lib "gdiplus.dll" (ByVal hEmf As Long, ByVal deleteEmf As Long, Metafile As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus.dll" (ByVal hIcon As Long, GpBitmap As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Type PICTDESC
   cbSizeOfStruct As Long
   picType As Long
   hgdiObj As Long
   hPalOrXYExt As Long
End Type

Private Type IID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7)  As Byte
End Type

Private Type PWMFRect16
    Left   As Integer
    Top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Private Type wmfPlaceableFileHeader
    Key         As Long
    hMf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    Reserved    As Long
    CheckSum    As Integer
End Type

Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" ( _
   lpPictDesc As PICTDESC, _
   riid As IID, _
   ByVal fOwn As Boolean, _
   lplpvObj As Object)

'------------------------------------------------------
' Procedure : LoadPicturePlus
' Purpose   : Loads an image using GDI+
' Returns   : The image loaded in a StdPicture object
' Author    : Eduardo A. Morcillo
'------------------------------------------------------
'
Public Function LoadPicturePlus(ByVal FileName As String) As StdPicture
Dim tSI As GdiplusStartupInput
Dim lGDIP As Long
Dim lRes As Long
Dim lBitmap As Long

   ' Initialize GDI+
   tSI.GdiplusVersion = 1
   lRes = GdiplusStartup(lGDIP, tSI)
   
   If lRes = 0 Then
   
      ' Open the image file
      lRes = GdipCreateBitmapFromFile(StrPtr(FileName), lBitmap)
   
      If lRes = 0 Then
      
         Dim hBitmap As Long
      
         ' Create a GDI bitmap
         lRes = GdipCreateHBITMAPFromBitmap(lBitmap, hBitmap, 0)
      
         ' Create the StdPicture object
         Set LoadPicturePlus = HandleToPicture(hBitmap, vbPicTypeBitmap)
         
         ' Dispose the image
         GdipDisposeImage lBitmap
      
      End If
      
      ' Shutdown GDI+
      GdiplusShutdown lGDIP
   End If
   
   If lRes Then Err.Raise 5, , "Cannot load file"
   
End Function

'------------------------------------------------------
' Procedure : HandleToPicture
' Purpose   : Creates a StdPicture object to wrap a GDI
'             image handle
'------------------------------------------------------
'
Public Function HandleToPicture( _
   ByVal hGDIHandle As Long, _
   ByVal ObjectType As PictureTypeConstants, _
   Optional ByVal hPal As Long = 0) As StdPicture
Dim tPictDesc As PICTDESC
Dim IID_IPicture As IID
Dim oPicture As IPicture
    
   ' Initialize the PICTDESC structure
   With tPictDesc
      .cbSizeOfStruct = Len(tPictDesc)
      .picType = ObjectType
      .hgdiObj = hGDIHandle
      .hPalOrXYExt = hPal
   End With
    
   ' Initialize the IPicture interface ID
   With IID_IPicture
      .Data1 = &H7BF80981
      .Data2 = &HBF32
      .Data3 = &H101A
      .Data4(0) = &H8B
      .Data4(1) = &HBB
      .Data4(3) = &HAA
      .Data4(5) = &H30
      .Data4(6) = &HC
      .Data4(7) = &HAB
   End With
    
   ' Create the object
   OleCreatePictureIndirect tPictDesc, IID_IPicture, _
                            True, oPicture
    
   ' Return the picture object
   Set HandleToPicture = oPicture
        
End Function

' Resize the picture using GDI plus
Public Sub gdipResize(Img As Long, hDC As Long, Width As Long, Height As Long, Optional RetainRatio As Boolean = False)
    Dim Graphics   As Long      ' Graphics Object Pointer
    Dim OrWidth    As Long      ' Original Image Width
    Dim OrHeight   As Long      ' Original Image Height
    Dim OrRatio    As Double    ' Original Image Ratio
    Dim DesRatio   As Double    ' Destination rect Ratio
    Dim DestX      As Long      ' Destination image X
    Dim DestY      As Long      ' Destination image Y
    Dim DestWidth  As Long      ' Destination image Width
    Dim DestHeight As Long      ' Destination image Height
    
    GdipCreateFromHDC hDC, Graphics
    GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic
    
    If RetainRatio Then
        GdipGetImageWidth Img, OrWidth
        GdipGetImageHeight Img, OrHeight
        
        OrRatio = OrWidth / OrHeight
        DesRatio = Width / Height
        
        ' Calculate destination coordinates
        DestWidth = IIf(DesRatio < OrRatio, Width, Height * OrRatio)
        DestHeight = IIf(DesRatio < OrRatio, Width / OrRatio, Height)
        DestX = (Width - DestWidth) / 2
        DestY = (Height - DestHeight) / 2

        GdipDrawImageRectRectI Graphics, Img, DestX, DestY, DestWidth, DestHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
    Else
        GdipDrawImageRectI Graphics, Img, 0, 0, Width, Height
    End If
    GdipDeleteGraphics Graphics
End Sub

Public Function Resize(Handle As Long, picType As PictureTypeConstants, Width As Long, Height As Long, Optional BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    Dim Img       As Long
    Dim hDC       As Long
    Dim hBitmap   As Long
    Dim WmfHeader As wmfPlaceableFileHeader
    
    ' Determine pictyre type
    Select Case picType
    Case vbPicTypeBitmap
         GdipCreateBitmapFromHBITMAP Handle, ByVal 0&, Img
    Case vbPicTypeMetafile
         FillInWmfHeader WmfHeader, Width, Height
         GdipCreateMetafileFromWmf Handle, False, WmfHeader, Img
    Case vbPicTypeEMetafile
         GdipCreateMetafileFromEmf Handle, False, Img
    Case vbPicTypeIcon
         ' Does not return a valid Image object
         GdipCreateBitmapFromHICON Handle, Img
    End Select
    
    ' Continue with resizing only if we have a valid image object
    If Img Then
        InitDC hDC, hBitmap, BackColor, Width, Height
        gdipResize Img, hDC, Width, Height, RetainRatio
        GdipDisposeImage Img
        GetBitmap hDC, hBitmap
        'Set Resize = CreatePicture(hBitmap)
        Set Resize = HandleToPicture(hBitmap, vbPicTypeBitmap)
    End If
End Function

Private Sub GetBitmap(hDC As Long, hBitmap As Long)
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
End Sub

Private Sub InitDC(hDC As Long, hBitmap As Long, BackColor As Long, Width As Long, Height As Long)
    Dim hBrush As Long
        
    ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hDC = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hDC, hBitmap)
    hBrush = CreateSolidBrush(BackColor)
    hBrush = SelectObject(hDC, hBrush)
    PatBlt hDC, 0, 0, Width, Height, PATCOPY
    DeleteObject SelectObject(hDC, hBrush)
End Sub

Private Sub FillInWmfHeader(WmfHeader As wmfPlaceableFileHeader, Width As Long, Height As Long)
    WmfHeader.BoundingBox.Right = Width
    WmfHeader.BoundingBox.Bottom = Height
    WmfHeader.Inch = 1440
    WmfHeader.Key = GDIP_WMF_PLACEABLEKEY
End Sub
