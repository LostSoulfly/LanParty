VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000E&
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = DI_MASK Or DI_IMAGE

Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long

Private Sub Command1_Click()

Dim lPic As Picture
Dim FilePath As String
Dim lngIcon As Long
Dim picTest As Picture

FilePath = App.Path & "\Parser.exe"

Dim lIconCount As Long
lIconCount = ExtractIconEx(FilePath, -1, 0, 0, 0)
If lIconCount = 0 Then
   'there is no icon in the file
   Exit Sub
End If

    'Me.Picture1.AutoRedraw = True
    'Set lPic = LoadPicture(App.Path & "\test.jpg") 'Use the correct path and filename here
    'ResizePicture Me.Picture1, lPic

'Picture1.Picture = LoadPicture(vbNullString)
Picture1.AutoRedraw = True
Picture1.Refresh
ExtractIconEx FilePath, 0, lngIcon, 0, 1
DrawIconEx Picture1.hDC, 0&, 0&, lngIcon, Fix(Picture1.Width / Screen.TwipsPerPixelX) - 10, Fix(Picture1.Height / Screen.TwipsPerPixelY) - 10, 0&, 0&, DI_NORMAL
SavePicture Picture1.Image, App.Path & "\pic.jpg"
DestroyIcon lngIcon

Image1.Picture = Picture1.Picture
Image1.Refresh

End Sub

Private Sub ResizePicture(pBox As PictureBox, pPic As Picture)
Dim lWidth      As Single, lHeight    As Single
Dim lnewWidth   As Single, lnewHeight As Single
 
    'Clear the Picture in the PictureBox
    pBox.Picture = Nothing
    
    'Clear the Image  in the Picturebox
    pBox.Cls
    
    'Get the size of the Image, but in the same Scale than the scale used by the PictureBox
    lWidth = pBox.ScaleX(pPic.Width, vbHimetric, pBox.ScaleMode)
    lHeight = pBox.ScaleY(pPic.Height, vbHimetric, pBox.ScaleMode)
    
    'If image Width > pictureBox Width, resize Width
    If lWidth > pBox.ScaleWidth Then
        lnewWidth = pBox.ScaleWidth              'new Width = PB width
        lHeight = lHeight * (lnewWidth / lWidth) 'Risize Height keeping proportions
    Else
        lnewWidth = lWidth                       'If not, keep the original Width value
    End If
    
    'If the image Height > The pictureBox Height, resize Height
    If lHeight > pBox.ScaleHeight Then
        lnewHeight = pBox.ScaleHeight                   'new Height = PB Height
        lnewWidth = lnewWidth * (lnewHeight / lHeight)  'Risize Width keeping proportions
    Else
        lnewHeight = lHeight                            'If not, use the same value
    End If
    
    'add resized and centered to Picturebox
    pBox.PaintPicture pPic, (pBox.ScaleWidth - lnewWidth) / 2, _
                            (pBox.ScaleHeight - lnewHeight) / 2, _
                            lnewWidth, lnewHeight
                            
    'Update the Picture with the new image if you need it
    Set pBox.Picture = pBox.Image
End Sub

