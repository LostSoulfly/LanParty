VERSION 5.00
Begin VB.Form frmProgram 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   275
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   354
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Select"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox lstOptions 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelectedListText As String
Public blDidCancel As Boolean

Public Function ShowList(ListOptions() As String, Title As String) As String
Dim i As Long

lstOptions.Clear
Me.Caption = Title

For i = LBound(ListOptions) To UBound(ListOptions)

    lstOptions.AddItem ListOptions(i)

Next i

End Function

Private Sub cmdCancel_Click()
    blDidCancel = True
    Me.Visible = False
End Sub

Private Sub cmdOkay_Click()
    SelectedListText = lstOptions.Text
    Me.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    blDidCancel = True
    Me.Visible = False
End Sub
