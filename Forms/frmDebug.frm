VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Debug"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   7095
   End
   Begin VB.TextBox txtOutput 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    txtOutput.Height = Me.Height - 1100
    txtOutput.Width = Me.Width - 350
    txtInput.Top = txtOutput.Height + txtOutput.Top + 100
    txtInput.Width = txtOutput.Width
End Sub

Private Sub txtInput_Change()
'restart
'stop
'var, dim
'close, exit
End Sub

Public Sub PrintDebug(Msg As String)
    txtOutput.Text = txtOutput.Text & Msg
End Sub

Private Sub txtOutput_Change()
txtOutput.SelStart = 65535
End Sub
