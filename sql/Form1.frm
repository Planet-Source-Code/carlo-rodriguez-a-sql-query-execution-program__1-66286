VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1380
      TabIndex        =   2
      Top             =   1680
      Width           =   1665
   End
   Begin VB.TextBox TXT2 
      Height          =   375
      Left            =   540
      TabIndex        =   1
      Top             =   1740
      Width           =   645
   End
   Begin VB.TextBox TXT 
      Height          =   1305
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3195
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    MsgBox decrypt(TXT.Text, TXT2.Text)
End Sub
