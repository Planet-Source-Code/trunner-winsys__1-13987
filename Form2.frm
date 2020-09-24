VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4320
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form2
End Sub
