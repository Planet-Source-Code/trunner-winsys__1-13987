VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Open a .BMP logo"
   ClientHeight    =   1140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   1140
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Image Must be 114x180 in size or there will be a error, also the file MUST be a BMP"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo errorhandler
Form1.pic1.Picture = LoadPicture(Text1.Text)
Form1.Label12.Caption = Text1.Text
Unload Form3
Exit Sub
errorhandler:
i = MsgBox(Err.Description, vbOKOnly, Err.Number)
End Sub
