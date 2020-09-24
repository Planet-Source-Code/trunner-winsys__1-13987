VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mid Town Computer Systems (www.mid-townonline.com) mtc@mail.ocis.net"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Support Information..."
      Height          =   350
      Left            =   8160
      TabIndex        =   14
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Apply to Windows"
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   5880
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   1560
      TabIndex        =   11
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Logo"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   1665
      Left            =   5505
      Picture         =   "Form1.frx":0742
      Top             =   1320
      Width           =   1755
   End
   Begin VB.Label Label6 
      Caption         =   "System:"
      Height          =   255
      Left            =   7560
      TabIndex        =   26
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Operating system name"
      Height          =   255
      Left            =   7920
      TabIndex        =   25
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "0.00.0000"
      Height          =   255
      Left            =   7920
      TabIndex        =   24
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Registered to:"
      Height          =   255
      Left            =   7560
      TabIndex        =   23
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "n/a"
      Height          =   255
      Left            =   7920
      TabIndex        =   22
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "00000-000-0000000-00000"
      Height          =   255
      Left            =   7920
      TabIndex        =   21
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label L1 
      Height          =   255
      Left            =   7920
      TabIndex        =   20
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label L2 
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label14 
      Caption         =   "Authentic CPU"
      Height          =   255
      Left            =   7920
      TabIndex        =   18
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label15 
      Caption         =   "CPU model(tm) Processor"
      Height          =   255
      Left            =   7920
      TabIndex        =   17
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label16 
      Caption         =   "000.0MB RAM"
      Height          =   255
      Left            =   7920
      TabIndex        =   16
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label17 
      Caption         =   "Manufactured and supported by:"
      Height          =   255
      Left            =   5160
      TabIndex        =   15
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Image pic1 
      Height          =   1755
      Left            =   5040
      Top             =   3840
      Width           =   2730
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   4680
      Picture         =   "Form1.frx":12F5
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "System Information:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Serial Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Sub Model:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Model:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Manufacturer:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Visible = True
End Sub

Private Sub Command2_Click()
Form2.Visible = True
Form2.Caption = L1
Dim numb As String
numb = 0
Do Until numb = List1.ListCount
Form2.Text1.Text = Form2.Text1.Text & List1.List(numb) & vbCrLf
numb = numb + 1
Loop
End Sub

Private Sub Command3_Click()
On Error GoTo errorhandler
Dim file1 As String
Dim file2 As String
Dim numb1 As String
Dim numb As String
file1 = Label12.Caption
file2 = "c:\windows\system\oemlogo.bmp"
FileCopy file1, file2

Open "c:\windows\system\oeminfo.ini" For Output As #1
Print #1, "[GENERAL]"
Print #1, "MANUFACTURER=" & Text2.Text
Print #1, "MODEL=" & Text3.Text
Print #1, ""
Print #1, "[OEMSPECIFIC]"
Print #1, "SUBMODEL=" & Text4.Text
Print #1, "SERIALNO=" & Text5.Text
Print #1, ""
Print #1, "[SUPPORT INFORMATION]"
numb = 0
numb1 = 1
Do Until numb = List1.ListCount
Print #1, "LINE" & numb1 & "=" & List1.List(numb)
numb = numb + 1
numb1 = numb1 + 1
Loop
Close #1
Exit Sub
errorhandler:
If Err.Number = "53" Then
i = MsgBox("No logo was specified, all other requested changes have been made", vbOKOnly, "Saved. No Logo")
Open "c:\windows\system\oeminfo.ini" For Output As #1
Print #1, "[GENERAL]"
Print #1, "MANUFACTURER=" & Text2.Text
Print #1, "MODEL=" & Text3.Text
Print #1, ""
Print #1, "[OEMSPECIFIC]"
Print #1, "SUBMODEL=" & Text4.Text
Print #1, "SERIALNO=" & Text5.Text
Print #1, ""
Print #1, "[SUPPORT INFORMATION]"
numb = 0
numb1 = 1
Do Until numb = List1.ListCount
Print #1, "LINE" & numb1 & "=" & List1.List(numb)
numb = numb + 1
numb1 = numb1 + 1
Loop
Close #1
End If
If Err.Number = "70" Then
Open "c:\windows\system\oeminfo.ini" For Output As #1
Print #1, "[GENERAL]"
Print #1, "MANUFACTURER=" & Text2.Text
Print #1, "MODEL=" & Text3.Text
Print #1, ""
Print #1, "[OEMSPECIFIC]"
Print #1, "SUBMODEL=" & Text4.Text
Print #1, "SERIALNO=" & Text5.Text
Print #1, ""
Print #1, "[SUPPORT INFORMATION]"
numb = 0
numb1 = 1
Do Until numb = List1.ListCount
Print #1, "LINE" & numb1 & "=" & List1.List(numb)
numb = numb + 1
numb1 = numb1 + 1
Loop
Close #1
End If

End Sub




Private Sub Text2_Change()
L1 = Text2
End Sub

Private Sub Text3_Change()
L2 = Text3
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    List1.AddItem Text6.Text
    Text6.Text = ""
    Text6.SetFocus
End If
End Sub
Private Sub List1_DblClick()
Dim ListNumber As Integer
ListNumber = List1.ListIndex
i = InputBox("Enter New Line", "Enter Data")
If i = "" Then
Else
List1.List(ListNumber) = i
End If
End Sub

