VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "DeterminingDay"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10785
   LinkTopic       =   "Form6"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Determining Day"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim d As Integer
d = Val(Text1.Text)
Select Case d

Case 1
Text2.Text = "Monday"
Case 2
Text2.Text = "Tuesday"
Case 3
Text2.Text = "Wednesday"
Case 4
Text2.Text = "Thursday"
Case 5
Text2.Text = "Friday"
Case 6
Text2.Text = "Saturday"
Case Else
Text2.Text = "Enter a number bw 1 to 7"
End Select

End Sub

Private Sub Form_Load()

End Sub
