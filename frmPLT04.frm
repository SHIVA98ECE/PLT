VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12945
   LinkTopic       =   "Form6"
   ScaleHeight     =   7125
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes2 
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtRes1 
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdSeparate 
      Caption         =   "Separate"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtNum 
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Double value"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSeparate_Click()
Dim num As Double
num = Val(txtNum.Text)
txtRes1.Text = Int(num)
txtRes2.Text = num - Int(num)

End Sub
