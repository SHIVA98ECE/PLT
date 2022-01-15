VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Even/Odd"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11910
   LinkTopic       =   "Form4"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdEvenOdd 
      Caption         =   "Even/Odd"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a number"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEvenOdd_Click()
Dim n As Integer
n = Val(txtNumber.Text)

If n Mod 2 = 0 Then
txtResult.Text = "EVEN"

Else
txtResult.Text = "ODD"
End If


End Sub










Private Sub Form_Load()

End Sub
