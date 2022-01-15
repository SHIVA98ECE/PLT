VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "SumOdd"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10905
   LinkTopic       =   "Form5"
   ScaleHeight     =   6885
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNum 
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   975
      Left            =   840
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtSum 
      Height          =   855
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalc_Click()

Dim N, sum, i As Double
N = Val(txtNum.Text)
sum = 0
For i = 0 To N
    If i Mod 2 <> 0 Then
    sum = sum + i
    End If
Next
txtSum.Text = sum
End Sub


