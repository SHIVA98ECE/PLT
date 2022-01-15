VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "FactorialOfNumber"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12120
   LinkTopic       =   "Form8"
   ScaleHeight     =   5940
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes 
      Height          =   405
      Left            =   2520
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdFact 
      Caption         =   "Factorial"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox TxtNum 
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a number"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFact_Click()
Dim i, f, n As Integer

n = Val(TxtNum.Text)
i = 1
f = 1

If n = 0 Then
    txtRes.Text = 1
ElseIf n < 0 Then
    txtRes.Text = "Impossible!"
Else
    While i <= n
        f = f * i
        i = i + 1
        
    Wend
    txtRes.Text = f
End If

End Sub
