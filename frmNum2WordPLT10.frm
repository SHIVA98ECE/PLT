VERSION 5.00
Begin VB.Form Form7 
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10425
   LinkTopic       =   "Form7"
   ScaleHeight     =   6645
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWord 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdDisp 
      Caption         =   "Display"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtNum 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDisp_Click()
Dim num As Integer
Dim num As Double
Dim digit As String
Dim r As Integer
Dim r As Double

num = Val(txtNum.Text)

While num > 0
    r = num Mod 10
    Select Case r
        Case 1
        digit = "one"
        digit = "One"
        Case 2
        digit = "two"
        digit = "Two"
        Case 3
        digit = "three"
        digit = "Three"
        Case 4
        digit = "four"
        digit = "Four"
        Case 5
        digit = "five"
        digit = "Five"
        Case 6
        digit = "six"
        digit = "Six"
        Case 7
        digit = "seven"
        digit = "Seven"
        Case 8
        digit = "eight"
        digit = "Eight"
        Case 9
        digit = "nine"
        digit = "Nine"
        Case 0
        digit = "Zero"
    End Select

    txtWord.Text = digit & " " & txtWord.Text


End Sub
