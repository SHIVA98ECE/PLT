VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "PLT16"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11205
   LinkTopic       =   "Form11"
   ScaleHeight     =   5940
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes 
      Height          =   765
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
Dim n, i, j, c As Integer
c = 0
For i = 1 To 10
    n = 7 * i
    For j = 2 To 7
        If (n Mod j = 1) Then
        c = c + 1
        If c = 1 Or c = 2 Or c = 4 Then
            txtRes.Text = txtRes.Text & " " & n
        End If
        Exit For
        End If
    Next
Next

End Sub
