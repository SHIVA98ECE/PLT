VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   2600
      Left            =   10320
      TabIndex        =   9
      Top             =   3480
      Width           =   2200
   End
   Begin VB.TextBox Text3 
      Height          =   2600
      Left            =   7080
      TabIndex        =   8
      Top             =   3480
      Width           =   2200
   End
   Begin VB.TextBox Text2 
      Height          =   2600
      Left            =   4080
      TabIndex        =   7
      Top             =   3480
      Width           =   2200
   End
   Begin VB.TextBox Text1 
      Height          =   2600
      Left            =   840
      TabIndex        =   6
      Top             =   3480
      Width           =   2200
   End
   Begin VB.CommandButton cmdGen4 
      Caption         =   "Generate4"
      Height          =   350
      Left            =   7080
      TabIndex        =   5
      Top             =   2160
      Width           =   1300
   End
   Begin VB.CommandButton cmdGen3 
      Caption         =   "Generate3"
      Height          =   350
      Left            =   4920
      TabIndex        =   4
      Top             =   2160
      Width           =   1300
   End
   Begin VB.CommandButton cmdGen2 
      Caption         =   "Generate2"
      Height          =   350
      Left            =   3000
      TabIndex        =   3
      Top             =   2160
      Width           =   1300
   End
   Begin VB.CommandButton cmdGen1 
      Caption         =   "Generate1"
      Height          =   350
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   1300
   End
   Begin VB.TextBox txtN 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "N"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim j As Integer
Dim r As Integer
Dim s As String

Private Sub cmdGen1_Click()
s = ""

Dim n1 As Double
Dim n2 As Double
Dim n3 As Double

n1 = 0
n2 = 1

txtRes1.Text = Str(n2)

For i = 2 To r
    n3 = n1 + n2
    If n3 <= r Then
     s = s & Str(n3)
    End If
    n1 = n2
    n2 = n3
Next

txtRes1.Text = txtRes1.Text & s

End Sub

Private Sub cmdGen2_Click()
s = ""

Dim n1 As Double
Dim n2 As Double
Dim n3 As Double

n1 = 0
n2 = 1

txtRes1.Text = Str(n2)

For i = 2 To r
    n3 = n1 + n2
    If n3 <= r Then
     s = s & Str(n3)
    End If
    n1 = n2
    n2 = n3
Next

txtRes1.Text = txtRes1.Text & s

End Sub

Private Sub cmdGen3_Click()

End Sub

Private Sub cmdGen4_Click()
s = ""

Dim n1, n2, n3, n4 As Integer
Dim count As Integer

n1 = 1
n2 = 5
n3 = 8

txtRes4.Text = Str(n1) & " " & Str(n2) & " " & Str(n3)

For i = 1 To r
        
        n4 = n1 + n2 + n3
        If n4 <= r Then
        s = s & " " & Str(n4)
        
        n1 = n2
        n2 = n3
        n3 = n4
        End If
    txtRes4.Text = txtRes4.Text & s
Next
End Sub

Private Sub Form_Load()

End Sub
