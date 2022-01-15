VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Output"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14085
   LinkTopic       =   "Form6"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes4 
      Height          =   2600
      Left            =   10200
      TabIndex        =   9
      Top             =   2640
      Width           =   2200
   End
   Begin VB.TextBox txtRes3 
      Height          =   2600
      Left            =   6960
      TabIndex        =   8
      Top             =   2640
      Width           =   2200
   End
   Begin VB.TextBox txtRes2 
      Height          =   2600
      Left            =   3720
      TabIndex        =   7
      Top             =   2640
      Width           =   2200
   End
   Begin VB.TextBox txtRes1 
      Height          =   2600
      Left            =   720
      TabIndex        =   6
      Top             =   2640
      Width           =   2200
   End
   Begin VB.CommandButton cmdGen4 
      Caption         =   "Generate4"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdGen3 
      Caption         =   "Generate3"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdGen2 
      Caption         =   "Generate2'"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdGen1 
      Caption         =   "Generate1"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtN 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "N"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   735
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
r = Val(txtN.Text)

For i = 1 To r
    For j = 1 To i
        s = s & Str(j) & " "
    Next
    s = s & vbCrLf
Next
txtRes1.Text = s
End Sub

Private Sub cmdGen2_Click()
s = ""

For i = 1 To r
    For j = 1 To i
        s = s & Str(i) & " "
    Next
    s = s & vbCrLf
Next
txtRes2.Text = s
End Sub

Private Sub cmdGen3_Click()
s = ""

Dim k As Integer
k = 1

For i = 1 To r
    For j = 1 To i
        s = s & Str(k) & " "
        k = k + 1
    Next
    s = s & vbCrLf
Next
txtRes3.Text = s
End Sub



Private Sub cmdGen4_Click()
s = ""

Dim N1 As Integer
Dim N2 As Integer
Dim N3 As Integer

N1 = 0
N2 = 1

txtRes4.Text = Str(N2) & vbCrLf


For i = 2 To r
    For j = 1 To i
        N3 = N1 + N2
            s = s & Str(N3)
            N1 = N2
            N2 = N3
    Next
    txtRes4.Text = txtRes4.Text & vbCrLf & s & vbCrLf
    s = ""
Next

End Sub
