VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Generate Series"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult6 
      Height          =   500
      Left            =   250
      TabIndex        =   14
      Top             =   9360
      Width           =   6750
   End
   Begin VB.TextBox txtResult5 
      Height          =   500
      Left            =   250
      TabIndex        =   13
      Top             =   8040
      Width           =   6750
   End
   Begin VB.TextBox txtResul4 
      Height          =   500
      Left            =   250
      TabIndex        =   12
      Top             =   6600
      Width           =   6750
   End
   Begin VB.TextBox txtResult3 
      Height          =   500
      Left            =   250
      TabIndex        =   11
      Top             =   5040
      Width           =   6750
   End
   Begin VB.TextBox txtResult2 
      Height          =   500
      Left            =   250
      TabIndex        =   10
      Top             =   3720
      Width           =   6750
   End
   Begin VB.TextBox txtResult1 
      Height          =   500
      Left            =   250
      TabIndex        =   9
      Top             =   2160
      Width           =   6750
   End
   Begin VB.CommandButton cmdStart6 
      Cancel          =   -1  'True
      Caption         =   "Start 6"
      Height          =   600
      Left            =   180
      TabIndex        =   8
      Top             =   8640
      Width           =   900
   End
   Begin VB.CommandButton cmdStart5 
      Caption         =   "Start 5"
      Height          =   600
      Left            =   180
      TabIndex        =   7
      Top             =   7200
      Width           =   900
   End
   Begin VB.CommandButton cmdStart4 
      Caption         =   "Start 4"
      Height          =   600
      Left            =   180
      TabIndex        =   6
      Top             =   5760
      Width           =   900
   End
   Begin VB.CommandButton cmdStart3 
      Caption         =   "Start 3"
      Height          =   600
      Left            =   180
      TabIndex        =   5
      Top             =   4320
      Width           =   900
   End
   Begin VB.CommandButton cmdStart2 
      Caption         =   "Start 2"
      Height          =   600
      Left            =   180
      TabIndex        =   4
      Top             =   2880
      Width           =   900
   End
   Begin VB.CommandButton cmdStart1 
      Caption         =   "Start 1"
      Height          =   600
      Left            =   180
      TabIndex        =   3
      Top             =   1320
      Width           =   900
   End
   Begin VB.TextBox txtN 
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "        N    "
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, n As Integer
Dim s As String
Private Sub cmdStart1_Click()

n = Val(txtN.Text)

For i = 1 To n
    If i Mod 2 = 0 Then
    If i * i <= n Then
        s = s & Str(i * i)
    End If
    End If
Next

txtRes1.Text = s
End Sub

Private Sub cmdStart2_Click()
s = ""

For i = 1 To n
    If i <= n Then
    If i Mod 2 = 0 Then
        s = s & " " & Str(i * -1)
        
    Else
        s = s & " " & Str(i)
    End If
    End If
Next

txtRes2.Text = s
End Sub

Private Sub cmdStart3_Click()
s = ""

For i = 1 To n
    If i ^ i <= n Then
        s = s & " " & Str(i ^ i)
    End If
Next

txtRes3.Text = s
End Sub


Private Sub cmdStart4_Click()
s = ""

Dim count As Integer
count = 0

For i = 1 To n
    count = count + 1
    If count = 4 Then
        s = s
        count = 0
    Else
        If (i ^ 2) <= n Then
        s = s & " " & Str(i ^ 2)
        End If
    End If
Next

txtRes4.Text = s
End Sub

Private Sub cmdStart5_Click()
s = " "

Dim count As Integer
Dim b, n1, n2, n3 As Integer
count = 0
b = 1

n3 = 12

s = Str(1) & Str(4) & Str(7) & Str(n3)

For i = 3 To 10
    n3 = n3 + (3 + 2 ^ i)
    If n3 < n Then
    s = s & Str(n3)
    End If
Next

txtRes5.Text = s
End Sub


Private Sub cmdStart6_Click()
s = "1"

Dim count As Integer
Dim b As Integer
count = 0
b = 1

For i = 1 To n
    count = count + 1
    If count = 3 Then
        b = b
        s = s
        count = 0
    Else
        b = b + 4 * i
        If b <= n Then
        s = s & " " & Str(b)
        End If
    End If
Next

txtRes6.Text = s
End Sub

