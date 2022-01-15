VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "Form13"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form13"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes4 
      Height          =   2600
      Left            =   11280
      TabIndex        =   9
      Top             =   2760
      Width           =   2200
   End
   Begin VB.TextBox txtRes3 
      Height          =   2600
      Left            =   7800
      TabIndex        =   8
      Top             =   2760
      Width           =   2200
   End
   Begin VB.TextBox txtRes2 
      Height          =   2600
      Left            =   4320
      TabIndex        =   7
      Top             =   2760
      Width           =   2200
   End
   Begin VB.TextBox txtRes1 
      Height          =   2600
      Left            =   840
      TabIndex        =   6
      Top             =   2760
      Width           =   2200
   End
   Begin VB.CommandButton cmdGen4 
      Caption         =   "Generate4"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdGen3 
      Caption         =   "Generate3"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdGen2 
      Caption         =   "Generate2"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdGen1 
      Caption         =   "Generate1"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtN 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "N"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j, n As Integer
Dim s As String

Private Sub cmdGen1_Click()
n = Val(txtN.Text)

For i = 1 To n
    For j = 0 To 4
        s = s & "*"
    Next
    s = s & vbCrLf
Next
txtRes1.Text = s
End Sub

Private Sub cmdGen2_Click()
s = ""

For i = 1 To n
    For j = 0 To 4
        s = s & Str(i)
    Next
    s = s & vbCrLf
Next
txtRes2.Text = s
End Sub

Private Sub cmdGen3_Click()
s = ""

For i = 1 To n
    For j = 0 To 4
        s = s & Str(j + 1)
    Next
    s = s & vbCrLf
Next
txtRes3.Text = s
End Sub

Private Sub cmdGen4_Click()
s = ""

Dim k As Integer

For i = 1 To n
    For j = 1 To i
        s = s & "*" & " "
    Next
    s = s & vbCrLf
Next
txtRes4.Text = s
End Sub
