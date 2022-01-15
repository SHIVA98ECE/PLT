VERSION 5.00
Begin VB.Form Form15 
   Caption         =   "Form15"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13590
   LinkTopic       =   "Form15"
   ScaleHeight     =   7755
   ScaleWidth      =   13590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes2 
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtRes1 
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSrch 
      Caption         =   "Search"
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "enter "
      Height          =   615
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnterN 
      Caption         =   "enter elements"
      Height          =   615
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtN 
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtSrch 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtEnter 
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "Size"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Search Element"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Enter array elements"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim n As Integer
Dim a() As Integer


Private Sub cmdEnter_Click(Index As Integer)
a(i) = Val(txtEnter.Text)
txtRes1.Text = txtRes1.Text & " " & a(i)
txtEnter.Text = ""
i = i + 1
End Sub

Private Sub cmdEnterN_Click(Index As Integer)
n = Val(txtN.Text) - 1
ReDim a(n)
End Sub

Private Sub cmdSrch_Click()
Dim srch As Integer
Dim c As Integer
srch = Val(txtSrch.Text)

For c = 0 To n
    If a(c) = srch Then
        txtRes2.Text = "Presnt at " & (c + 1)
        Exit For
    End If
Next
If c = n + 1 Then
    txtRes2.Text = "Not Present"
End If
End Sub
