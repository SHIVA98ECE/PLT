VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14520
   LinkTopic       =   "Form2"
   ScaleHeight     =   7095
   ScaleWidth      =   14520
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdSimpleIntrest 
      Caption         =   "SI"
      Height          =   615
      Left            =   4560
      TabIndex        =   6
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox txtTime 
      Height          =   615
      Left            =   9360
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtRate 
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtPrincipal 
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "The Simple Intrest is"
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   5520
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Time"
      Height          =   200
      Left            =   9240
      TabIndex        =   2
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "Rate"
      Height          =   200
      Left            =   4680
      TabIndex        =   1
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "Principal"
      Height          =   200
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   1500
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSimpleIntrest_Click()
Dim p As Integer
Dim r As Single
Dim t As Integer
Dim si As Integer
p = Val(txt.Principle.Text)
r = Val(txt.Rate.Text)
t = Val(txt.Time.Text)
si = p * r * t / 100
txt.Result.Text = si

End Sub

