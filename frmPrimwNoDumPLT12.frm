VERSION 5.00
Begin VB.Form PrimeNoSum 
   Caption         =   "Form13"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11715
   LinkTopic       =   "Form13"
   ScaleHeight     =   6000
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtRes 
      Height          =   350
      Left            =   1920
      TabIndex        =   3
      Top             =   3480
      Width           =   900
   End
   Begin VB.TextBox txtTotal 
      Height          =   350
      Left            =   1920
      TabIndex        =   2
      Top             =   2520
      Width           =   900
   End
   Begin VB.TextBox txtN 
      Height          =   350
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   1560
      Width           =   900
   End
   Begin VB.TextBox txtM 
      Height          =   350
      Index           =   1
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "Prime"
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "Sum"
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "N"
      Height          =   300
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "M"
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   900
   End
End
Attribute VB_Name = "PrimeNoSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFind_Click()
Dim m, n, i, f, j, sum As Integer
m = Val(txtM.Text)
n = Val(txtN.Text)
sum = 0
f = 1

For i = m To n
    
    For j = 2 To i \ 2
    
        If i Mod j = 0 Then
            f = 0
            Exit For
        Else
            f = 1
        End If
    Next
    
    If f = 1 Then
       txtRes.Text = txtRes.Text & " " & i
       sum = sum + i
    End If
    
Next

txtTotal.Text = sum
End Sub

