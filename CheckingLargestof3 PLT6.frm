VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12765
   LinkTopic       =   "Form5"
   ScaleHeight     =   7635
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Largest"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtNumber3 
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtNumber2 
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtNumber1 
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "2nd largest"
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "1st largest"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Enter 3rd number"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Enter 2nd number"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Enter 1st number"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim n1, n2, n3 As Integer
n1 = Val(txtNumber1.Text)
n2 = Val(txtNumber2.Text)
n3 = Val(txtNumber3.Text)
If n1 > n2 Then
   If n1 > n3 Then
      Text4.Text = n1
   Else
       Text4.Text = n3
   End If
   
Else
   If n2 > n3 Then
   Text4.Text = n2
   Else
   Text4.Text = n3
   End If
End If

End Sub
