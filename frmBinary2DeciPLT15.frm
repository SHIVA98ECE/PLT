VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Binary2Deci"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8895
   LinkTopic       =   "Form10"
   ScaleHeight     =   5220
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtNum 
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Decimal Number"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a Binary No"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConvert_Click()
Dim n, dec, base, re As Integer
n = Val(txtNum.Text)
dec = 0
base = 1

While n > 0
    re = n Mod 10
    dec = dec + re * base
    n = n \ 10
    base = base * 2
Wend

txtRes.Text = dec

End Sub
