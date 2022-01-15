VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Decimal2Binary"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10980
   LinkTopic       =   "Form9"
   ScaleHeight     =   5700
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   555
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Binary Number"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Entera decimal no"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConvert_Click()
Dim a(10), n, i, j As Integer
n = Val(txtNum.Text)

i = 0
While n > 0
    a(i) = n Mod 2
    n = n \ 2
    i = i + 1
    
Wend

j = i - 1
While j >= 0
    
    txtRes = a(j) & " " & txtRes
    
    j = j - 1
    
Wend
End Sub
