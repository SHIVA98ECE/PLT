VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "ReverseA|Number"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10740
   LinkTopic       =   "Form4"
   ScaleHeight     =   7005
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReverse 
      Caption         =   "Reverse"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtNum 
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Enter a Number"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   4800
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReverse_Click()
Dim num As Integer
Dim re As Integer
Dim rev As Integer
num = Val(txtNum.Text)

re = 0
rev = 0

 While num > 0
    re = num Mod 10
    rev = (rev * 10) + re
    num = num \ 10
    
 Wend
 
MsgBox rev



End Sub
