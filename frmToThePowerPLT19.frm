VERSION 5.00
Begin VB.Form Form14 
   Caption         =   "Form14"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11730
   LinkTopic       =   "Form14"
   ScaleHeight     =   6135
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calc"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtBase 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Exponent"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a Base"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Base, Expo As Integer
Dim v1 As Double
v1 = 1
Base = Val(txtBase.Text)
Expo = Val(txtExp.Text)

While Expo <> 0
    v1 = v1 * Base
    Expo = Expo - 1
Wend
txtRes.Text = v1
End Sub
