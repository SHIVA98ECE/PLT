VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "StrRev"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
   LinkTopic       =   "Form13"
   ScaleHeight     =   6000
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes1 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton cmdRev 
      Caption         =   "Rev"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtStr 
      Height          =   405
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a string"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRev_Click()
Dim len1 As Integer, i As Integer
Dim s As String
Dim res As String
 
len1 = Len(txtStr.Text)

For i = len1 To 1 Step -1
    s = Mid(txtStr.Text, i, 1)
    res = res & s
Next
 
txtRes1.Text = res
End Sub
