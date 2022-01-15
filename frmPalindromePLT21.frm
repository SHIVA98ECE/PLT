VERSION 5.00
Begin VB.Form Form14 
   Caption         =   "PalindromeChecking"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12570
   LinkTopic       =   "Form14"
   ScaleHeight     =   7065
   ScaleWidth      =   12570
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes 
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtStr 
      Height          =   405
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a String"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheck_Click()
Dim len1 As Integer, i As Integer
Dim s As String
Dim res As String
 
len1 = Len(txtStr.Text)

For i = len1 To 1 Step -1
    s = Mid(txtStr.Text, i, 1)
    res = res & s
Next
 
If txtStr.Text = res Then
    txtRes.Text = "yes"
Else
    txtRes.Text = "no"
End If

End Sub
