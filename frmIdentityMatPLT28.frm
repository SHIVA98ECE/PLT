VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Columns"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12240
   LinkTopic       =   "Form6"
   ScaleHeight     =   6555
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtElements 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   5295
   End
   Begin VB.TextBox txtColumns 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtRows 
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Elements"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Columns"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Rows"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(10, 10) As Integer
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim y As Integer

Private Sub cmdAdd_Click()
Dim v1 As Integer

i = Val(txtColumns.Text)
j = Val(txtRows.Text)

v1 = Val(txtElements.Text)

a(i, j) = v1

End Sub

Private Sub cmdCheck_Click()
For x = 0 To i
    For y = 0 To j
        If a(x, x) = 1 And a(y, y) = 1 Then
            MsgBox "Identity Matrix"
            Exit For
        Else
            MsgBox "Not Identity Matrix"
            Exit For
        End If
    Next
Next
End Sub
