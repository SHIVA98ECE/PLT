VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "Form13"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form13"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtElem 
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   2280
      Width           =   6375
   End
   Begin VB.TextBox txtColumn 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtRow 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Elements"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Columns"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Rows"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(10, 10) As Integer
Dim i, j As Integer
Dim x, y As Integer
Dim m, n As Integer
Private Sub cmdAdd_Click()
Dim v1 As Integer

i = Val(txtRow.Text)
j = Val(txtColumn.Text)

v1 = Val(txtElem.Text)

a(i, j) = v1

m = i

n = j

End Sub

Private Sub cmdCheck_Click()
Dim t(10, 10) As Integer

For x = 0 To m
    For y = 0 To n
        t(y, x) = a(x, y)
    Next
Next

Dim f As Integer
f = 1

For x = 0 To m
    For y = 0 To n
        If a(x, y) <> t(x, y) Then
            f = 0
            Exit For
        End If
    Next
    If f = 0 Then
        MsgBox "Not Symmetric"
        Exit For
    Else
        MsgBox "Symmetric"
        Exit For
    End If
Next

End Sub
