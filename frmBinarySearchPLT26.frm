VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "BinarySearch"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11415
   LinkTopic       =   "Form6"
   ScaleHeight     =   5310
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes2 
      Height          =   405
      Left            =   3480
      TabIndex        =   9
      Top             =   6360
      Width           =   4215
   End
   Begin VB.CommandButton cmdSrch 
      Caption         =   "Search"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter Elements"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnterN 
      Caption         =   "Enter N"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtSort 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   7200
      Width           =   2175
   End
   Begin VB.TextBox txtRes1 
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   5520
      Width           =   4215
   End
   Begin VB.TextBox txtSrch 
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtEnter 
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtN 
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Search Elements"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Array Elements"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter N"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j As Integer

Dim n As Integer
Dim a() As Integer

Private Sub cmdEnter_Click()
a(i) = Val(txtEnter.Text)
txtRes1.Text = txtRes1.Text & " " & a(i)
txtEnter.Text = ""
i = i + 1

End Sub

Private Sub cmdEnterN_Click()
n = Val(txtN.Text) - 1
ReDim a(n)
End Sub

Private Sub cmdSort_Click()
Dim t As Integer

For i = 0 To n
    For j = i + 1 To n
        If a(i) > a(j) Then
            t = a(i)
            a(i) = a(j)
            a(j) = t
        End If
    Next
Next
For i = 0 To n
    txtResSort.Text = txtResSort.Text & " " & a(i)
Next
End Sub

Private Sub cmdSrch_Click()
Dim first, last, midl, srch As Integer

Dim s As String

s = ""

srch = Val(txtSrch.Text)

first = 0
last = n

While first <= last
    midl = first + (last - first) \ 2
    If (a(midl) = srch) Then
        s = "found in " & Str(midl + 1)
    End If
    If a(midl) < srch Then
        first = midl + 1
    Else
        last = midl - 1
    End If
    txtRes2.Text = s
    If s = "" Then
    txtRes2.Text = "not found"
    End If
Wend

End Sub

