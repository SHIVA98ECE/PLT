VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10365
   LinkTopic       =   "Form6"
   ScaleHeight     =   7560
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes2 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox txtRes1 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdTranspose 
      Caption         =   "Transpose"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisp 
      Caption         =   "Display"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtElem 
      Height          =   405
      Index           =   2
      Left            =   2880
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtColumn 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtRows 
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Columns1"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Rows1"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(10, 10) As Integer
Dim i, j As Integer
Dim x, y As Integer
Dim s As String

Private Sub cmdAdd_Click()
Dim v1 As Integer

i = Val(txtRows1.Text)
j = Val(txtCollumns1.Text)

v1 = Val(txtElem.Text)

a(i, j) = v1
End Sub

Private Sub cmdDisp_Click()
For x = 0 To i
    For y = 0 To j
        s = s & " " & a(x, y)
    Next
    txtRes1.Text = txtRes1.Text & s & vbCrLf
    s = ""
Next
End Sub

Private Sub cmdTranspose_Click()
Dim t(10, 10) As Integer

For x = 0 To i
    For y = 0 To j
        t(y, x) = a(x, y)
    Next
Next

For x = 0 To i
    For y = 0 To j
        s = s & " " & t(x, y)
    Next
    txtRes2.Text = txtRes2.Text & s & vbCrLf
    s = ""
Next

End Sub
