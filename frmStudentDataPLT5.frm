VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "StudentData"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11010
   LinkTopic       =   "Form12"
   ScaleHeight     =   6525
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   14
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">>"
      Height          =   300
      Left            =   9720
      TabIndex        =   13
      Top             =   3120
      Width           =   600
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<<"
      Height          =   300
      Left            =   6600
      TabIndex        =   12
      Top             =   3120
      Width           =   600
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "save"
      Height          =   435
      Index           =   0
      Left            =   7800
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student Database"
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   495
         Index           =   0
         Left            =   840
         TabIndex        =   15
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtSub3 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txtSub2 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtSub1 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtStud 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "subject3"
         Height          =   300
         Left            =   140
         TabIndex        =   4
         Top             =   3360
         Width           =   1000
      End
      Begin VB.Label Label3 
         Caption         =   "Subject2"
         Height          =   300
         Left            =   135
         TabIndex        =   3
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Subject1"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Student"
         Height          =   300
         Left            =   140
         TabIndex        =   1
         Top             =   600
         Width           =   1000
      End
   End
   Begin VB.Label txtTot 
      Caption         =   "Total"
      Height          =   300
      Left            =   5520
      TabIndex        =   10
      Top             =   1920
      Width           =   800
   End
   Begin VB.Label txtAvg 
      Caption         =   "Average"
      Height          =   300
      Left            =   5520
      TabIndex        =   9
      Top             =   840
      Width           =   800
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type emp
    Name As String
    Id As String
    BasicS As Double
    SpecialA As Double
    Bonus As Double
    TaxSI As Double
    Gross As Double
    Annual As Double
    AnnualNet As Double
    ATax As Double
End Type

Dim E(20) As emp
Dim Index As Integer
Dim ci As Integer
Private Sub Command1_Click()

End Sub



Private Sub cmdClear_Click(Index As Integer)
txtName.Text = ""
txtSub1.Text = ""
txtSub2.Text = ""
txtSub3.Text = ""
txtAvg.Text = ""
txtTotal.Text = ""
lblRes.Caption = ""
End Sub

Private Sub cmdLeft_Click()
If ci > 0 Then
ci = ci - 1
getrecord ci
End If
End Sub

Private Sub cmdRight_Click()
If ci > 0 Then
ci = ci + 1
getrecord ci
End If
End Sub

Private Sub cmdSave_Click(Index As Integer)
Index = Index + 1
ci = ci + 1
Update (Index)
End Sub

