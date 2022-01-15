VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "Form13"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form13"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Controls"
      Height          =   1695
      Left            =   6240
      TabIndex        =   2
      Top             =   4080
      Width           =   6015
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   360
         Left            =   4680
         TabIndex        =   24
         Top             =   720
         Width           =   1000
      End
      Begin VB.CommandButton cmdRight 
         Caption         =   " >>"
         Height          =   360
         Left            =   3240
         TabIndex        =   23
         Top             =   720
         Width           =   1000
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   360
         Left            =   1680
         TabIndex        =   22
         Top             =   720
         Width           =   1000
      End
      Begin VB.CommandButton cmdLeft 
         Caption         =   " <<"
         Height          =   360
         Left            =   360
         TabIndex        =   21
         Top             =   720
         Width           =   1000
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Salary"
      Height          =   2775
      Left            =   6120
      TabIndex        =   1
      Top             =   600
      Width           =   5535
      Begin VB.TextBox txtNet 
         Height          =   400
         Left            =   2300
         TabIndex        =   20
         Top             =   2280
         Width           =   1500
      End
      Begin VB.TextBox txtAnnual 
         Height          =   400
         Left            =   2300
         TabIndex        =   19
         Top             =   1440
         Width           =   1500
      End
      Begin VB.TextBox txtGross 
         Height          =   400
         Left            =   2300
         TabIndex        =   18
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label9 
         Caption         =   "Annual Net Salary"
         Height          =   300
         Left            =   160
         TabIndex        =   11
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label Label8 
         Caption         =   "Annual Salary"
         Height          =   300
         Left            =   160
         TabIndex        =   10
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label7 
         Caption         =   "Gross Salary"
         Height          =   300
         Left            =   160
         TabIndex        =   9
         Top             =   600
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee"
      Height          =   5535
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   4815
      Begin VB.TextBox txtTaxInv 
         Height          =   400
         Left            =   2200
         TabIndex        =   17
         Top             =   4800
         Width           =   2400
      End
      Begin VB.TextBox txtBonus 
         Height          =   400
         Left            =   2200
         TabIndex        =   16
         Top             =   3840
         Width           =   2400
      End
      Begin VB.TextBox txtAllo 
         Height          =   400
         Left            =   2200
         TabIndex        =   15
         Top             =   3000
         Width           =   2400
      End
      Begin VB.TextBox txtBsal 
         Height          =   400
         Left            =   2200
         TabIndex        =   14
         Top             =   2160
         Width           =   2400
      End
      Begin VB.TextBox txtID 
         Height          =   400
         Left            =   2200
         TabIndex        =   13
         Top             =   1320
         Width           =   2400
      End
      Begin VB.TextBox txtName 
         Height          =   400
         Left            =   2200
         TabIndex        =   12
         Top             =   600
         Width           =   2400
      End
      Begin VB.Label Label6 
         Caption         =   "Monthly tax"
         Height          =   310
         Left            =   150
         TabIndex        =   8
         Top             =   4800
         Width           =   1400
      End
      Begin VB.Label Label5 
         Caption         =   "% of Bonus"
         Height          =   310
         Left            =   150
         TabIndex        =   7
         Top             =   3720
         Width           =   1400
      End
      Begin VB.Label Label4 
         Caption         =   "Special Allowances"
         Height          =   310
         Left            =   150
         TabIndex        =   6
         Top             =   2880
         Width           =   1400
      End
      Begin VB.Label Label3 
         Caption         =   "Basic Salary"
         Height          =   310
         Left            =   150
         TabIndex        =   5
         Top             =   2040
         Width           =   1400
      End
      Begin VB.Label Label2 
         Caption         =   "Employee ID"
         Height          =   310
         Left            =   150
         TabIndex        =   4
         Top             =   1320
         Width           =   1400
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   310
         Left            =   150
         TabIndex        =   3
         Top             =   600
         Width           =   1400
      End
   End
End
Attribute VB_Name = "Form13"
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
Dim index As Integer
Dim ci As Integer

Private Sub cmdClear_Click()
cmdSave.Enabled = True
txtName.Text = ""
txtID.Text = ""
txtBsal.Text = ""
txtAllo.Text = ""
txtBonus.Text = ""
txtTaxInv.Text = ""
End Sub

Private Sub cmdLeft_Click()
cmdSave.Enabled = False
If ci > 0 Then
ci = ci - 1
getrecord ci
End If
End Sub

Private Sub cmdRight_Click()
cmdSave.Enabled = False
If ci > 0 Then
ci = ci + 1
getrecord ci
End If
End Sub

Private Sub cmdSave_Click()
cmdLeft.Enabled = True
cmdRight.Enabled = True
index = index + 1
ci = ci + 1
update (index)
End Sub

Private Sub update(index As Integer)
With E(index)
.Name = txtName.Text
.Id = txtID.Text
.BasicS = Val(txtBsal.Text)
.SpecialA = Val(txtAllo.Text)
.Bonus = Val(txtBonus.Text) * .BasicS \ 100
.TaxSI = Val(txtTaxInv.Text)
.Gross = .BasicS + .SpecialA
.Annual = .BasicS + .SpecialA + .Bonus

If (.TaxSI <= 100000) Then
    .ATax = .Annual - .TaxSI
ElseIf (.TaxSI > 100000) Then
    .ATax = .Annual - 100000
End If

If (.ATax <= 100000) Then
    .AnnualNet = .ATax
ElseIf (.ATax > 100000 And .ATax <= 150000) Then
    .AnnualNet = .ATax - (.ATax * 20 \ 100)
ElseIf (.ATax > 150000) Then
    .AnnualNet = .ATax - (.ATax * 30 \ 100)
End If

txtGross.Text = .Gross
txtAnnual.Text = .Annual
txtNet.Text = .AnnualNet
lblIndex.Caption = "The Index is: " & index

End With
End Sub

Private Sub getrecord(index As Integer)
With E(index)
txtName.Text = .Name
txtID.Text = .Id
txtBsal.Text = .BasicS
txtAllo.Text = .SpecialA
txtBonus.Text = .Bonus * 100 \ .BasicS
txtTaxInv.Text = .TaxSI
txtGross.Text = .Gross
txtAnnual.Text = .Annual
txtNet.Text = .AnnualNet
lblIndex.Caption = "The Index is: " & index
End With
End Sub

Private Sub Form_Load()
cmdSave.Enabled = False
cmdClear.Enabled = False
cmdLeft.Enabled = False
cmdRight.Enabled = False
End Sub

Private Sub txtTaxInv_Change()
cmdSave.Enabled = True
cmdClear.Enabled = True
End Sub
