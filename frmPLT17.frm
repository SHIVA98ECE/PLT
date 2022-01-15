VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Total"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTotPrice 
      Caption         =   "TotPrice"
      Height          =   400
      Left            =   5295
      TabIndex        =   16
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   400
      Left            =   5295
      TabIndex        =   15
      Top             =   3360
      Width           =   1000
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   400
      Index           =   0
      Left            =   5295
      TabIndex        =   14
      Top             =   2400
      Width           =   1000
   End
   Begin VB.CommandButton cmdCash 
      Caption         =   "Cash"
      Height          =   400
      Left            =   5175
      TabIndex        =   13
      Top             =   1440
      Width           =   1000
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "Card"
      Height          =   400
      Left            =   5175
      TabIndex        =   12
      Top             =   600
      Width           =   1000
   End
   Begin VB.TextBox txtIsNext 
      Height          =   400
      Left            =   2000
      TabIndex        =   11
      Top             =   5280
      Width           =   2060
   End
   Begin VB.TextBox txtGrandTotal 
      Height          =   400
      Left            =   2000
      TabIndex        =   10
      Top             =   4440
      Width           =   2060
   End
   Begin VB.TextBox txtPrice 
      Height          =   400
      Left            =   2000
      TabIndex        =   9
      Top             =   3360
      Width           =   2060
   End
   Begin VB.TextBox txtQty 
      Height          =   400
      Left            =   2000
      TabIndex        =   8
      Top             =   2520
      Width           =   2060
   End
   Begin VB.TextBox txtDesc 
      Height          =   400
      Left            =   2000
      TabIndex        =   7
      Top             =   1680
      Width           =   2060
   End
   Begin VB.TextBox txtCode 
      Height          =   400
      Left            =   2000
      TabIndex        =   6
      Top             =   480
      Width           =   2060
   End
   Begin VB.Label Label6 
      Caption         =   "OK"
      Height          =   300
      Left            =   300
      TabIndex        =   5
      Top             =   5280
      Width           =   1000
   End
   Begin VB.Label Label5 
      Caption         =   "Total Price"
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label Label4 
      Caption         =   "Price"
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Qty"
      Height          =   300
      Left            =   300
      TabIndex        =   2
      Top             =   2400
      Width           =   1000
   End
   Begin VB.Label Label2 
      Caption         =   "Description"
      Height          =   300
      Left            =   300
      TabIndex        =   1
      Top             =   1560
      Width           =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "Item Code"
      Height          =   300
      Left            =   300
      TabIndex        =   0
      Top             =   480
      Width           =   1000
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type item
    code As String
    desc As String
    qty As Double
    price As Double
    total As Double
    
End Type

Dim i(20) As item
Dim index As Integer
Dim ci As Integer
Dim gTotal As Double

Private Sub cmdCard_Click()
If gTotal > 10000 Then
    MsgBox "you paid " & Str(gTotal - (gTotal * 10 \ 100))
ElseIf gTotal < 1000 Then
    MsgBox "you paid " & Str(gTotal + (gTotal * 2.5 \ 100))
Else
    MsgBox "you paid " & Str(gTotal)
End If
End Sub

Private Sub cmdCash_Click()
If gTotal > 10000 Then
    MsgBox "you paid " & Str(gTotal - (gTotal * 10 \ 100))
Else
    MsgBox "you paid " & Str(gTotal)
End If
End Sub

Private Sub cmdNext_Click(index As Integer)
index = index + 1
If txtIsNext.Text = "y" Then
update (index)
txtCode.Text = ""
txtDesc.Text = ""
txtQty.Text = ""
txtPrice.Text = ""
txtTotPrice.Text = ""
End If
End Sub
Private Sub update(index As Integer)
With i(index)
.code = txtCode.Text
.desc = txtDesc.Text
.qty = Val(txtQty.Text)
.price = Val(txtPrice.Text)
.total = Val(txtTotPrice.Text)
End With
End Sub

Private Sub cmdTotal_Click()
For c = 0 To index
    gTotal = gTotal + i(c).total
Next
txtGrandTotal.Text = gTotal
End Sub

Private Sub cmdTotPrice_Click()
txtTotPrice.Text = Val(txtQty.Text) * Val(txtPrice.Text)
End Sub
