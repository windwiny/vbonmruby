VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   7725
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command3 
      Caption         =   "close"
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   1560
      Width           =   6855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "run string"
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "init"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command2_Click()
On Error GoTo tk
    Dim va() As String
    va = VB_run(Text1.Text, 300, 50.01, "---abcd---")
    
    Dim i As Integer
    For i = 0 To UBound(va) - LBound(va)
        MsgBox va(i), vbOKOnly
    Next
    Exit Sub
tk:
    MsgBox "error"
End Sub

