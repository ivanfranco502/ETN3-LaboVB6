VERSION 5.00
Begin VB.Form Swap 
   Caption         =   "Swap"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSwap 
      Caption         =   "a=1 : b=2"
      Default         =   -1  'True
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "Swap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Swap(ByRef intx As Integer, ByRef inty As Integer)

Dim intaux As Integer
intaux = intx
intx = inty
inty = intaux

End Sub

Private Sub btnSwap_Click()
Dim inta As Integer
Dim intb As Integer
inta = 1
intb = 2
Print inta, intb
Swap inta, intb
Print inta, intb

End Sub
