VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5565
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8175
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu nuevo 
         Caption         =   "Nuevo Editor"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bytindice As Byte
Private Sub nuevo_Click()
Static formX(300) As New Form1

   formX(bytindice).Show
bytindice = bytindice + 1
End Sub
