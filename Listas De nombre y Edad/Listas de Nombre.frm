VERSION 5.00
Begin VB.Form Listar 
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8790
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEdad 
      Height          =   495
      Index           =   7
      Left            =   2280
      TabIndex        =   15
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtEdad 
      Height          =   495
      Index           =   6
      Left            =   2280
      TabIndex        =   13
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtEdad 
      Height          =   495
      Index           =   5
      Left            =   2280
      TabIndex        =   11
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtEdad 
      Height          =   495
      Index           =   4
      Left            =   2280
      TabIndex        =   9
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtEdad 
      Height          =   495
      Index           =   3
      Left            =   2280
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtEdad 
      Height          =   495
      Index           =   2
      Left            =   2280
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtEdad 
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtNombre 
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtNombre 
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox txtNombre 
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtNombre 
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtNombre 
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtNombre 
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtNombre 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton btnDesordenar 
      Caption         =   "Desordenar"
      Height          =   615
      Left            =   3120
      TabIndex        =   17
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton btnOrdenar 
      Caption         =   "Ordenar Alfabeticamente"
      Default         =   -1  'True
      Height          =   615
      Left            =   3120
      TabIndex        =   16
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtEdad 
      Height          =   495
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtNombre 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblEdad 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   7
      Left            =   7800
      TabIndex        =   33
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label lblEdad 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   6
      Left            =   7800
      TabIndex        =   32
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label lblEdad 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   5
      Left            =   7800
      TabIndex        =   31
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblEdad 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   4
      Left            =   7800
      TabIndex        =   30
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblEdad 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   3
      Left            =   7800
      TabIndex        =   29
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblEdad 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   2
      Left            =   7800
      TabIndex        =   28
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblEdad 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   1
      Left            =   7800
      TabIndex        =   27
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   7
      Left            =   5160
      TabIndex        =   26
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label lblNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   6
      Left            =   5160
      TabIndex        =   25
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label lblNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   5
      Left            =   5160
      TabIndex        =   24
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label lblNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   4
      Left            =   5160
      TabIndex        =   23
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   3
      Left            =   5160
      TabIndex        =   22
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   2
      Left            =   5160
      TabIndex        =   21
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   1
      Left            =   5160
      TabIndex        =   20
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblEdad 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   7800
      TabIndex        =   19
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblNombre 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   5160
      TabIndex        =   18
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Listar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Swap(ByRef VarX As Variant, ByRef varY As Variant)

Dim varAux As Variant

varAux = VarX
VarX = varY
varY = varAux

End Sub
Private Sub Alfabe(ByRef strT() As String * 10, ByRef strEdad() As String * 2, bytComienzo As Byte, bytFinal As Byte)
Dim bytN As Byte
Dim bytX As Byte

For bytN = bytComienzo To bytFinal - 1
    For bytX = bytN + 1 To bytFinal
        If strT(bytX) < strT(bytN) Then
            Swap strT(bytX), strT(bytN)
            Swap strEdad(bytX), strEdad(bytN)
        End If
    Next bytX
Next bytN
End Sub
Private Sub desordenar(ByRef strT() As String * 10, ByRef strEdad() As String * 2, bytN As Byte)
Dim bytR As Byte
Dim bytX As Byte
    For bytR = 0 To bytN
        bytX = (Fix(10 * Rnd))
            Swap strT(bytR), strT(bytX)
            Swap strEdad(bytR), strEdad(bytX)
    Next bytR
End Sub

Private Sub btnDesordenar_Click()
Dim strVNom(7) As String * 15
Dim strVEdad(7) As String * 2
Dim n As Byte
Dim bytNumer  As Byte

For n = 0 To 7
    strVNom(n) = txtNombre(n).Text
    strVEdad(n) = txtEdad(n).Text
Next n

desordenar strVNom(), strVEdad(), 7

For bytNumer = 0 To 7
    lblNombre(bytNumer).Caption = strVNom(bytNumer)
    lblEdad(bytNumer).Caption = strVEdad(bytNumer)
Next bytNumer
End Sub

Private Sub btnOrdenar_Click()
Dim strVNom(7) As String * 15
Dim strVEdad(7) As String * 2
Dim n As Byte
Dim bytNumer  As Byte

For n = 0 To 7
    strVNom(n) = txtNombre(n).Text
    strVEdad(n) = txtEdad(n).Text
Next n

Alfabe strVNom(), strVEdad(), 0, 7

For bytNumer = 0 To 7
    lblNombre(bytNumer).Caption = strVNom(bytNumer)
    lblEdad(bytNumer).Caption = strVEdad(bytNumer)
Next bytNumer
End Sub
