VERSION 5.00
Begin VB.Form TP4 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Van's Show Name"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "x"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton btnNombre 
      BackColor       =   &H00808080&
      Caption         =   "&Nombre y Apellido"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblApellido 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Apellido:"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ingrese nombre y apellido"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "TP4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnNombre_Click()
Dim bytPosi As Byte, Ape As String, nom As String

bytPosi = 0: Do: DoEvents 'Busca el blanco en la cadena
                bytPosi = bytPosi + 1
             Loop While Not Mid(txtNombre.Text, bytPosi, 1) = " "

nom = Left(txtNombre.Text, bytPosi - 1)
Ape = Right(txtNombre.Text, (Len(txtNombre.Text) - bytPosi))

nom = UCase(Left(nom, 1)) & LCase(Right(nom, Len(nom) - 1))
Ape = UCase(Left(Ape, 1)) & LCase(Right(Ape, Len(Ape) - 1))

lblApellido.Caption = Ape
lblNombre.Caption = nom
End Sub

Private Sub btnSalir_Click()
End
End Sub
