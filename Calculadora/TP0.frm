VERSION 5.00
Begin VB.Form TP0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Van's Calculator v.1.0"
   ClientHeight    =   3090
   ClientLeft      =   4965
   ClientTop       =   3945
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnResta 
      Caption         =   "&Restar"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton btnDivide 
      Caption         =   "&Dividir"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton btnMultiplique 
      Caption         =   "&Multiplicar"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton btnSalida 
      Caption         =   "Sali&r"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton btnSumar 
      Caption         =   "&Sumar"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Clikea aquí para sumar los valores introducidos."
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtEntrada2 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      ToolTipText     =   "Ingresar aquí un valor"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtEntrada1 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Ingresar aquí un valor"
      Top             =   480
      Width           =   1815
   End
   Begin VB.Frame frmCalculo 
      Height          =   1575
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   4095
      Begin VB.CommandButton btnPotencia 
         Caption         =   "&Potenciar"
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame frmAcerca 
      Caption         =   "Acerca de"
      Height          =   2055
      Left            =   4560
      TabIndex        =   11
      Top             =   960
      Width           =   3615
      Begin VB.Label LblPotenciarhow 
         Caption         =   "Para potenciar el primer numero es la base."
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label lblDividirhow 
         Caption         =   "Para dividir, el primer numero es el dividendo, y el segundo es el divisor."
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label LblHow 
         Caption         =   "Para sumar, restar y multiplicar es necesario ingresar dos valores."
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Label Label1 
      Caption         =   "="
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblSalida 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      ToolTipText     =   "Aquí se mostrara el resultado de la operación"
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "TP0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub btnDivide_Click()
Dim priVal As Single, segVal As Single
priVal = Val(txtEntrada1.Text)
segVal = Val(txtEntrada2.Text)
    If (segVal = 0) Then
        lblSalida.Caption = "El divisor es 0 - Imposible de operar"
    Else
        lblSalida.Caption = Str(priVal / segVal)
    End If
End Sub

Private Sub btnMultiplique_Click()
Dim priVal As Single, segVal As Single
priVal = Val(txtEntrada1.Text)
segVal = Val(txtEntrada2.Text)
lblSalida.Caption = Str(priVal * segVal)
End Sub

Private Sub btnPotencia_Click()
Dim priVal As Single, segVal As Single
priVal = Val(txtEntrada1.Text)
segVal = Val(txtEntrada2.Text)
lblSalida.Caption = Str(priVal ^ segVal)
End Sub

Private Sub btnResta_Click()
Dim priVal As Single, segVal As Single
priVal = Val(txtEntrada1.Text)
segVal = Val(txtEntrada2.Text)
lblSalida.Caption = Str(priVal - segVal)
End Sub

Private Sub btnSalida_Click()
End
End Sub

Private Sub btnSumar_Click()
Dim priVal As Single, segVal As Single
priVal = Val(txtEntrada1.Text)
segVal = Val(txtEntrada2.Text)
lblSalida.Caption = Str(priVal + segVal)

End Sub

Private Sub Label2_Click()

End Sub
