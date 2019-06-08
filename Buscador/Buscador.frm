VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   FillColor       =   &H008080FF&
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VSBlue 
      Height          =   2775
      Left            =   5640
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.VScrollBar VSGreen 
      Height          =   2775
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.VScrollBar VSRed 
      Height          =   2775
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton btnBorrar 
      Caption         =   "Borrar Seleccion"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton btnAgregar 
      Caption         =   "Agregar"
      Default         =   -1  'True
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.ListBox lstLista 
      Height          =   1410
      ItemData        =   "Buscador.frx":0000
      Left            =   120
      List            =   "Buscador.frx":0025
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAgregar_Click()
Dim palabra As String

If txtBuscar.Text <> "" Then
    palabra = txtBuscar.Text
    lstLista.AddItem (palabra)
btnBorrar.Enabled = True
End If
txtBuscar.Text = ""

End Sub

Private Sub btnBorrar_Click()
Dim bytn As Byte

Do: DoEvents
    If lstLista.Selected(bytn) = True Then
        lstLista.RemoveItem (bytn)
    Else
        bytn = bytn + 1
    End If

Loop Until (bytn >= lstLista.ListCount)

Do: DoEvents
    If lstLista.ListCount = 0 Then
        btnBorrar.Enabled = False
    End If
    bytn = bytn + 1
Loop Until (bytn >= lstLista.ListCount)
End Sub

Private Sub Form_Load()
lstLista.ListIndex = -1
VSBlue.Min = 0
VSBlue.Max = 255
VSGreen.Min = 0
VSGreen.Max = 255
VSRed.Min = 0
VSRed.Max = 255
VSRed.LargeChange = 3
VSGreen.LargeChange = 3
VSBlue.LargeChange = 3
Form1.BackColor = 0
End Sub

Private Sub txtBuscar_Change()
Dim buscar As Byte
Dim boolesta  As Boolean
Dim palabra As String
Dim extension As Byte
Dim mensaje As VbMsgBoxResult

boolesta = False
extension = Len(Trim(txtBuscar.Text))
palabra = Trim(txtBuscar.Text)
buscar = 0
lstLista.ListIndex = -1

If txtBuscar.Text = "" Then
    lstLista.ListIndex = -1
Else

    Do: DoEvents
        
    If palabra = Left(lstLista.List(buscar), extension) Then
        boolesta = True
        lstLista.ListIndex = buscar
    Else
        lstLista.ListIndex = -1
    End If
    buscar = buscar + 1
    
    Loop Until (boolesta = True Or buscar >= lstLista.ListCount - 1)
End If
End Sub

Private Sub VSBlue_Change()
Form1.BackColor = RGB(VSRed.Value, VSGreen.Value, VSBlue.Value)
End Sub

Private Sub VSGreen_Change()
Form1.BackColor = RGB(VSRed.Value, VSGreen.Value, VSBlue.Value)
End Sub

Private Sub VSRed_Change()
Form1.BackColor = RGB(VSRed.Value, VSGreen.Value, VSBlue.Value)
End Sub
