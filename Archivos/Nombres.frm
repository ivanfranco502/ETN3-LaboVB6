VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listas de "
   ClientHeight    =   3585
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCargar 
      Caption         =   "Cargar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtPrecio 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Text            =   "Precio"
      ToolTipText     =   "Precio"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtAño 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "Año"
      ToolTipText     =   "Año"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtDisco 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Text            =   "Discografica"
      ToolTipText     =   "Discografica"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Text            =   "Nombre de Disco"
      ToolTipText     =   "Nombre de Disco"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtArtista 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Artista"
      ToolTipText     =   "Artista"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.ListBox lstDiscos 
      Height          =   1185
      ItemData        =   "Nombres.frx":0000
      Left            =   120
      List            =   "Nombres.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu abrir 
         Caption         =   "Abrir"
         Shortcut        =   ^A
      End
      Begin VB.Menu eliminar 
         Caption         =   "Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cddisco As CompacDisk
Private Type CompacDisk
    strArtista As String * 20
    strNombre As String * 15
    strDisco As String * 10
    intAño As Integer
    bytPrecio As Currency
End Type
Private Sub Cargar(ByRef cddisco As CompacDisk)
lstDiscos.AddItem (cddisco.strArtista & cddisco.strNombre & cddisco.strDisco _
& Str(cddisco.intAño) & Str(cddisco.bytPrecio))

End Sub

Private Sub abrir_Click()
lstDiscos.Clear
Dim intUltiPos As Integer
Dim bytn As Byte

Open "U:\archivos\discos.dat" For Random As #1 Len = Len(cddisco)
    intUltiPos = LOF(1) / Len(cddisco)
    For bytn = 1 To intUltiPos
        Get #1, bytn, cddisco
        Cargar cddisco
    Next
Close #1
End Sub

Private Sub btnCargar_Click()
       
If txtArtista.Text = "Artista" Then
    MsgBox "Escriba el nombre del artista en el campo", vbOKOnly, "Error"
    txtArtista.SetFocus
Else
    If txtNombre.Text = "Nombre de Disco" Then
        MsgBox "Escriba el nombre del Disco en el campo", vbOKOnly, "Error"
        txtNombre.SetFocus
    Else
        If txtDisco.Text = "Discografica" Then
            MsgBox "Escriba el nombre de la discografica en el campo", vbOKOnly, "Error"
            txtDisco.SetFocus
        Else
            If txtAño.Text = "Año" Then
                MsgBox "Escriba el año correcto en el campo", vbOKOnly, "Error"
                txtAño.SetFocus
            Else
                If txtPrecio.Text = "Precio" Then
                    MsgBox "Escriba el monto correcto en el campo", vbOKOnly, "Error"
                    txtPrecio.SetFocus
                Else
                    If IsNumeric(txtAño.Text) Then
                        If IsNumeric(txtPrecio.Text) Then
                            If Not txtAño.Text > 1200 And txtAño.Text < 3000 Then
                                MsgBox "El Año es un valor incorrecto", vbOKOnly, "Error"
                                txtAño.SetFocus
                            Else
                                cddisco.strArtista = txtArtista.Text
                                cddisco.strNombre = txtNombre.Text
                                cddisco.strDisco = txtDisco.Text
                                cddisco.intAño = txtAño.Text
                                cddisco.bytPrecio = txtPrecio.Text
                                
                                Cargar cddisco
                                guardar cddisco
                            End If
                        Else
                            MsgBox "Ingrese un valor numerico para Precio", vbOKOnly, "Error"
                            txtPrecio.SetFocus
                        End If
                    Else
                        MsgBox "Ingrese un valor numerico para año", vbOKOnly, "Error"
                        txtAño.SetFocus
                    End If
                End If
            End If
        End If
    End If
End If
    
End Sub

Private Sub eliminar_Click()
Dim bytn As Byte
Dim intUltiPosi As Byte
Dim intn As Integer

intUltiPosi = 1

Open "U:\Archivos\Discos.dat" For Random As #1 Len = Len(cddisco)
Open "U:\Archivos\Discos.tmp" For Random As #2 Len = Len(cddisco)

For bytn = 0 To lstDiscos.ListCount - 1
    If lstDiscos.Selected(bytn) Then
        Get #1, intn + 1, cddisco
        Put #2, intUltiPosi, cddisco
        intUltiPosi = intUltiPosi + 1
    End If
Next bytn
Close #1, #2
Kill "U:\Archivos\Discos.dat"
Name "U:\Archivos\Discos.tmp" As "U:\Archivos\Discos.dat"
End Sub

Private Sub Form_Unload(cancel As Integer)
Dim salir As VbMsgBoxResult

salir = MsgBox("Realmente desea abandonar la aplicacion", vbOKCancel, "Abandonar")
If salir = vbOK Then
    End
Else
cancel = 1
End If
End Sub

Private Sub guardar(ByRef cddisco As CompacDisk)
Static Contador As Integer

If Not lstDiscos.ListCount = 0 Then
        Contador = Contador + 1
        Open "U:\Archivos\Discos.dat" For Random As #1 Len = Len(cddisco)
            Put #1, Contador, cddisco
        Close #1
Else
    MsgBox "Para guardar debe ingresar algun disco", vbOKOnly, "Error Grabacion"
End If
    
End Sub

Private Sub salir_Click()
Form_Unload (0)
End Sub

Private Sub txtNombre_GotFocus()
    txtNombre.SelStart = 0
    txtNombre.SelLength = Len(txtNombre.Text)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtDisco.SetFocus
End If
End Sub

Private Sub txtDisco_GotFocus()
    txtDisco.SelStart = 0
    txtDisco.SelLength = Len(txtDisco.Text)
End Sub
Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtArtista.SetFocus
End If
End Sub
Private Sub txtPrecio_GotFocus()
    txtPrecio.SelStart = 0
    txtPrecio.SelLength = Len(txtPrecio.Text)
End Sub
Private Sub txtAño_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPrecio.SetFocus
End If
End Sub

Private Sub txtAño_GotFocus()
    txtAño.SelStart = 0
    txtAño.SelLength = Len(txtAño.Text)
End Sub
Private Sub txtDisco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtAño.SetFocus
End If
End Sub


Private Sub txtArtista_GotFocus()
    txtArtista.SelStart = 0
    txtArtista.SelLength = Len(txtArtista.Text)
End Sub

Private Sub txtArtista_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtNombre.SetFocus
End If
End Sub
