VERSION 5.00
Begin VB.Form Evaluacion2eva2 
   Caption         =   "2da Evaluacion 2do trimestre"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "X"
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton btnGenerar 
      Caption         =   "Generar"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label25 
      Caption         =   "10"
      Height          =   255
      Left            =   3840
      TabIndex        =   129
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label24 
      Caption         =   "9"
      Height          =   255
      Left            =   3480
      TabIndex        =   128
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label23 
      Caption         =   "8"
      Height          =   255
      Left            =   3120
      TabIndex        =   127
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label22 
      Caption         =   "7"
      Height          =   255
      Left            =   2760
      TabIndex        =   126
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label21 
      Caption         =   "6"
      Height          =   255
      Left            =   2400
      TabIndex        =   125
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label20 
      Caption         =   "5"
      Height          =   255
      Left            =   2040
      TabIndex        =   124
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label19 
      Caption         =   "4"
      Height          =   255
      Left            =   1680
      TabIndex        =   123
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label18 
      Caption         =   "3"
      Height          =   255
      Left            =   1320
      TabIndex        =   122
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label17 
      Caption         =   "2"
      Height          =   255
      Left            =   960
      TabIndex        =   121
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label16 
      Caption         =   "1"
      Height          =   255
      Left            =   600
      TabIndex        =   120
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label15 
      Caption         =   "10"
      Height          =   255
      Left            =   120
      TabIndex        =   119
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label14 
      Caption         =   "9"
      Height          =   255
      Left            =   240
      TabIndex        =   118
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label13 
      Caption         =   "8"
      Height          =   255
      Left            =   240
      TabIndex        =   117
      Top             =   3720
      Width           =   135
   End
   Begin VB.Label Label12 
      Caption         =   "7"
      Height          =   255
      Left            =   240
      TabIndex        =   116
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label11 
      Caption         =   "6"
      Height          =   255
      Left            =   240
      TabIndex        =   115
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label10 
      Caption         =   "5"
      Height          =   255
      Left            =   240
      TabIndex        =   114
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "4"
      Height          =   255
      Left            =   240
      TabIndex        =   113
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "3"
      Height          =   255
      Left            =   240
      TabIndex        =   112
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "2"
      Height          =   255
      Left            =   240
      TabIndex        =   111
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "1"
      Height          =   255
      Left            =   240
      TabIndex        =   110
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label lblRepetido 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4560
      TabIndex        =   109
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Alrededor"
      Height          =   375
      Left            =   5280
      TabIndex        =   108
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Con "
      Height          =   375
      Left            =   4200
      TabIndex        =   107
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblColumna 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5280
      TabIndex        =   106
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblFila 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5280
      TabIndex        =   105
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Columna:"
      Height          =   255
      Left            =   4320
      TabIndex        =   104
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Fila:"
      Height          =   255
      Left            =   4320
      TabIndex        =   103
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "El cero que esta mas rodeado por 1 es:"
      Height          =   375
      Left            =   4320
      TabIndex        =   102
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   99
      Left            =   3720
      TabIndex        =   101
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   98
      Left            =   3360
      TabIndex        =   100
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   97
      Left            =   3000
      TabIndex        =   99
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   96
      Left            =   2640
      TabIndex        =   98
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   95
      Left            =   2280
      TabIndex        =   97
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   94
      Left            =   1920
      TabIndex        =   96
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   93
      Left            =   1560
      TabIndex        =   95
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   92
      Left            =   1200
      TabIndex        =   94
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   91
      Left            =   840
      TabIndex        =   93
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   90
      Left            =   480
      TabIndex        =   92
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   89
      Left            =   3720
      TabIndex        =   91
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   88
      Left            =   3360
      TabIndex        =   90
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   87
      Left            =   3000
      TabIndex        =   89
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   86
      Left            =   2640
      TabIndex        =   88
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   85
      Left            =   2280
      TabIndex        =   87
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   84
      Left            =   1920
      TabIndex        =   86
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   83
      Left            =   1560
      TabIndex        =   85
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   82
      Left            =   1200
      TabIndex        =   84
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   81
      Left            =   840
      TabIndex        =   83
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   80
      Left            =   480
      TabIndex        =   82
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   79
      Left            =   3720
      TabIndex        =   81
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   78
      Left            =   3360
      TabIndex        =   80
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   77
      Left            =   3000
      TabIndex        =   79
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   76
      Left            =   2640
      TabIndex        =   78
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   75
      Left            =   2280
      TabIndex        =   77
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   74
      Left            =   1920
      TabIndex        =   76
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   73
      Left            =   1560
      TabIndex        =   75
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   72
      Left            =   1200
      TabIndex        =   74
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   71
      Left            =   840
      TabIndex        =   73
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   70
      Left            =   480
      TabIndex        =   72
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   69
      Left            =   3720
      TabIndex        =   71
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   68
      Left            =   3360
      TabIndex        =   70
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   67
      Left            =   3000
      TabIndex        =   69
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   66
      Left            =   2640
      TabIndex        =   68
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   65
      Left            =   2280
      TabIndex        =   67
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   64
      Left            =   1920
      TabIndex        =   66
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   63
      Left            =   1560
      TabIndex        =   65
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   62
      Left            =   1200
      TabIndex        =   64
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   61
      Left            =   840
      TabIndex        =   63
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   60
      Left            =   480
      TabIndex        =   62
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   59
      Left            =   3720
      TabIndex        =   61
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   58
      Left            =   3360
      TabIndex        =   60
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   57
      Left            =   3000
      TabIndex        =   59
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   56
      Left            =   2640
      TabIndex        =   58
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   55
      Left            =   2280
      TabIndex        =   57
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   54
      Left            =   1920
      TabIndex        =   56
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   53
      Left            =   1560
      TabIndex        =   55
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   52
      Left            =   1200
      TabIndex        =   54
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   51
      Left            =   840
      TabIndex        =   53
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   50
      Left            =   480
      TabIndex        =   52
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   49
      Left            =   3720
      TabIndex        =   51
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   48
      Left            =   3360
      TabIndex        =   50
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   47
      Left            =   3000
      TabIndex        =   49
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   46
      Left            =   2640
      TabIndex        =   48
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   45
      Left            =   2280
      TabIndex        =   47
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   44
      Left            =   1920
      TabIndex        =   46
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   43
      Left            =   1560
      TabIndex        =   45
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   42
      Left            =   1200
      TabIndex        =   44
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   41
      Left            =   840
      TabIndex        =   43
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   40
      Left            =   480
      TabIndex        =   42
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   39
      Left            =   3720
      TabIndex        =   41
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   38
      Left            =   3360
      TabIndex        =   40
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   37
      Left            =   3000
      TabIndex        =   39
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   36
      Left            =   2640
      TabIndex        =   38
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   35
      Left            =   2280
      TabIndex        =   37
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   34
      Left            =   1920
      TabIndex        =   36
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   33
      Left            =   1560
      TabIndex        =   35
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   32
      Left            =   1200
      TabIndex        =   34
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   31
      Left            =   840
      TabIndex        =   33
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   30
      Left            =   480
      TabIndex        =   32
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   29
      Left            =   3720
      TabIndex        =   31
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   28
      Left            =   3360
      TabIndex        =   30
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   27
      Left            =   3000
      TabIndex        =   29
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   26
      Left            =   2640
      TabIndex        =   28
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   25
      Left            =   2280
      TabIndex        =   27
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   24
      Left            =   1920
      TabIndex        =   26
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   23
      Left            =   1560
      TabIndex        =   25
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   22
      Left            =   1200
      TabIndex        =   24
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   21
      Left            =   840
      TabIndex        =   23
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   20
      Left            =   480
      TabIndex        =   22
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   19
      Left            =   3720
      TabIndex        =   21
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   18
      Left            =   3360
      TabIndex        =   20
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   17
      Left            =   3000
      TabIndex        =   19
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   16
      Left            =   2640
      TabIndex        =   18
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   15
      Left            =   2280
      TabIndex        =   17
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   14
      Left            =   1920
      TabIndex        =   16
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   13
      Left            =   1560
      TabIndex        =   15
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   12
      Left            =   1200
      TabIndex        =   14
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   11
      Left            =   840
      TabIndex        =   13
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   10
      Left            =   480
      TabIndex        =   12
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   9
      Left            =   3720
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   8
      Left            =   3360
      TabIndex        =   10
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   7
      Left            =   3000
      TabIndex        =   9
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   6
      Left            =   2640
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   4
      Left            =   1920
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblMatriz 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "Evaluacion2eva2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Generar una matriz de 10*10 llena de 0 y 1 al azar.
'Mostrarla e indicar el 0 con mayor cantidad de 1 alrededor.
'Una función deberá recibir la matriz y la posición del 0 para retornar
'la suma de unos alrededor de ese cero. En caso de haber varios 0 con la misma
'cantidad de 1 a su alrededor, indicar sólo uno de ellos.
'Guardar como 2eva2
'Alumno: Iván Franco
'Curso: 4to año
'Division: 2da

Private Sub GenerarBin(ByRef bytmatriz() As Byte)
Dim bytA As Byte
Dim fila As Byte
Dim col As Byte

For bytA = 0 To 99      'Inicia sistema de repeticion para ir llenando la matriz
    DoEvents
    fila = bytA \ 10    'se obtiene el valor de la fila
    col = bytA Mod 10   'se obtiene el valor de la columna
    bytmatriz(fila, col) = Rnd * 1      'Se genera el numero al azar
Next bytA               'Fin sistema de repeticion

End Sub
Private Sub Mostrar(ByRef bytmatriz() As Byte)
Dim bytN As Byte
Dim fila As Byte
Dim col As Byte

For bytN = 0 To 99      'Inicio sistema de repeticion para ir mostrando la matriz
    DoEvents
    fila = bytN \ 10    'se obtiene fila
    col = bytN Mod 10   'se obtiene columna
    lblMatriz(bytN).Caption = bytmatriz(fila, col) 'Se va imprimiendo en los Labels
Next bytN               'Fin sistema de repeticion

End Sub
Private Sub Swap(ByRef bytX As Byte, ByRef bytY As Byte)
Dim bytZ As Byte

bytZ = bytX 'Variable auxiliar toma valor de A
bytX = bytY 'A toma valor de B
bytY = bytZ 'B toma valor de Variable auxiliar

End Sub
Private Function Alrededor(ByRef matriz() As Byte, ByVal bytposi As Byte) As Byte
Dim bytFila As Byte
Dim bytCol As Byte
Dim bytContador As Byte
Dim bytA As Byte
Dim bytB As Byte

bytFila = bytposi \ 10  'Se obtiene el valor de fila
bytCol = bytposi Mod 10 'Se obtiene el valor de columna

If bytFila = 0 Then     'Si fila Vale 0
    
    If bytCol = 0 Then      'Si columna vale 0
        
        If matriz(bytFila, bytCol + 1) = 1 Then 'si el valor a la derecha es unos
            bytContador = bytContador + 1   'contador suma
        End If
        
        For bytA = bytCol To bytCol + 1 'busca unos abajo del valor
        DoEvents
            If matriz(bytFila + 1, bytA) = 1 Then   'si el valor de arriba es unos
                bytContador = bytContador + 1   'contador suma
            End If
        Next bytA
        
    ElseIf bytCol = 9 Then  'Si columna vale 9
        
        If matriz(bytFila, bytCol - 1) = 1 Then 'si el valor a la izquierda es unos
            bytContador = bytContador + 1   'contador suma
        End If
        
        For bytA = bytCol - 1 To bytCol 'busca unos abajo del valor
        DoEvents
            If matriz(bytFila + 1, bytA) = 1 Then   'si el valor de abajo es unos
                bytContador = bytContador + 1   'contador suma
            End If
        Next bytA
        
    Else    'si columna no vale 0 ó 9
        
        For bytA = bytCol - 1 To bytCol + 1 'busca unos abajo del valor
        DoEvents
            If matriz(bytFila + 1, bytA) = 1 Then   ' si el valor de abajo es unos
                bytContador = bytContador + 1   'contador suma
            End If
        Next bytA
        
        If matriz(bytFila, bytCol - 1) = 1 Then 'si el valor de a la izq es unos
            bytContador = bytContador + 1   'contador suma
        End If
        If matriz(bytFila, bytCol + 1) = 1 Then 'si el valor de a la derecha es unos
            bytContador = bytContador + 1   'contador suma
        End If
        
    End If
    
ElseIf bytFila = 9 Then 'Si el valor de fila es 9
    
    If bytCol = 0 Then      'Si columna vale 0
        
        For bytA = bytCol To bytCol + 1 'busca unos arriba del valor
        DoEvents
            If matriz(bytFila - 1, bytA) = 1 Then   'si el valor arriba es unos
                bytContador = bytContador + 1   'contador suma
            End If
        Next bytA
        If matriz(bytFila, bytCol + 1) = 1 Then 'si el valor a la derecha es unos
            bytContador = bytContador + 1   'contador suma
        End If
                        
    ElseIf bytCol = 9 Then  'si columna vale 9
    
        For bytA = bytCol - 1 To bytCol 'busca unos arriba del valor
        DoEvents
            If matriz(bytFila - 1, bytA) = 1 Then 'si el valor arriba es unos
                bytContador = bytContador + 1   'contador suma
            End If
        Next bytA
        If matriz(bytFila, bytCol - 1) = 1 Then 'si el valor a la izq es unos
            bytContador = bytContador + 1   'contador suma
        End If
    
    Else    'si columna no vale 0 ni 9
    
        For bytA = bytCol - 1 To bytCol + 1 'busca unos arriba del valor
        DoEvents
            If matriz(bytFila - 1, bytA) = 1 Then   'si el valor arriba es unos
                bytContador = bytContador + 1   'contador suma
            End If
        Next bytA
        
        If matriz(bytFila, bytCol - 1) = 1 Then 'si el valor a la izq es unos
            bytContador = bytContador + 1   'contador suma
        End If
        If matriz(bytFila, bytCol + 1) = 1 Then 'si el valor a la derecha es unos
            bytContador = bytContador + 1   'contador suma
        End If
    
    End If
    
Else    'si el valor de fila no es 0 ni 9

    If bytCol = 0 Then  'si columna vale 0

        For bytA = bytCol To bytCol + 1 'busca unos arriba y abajo del valor
        DoEvents
            If matriz(bytFila - 1, bytA) = 1 Then   'si el valor arriba es unos
                bytContador = bytContador + 1   'contador suma
            End If
            If matriz(bytFila + 1, bytA) = 1 Then   'si el valor abajo es unos
                bytContador = bytContador + 1   'contador suma
            End If
        Next bytA
            If matriz(bytFila, bytCol + 1) = 1 Then 'si el valor de derecha es unos
                bytContador = bytContador + 1   'contador suma
            End If

    ElseIf bytCol = 9 Then  'si columna vale 9
    
        For bytA = bytCol - 1 To bytCol 'busca unos arriba y abajo del valor
        DoEvents
            If matriz(bytFila - 1, bytA) = 1 Then   'si el valor arriba es unos
                bytContador = bytContador + 1   'contador suma
            End If
            If matriz(bytFila + 1, bytA) = 1 Then   'si el valor abajo es unos
                bytContador = bytContador + 1   'contador suma
            End If
        Next bytA
            If matriz(bytFila, bytCol - 1) = 1 Then 'si el valor de derecha es unos
                bytContador = bytContador + 1   'contador suma
            End If
    Else    'si columna no vale 0 ni 9
    
        For bytA = bytCol - 1 To bytCol + 1 'busca unos arriba y abajo del valor
        DoEvents
            If matriz(bytFila - 1, bytA) = 1 Then   'si el valor arriba es unos
                bytContador = bytContador + 1   'contador suma
            End If
            If matriz(bytFila + 1, bytA) = 1 Then   'si el valor abajo es unos
                bytContador = bytContador + 1   'contador suma
            End If
        Next bytA
        If matriz(bytFila, bytCol - 1) = 1 Then 'si el valor a izq es unos
            bytContador = bytContador + 1   'contador suma
        End If
        If matriz(bytFila, bytCol + 1) = 1 Then 'si el valor a der es unos
            bytContador = bytContador + 1   'contador suma
        End If
    
    End If
End If

Alrededor = bytContador ' se iguala la funcion
End Function
Private Sub btnGenerar_Click()
Dim bytBin(9, 9) As Byte
Dim bytO As Byte
Dim fila As Byte
Dim col As Byte
Dim bytRepetido(1) As Byte
Dim bytPosRep(1) As Byte


GenerarBin bytBin()         'Manda al procedimiento que genera los 0 y 1

For bytO = 0 To 99  'Sistema de repeticion para ir buscando 0
    DoEvents
    fila = bytO \ 10    'se obtiene la fila
    col = bytO Mod 10   'se obtiene la col
    If bytBin(fila, col) = 0 Then   'Pregunta si en la posicion hay un cero
        bytRepetido(1) = Alrededor(bytBin(), bytO) 'Se obtienen los ceros alrededor
        bytPosRep(1) = bytO 'Se guarda la posicion
        If bytRepetido(1) > bytRepetido(0) Then 'Se pregunta si el nuevo es mayor
                                                'al anterior guardado
            Swap bytRepetido(1), bytRepetido(0) 'si es mayor se cambian los valores
            Swap bytPosRep(1), bytPosRep(0)
            
        End If
    End If
Next bytO

lblFila.Caption = (bytPosRep(0) \ 10) + 1       'Se muestra la fila del mas repetido
lblColumna.Caption = (bytPosRep(0) Mod 10) + 1  'Se muestra la columna del mas repetido
lblRepetido.Caption = bytRepetido(0)        'se muestra la cant de veces rept

Mostrar bytBin()            'Manda al procedimiento que imprime la matriz
End Sub

Private Sub btnSalir_Click()
End
End Sub

Private Sub Form_Load()
Randomize Timer
End Sub
